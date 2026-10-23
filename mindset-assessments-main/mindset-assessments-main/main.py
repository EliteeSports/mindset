"""Mindset assessment automation workflow."""

import base64
import json
import logging
import re
import time
import traceback
from datetime import datetime, timedelta
from io import BytesIO, StringIO

import openpyxl
import requests
from google.cloud import bigquery, bigquery_storage, secretmanager, storage
from google.cloud.exceptions import GoogleCloudError
from simple_salesforce import Salesforce, SalesforceMalformedRequest

# In-memory log stream used for status email attachments.
log_stream = StringIO()

# Shared logging configuration for file, stdout, and in-memory handlers.
default_log_args = {
	"level": logging.INFO,
	"format": '%(asctime)s.%(msecs)03d %(levelname)s\t\t%(message)s',
	"datefmt": "%d-%m-%Y %H:%M:%S"
}
root = logging.getLogger()
if root.handlers:
	for handler in root.handlers:
		root.removeHandler(handler)

logging.basicConfig(**default_log_args,handlers = [logging.StreamHandler(log_stream),logging.StreamHandler(),logging.FileHandler(f"log.log")])
log = logging.getLogger("")

# Google Cloud project configuration.
project_id = "640977351943"
project_name = "focal-rig-396713"
secret_version = "latest"

# Script name used in logs and notifications.
SCRIPT_NAME = "LIGHTHOUSE - MINDSET ASSESSMENTS"
secret_manager_client = secretmanager.SecretManagerServiceClient()

# Salesforce secret reference and credentials payload.
secret_id = "leathwaite-ai-sf-prod"
secret_name = f"projects/{project_id}/secrets/{secret_id}/versions/{secret_version}"
response = secret_manager_client.access_secret_version(name=secret_name)
sf_creds = json.loads(response.payload.data.decode('UTF-8'))

# Microsoft 365 / Thrive secret reference and credentials payload.
secret_id = "leathwaite"
secret_name = f"projects/{project_id}/secrets/{secret_id}/versions/{secret_version}"
response = secret_manager_client.access_secret_version(name=secret_name)
th_creds = json.loads(response.payload.data.decode('UTF-8'))

# Shared BigQuery clients.
client = bigquery.Client(project=project_name)
bqstorage_client = bigquery_storage.BigQueryReadClient()

# Expected report names produced by the assessment workflow.
report_names = [
	'Mindset Advantage Consultant Report',
	'Mindset Advantage Client Report_Portrait',
	'Mindset Advantage Client Report_Landscape',
	'Mindset Advantage Personal Report'
]
filename_regex = [
	r'[0-9]{1,50}–[0-9]{1,50}–Mindset Advantage Consultant Report.pdf',
	r'[0-9]{1,50}–[0-9]{1,50}–Mindset Advantage Client Report_Landscape.pdf',
	r'[0-9]{1,50}–[0-9]{1,50}–Mindset Advantage Client Report_Portrait.pdf',
	r'[0-9]{1,50}–[0-9]{1,50}–Mindset Advantage Personal Report.pdf'
]

# Compiled file-name validators for inbound reports.
compiled_patterns = [re.compile(pattern) for pattern in filename_regex]

# Storage and API configuration shared across the module.
bucket_name = "lw-ms"
gcp_table = "candidate_assessments"
base_url = "https://graph.microsoft.com/v1.0/"
report_file_name = "report_results.xlsx"

# Tracks whether any step failed during a full run.
script_error = False

# Authenticated Salesforce session for Lighthouse updates.
sf = Salesforce(username=sf_creds['username'], password=sf_creds['password'], security_token=sf_creds['token'], domain=sf_creds['domain'])

def lighthouse_update_candidate_stage(can_details):
	"""Mark a candidate assessment as completed in Salesforce."""
	today = datetime.today().strftime('%Y-%m-%d')
	fields = {
		'Candidate_Assessment_complete__c': True,
		'Assessment_Completed_Date__c': today
	}
	ap_result = sf.seven20__Application__c.update(can_details['candidacy_id'],fields)
	log.info(f"SELECTION PROCESS UPDATED: {int(ap_result)}")
	if ap_result == 204:
		return True
	else:
		return False


def get_inv_id():
	"""Build the invoice identifier using the month rollover rule."""
	today = datetime.today()
	if today.day > 25:
		next_month = today.replace(day=28) + timedelta(days=4)
		next_month = next_month.replace(day=1)
		formatted_date = next_month.strftime("%m%y")
	else:
		formatted_date = today.strftime("%m%y")
	return f"LW-MS-{formatted_date}"

def move_completed_candidates(lh_candidate_details):
	"""Advance candidates whose four reports are fully processed."""

	log.info("")
	log.info("-" * 60)
	log.info("MOVE COMPLETED CANDIDATES TO 'Mindset Assessment Complete' STAGE")
	log.info("-" * 60)
	log.info("")

	query = f'''SELECT DISTINCT search_id, contact_id
			FROM `focal-rig-396713.CandidateAssessment.{gcp_table}`
			WHERE stage_updated = FALSE
			  AND thrive_uploaded = TRUE
			QUALIFY COUNT(*) OVER (PARTITION BY search_id, contact_id) = 4;
			'''

	q1 = exec_bq_query(query, log_label="GET COMPLETED CANDIDATES FROM BQ")
	if len(q1[2]) > 0:
		for x in q1[2]:
			can_details = None
			for lh_can in lh_candidate_details:
				if (int(lh_can['search_number']) == int(x['search_id'])) and (int(lh_can['candidacy_number']) == int(x['contact_id'])):
					can_details = lh_can
					break
			log.info(f"SEARCH_ID: {x['search_id']}, CONTACT_ID: {x['contact_id']}")
			if can_details is None:
				log.error("	CANDIDATE NOT FOUND IN CURRENT LIGHTHOUSE ASSESSMENT LIST, SKIPPING STAGE UPDATE")
				continue
			inv_id = get_inv_id()
			log.info(f"	INVOICE ID IS: {inv_id}")
			query = f"UPDATE `focal-rig-396713.CandidateAssessment.{gcp_table}` SET `inv_id` = @inv_id WHERE search_id = @search_id AND contact_id = @contact_id"
			q2 = exec_bq_query(
				query,
				params={"inv_id": inv_id, "search_id": int(x["search_id"]), "contact_id": int(x["contact_id"])},
				log_label="UPDATING INV ID IN BQ",
			)

			if lighthouse_update_candidate_stage(can_details):
				query = f"UPDATE `focal-rig-396713.CandidateAssessment.{gcp_table}` SET `stage_updated` = True WHERE search_id = @search_id AND contact_id = @contact_id"
				q3 = exec_bq_query(
					query,
					params={"search_id": int(x["search_id"]), "contact_id": int(x["contact_id"])},
					log_label="UPDATING STAGE DATA IN BQ",
				)
	else:
		log.info(f"	NO COMPLETED CANDIDATES")


def send_email_log(em_receipient, em_subject, em_content,  at_data, at_name):
	"""Send a status email with an attached log payload."""
	access_token = get_o365_token()
	if isinstance(at_data, str):
		encoded_attachment = base64.b64encode(at_data.encode("utf-8")).decode("utf-8")
	else:
		encoded_attachment = base64.b64encode(at_data.read()).decode("utf-8")

	# Create email message
	shared_mailbox_email = "leathwaite.ai@leathwaite.com"
	message_url = f"https://graph.microsoft.com/v1.0/users/{shared_mailbox_email}/sendMail"
	message_headers = {
		"Authorization": f"Bearer {access_token}",
		"Content-Type": "application/json"
	}
	email_data = {
		"message": {
			"subject": em_subject,
			"body": {
				"content": em_content,
				"contentType": "Text",
			},
			"toRecipients": [
				{"emailAddress": {"address": em_receipient}}
			],
			"attachments": [
				{
					"@odata.type": "#microsoft.graph.fileAttachment",
					"contentBytes": encoded_attachment,
					"name": at_name,
				}
			]
		}
	}
	response = requests.post(message_url, headers=message_headers, data=json.dumps(email_data))
	if response.status_code == 202:
		log.info(f"O365 EMAIL RESPONSE: {response.status_code}")
		log.info(f"EMAIL SENT")
		return True
	else:
		log.error(f"O365 EMAIL RESPONSE: {response.status_code}")
		log.error(f"O365 EMAIL RESPONSE: {response.text}")
		log.error(f"EMAIL FAILED TO SEND")
		return False


def build_report_spreadsheet():
	"""Rebuild the reporting spreadsheet and publish it to SP and GCS."""
	log.info("")
	log.info("-" * 60)
	log.info(f"REBUILDING REPORT SPREADSHEET")
	log.info("-" * 60)
	log.info("")

	sp_creds = sp_get_creds()
	inbound_files = sp_get_item_list(sp_creds, "CandidateAssessment/Reports/Inbound")
	report_files = sp_get_item_list(sp_creds, "CandidateAssessment/Reports")

	for x in inbound_files:
		if x['name'] == report_file_name:
			delete_result = sp_delete_item(sp_creds['headers'], sp_creds['site_id'], x['id'])
			if delete_result[0]:
				log.info(f"OLD SPREADSHEET DELETED")
			else:
				log.error(f"CANNOT DELETE OLD SPREADSHEET: {delete_result[1]}")
				return False

	for x in report_files:
		if x['name'] == report_file_name:
			delete_result = sp_delete_item(sp_creds['headers'], sp_creds['site_id'], x['id'])
			if delete_result[0]:
				log.info(f"OLD SPREADSHEET DELETED")
			else:
				log.error(f"CANNOT DELETE OLD SPREADSHEET: {delete_result[1]}")
				return False

	query = (
		f"SELECT distinct search_id,contact_id, candidacy_id,criteria_1,criteria_2,criteria_3,candidate_first_name,candidate_last_name,candidate_personal_email,cast(report_requested_at as DATE) as date_only,inv_id,company FROM `focal-rig-396713.CandidateAssessment.{gcp_table}`")
	query_r1 = exec_bq_query(query, log_label="GETTING REPORT DATA FROM BQ")
	query = (
		f"SELECT search_id, contact_id, candidacy_id,report_name, file_created_at FROM `focal-rig-396713.CandidateAssessment.{gcp_table}`")
	query_r2 = exec_bq_query(query, log_label="GETTING REPORT DATA FROM BQ")

	results = []
	for x in query_r1[2]:
		rep1 = ""
		rep2 = ""
		rep3 = ""
		rep4 = ""
		for y in query_r2[2]:
			if x['candidacy_id'] == y['candidacy_id']:
				if y['report_name'] == "Mindset Advantage Client Report_Landscape":
					rep1 = y['file_created_at']
				if y['report_name'] == "Mindset Advantage Client Report_Portrait":
					rep2 = y['file_created_at']
				if y['report_name'] == "Mindset Advantage Client Report":
					rep1 = y['file_created_at']
					rep2 = y['file_created_at']
				if y['report_name'] == "Mindset Advantage Personal Report":
					rep3 = y['file_created_at']
				if y['report_name'] == "Mindset Advantage Consultant Report":
					rep4 = y['file_created_at']
		entry = {
			"search_id": x['search_id'],
			"contact_id": x['contact_id'],
			"candidacy_id": x['candidacy_id'],
			"criteria_1": re.sub(r'^[0-9]{1,2}\. ', '', x['criteria_1']),
			"criteria_2": re.sub(r'^[0-9]{1,2}\. ', '', x['criteria_2']),
			"criteria_3": re.sub(r'^[0-9]{1,2}\. ', '', x['criteria_3']),
			"candidate_first_name": x['candidate_first_name'],
			"candidate_last_name": x['candidate_last_name'],
			"candidate_personal_email": x['candidate_personal_email'],
			"report_requested_at": x['date_only'],
			"report_1_received": rep1,
			"report_2_received": rep2,
			"report_3_received": rep3,
			"report_4_received": rep4,
			"invoice_id": x['inv_id'],
			"company": x['company']
		}
		results.append(entry)

	workbook = openpyxl.Workbook()
	sheet = workbook.active
	sheet.title = "Sheet1"
	if results:
		header_row = list(results[0].keys())
		sheet.append(header_row)
		for data_row in results:
			sheet.append(list(data_row.values()))

	excel_file_binary = BytesIO()
	workbook.save(excel_file_binary)
	excel_file_binary.seek(0)
	new_file_name = report_file_name
	log.info(f"	FILENAME IS: {new_file_name}")
	update_url = f"https://graph.microsoft.com/v1.0/drives/{sp_creds['drive_id']}/root:/CandidateAssessment/{new_file_name}:/content"
	response = requests.put(
		update_url,
		headers=sp_creds['headers'],
		data=excel_file_binary,
	)

	if response.status_code in (200, 201):
		log.info("	FILE UPLOAD TO SP: OK")
		sp_upload = True
	else:
		log.error(f"	FILE UPLOAD TO SP: ERR ({response.status_code} - {response.text})")
		sp_upload = False

	# add file to GCP bucket
	storage_client = storage.Client(project=project_name)
	bucket = storage_client.get_bucket(bucket_name)
	blob = bucket.blob(new_file_name)

	try:
		blob.upload_from_string(excel_file_binary.getvalue(),
								content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
		log.info("	FILE UPLOAD TO GCP: OK")
		gcp_upload = True
	except GoogleCloudError as e:
		log.error(f"	FILE UPLOAD TO GCP: ERR ({e})")
		gcp_upload = False

	if (not sp_upload) or (not gcp_upload):
		return False
	else:
		return True

def notify_search_lead(lh_candidate_details):
	"""Notify search leads when all reports are available."""

	log.info("")
	log.info("-" * 60)
	log.info("NOTIFY SEARCH LEADS OF COMPLETED REPORTS")
	log.info("-" * 60)
	log.info("")

	query = f'''SELECT DISTINCT search_id, contact_id
		FROM `focal-rig-396713.CandidateAssessment.{gcp_table}`
		WHERE notified_lead = FALSE
		  AND thrive_uploaded = TRUE
		QUALIFY COUNT(*) OVER (PARTITION BY search_id, contact_id) = 4;
		'''
	query_r = exec_bq_query(query, log_label="GETTING COMPLETED SEARCHES THAT HAVE NOT BEEN NOTIFIED")

	for x in query_r[2]:
		can_details = None
		for lh_can in lh_candidate_details:
			if (int(lh_can['search_number']) == int(x['search_id'])) and (int(lh_can['candidacy_number']) == int(x['contact_id'])):
				can_details = lh_can
				break
		if can_details is None:
			log.error(
				f"	CANDIDATE NOT FOUND IN CURRENT LIGHTHOUSE ASSESSMENT LIST, "
				f"SKIPPING NOTIFICATION FOR SEARCH_ID: {x['search_id']}, CONTACT_ID: {x['contact_id']}"
			)
			continue
		search_lead = get_search_lead(can_details['search_id'])
		if process_notification(can_details, search_lead,6):
			log.info(f"	NOTIFICATION PROCESSED: OK")
			query = f'''UPDATE `focal-rig-396713.CandidateAssessment.{gcp_table}`
							SET notified_lead = True
							WHERE search_id = @search_id AND
							contact_id = @contact_id'''
			q2 = exec_bq_query(
				query,
				params={"search_id": int(x["search_id"]), "contact_id": int(x["contact_id"])},
				log_label="UPDATING SEARCH LEAD NOTIFICATION STATUS",
			)
			if q2[0]:
				log.info("	SEARCH LEAD UPDATED: OK")
			else:
				log.error("	SEARCH LEAD UPDATED: ERROR")
		else:
			log.error(f"	NOTIFICATION PROCESSED: ERROR")

# File transfer and storage helpers.

def lh_upload_document(lh_candidate_details):
	"""Upload completed report files into Salesforce records."""
	log.info("")
	log.info("-" * 60)
	log.info(f"UPLOADING REPORT DOCUMENTS TO LIGHTHOUSE")
	log.info("-" * 60)
	log.info("")

	client = bigquery.Client(project=project_name)
	query = (
		f"SELECT `id`,`search_id`,`report_name`,`contact_id`, `file_name`, `mime_type`, `download_url` FROM `{project_name}.CandidateAssessment.{gcp_table}` WHERE `sp_deleted` = False and `thrive_uploaded` = False and `report_received` = True and `in_lh` = True")
	query_r1 = exec_bq_query(query)
	if len(query_r1[2]):
		successful_files = []
		for x in query_r1[2]:
			log.info(
				f"UPLOADING DOCUMENT: {x['file_name']} FOR CONTACT {x['contact_id']}")
			response = requests.get(x['download_url'])
			if response.status_code == 200:
				log.info(f"	DOWNLOAD FROM SP OK")
				base64_data = base64.b64encode(response.content).decode('utf-8')

				for lh_can in lh_candidate_details:
					if (int(lh_can['search_number']) == x['search_id']) and (int(lh_can['candidacy_number']) == int(x['contact_id'])):
						sp_id = lh_can['candidacy_id']
						can_id = lh_can['contact_id']

				file_name_short = x['file_name'].split(".")[0]
				cv_body = {
					"Title":file_name_short,
					"PathOnClient":x['file_name'],
					"VersionData":base64_data,
					"Description":file_name_short
				}
				try:
					cv_result = sf.ContentVersion.create(cv_body)
					cv_id = cv_result.get("id")
					cv = sf.ContentVersion.get(cv_id)
					content_document_id = cv["ContentDocumentId"]
					cd_body = {
						"ContentDocumentId":content_document_id,
						"LinkedEntityId":sp_id
					}
					cd_result = sf.ContentDocumentLink.create(cd_body)
					log.info(f"FILE UPLOAD TO LH IS: {cd_result['success']}")
					cd_body = {
						"ContentDocumentId": content_document_id,
						"LinkedEntityId": can_id
					}
					cd_result = sf.ContentDocumentLink.create(cd_body)
					log.info(f"FILE UPLOAD TO LH IS: {cd_result['success']}")
					log.info("		ADDED TO LH OK")
					successful_files.append(x["id"])
				except SalesforceMalformedRequest as e:
					log.error(f"		FAILED TO UPLOAD FILE: {e.content}")
					script_error

			else:
				log.error(f"	DOWNLOAD FROM SP FAILED")
				script_error


		if successful_files:
			query = f"UPDATE `{project_name}.CandidateAssessment.{gcp_table}` SET `thrive_uploaded` = True WHERE `id` IN UNNEST(@ids)"
			query_r1 = exec_bq_query(query, params={"ids": successful_files})

	else:
		log.info("	THERE ARE NO FILES READY TO BE UPLOADED")

def update_sp_urls():
	"""Refresh expired SharePoint download URLs stored in BigQuery."""

	log.info("")
	log.info("-" * 60)
	log.info(f"UPDATING EXPIRED SP URLS")
	log.info("-"*60)
	log.info("")

	sp_creds = sp_get_creds()

	query =f"SELECT `id`,`site_id`,`download_url` FROM `focal-rig-396713.CandidateAssessment.{gcp_table}` where download_url is not null"
	query_r1 = exec_bq_query(query, log_label="GETTING EXPIRED DOWNLOAD URLS")
	for x in query_r1[2]:
		response = requests.get(x['download_url'])
		if response.status_code != 200:
			report_details = sp_get_item_details(sp_creds['headers'], sp_creds['site_id'], x['id'])
			query = f"UPDATE `{project_name}.CandidateAssessment.{gcp_table}` SET `download_url` = @download_url WHERE `id` = @id"
			query_r2 = exec_bq_query(
				query,
				params={"download_url": report_details["@microsoft.graph.downloadUrl"], "id": x["id"]},
				log_label="UPDATING EXPIRED DOWNLOAD URLS",
			)

def update_bq_report_data(reports):
	"""Write the latest SharePoint report metadata into BigQuery."""
	log.info("")
	log.info("-" * 60)
	log.info(f"UPDATING BQ WITH NEW REPORT DATA")
	log.info("-" * 60)
	log.info("")

	sp_creds = sp_get_creds()

	for x in reports:
		log.info(f"{x['search_id']},{x['contact_id']},{x['report_name']}")
		report_details = sp_get_item_details(sp_creds['headers'], sp_creds['site_id'], x['file_id'])
		query = f'''UPDATE `{project_name}.CandidateAssessment.{gcp_table}` SET
		`id` = @id,
		`file_name` = @file_name,
		`drive_id` = @drive_id,
		`folder_id` = @folder_id,
		`site_id`= @site_id,
		`mime_type` = @mime_type,
		`download_url`= @download_url,
		`file_created_by`= @file_created_by,
		`file_created_at`= @file_created_at,
		`file_modified_at`= @file_modified_at,'''
		params = {
			"id": x["file_id"],
			"file_name": report_details["name"],
			"drive_id": sp_creds["drive_id"],
			"folder_id": report_details["parentReference"]["id"],
			"site_id": sp_creds["site_id"],
			"mime_type": report_details["file"]["mimeType"],
			"download_url": report_details["@microsoft.graph.downloadUrl"],
			"file_created_by": report_details["lastModifiedBy"]["user"]["displayName"],
			"file_created_at": report_details["fileSystemInfo"]["createdDateTime"].rstrip("Z"),
			"file_modified_at": report_details["fileSystemInfo"]["lastModifiedDateTime"].rstrip("Z"),
			"search_id": int(x["search_id"]),
			"contact_id": int(x["contact_id"]),
			"report_name": x["report_name"],
		}
		if x['new_file']:
			query = query + f'''
			`sp_deleted` = False,
			`thrive_uploaded` = False,
			`thrive_candidacy` = False,
			`notified_lead` = False,
			`notified_finance` = False,
			`report_received` = True
			WHERE
			`search_id` = @search_id AND
			`contact_id` = @contact_id AND
			`report_name` = @report_name
			'''
		else:
			query = query + f'''
			`thrive_uploaded` = False,
			`notified_lead` = False,
			`report_received` = True
			WHERE
			`search_id` = @search_id AND
			`contact_id` = @contact_id AND
			`report_name` = @report_name
			'''
		query_r1 = exec_bq_query(query, params=params, log_label="UPDATING REPORT DATA IN BQ")

def delete_from_bucket(file_name):
	"""Delete a processed report file from the GCS bucket."""
	storage_client = storage.Client(project=project_name)
	bucket = storage_client.get_bucket(bucket_name)
	blob = bucket.blob(file_name)
	try:
		blob.delete()
		return True
	except GoogleCloudError as e:
		log.error(f"		GCP ERR: {e}")
		return False

def sp_get_creds():
	"""Get authenticated SharePoint headers plus site and drive ids."""
	access_token = get_o365_token()
	headers = {
		"Authorization": f"Bearer {access_token}",
		"Content-Type": "application/json"
	}
	site_id = sp_get_site_id(headers, "LW IT (Internal) All")
	if site_id == "":
		log.error("SITE ID NOT FOUND")
		return False

	drive_id = sp_get_drive_id(headers, "Documents", site_id)
	if drive_id == "":
		log.error("DRIVE ID NOT FOUND")
		return False
	creds = {
		'headers': headers,
		'site_id': site_id,
		'drive_id': drive_id
	}
	return creds


def sp_get_item_list(creds, folder):
	"""List the contents of a SharePoint folder path."""
	folder = folder.split("/")

	dir_id_list = sp_get_list_of_folders_root(creds['headers'], creds['drive_id'])
	if dir_id_list == "":
		log.error("DIRECTORY ID LIST NOT FOUND")
		return False

	folder_id = 0
	while (folder_id < len(folder)):
		for x in dir_id_list:
			if x['name'] == folder[folder_id]:
				dir_id = x['id']
		dir_id_list = sp_get_list_of_items(creds['headers'], creds['drive_id'], dir_id)
		folder_id = folder_id + 1

	return dir_id_list


def sp_get_site_id(header, site):
	"""Return the SharePoint site id for a site name."""
	response = requests.get(base_url + "sites/getAllSites", headers=header)
	for x in json.loads(response.text)['value']:
		if x['name'] == site:
			return x['id']


def sp_get_drive_id(header, drive, site_id):
	"""Return the SharePoint drive id for a drive name."""
	response = requests.get(base_url + "sites/" + site_id + "/drives", headers=header)
	for y in json.loads(response.text)['value']:
		if y['name'] == drive:
			return (y['id'])


def sp_get_list_of_folders_root(header, drive_id):
	"""List root-level folders in a SharePoint drive."""
	response = requests.get("https://graph.microsoft.com/v1.0/drives/" + drive_id + "/root/children", headers=header)
	results = []
	for x in json.loads(response.text)['value']:
		result = {
			'id': x['id'],
			'name': x['name']
		}
		results.append(result)
	return (results)

def sp_get_list_of_items(header, drive_id, item_id):
	"""List child items for a SharePoint drive item."""
	response = requests.get("https://graph.microsoft.com/v1.0/drives/" + drive_id + "/items/" + item_id + "/children",
							headers=header)
	results = []
	for x in json.loads(response.text)['value']:
		result = {
			'id': x['id'],
			'name': x['name']
		}
		results.append(result)
	return (results)


def sp_get_item_details(header, site_id, item_id):
	"""Fetch metadata for a specific SharePoint drive item."""
	response = requests.get("https://graph.microsoft.com/v1.0/sites/" + site_id + "/drive/items/" + item_id,
							headers=header)
	return (json.loads(response.text))


def sp_delete_item(header, site_id, item_id):
	"""Delete a SharePoint drive item by id."""
	response = requests.delete("https://graph.microsoft.com/v1.0/sites/" + site_id + "/drive/items/" + item_id,
							   headers=header)
	if response.status_code == 204:
		return True, "OK"
	else:
		return False, f"{response.status_code} - {response.text}"

def upload_to_sp(file_bytes,file_name):
	"""Upload a file into the SharePoint reports folder."""
	access_token = get_o365_token()

	headers = {
		"Authorization": f"Bearer {access_token}",
		"Content-Type": "application/octet-stream"
	}

	site_id = sp_get_site_id(headers, "LW IT (Internal) All")
	if site_id == "":
		log.error("SITE ID NOT FOUND")
		return False

	drive_id = sp_get_drive_id(headers, "Documents", site_id)
	if drive_id == "":
		log.error("DRIVE ID NOT FOUND")
		return False

	update_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/CandidateAssessment/Reports/{file_name}:/content"

	response = requests.put(
		update_url,
		headers=headers,
		data=file_bytes
	)
	file_id = json.loads(response.text)['id']
	if response.status_code in (200,201):
		log.info(F"		UPLOAD TO SP: OK")
		return True,response.status_code,file_id
	else:
		log.info(response.status_code)
		log.info(f"		UPLOAD TO SP: ERR ({response.text})")
		return False,response.status_code,file_id

def move_from_bucket_to_sp(current_bucket_files):
	"""Move valid inbound report files from GCS to SharePoint."""
	file_list = []
	storage_client = storage.Client(project=project_name)
	bucket = storage_client.get_bucket(bucket_name)
	for x in current_bucket_files:
		log.info(f"	MOVING {x} FROM GCP TO SP")
		blob = bucket.blob(x)
		upload_state = upload_to_sp(blob.download_as_bytes(), x)
		if upload_state[0]:
			name_split = blob.name.split("–")
			report_name = name_split[2].split(".")[0]
			new_file = True
			if upload_state[1] == 200:
				new_file = False

			entry = {
				'search_id': name_split[0],
				'contact_id': name_split[1],
				'report_name': report_name,
				'file_name': x,
				'new_file': new_file,
				'file_id': upload_state[2]
			}
			file_list.append(entry)
			if delete_from_bucket(x):
				log.info("		DELETED FROM GCP: OK")
			else:
				log.error("		DELETED FROM GCP: ERR")
	return file_list

def valid_file_name(text):
	"""Check whether a filename matches an allowed report pattern."""
	for pattern in compiled_patterns:
		if pattern.search(text):
			return True
	return False

def get_bucket_files():
	"""Return valid report filenames currently in the GCS bucket."""
	log.info("")
	log.info("-" * 60)
	log.info("GETTING FILES FROM GCP BUCKET")
	log.info("-" * 60)
	log.info("")

	file_list = []

	client = storage.Client(project=project_name)
	bucket = client.get_bucket(bucket_name)
	blobs = bucket.list_blobs()
	for blob in blobs:
		log.info(f"	GCP FILENAME: {blob.name}")
		if valid_file_name(blob.name):
			file_list.append(blob.name)
		else:
			log.warning(f"		INVALID FILE NAME STRUCTURE: {blob.name}, IGNORING")

	return file_list

def get_o365_token():
	"""Fetch an application token for Microsoft Graph."""
	token_url = f"https://login.microsoftonline.com/{th_creds['tenant_id']}/oauth2/v2.0/token"
	payload = {
		"grant_type": "client_credentials",
		"client_id": th_creds['client_id'],
		"client_secret": th_creds['client_secret'],
		"scope": "https://graph.microsoft.com/.default"
	}
	response = requests.post(token_url, data=payload)
	if response.status_code == 200:
		return response.json()["access_token"]
	else:
		log.error(f"O365 TOKEN RESPONSE: {response.status_code}")
		log.error(f"O365 TOKEN RESPONSE: {response.text}")
		return ""


def send_email(em_receipient, em_subject, em_content, em_format, src_mailbox="mindsetassessments@leathwaite.com"):
	"""Send an email through Microsoft Graph."""
	access_token = get_o365_token()

	message_url = f"https://graph.microsoft.com/v1.0/users/{src_mailbox}/sendMail"
	message_headers = {
		"Authorization": f"Bearer {access_token}",
		"Content-Type": "application/json"
	}
	message_body = {'Message': {'Subject': em_subject,
								'Body': {'ContentType': em_format, 'Content': em_content},
								'ToRecipients': [
									{'EmailAddress': {'Address': em_receipient}}
								],
								},
					'SaveToSentItems': 'true'}
	response = requests.post(message_url, headers=message_headers, json=message_body)
	if response.status_code == 202:
		return True
	else:
		log.error(f"O365 EMAIL RESPONSE: {response.status_code}")
		log.error(f"O365 EMAIL RESPONSE: {response.text}")
		log.error(f"EMAIL FAILED TO SEND")
		return False

def _bq_param_type(value):
	"""Map a Python value to the matching BigQuery parameter type."""
	if isinstance(value, bool):
		return "BOOL"
	if isinstance(value, int):
		return "INT64"
	if isinstance(value, float):
		return "FLOAT64"
	return "STRING"


def _bq_param_from_value(name, value):
	"""Build a BigQuery query parameter from a Python value."""
	if isinstance(value, list):
		item_type = "STRING"
		for item in value:
			if item is not None:
				item_type = _bq_param_type(item)
				break
		return bigquery.ArrayQueryParameter(name, item_type, value)
	return bigquery.ScalarQueryParameter(name, _bq_param_type(value), value)


def exec_bq_query(query, params=None, log_label=None, attempts=5):
	"""Execute a BigQuery statement with retries and optional params."""
	try:
		if log_label:
			log.info(log_label)
		job_config = None
		if params:
			job_config = bigquery.QueryJobConfig(
				query_parameters=[_bq_param_from_value(k, v) for k, v in params.items()]
			)
		query_job = client.query(query, job_config=job_config)
		df = query_job.to_dataframe(bqstorage_client=bqstorage_client)
		if not df.empty:
			result_dicts = df.to_dict(orient='records')
			return True, None, result_dicts
		else:
			return True, None, []

	except GoogleCloudError as e:
		log.error(f"	BQ EXEC ERR: {e}")
		log.error(f"	BQ EXEC ERR: {query}")
		if attempts > 0:
			log.info(f"	BQ EXEC WILL ATTEMPT {attempts-1} MORE TIME(S)")
			log.info(f"	BQ EXEC SLEEPING FOR {(10-attempts)*2} SECONDS")
			time.sleep((10-attempts)*2)
			return exec_bq_query(query, params=params, log_label=log_label, attempts=attempts-1)
		else:
			send_email("andy.harris@leathwaite.com", f"{SCRIPT_NAME}: QUERY ERROR", str(e) + "\n" + query,
					   "Text", "application/json", log_stream.getvalue(), "log.txt")
			return False, 0, []

# Notification and Salesforce workflow helpers.

def process_notification(lh_can, search_lead_data, notif_type):
	"""Send the correct assessment notification email for a workflow event."""
	if notif_type == 2:
		subject = f'''
	ACTION REQUIRED - Cannot send assessment for {lh_can['candidate_first_name']} {lh_can['candidate_last_name']} for Lighthouse Role '{lh_can['role_name']}'
		'''

		body = f'''
		Dear all,
		<br><br>
		We are unable to send the assessment requested for {lh_can['candidate_first_name']} {lh_can['candidate_last_name']} for Lighthouse Role '{lh_can['role_name']}'. This is because the candidate’s contact record doesn’t have an entry in the 'Preferred Personal Email' field.
		<br><br>
		The candidate has been returned to their previous candidate stage. Please update the Candidate Contact record <a href =\"https://leathwaite.lightning.force.com/lightning/r/Contact/{lh_can['contact_id']}/view\">here</a> and then click "Send for Assessments" on the candidate Selection Process Record <a href =\"https://leathwaite.lightning.force.com/lightning/r/seven20__Application__c/{lh_can['candidacy_id']}/view\">here</a>
		<br><br>
		If you have any questions, please speak to Nicola Lukes or Chris Smith
		<br><br>
		Leathwaite AI 
		<br><br>
		Please don't reply to this email, I'm not a real person
		<br><br>
		Don't know why you've had this email or what to do with it? click <a href=\"https://leathwaite.sharepoint.com/sites/CompanyICT/SitePages/LW_AI_emails_FAQ.aspx\">here</a> for help
		'''

	if notif_type == 3:
		subject = f'''
	Leathwaite Assessment
		'''

		body = f'''
		Dear {lh_can['candidate_first_name']},
		<br><br>
		As you know, Leathwaite have introduced a growth mindset assessment into their search process. Thank you for agreeing to complete your Mindset Advantage assessment.
		<br><br> 
		Please complete the assessment at your earliest convenience by visiting <a href=\"https://matthewsyed.typeform.com/to/sAJUEEAN\">https://matthewsyed.typeform.com/to/sAJUEEAN</a>
		<br><br>
		Many thanks for your time
		<br><br>
		The Leathwaite Team
		<br><br>
		'''

	if notif_type == 4:
		subject = f'''
	Leathwaite Assessment sent for {lh_can['candidate_first_name']} {lh_can['candidate_last_name']} for Lighthouse Role '{lh_can['role_name']}'
		'''
		body = f'''
		Dear all,
		<br><br>
		This is a notification to let you know that an assessment invitation has been sent out to {lh_can['candidate_first_name']} {lh_can['candidate_last_name']} with regards to Lighthouse Role – <a href =\"https://leathwaite.lightning.force.com/lightning/r/seven20__Job__c/{lh_can['search_id']}/view\">{lh_can['role_name']}</a>. If there is any issue with this please notify Nicola Lukes as soon as possible.
		<br><br>
		Many thanks for your time
		<br><br>
		Leathwaite AI 
		<br><br>
		Please don't reply to this email, I'm not a real person
		<br><br>
		Don't know why you've had this email or what to do with it? click <a href=\"https://leathwaite.sharepoint.com/sites/CompanyICT/SitePages/LW_AI_emails_FAQ.aspx\">here</a> for help
		'''

	if notif_type == 6:
		subject = f"Leathwaite Assessment ready for {lh_can['candidate_first_name']} {lh_can['candidate_last_name']} for Lighthouse Role '{lh_can['role_name']}'"

		body = f'''
		Dear all
		<br><br>
		The assessment reports for {lh_can['candidate_first_name']} {lh_can['candidate_last_name']} for Lighthouse Role '{lh_can['role_name']}' have been received and are accessible in the 'Files' section at the following link:
		<br><br>
		<a href =\"https://leathwaite.lightning.force.com/lightning/r/Contact/{lh_can['contact_id']}/view\">here</a>
		<br><br>
		Please ensure you send the candidate their report, add the client report into your assessment pack, and review the questions report
		<br><br>
		Many thanks
		<br><br>
		Leathwaite AI Please don't reply to this email, I'm not a real person'''

	if notif_type in (2, 4, 6):
		for user in search_lead_data:
			log.info(f"		SENDING NOTIFICATION TYPE {notif_type} TO: {user['email']}")
			send_email(
				user['email'],
				subject,
				body,
				"HTML",
				src_mailbox="mindsetassessments@leathwaite.com",
			)
	if notif_type == 3:
		log.info(f"		SENDING NOTIFICATION TYPE {notif_type} TO: {lh_can['candidate_personal_email']}")
		send_email(
			lh_can['candidate_personal_email'],
			subject,
			body,
			"HTML",
			src_mailbox="mindsetassessments@leathwaite.com",
		)
	return True

def get_candidate_details_from_lh():
	"""Fetch Salesforce candidates currently in the assessment queue."""

	log.info("")
	log.info("-" * 60)
	log.info("GETTING ALL CANDIDATES IN 'Send for Assessment' STATE")
	log.info("-" * 60)
	log.info("")

	mindset_candidate_details = []

	##### GET ROLES WITH MINDSET GROWTH SET
	mindset_roles = list(sf.query_all_iter(
		"SELECT Id,Name,Growth_Mindset_Behavioural_Criteria__c,Role_Number__c,seven20__Account__c FROM seven20__Job__c WHERE Growth_Mindset_Behavioural_Criteria__c != ''"))
	if len(mindset_roles) == 0:
		log.warning("	THERE ARE NO ROLES CURRENTLY WITH MINDSET CRITERIA SET")
	else:
		for role in mindset_roles:
			log.info(f"	ROLE NAME IS: '{role['Name']}' ({role['Id']})")

			role_company = sf.query(f"SELECT Name FROM Account WHERE Id = '{role['seven20__Account__c']}'")['records'][0]['Name']

			mindset_criteria = role['Growth_Mindset_Behavioural_Criteria__c'].split(";")
			role_number = role['Role_Number__c'].split("-")[1]

			#### GETTING ALL CANDIDATES THAT HAVE BEEN SENT FOR ASSESSMENT BUT NOT COMPLETED
			mindset_candidates = list(sf.query_all_iter(
				f"SELECT Id,Name,seven20__Candidate__c FROM seven20__Application__c WHERE seven20__Job__c = '{role['Id']}' AND Candidate_Sent_for_Assessment__c = True AND Candidate_Assessment_complete__c = False AND Assessment_Sent_Date__c > 2025-10-12"))
			if len(mindset_candidates) == 0:
				log.warning("		THERE ARE NO CANDIDATES READY FOR ASSESSMENT")
			else:
				for candidate in mindset_candidates:
					#### GET CANDIDATE DETAILS FROM CONTACT RECORD

					sp_number = candidate['Name'].split("-")[1]

					can_details = sf.query(
						f"SELECT Id,FirstName,LastName,Email,Title,seven20__Employer_Name__c FROM Contact WHERE Id = '{candidate['seven20__Candidate__c']}'")[
						'records'][0]
					log.info(
						f"		CANDIDATE DETAILS, FIRST NAME: {can_details['FirstName']}, LAST NAME: {can_details['LastName']}, ID: {can_details['Id']}")
					entry = {
						'role_name': role['Name'],
						'search_number': role_number,
						'candidacy_number': sp_number,
						'search_id': role['Id'],
						'candidacy_id': candidate['Id'],
						'contact_id': can_details['Id'],
						'job_title': can_details['Title'],
						'company': role_company,
						'criteria_1': mindset_criteria[0],
						'criteria_2': mindset_criteria[1],
						'criteria_3': mindset_criteria[2],
						'candidate_first_name': can_details['FirstName'],
						'candidate_last_name': can_details['LastName'],
						'candidate_personal_email': can_details['Email']
					}
					mindset_candidate_details.append(entry)

	return mindset_candidate_details

def get_search_lead(search_id):
	"""Return the users assigned to a Salesforce search team."""

	search_team_details = []
	search_lead_ids = sf.query(f"SELECT User__c,Team_Role__c FROM Role_Team__c WHERE Role__c = '{search_id}'")
	for lead in search_lead_ids['records']:
		search_lead_email = sf.query(f"SELECT Name,Email FROM User WHERE Id = '{lead['User__c']}'")
		search_team_details.append({
			'name':search_lead_email['records'][0]['Name'],
			'email':search_lead_email['records'][0]['Email'],
			'role':lead['Team_Role__c']
		})
	return search_team_details

def revert_candidate_status(lh_can):
	"""Reset a candidate's assessment-sent status in Salesforce."""

	fields = [{
		"Id": lh_can['candidacy_id'],
		"Candidate_Sent_for_Assessment__c": False,
		"Assessment_Sent_Date__c":None
	}]
	sf_results = sf.bulk.seven20__Application__c.update(fields,batch_size=10000,use_serial=True)
	if not sf_results[0]['success']:
		log.error(sf_results)
		return False
	else:
		return True

def add_candidate_to_bq(can_data):
	"""Insert the expected report rows for a candidate into BigQuery."""
	candidate_added_ok = False
	for report in report_names:
		query = f'''
		INSERT INTO `{project_name}.CandidateAssessment.{gcp_table}`
		(`search_id`,`contact_id`,`candidacy_id`,`company`,`job_title`,`criteria_1`,`criteria_2`,`criteria_3`,`candidate_first_name`,`candidate_last_name`,`candidate_personal_email`,`report_name`)
		VALUES
		(@search_id,@contact_id,@candidacy_id,@company,@job_title,@criteria_1,@criteria_2,@criteria_3,@candidate_first_name,@candidate_last_name,@candidate_personal_email,@report_name)
		'''
		params = {
			"search_id": int(can_data["search_number"]),
			"contact_id": int(can_data["candidacy_number"]),
			"candidacy_id": int(can_data["candidacy_number"]),
			"company": can_data["company"],
			"job_title": can_data["job_title"],
			"criteria_1": can_data["criteria_1"],
			"criteria_2": can_data["criteria_2"],
			"criteria_3": can_data["criteria_3"],
			"candidate_first_name": can_data["candidate_first_name"],
			"candidate_last_name": can_data["candidate_last_name"],
			"candidate_personal_email": can_data["candidate_personal_email"],
			"report_name": report,
		}
		qr = exec_bq_query(query, params=params)
		if qr[0] == True:
			candidate_added_ok = True
	return candidate_added_ok

def mindset_assessments(request=None):
	"""Run the end-to-end mindset assessment workflow for local or HTTP execution."""

	try:

		script_error = False

		log_stream.truncate(0)
		log_stream.seek(0)

		log.info("")
		log.info("-" * 60)
		log.info("-" * 60)
		log.info(SCRIPT_NAME)
		log.info("-" * 60)
		log.info("-" * 60)
		log.info("")

		lh_candidate_details = get_candidate_details_from_lh()
		gcp_candidate_details = exec_bq_query(f"SELECT * FROM `focal-rig-396713.CandidateAssessment.{gcp_table}`")
		log.info("")
		log.info("-" * 60)
		log.info("PROCESSING CANDIDATE DATA INTO GCP")
		log.info("-" * 60)
		log.info("")

		if len(lh_candidate_details):
			for lh_can in lh_candidate_details:

				log.info("")
				log.info(
					f"	ROLE NUMBER: {lh_can['search_number']} ({lh_can['search_id']}), SELECTION PROCESS NUMBER: {lh_can['candidacy_number']} ({lh_can['candidacy_id']}), NAME: {lh_can['candidate_first_name']} {lh_can['candidate_last_name']} ({lh_can['contact_id']})")

				# check if candidate is in GCP already or not
				candidate_in_gcp = False
				for gcp_can in gcp_candidate_details[2]:
					if (int(gcp_can['search_id']) == int(lh_can['search_number'])) and (int(gcp_can['candidacy_id']) == int(lh_can['candidacy_number'])):
						candidate_in_gcp = True

				if not candidate_in_gcp:
					search_lead_data = get_search_lead(lh_can['search_id'])
					if lh_can['candidate_personal_email'] is not None:
						log.info(f"		CANDIDATE IS READY FOR ASSESSMENT")
						if process_notification(lh_can, search_lead_data, 3):
							add_candidate_to_bq(lh_can)
							process_notification(lh_can, search_lead_data, 4)
					else:
						log.error(f"	CANDIDATE IS NOT READY FOR ASSESSMENT, NO EMAIL")
						process_notification(lh_can, search_lead_data, 2)
						log.info("	REVERTING CANDIDATE IN LIGHTHOUSE")
						if not revert_candidate_status(lh_can):
							script_error = True
				else:
					log.info(f"		CANDIDATE IS ALREADY IN ASSESSMENT")
		else:
			log.warning(f"	NO CANDIDATES FOUND")

		### get any new files that have been added to GCP bucket since last run
		current_bucket_files = get_bucket_files()

		if len(current_bucket_files) > 0:
			log.info("NEW FILES FOUND IN GCP BUCKET")
			log.info("MOVING FILES FROM BUCKET TO SHAREPOINT")
			files_added_to_sp = move_from_bucket_to_sp(current_bucket_files)

			# update BQ with latest files
			update_bq_report_data(files_added_to_sp)
		else:
			log.info("	NO NEW VALID FILES FOUND IN GCP BUCKET")
			log.info("")

		### update any expired sp urls
		# update_sp_urls()

		### upload new files to thrive contact record
		lh_upload_document(lh_candidate_details)
		#
		### notify search leads
		notify_search_lead(lh_candidate_details)
		#
		### move completed candidates to 'Candidate Assessment complete?' stage
		move_completed_candidates(lh_candidate_details)
		#
		### notify search leads
		notify_search_lead(lh_candidate_details)
		#
		# ### rebuild report spreadsheet in SP and GCP Bucket
		build_report_spreadsheet()

		if script_error:
			send_email_log("andy.harris@leathwaite.com", f"{SCRIPT_NAME}: FAILED", log_stream.getvalue(),
					    	log_stream.getvalue(), "log.txt")
		# else:
		# 	send_email_log("andy.harris@leathwaite.com", f"{SCRIPT_NAME}: PASSED", log_stream.getvalue(),
		# 				   log_stream.getvalue(), "log.txt")

		return "OK", 200

	except Exception as e:

		log.error(str(traceback.format_exc()))
		send_email_log("andy.harris@leathwaite.com", f"{SCRIPT_NAME}: FAILED", log_stream.getvalue(), log_stream.getvalue(), "log.txt")
		return "ERROR", 500

if __name__ == '__main__':
	mindset_assessments()
