FROM python:3.11
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY main.py .
ENTRYPOINT ["python","-W","ignore","-u"]
CMD ["main.py"]
#CMD ["py-spy", "record", "--output", "profile.svg", "--", "python", "-W","ignore","-u", "main.py"]
