import os
from dotenv import load_dotenv
load_dotenv()

MINIO_ENDPOINT    = os.getenv("MINIO_ENDPOINT",    "203.113.132.48:8008")
MINIO_ACCESS_KEY  = os.getenv("MINIO_ACCESS_KEY",  "course2")
MINIO_SECRET_KEY  = os.getenv("MINIO_SECRET_KEY",  "course2-s3-uiauia")
MINIO_SECURE      = os.getenv("MINIO_SECURE",      "false").lower() == "true"
MINIO_BUCKET      = os.getenv("MINIO_BUCKET",      "syllabus")
MINIO_INPUT_PREFIX  = os.getenv("MINIO_INPUT_PREFIX",  "courses-raw/syllabus/")
MINIO_OUTPUT_PREFIX = os.getenv("MINIO_OUTPUT_PREFIX",  "courses-raw/qb/")
MAX_WORKERS = int(os.getenv("MAX_WORKERS", "5"))
