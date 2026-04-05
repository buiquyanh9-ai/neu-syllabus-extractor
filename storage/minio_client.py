"""
storage/minio_client.py
Đọc .docx từ MinIO và ghi JSON đã parse ngược lại MinIO.
"""

import json
import logging
import re
from io import BytesIO
from pathlib import Path, PurePosixPath

from minio import Minio
from minio.error import S3Error

log = logging.getLogger(__name__)


class MinIOStore:
    def __init__(
        self,
        endpoint:    str,
        access_key:  str,
        secret_key:  str,
        bucket:      str,
        input_prefix:  str,   # VD: "courses-raw/syllabus/"
        output_prefix: str,   # VD: "courses-processed/"
        secure: bool = False,
    ):
        self.client         = Minio(endpoint, access_key=access_key, secret_key=secret_key, secure=secure)
        self.bucket         = bucket
        self.input_prefix   = input_prefix.rstrip("/") + "/"
        self.output_prefix  = output_prefix.rstrip("/") + "/"

    # ── Đọc ──────────────────────────────────────────────────────────────

    def list_docx(self) -> list[str]:
        """Trả về danh sách object_name của tất cả .docx trong input_prefix."""
        result = []
        objects = self.client.list_objects(self.bucket, prefix=self.input_prefix, recursive=True)
        for obj in objects:
            if obj.object_name.lower().endswith(".docx"):
                result.append(obj.object_name)
        log.info(f"Tìm thấy {len(result)} file .docx tại {self.bucket}/{self.input_prefix}")
        return result

    def download(self, object_name: str) -> bytes:
        """Tải nội dung binary của một object."""
        resp = self.client.get_object(self.bucket, object_name)
        try:
            return resp.read()
        finally:
            resp.close()
            resp.release_conn()

    # ── Ghi ──────────────────────────────────────────────────────────────

    def _output_key(self, object_name: str) -> str:
        """
        Chuyển đổi input path → output path.
        VD: courses-raw/syllabus/Tái_lập_QTTH1120.docx
            → courses-processed/Tái_lập_QTTH1120.json
        """
        filename = PurePosixPath(object_name).name
        stem     = re.sub(r"\.docx$", "", filename, flags=re.IGNORECASE)
        return f"{self.output_prefix}{stem}.json"

    def already_processed(self, object_name: str) -> bool:
        """Kiểm tra file JSON đầu ra đã tồn tại chưa (skip logic)."""
        key = self._output_key(object_name)
        try:
            self.client.stat_object(self.bucket, key)
            return True
        except S3Error as e:
            if e.code == "NoSuchKey":
                return False
            raise

    def upload_json(self, object_name: str, data: dict) -> str:
        """Upload dict dưới dạng JSON vào MinIO. Trả về output key."""
        key     = self._output_key(object_name)
        payload = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
        buf     = BytesIO(payload)
        self.client.put_object(
            bucket_name  = self.bucket,
            object_name  = key,
            data         = buf,
            length       = len(payload),
            content_type = "application/json",
        )
        return key

    # ── Tiện ích ──────────────────────────────────────────────────────────

    def filename_of(self, object_name: str) -> str:
        return PurePosixPath(object_name).name
