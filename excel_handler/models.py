from django.db import models

class UploadedExcelFile(models.Model):
    file = models.FileField(upload_to='uploads/')
    uploaded_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"Uploaded at {self.uploaded_at}"

class ProcessedData(models.Model):
    original_file = models.ForeignKey(UploadedExcelFile, on_delete=models.CASCADE)
    data = models.JSONField()  # Store manipulated data as JSON
    processed_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"Processed data for {self.original_file}"
