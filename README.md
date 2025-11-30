# Blast Project

A Django web application for processing Excel files. Upload Excel files, filter data based on the first column values (if numeric, keeps rows where value > 10), and download the processed results.

## Prerequisites

- Python 3.13
- Virtual environment (already set up in `venv/`)

## Setup and Running

1. Activate the virtual environment:
   ```bash
   source venv/bin/activate
   ```

2. Run database migrations (if needed):
   ```bash
   python manage.py migrate
   ```

3. Start the development server:
   ```bash
   python manage.py runserver
   ```

4. Open your browser and go to `http://127.0.0.1:8000/`

## API Endpoints

- `POST /excel/upload/` - Upload an Excel file for processing
- `GET /excel/download/<file_id>/` - Download the processed Excel file
- `/admin/` - Django admin interface

## Dependencies

- Django 5.2.8
- pandas
- openpyxl

## Project Structure

- `blast_project/` - Main Django project settings
- `excel_handler/` - App for handling Excel file uploads and processing
- `db.sqlite3` - SQLite database file
- `manage.py` - Django management script