# Database Setup Script

## Description

Python script that automatically creates a PostgreSQL database with traffic management system tables and imports data from Excel files.

## Requirements

- Python 3.7+
- PostgreSQL
- Excel file named `CODD.xlsx`

## Installation

1. Install required packages:

```bash
pip install pandas psycopg2-binary openpyxl
```
2. Make sure PostgreSQL is running
3. Place files in same folder:

- [create_database.py](\create_database.py)
- [CODD.xlsx](\CODD.xlsx)

4. Run the script:

```bash
python create_database.py
```
## Default PostgreSQL Settings

- Host: localhost
- User: postgres
- Password: 1234
- Database: codd

## What It Does

- Creates database and tables automatically
- Imports data from all Excel sheets
- Handles errors with detailed logging

## Troubleshooting

If installation fails, try:

```bash
python -m pip install --upgrade pip
pip install pandas psycopg2-binary openpyxl --user
```
***Check that PostgreSQL service is running before executing the script!!!***

