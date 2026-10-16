# Data Click Education — ERP-Level Student Management System

## Stack
Frontend: HTML5 + CSS3 + Vanilla JavaScript  
Backend: Python + Flask  
Data: Pandas + CSV  
Reports: Excel (`openpyxl`) + PDF (`reportlab`)

## ERP modules
- Admin login / logout
- Dashboard with live counters
- New Admission
- Student search
- Edit student
- Delete student
- Student photo upload
- Student ID card PDF
- Admission Enquiry
- Fee Collection
- Fee Receipt PDF
- Attendance
- Staff Management
- Exams & Marks
- Automatic grade calculation
- Student Promotion
- Reports & Excel export
- Student academic/fee/attendance report
- Database backup ZIP
- Responsive sidebar UI
- Local CSV database generated automatically by Pandas

## Default login
Username: `admin`
Password: `admin123`

Change the default credentials and Flask secret key before using this beyond local development.

## Windows run
```powershell
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Open:
`http://127.0.0.1:5000`

## Database
No MySQL is required. CSV files are stored in `data/`.
Backups are generated in `backups/`.

## Production hardening
Before public deployment, add a real password hash system, CSRF protection, database transactions, role/permission enforcement, secure secret management, HTTPS, audit logs and automated backups.
