# 🧾 Attendance Report Generator

This project automates the creation of **styled Excel attendance reports** from a **MySQL database**.  
It uses **OpenPyXL** for Excel formatting and **Pandas + SQLAlchemy** for data retrieval and transformation.


## 📦 Features

- 🎨 Generates a pre-styled Excel **attendance template** (`template.xlsx`)
- 📅 Groups attendance records **by month**
- 🌈 Applies **color-coded highlights** based on attendance status
- 📊 Calculates key metrics:
  - Attendance rate (performance %)
  - Total working days
  - Late arrivals
  - Work-from-home days
  - Public holidays
  - Leaves not applied
- 🧠 Configurable color themes and logic via `utils.py`


## 🧰 Tech Stack

- **Python 3.9+**
- **Pandas**
- **SQLAlchemy**
- **OpenPyXL**
- **MySQL (with PyMySQL driver)**


## ⚙️ Setup

1. **Clone the repository**
```
  git clone https://github.com/your-username/attendance-report-generator.git
  cd attendance-report-generator
```

2. **Install dependencies**
```
  pip install -r requirements.txt
```

3. **Configure database credentials**
```
  DB_HOST = "your-host"
  DB_USER = "your-username"
  DB_PASS = "your-password"
  DB_NAME = "your-database"
```

4. **Run template generator**
```
  python create_template.py
```

5. **Generate attendance report**
```
  python generate_report.py
```


## 🧾 Example Output
```
⏳ Initiating attendance report for employee ID: 131...
📋 Attendance report exported as Attendance_EmployeeName_Jan01-Oct31_2025.xlsx
📊 Attendance Summary:
Performance         : 48.07
Days Present        : 201.0/209
Late Arrivals       : 28
Public Holidays     : 9
Work From Home      : 22.0
Leaves Not Applied  : 1
```

It includes:
- Monthly attendance breakdown
- Performance summary widgets
- Automatically formatted color-coded rows