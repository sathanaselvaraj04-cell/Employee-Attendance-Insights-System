import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta

# -----------------------------
# CONFIGURATION
# -----------------------------
num_employees = 20
num_days = 30

employees = [f"E{str(i).zfill(3)}" for i in range(1, num_employees + 1)]
names =["Arun", "Priya", "Rahul", "Sneha", "Kiran",
    "Ananya", "Vikram", "Meena", "Rohit", "Divya",
    "Ajay", "Neha", "Karthik", "Pooja", "Sanjay",
    "Aisha", "Manoj", "Keerthi", "Naveen", "Lakshmi"]

start_date = datetime(2026, 7, 1)

status_choices = ["Present", "Absent", "Leave"]

data = []

# -----------------------------
# GENERATE DATA
# -----------------------------
for emp_id, name in zip(employees, names):
    for day in range(num_days):
        date = start_date + timedelta(days=day)

        status = random.choices(
            status_choices,
            weights=[0.75, 0.15, 0.10]  # more realistic: mostly present
        )[0]

        if status == "Present":
            check_in = datetime.strptime("09:00", "%H:%M") + timedelta(minutes=random.randint(-30, 30))
            check_out = datetime.strptime("18:00", "%H:%M") + timedelta(minutes=random.randint(-30, 60))
        else:
            check_in = None
            check_out = None

        data.append([
            emp_id,
            name,
            "IT",
            date.date(),
            status,
            check_in.time() if check_in else None,
            check_out.time() if check_out else None
        ])

# -----------------------------
# CREATE DATAFRAME
# -----------------------------
df = pd.DataFrame(data, columns=[
    "Employee_ID",
    "Employee_Name",
    "Department",
    "Date",
    "Status",
    "Check_In",
    "Check_Out"
])

# -----------------------------
# SAVE TO EXCEL
# -----------------------------
df.to_excel("attendance.xlsx", index=False)

print("Dataset created successfully!")