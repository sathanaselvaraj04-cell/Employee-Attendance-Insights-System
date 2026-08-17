import pandas as pd
import numpy as np
import matplotlib.pyplot as plt 
from charts import generate_all_charts
print("Libraries imported successfully!")

df=pd.read_excel("attendance.xlsx")
# print(df.head())

#================================================================
#ATTENDANCE SUMMARY OF ALL EMPLOYEES
#================================================================

attendance_summary=df["Status"].value_counts()
# print("\nAttendance Summary:")
# print(attendance_summary)

#================================================================
#ATTENDANCE PERCENTAGE OF EACH EMPLYOEE
#================================================================

present_days=df[df["Status"]=="Present"].groupby("Employee_Name").size()
total_days=df.groupby("Employee_Name").size()
attendance_percentage=((present_days/total_days)*100).round(2)
attendance_percentage=attendance_percentage.reset_index()
attendance_percentage.columns=["Employee Name","Attendance percentage(%)"]
attendance_percentage.index=attendance_percentage.index+1
# print("/nAttendance percentage:")
# print(attendance_percentage.to_string(col_space=10))

#================================================================
# CONVERT CHECK_IN AND CHECK_OUT TO DATETIME
#================================================================

df["Check_In"] = pd.to_datetime(df["Check_In"], format="%H:%M:%S", errors="coerce")
df["Check_Out"] = pd.to_datetime(df["Check_Out"], format="%H:%M:%S", errors="coerce")


#================================================================
 # CALCULATE WORKING HOURS
 #================================================================

df["Working_Hours"] = (df["Check_Out"] - df["Check_In"]).dt.total_seconds() / 3600
df["Working_Hours"]=df["Working_Hours"].round(2)
print("\nWorking Hours:")
print(df[["Employee_Name", "Date", "Working_Hours"]].to_string(col_space=15))

#================================================================
 #STORING THIS DATA IN THE EXCEL FILE
 #================================================================

# with pd.ExcelWriter( "attendance.xlsx",engine="openpyxl",mode="a",
#     if_sheet_exists="replace"
# ) as writer:
#     df[["Employee_Name", "Date", "Working_Hours"]].to_excel(
#         writer,
#         sheet_name="Working_Hours",
#         index=False
# )
# print("Working Hours sheet added successfully!")

#================================================================
#CALCULATE AVERAGE WORKING HOURS
#================================================================

average_hours = df.groupby("Employee_Name")["Working_Hours"].mean().round(2)
average_hours = average_hours.reset_index()
average_hours.columns = ["Employee Name", "Average Working Hours"]
average_hours.index = average_hours.index + 1

print("\nAverage Working Hours:")
print(average_hours.to_string(col_space=10))

#================================================================
#STORING THAT SHEET IN EXCEL
#================================================================

# with pd.ExcelWriter("attendance.xlsx", engine="openpyxl",mode="a",
#     if_sheet_exists="replace"
# ) as writer:
#     average_hours.to_excel(
#         writer,
#         sheet_name="Average_Working_Hours",
#         index=False
#     )
# print("Average Working Hours sheet added successfully!")

# ===============================================================
# CREATING AN TABLE AND TO STORE THE EMPLOYEE SUMMARY TABLE
# ===============================================================

summary = df.groupby("Employee_Name").agg(
    Present=("Status", lambda x: (x == "Present").sum()),
    Absent=("Status", lambda x: (x == "Absent").sum()),
    Leave=("Status", lambda x: (x == "Leave").sum()),
    Average_Working_Hours=("Working_Hours", lambda x: x.dropna().mean())
).reset_index()

summary["Attendance_Percentage"] = (
    summary["Present"] /
    (summary["Present"] + summary["Absent"] + summary["Leave"])
) * 100

summary["Attendance_Percentage"] = summary["Attendance_Percentage"].round(2)
summary["Average_Working_Hours"] = summary["Average_Working_Hours"].round(2)

print("\nEMPLOYEE SUMMARY")
print(summary)

# =======================================================================
# STORING THIS IN A NEW EXCEL FILE 
# ========================================================================

with pd.ExcelWriter(
    "attendance.xlsx",
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace"
) as writer:

    summary.to_excel(
        writer,
        sheet_name="Employee_Summary",
        index=False
    )

print("Employee Summary sheet added successfully!")

print("\nEmployee Summary saved successfully!")

# ============================================================
# BEST EMPLOYEE (HIGHEST ATTENDANCE)
# ============================================================

best_employee = summary.loc[summary["Attendance_Percentage"].idxmax()]

print("\n==============================")
print("BEST EMPLOYEE")
print("==============================")
print(best_employee)

# ============================================================
# EMPLOYEE WITH MOST ABSENCES
# ============================================================

most_absent = summary.loc[summary["Absent"].idxmax()]

print("\n==============================")
print("EMPLOYEE WITH MOST ABSENCES")
print("==============================")
print(most_absent)

# ============================================================
# HIGHEST AVERAGE WORKING HOURS
# ============================================================

hard_worker = summary.loc[summary["Average_Working_Hours"].idxmax()]

print("\n==============================")
print("HIGHEST AVERAGE WORKING HOURS")
print("==============================")
print(hard_worker)

insights = pd.DataFrame({
    "Category": [
        "Best Employee",
        "Most Absent Employee",
        "Highest Average Working Hours"
    ],
    "Employee": [
        best_employee["Employee_Name"],
        most_absent["Employee_Name"],
        hard_worker["Employee_Name"]
    ]
})

with pd.ExcelWriter(
    "attendance.xlsx",
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace"
) as writer:

    insights.to_excel(
        writer,
        sheet_name="Project_Insights",
        index=False
    )

print("\nProject Insights sheet added successfully!")

# -----------------------------
# Generate charts
# -----------------------------
try:
    generated = generate_all_charts(df)
    print("Charts generated:")
    for k, v in generated.items():
        print(f" - {k}: {v}")
except Exception as e:
    print("Failed to generate charts:", e)
