import numpy as np
from openpyxl import Workbook 



expenses = np.array([100,2000,3470,2055,1000,500,600,14,5,15,60,10000,150,200,3000,6900,11,1111,200,367,44000,12,200,900,6700,7899,25000,7900,2,200])

wb = Workbook() 

sheet1 = wb.active
sheet1.title = "Daily Expenses"
sheet1.append(["Day","Expenses(₹)"])

for day, amount in enumerate(expenses, start = 1):
    sheet1.append([day, amount])

# week1 = np.sum(expanses[0:7])
# week2 = np.sum(expanses[7:14])
# week3 = np.sum(expanses[14:21])
# week4 = np.sum(expanses[21:28])
# last_days = np.sum(expanses[28:30])

# print("Week1 expanses:",week1)
# print("Week2 expanses:",week2)
# print("Week3 expanses:",week3)
# print("Week4 expanses:",week4)
# print("Last days expanses:",last_days)

# monthly_total = np.sum(expanses)

# print(f"Your monthly expanses is {monthly_total}")

# max_expanse = np.max(expanses)
# day = np.argmax(expanses) + 1

# print(f"Your highest spending was ₹{max_expanse} on Day {day}")

# average = np.mean(expanses)
# print("Average Daily Expanse is :", round(average ,2))

sheet2 = wb.create_sheet(title= "Monthly Summary")

summary = [
    ("Week 1 Total",np.sum(expenses[0:7])),
    ("Week 2 Total",np.sum(expenses[7:14])),
    ("Week 3 Total",np.sum(expenses[14:21])),
    ("Week 4 Total",np.sum(expenses[21:28])),
    ("Last Days Total",np.sum(expenses[28:30])),
    ("Monthly Expenses",np.sum(expenses)),
    ("Highest Expense", np.max(expenses)),
    ("Highest Expense Day", np.argmax(expenses) + 1),
    ("Average Daily Expense", round(np.mean(expenses), 2))
]

sheet2.append(["Metric","value"])

for item in summary:
    sheet2.append(item)

wb.save("Monthly_Expenses_Report.xlsx")

print("Excel file created successfully..!")