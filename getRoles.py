import requests as r
import pandas as p
import re
import json
import subprocess
import xlsxwriter as xls

# Aux method library
def create_xlsx():
    resultset = xls.Workbook("users_result_set.xlsx")
    return resultset

def write_in_excel(roles, currentSheet, email, row):
  for k in range(len(roles)):
      currentSheet.write(row,0,email)
      currentSheet.write(row,1,roles[k]["role"]["applicationCode"])
      currentSheet.write(row,2,roles[k]["role"]["roleName"])
      if roles[k]["salesOrgCode"] != None : 
        currentSheet.write(row,3,roles[k]["salesOrgCode"])
      else :
        currentSheet.write(row,3,"GLOBAL")

      row += 1
  return row
      
resultset = create_xlsx()
row = 0
currentSheet = resultset.add_worksheet("worksheet")
currentSheet.write(row,0,"email")
currentSheet.write(row,1,"application")
currentSheet.write(row,2,"role")
currentSheet.write(row,3,"salesOrgCode")
row += 1
filename = "cadi_users.xlsx"
collName = 'emails'
xlsx = p.read_excel(filename)
arr = xlsx[collName].tolist()
print(arr)
idls = []
salesOrg = {
  "DK": "A001",
  "UK": "B001",
  "NO": "C001",
  "SE": "D001",
  "FI": "E001",
  "FR": "F001",
  "DE": "L001"
}
roles = {
  "regular_ont": "APP_ONT_USER",
  "regular_oft": "APP_OFT_USER",
  "super_user_ont": "APP_ONT_SUP_USER",
  "super_user_oft": "APP_OFT_SUP_USER",
  "ordering": "APP_ORDER_USER"
}
action = {
  "add": "ENABLE",
  "remove": "DISABLE"
}

for i in range(len(arr)):
   response = r.get("http://localhost:8081/services/users?type=INTERNAL&email={user}".format(user = arr[i]))
   data = response.json()
   roles=(data.get("users")[0].get("roles"))
   print(roles)
   row = write_in_excel(roles["userRoles"],currentSheet,arr[i], row)
   row = write_in_excel(roles["userSalesOrgRoles"], currentSheet,arr[i], row)
resultset.close()
 

     

    
 