import math
import pandas as pd
from datetime import datetime
class employee:
    def __init__(self, firstName, lastName, gender, birthDate, worBegin, salary,
                 law14Date, propertyValue, deposits, workLeave, propertyPayment, checkComplet, leaveReason, law14per):
        self.compensation = 0

        self.gender = gender
        self.firstName = firstName
        self.lastName = lastName
        self.age = 2020 - datetime.strptime(birthDate, '%Y-%m-%d').year
        self.seniority = (datetime.strptime('2021-12-31', '%Y-%m-%d').year - datetime.strptime(worBegin,
                                                                                               '%Y-%m-%d').year) / 365.25
        self.salary = float(salary)
        if type(law14Date) != float:
            self.law14date = datetime.strptime(law14Date, '%Y-%m-%d')
        else:
            self.law14date = None

        if type(workLeave) != float:
            if workLeave == '-':
                self.workleave = None
            else:
                self.workleave = datetime.strptime(workLeave, '%Y-%m-%d %H:%M:%S')
        else:
            self.workleave = None

        try:
            self.propertyValue = float(propertyValue)
            if pd.isna(self.propertyValue):
                self.propertyValue = 0
        except:
            self.propertyValue = 0
        try:
            self.deposits = float(deposits)
            if pd.isna(self.deposits):
                self.deposits = 0
        except:
            self.deposits = 0
        try:
            self.checkComplet = float(checkComplet)
            if pd.isna(self.checkComplet):
                self.checkComplet = 0
        except:
            self.checkComplet = 0
        try:
            self.propertyPayment = float(propertyPayment)
            if pd.isna(self.propertyPayment):
                self.propertyPayment = 0
        except:
            self.propertyPayment = 0
        self.propertyNetWorth = self.propertyPayment + self.checkComplet

        try:
            self.law14per = float(law14per) / 100
            if pd.isna(self.law14per):
                self.law14per = 0
        except:
            self.law14per = 0

        self.leaveReason = leaveReason
######### importing Data from ExcelFile to CSV file #################
df = pd.read_excel(r'data4.xlsx', header=1)
df.to_csv('csvfile.csv', encoding='utf-8')
df = pd.read_csv('csvfile.csv', encoding='utf-8')
df = df.reset_index()
############################################################

########### creating death log dictionary #################
dl = pd.read_excel(r'deathlog.xlsx', header=1)
dl.to_csv('death.csv', encoding='utf-8')
dl = pd.read_csv('death.csv', encoding='utf-8')
dl = dl.reset_index()
deathdict = {}
for i in dl.iloc:
    deathdict[i['age']] = i['q(x)']
############################################################

########### creating DiscountRate dictionary #############
discountRate = pd.read_excel(r'data4.xlsx', header=2, sheet_name=1, usecols="A,B")
discountRate.to_csv('discountRate.csv', encoding='utf-8')
discountRate = pd.read_csv('discountRate.csv', encoding='utf-8')
discountRate = discountRate.reset_index()

discountRateDict = {}
for i in discountRate.iloc:
    discountRateDict[i['שנה ']] = i['שיעור היוון']
###################################################################
employeeArray = []
for i in df.iloc:
    e = employee(i['שם '], i['שם משפחה'], i['מין'], i['תאריך לידה'], i['תאריך תחילת עבודה '], i['שכר '],
                 i['תאריך  קבלת סעיף 14'], i['שווי נכס'], i['הפקדות'],
                 i['תאריך עזיבה '], i['תשלום מהנכס'], i["השלמה בצ'ק"], i['סיבת עזיבה'], i['אחוז סעיף 14'])
    employeeArray.append(e)
