import regex
import openpyxl as xl
# import os

class Company:
    def __init__(self, id, type, companyName, eligibility,branches, jobProfile):
        self.id = id
        self.type = type
        self.companyName = companyName
        self.eligibility = eligibility
        self.branches = branches.split(',')
        self.jobProfile = jobProfile
        self.cgpa = regex.split(r'CGPA',self.eligibility)[0]
    
    def printCompany(self):
        print("ID:",self.id)
        print("Offer Type:",self.type)
        print("Company Name:",self.companyName)
        print("Eligibility:",self.eligibility)
        print("Job Profile:",self.jobProfile)
        print("-----------------------------------------------------------------------------------------------------------------------------------")


    

wb = xl.load_workbook('Companies.xlsx')
data=wb['Sheet1']

companies=[]

for i in data:
    companies.append(
        Company(i[0].value,i[2].value,i[3].value,i[4].value,i[5].value, i[6].value)
    )
companies.remove(companies[0])
companies.sort(key=lambda x: x.companyName)

cgpa=float(7.28)
branch="CC"

# cgpa = float(input("Enter your CGPA:"))
# branch = input("Enter Branch:")
print("File written to {}_{}.txt".format(branch,cgpa))
# os.sleep(10)

with open('{}_{}.txt'.format(branch,cgpa),'w') as f:
    for i in companies:
        if (cgpa>= float(i.cgpa) and (branch in i.branches or i.branches[0]=='All Branches')):
            f.write("ID:"+str(i.id))
            f.write("\n")
            f.write("Offer Type:"+str(i.type))
            f.write("\n")
            f.write("Company Name:"+str(i.companyName))
            f.write("\n")
            f.write("Eligibility:"+str(i.eligibility))
            f.write("\n")
            f.write("Job Profile:"+str(i.jobProfile))
            f.write("\n")
            f.write("-----------------------------------------------------------------------------------------------------------------------------------")
            f.write("\n")

f.close()

input("Press Enter to exit")


