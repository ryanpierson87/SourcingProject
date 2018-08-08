#The backup
import zipfile as zp
import datetime as dt
from tkinter.filedialog import askopenfilename
from tkinter import filedialog
import pandas as pd
import os
import shutil as shtl
from tkinter import messagebox
import numpy as np

class Program:
###
######## Creates the interface|| links the reference .csv file|| populates data based on the .csv file
###
    def __init__(self, master):
        self.data = pd.read_csv("SourcingPython\\ProgramData.csv")
        self.data.Input[0] = ""
        if pd.isnull(self.data["User"][0]):
            path = os.getcwd()
            fair = path.split("\\")
            for i in fair:
                if i.isdigit():
                    self.data["User"][0] = i
            self.data.to_csv("SourcingPython\\ProgramData.csv", index=False) 


        self.input = Button(text="input",bg="brown", fg="white", relief=RIDGE,width=10, command=self.inputSelect).grid(row=1,column=0)
        self.inLabel = Label(text=self.data.Input[0], relief=SUNKEN,width=30).grid(row=1,column=1)
        self.output = Button(text="output",bg="brown", fg="white", relief=RIDGE,width=10, command=self.outputSelect).grid(row=2,column=0)
        self.outLabel = Label(text=self.data.Output[0][self.data.Output[0].find("/", 15)+ 1:], relief=SUNKEN,width=30).grid(row=2,column=1)
        self.quit = Button(text="Quit", fg="red", relief=RIDGE, width=40, command=root.quit).grid(row=4, column=0, columnspan=2)
        self.zipFile = Button(text="Start Automation", fg="green", relief=RIDGE, width=40, command=self.zippo).grid(row=5, column=0, columnspan=2)
        self.nextTest = Button(text="Secondary Function", fg="blue", relief=RIDGE, width=40, command=self.import_file_and_test).grid(row=6, column=0, columnspan=2)

###############################

###
#####       methods to select the input zip file and output directory
###
    def inputSelect(self):
        self.data.Input.at[0]  = askopenfilename()
        self.data.to_csv("SourcingPython\\ProgramData.csv", index=False)
        self.inLabel = Label(text=self.data.Input[0][self.data.Input[0].rfind("/")+ 1:], relief=SUNKEN,width=30).grid(row=1,column=1) 
        root.update()
    
    def outputSelect(self):
        self.data.Output.at[0]  = filedialog.askdirectory()
        self.data.to_csv("SourcingPython\\ProgramData.csv", index=False)
        self.outLabel = Label(text=self.data.Output[0][self.data.Output[0].find("/", 15)+ 1:], relief=SUNKEN,width=30).grid(row=2,column=1) 
        root.update()
#################################
    def import_file_and_test(self):
        if len(self.data.Input[0]) < 3:
             return
        final = pd.read_excel(self.data["Input"][0])     
            
        
    def zippo(self):
        multi_file = False
#####
##  Extracts a file into a specific location and deleting original 
#####This portion works within the object correctly
####
        if len(self.data.Input[0]) > 3:
            now = dt.datetime.now()
            current ="Sourcing-" + str(now.month) + "_" +str( now.day) + "-"+str(now.year)
            inpt = self.data.Input[0]
            print(inpt)
            self.reports_to_collect = {}
            self.zip_dest = self.data.Input[0][:self.data.Input[0].rfind("/")+ 1] + current
            if inpt[-3:] == "zip":
                multi_file = True
                print("it is zip")
                os.mkdir(self.zip_dest)
                zip_file = inpt
                file = zp.ZipFile(zip_file)
                file.extractall(path=self.zip_dest)
                file.close()
                os.remove(zip_file)
                #The dictionary to divvy up the Buyer Review documents
                #If there are more than one file
                

                ### Loops through the files in the directory of the extracted zip file
                for roots, dirs, files in os.walk(self.zip_dest):
                    for file in files:
                        print(file)
                        report = file
                        initial = report.find(" ") + 1
                        if initial > 0:
                            try:
                                org = report[initial: report.find("-", initial)]
                            except:
                                pass
                            if org in self.reports_to_collect:
                                if report not in self.reports_to_collect:
                                    self.reports_to_collect[org].append(file)
                            else:
                                self.reports_to_collect[org] = []
                                self.reports_to_collect[org].append(file)
            else:
                print("it is not")
                initial = inpt.find(" ") + 1
                org = inpt[initial: inpt.find("-", initial)]
                self.reports_to_collect[org] = []
                self.reports_to_collect[org].append(inpt)
            
        
        

#Moves the data from the files to duplicates of the empty tempalte documents
#Price from: Price(In issue UOM)
        self.columns = ['ORDER UNIT OF MEASURE', 'Order UOM Price','Supplier', \
           'Supplier Site\n(POI preferred)', 'Supplier Item Number','New/Existing Part Number (entered by Loading Team/Code)', 'Ship To']

        self.template = pd.read_excel("SourcingPython\SourcingTemplate.xlsx")
        self.template.dropna(inplace=True)
        #
        #Loops through the above dictionary and and creates the copy of the template for each org
        #For Org
            #Everything appears that it will be done within this loop
        for key in self.reports_to_collect:
            final = pd.DataFrame()
            template1 = self.template.copy()
            #For reports in Org
            print(multi_file)
            if multi_file:
                for i in self.reports_to_collect[key]:
                    print('a')
                    df = pd.read_excel(self.zip_dest+"\\" + i, "New Item Entry")
                    df.drop_duplicates(keep=False, inplace= True)  
                    final = pd.concat([final, df])
            else:
                df = pd.read_excel(inpt, "New Item Entry")
                df.drop_duplicates(keep=False, inplace= True)
                final = df
            #Moves copies the data for each sheet per org into the created template duplicate
            for j in range(len(self.columns)):
                template1[self.columns[j]] = final[self.columns[j]]
            final.to_excel(key + ".xlsx", index=False)
            template1.to_excel(self.data.Output[0]+"//Sourcing "+key + ".xlsx", index=False)
            
            #Save the concatenation as a variable to be used in the first SQL Query
            self.sql1 = "'" + template1["New/Existing Part Number (entered by Loading Team/Code)"]+"'"
#             self.sql1 = self.sql1[np.isnan(self.sql1)]
#             self.sql.remove
            self.sql1 =','.join(str(num) for num in self.sql1)
        if multi_file:
            shtl.rmtree(self.zip_dest)
        else:
            os.remove(inpt)
        messagebox.showinfo("Complete", "Sourcing Templates are available")
            
        
root = Tk()
app = Program(root)
root.mainloop()
root.destroy()