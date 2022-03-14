import pandas as pd
import re
from tkinter import *
from tkinter import messagebox
from tkinter import messagebox
import os

#Window
root = Tk()
root.title("XLXS File TO SQL Converter")
root.geometry('650x450')
root.option_add("*Font", ('Arial', 10))

#Common Style
textbox_style={'relief':'sunken', 'highlightthickness':1, 'borderwidth':1,'font':'Arial 16'}

# Query Label and TextBox
QueryLabel = Label(root, text="SQL Template: ")
QueryLabel.grid(row=1,column=1)
QueryTextBox= Text(
    root,
    height=12,
    width=45,
    **textbox_style
)
QueryTextBox.grid(row=1,column=2,columnspan=2)

#Query Example Label

QueryExample = Label(root, text="Example: Update hsdl_application set reference_number='{ref}',status={state} where id={id};\n Note: For string type field use single quatation.")
QueryExample.grid(row=2,column=2)


#Browse XLX File
from tkinter import filedialog
filedir=''
def browseFiles():
    global filedir
    filedir = filedialog.askopenfilename(
                                          title = "Select XLXS File",
                                          filetypes = (("all files","*.*"),
                                                        ("XLXS files","*.xlxs*"),
                                                        ("XLX files","*.xlx*")
                                                       ))
    

 
    print(filedir)
    return filedir

browsebtn_style={'padx':5, 'pady':5,'borderwidth':5, 'relief':'raised','font':'Arial 9'}
BrowseFilebtn = Button(root, text = "Browse XLX File",command=browseFiles,**browsebtn_style)
BrowseFilebtn.grid(row=3,column=2)



#Run SQL Generation
def GenerateSQL():
    if QueryTextBox.get("1.0",'end-1c')=='':
        messagebox.showwarning('Alert',"SQL template is empty!")
    
    elif filedir=='':
        messagebox.showwarning('Alert',"Excell file is not browsed!")

    else:
        #Read XLXS File
        df = pd.read_excel(filedir,keep_default_na=False)
        columnnamexlxlist=df.columns.tolist()
        columnnamexlxlist = [each_string.lower() for each_string in columnnamexlxlist]
        print("Column Name Found in XLXS file: ")
        print(columnnamexlxlist)


        #Query Template
        #query="UPDATE HSDL_APPLICATION SET REFERENCE_NUMBER='{ref}',LICENSE_NUMBER_EN='{dl}',APPLICATION_STATUS={app_status},CARD_STATUS={card_status},AFIS_STATUS={afis_status},ISSUE_DATE=TO_TIMESTAMP('{issue_date}','DD-MON-YYYY HH12: MI:SS:FF AM'),EXPIRY_DATE=TO_TIMESTAMP('{expiry_date}','DD-MON-YYYY HH12: MI:SS:FF AM'),STATUS={status} WHERE ID={id};"
        query=QueryTextBox.get("1.0",'end-1c')
        query=query.lower()
        columnsFoundInQuery=re.findall(r"\{\w+\}", query) #regex for finding fields which close with second bracket in the SQL
        SQLFieldsDictionary = {}

        for columnnamesql in columnsFoundInQuery:
            columnnamesql=columnnamesql.replace("{","")
            columnnamesql=columnnamesql.replace("}","")
            #print("Column Name in SQL Query= "+columnnamesql)
            if columnnamesql not in columnnamexlxlist:

                result = messagebox.askquestion("Column Missing!", columnnamesql+" is not found or matched in browsed file. Please check SQL Template or File again. Do you want to continue with this problem?", icon='warning')
                if result == 'yes':
                    pass
                else:
                    return


            else:
                columnid=columnnamexlxlist.index(columnnamesql)
                SQLFieldsDictionary[columnnamesql] = columnid

        print("Column ID index dictionary which fields mentioned in the SQL Query: ")
        print(SQLFieldsDictionary)



        #Row Finding
        f = open("test.sql", "a",encoding='utf-8')
        for row in df.itertuples(index=False,name='eachrow'):
            queryreplaced=query
            for x, y in SQLFieldsDictionary.items():
                #print(str(x)+"=" )
                #print(row[y])
                x="{"+x+"}"
                if row[y]=='':
                    queryreplaced=queryreplaced.replace(x,"''")
                else:
                    queryreplaced=queryreplaced.replace(x,str(row[y]))
        
            #print(queryreplaced)
            f.write(queryreplaced)
            f.write("\n")

        f.close()




GenerateSQLFilebtn = Button(root, text = "Generate SQL File",command=GenerateSQL,**browsebtn_style)
GenerateSQLFilebtn.grid(row=4,column=2)

root.mainloop()











