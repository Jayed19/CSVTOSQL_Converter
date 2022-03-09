import pandas as pd
import re


#Read XLXS File
df = pd.read_excel('Test.xlsx',keep_default_na=False)
columnnamexlxlist=df.columns.tolist()
columnnamexlxlist = [each_string.lower() for each_string in columnnamexlxlist]

print("Column Name Found in XLXS file: ")
print(columnnamexlxlist)



#Query Template
query="UPDATE HSDL_APPLICATION SET REFERENCE_NUMBER='{ref}',LICENSE_NUMBER_EN='{dl}',APPLICATION_STATUS={app_status},CARD_STATUS={card_status},AFIS_STATUS={afis_status},ISSUE_DATE=TO_TIMESTAMP('{issue_date}','DD-MON-YYYY HH12: MI:SS:FF AM'),EXPIRY_DATE=TO_TIMESTAMP('{expiry_date}','DD-MON-YYYY HH12: MI:SS:FF AM'),STATUS={status} WHERE ID={id};"
query=query.lower()
columnsFoundInQuery=re.findall(r"\{\w+\}", query) #regex for finding fields which close with second bracket in the SQL

SQLFieldsDictionary = {}
for columnnamesql in columnsFoundInQuery:
    columnnamesql=columnnamesql.replace("{","")
    columnnamesql=columnnamesql.replace("}","")
    #print("Column Name in SQL Query= "+columnnamesql)
    if columnnamesql not in columnnamexlxlist:
        print("The variable name inside {}, name "+ columnnamesql+" is not found or matched with your browsed xlxs file column name. Please check your written SQL again.")
    else:
        columnid=columnnamexlxlist.index(columnnamesql)
        SQLFieldsDictionary[columnnamesql] = columnid

print("Column ID index dictionary which fields mentioned in the SQL Query: ")
print(SQLFieldsDictionary)



#Row Finding
f = open("test.sql", "a")
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