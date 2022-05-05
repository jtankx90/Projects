from selenium import webdriver
import pandas as pd
import datetime as datetime
import os
import pyodbc


#def main():
cwd = os.path.abspath(os.getcwd())

driver = webdriver.Chrome(cwd+"\\chromedriver.exe") ##path for chromedriver.exe
driver.get("https://www.mom.gov.sg/employment-practices/public-holidays")
finaloutput=[]
years=[]
years = [elems.get_attribute('id') for elems in driver.find_elements_by_css_selector('.ui-tabs div')] #retrieve the years available on the website
years = list(filter(None,years)) #remove Nonetype 

for year in years:
    outputlist =[]
    driver.get("https://www.mom.gov.sg/employment-practices/public-holidays#"+year)
    for j in [1,2]: #j for the first 2 columns in the table
        date_list =[]
        for rows in driver.find_elements_by_xpath("//div[@id='"+year+"']//table/tbody/tr/td[" +str(j)+str("]")):
            if rows is not None:
                old_rows = rows.text.replace("<br>","") #format data
                old_rows = old_rows.split("\n") #format data
                if len(old_rows) !=1: #for rows in table with only 1 entry
                    for i in old_rows:
                        date_list.append(datetime.datetime.strptime(i,'%d %B %Y').date()) if j==1 else date_list.append(i)
                else: #for rows in table that have >1 entry  for e.g. cny
                    temp = (str(w) for w in old_rows)
                    temp = "".join(temp)
                    date_list.append(datetime.datetime.strptime(temp,'%d %B %Y').date()) if j ==1 else date_list.append(temp)
            else:
                print("No data from Website")
                driver.close()
        for x,y in enumerate(date_list):
            if j==1: #insert monday as date as holiday if ph falls on sunday
                if y.isoweekday()==7:
                    date_list.insert(x+1,y+datetime.timedelta(days=1))
            elif j==2:
                if y =='Sunday': #insert monday as day as holiday if ph falls on sunday
                    date_list.insert(x+1,'Monday')

        outputlist.append(date_list)
        date_list=[]            
    df = pd.DataFrame(outputlist)
    df = df.transpose()
   # df.columns = ['Date','Day']
    finaloutput.append(df)        
    df.empty
driver.close()
finaloutput = pd.concat(finaloutput) #concat all tables into 1



   

 sql = "SELECT * from tbl_Public_Holidays"
 conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+r'C:\Users\joetan\Desktop\db.accdb',autocommit=True)
 cursor = conn.cursor()
 excel_results = cursor.execute(sql).fetchall()

 cursor.close()

 output = [list(j) for j in excel_results]
 df = pd.DataFrame(output,columns = ['Date','Day_'])
 df['Date'] = df['Date'].dt.date

 df_insert = finaloutput.merge(df, on=['Date'],how='left',indicator=True)
 df_insert = df_insert[df_insert['_merge'] =='left_only'][['Date','Day']]
 newdbnumber = len(df_insert['Date'])

 if newdbnumber > 0:

     sql = "{CALL qry_insert_Public_Holidays(?,?)}"
     cursor = conn.cursor()
     cursor.executemany(sql,df_insert.values.tolist())
     cursor.commit()
     cursor.close()
     conn.close()
 import win32com.client as win32
 outlook = win32.Dispatch('outlook.application')
 mail = outlook.CreateItem(0)
 mail.To = 
 mail.Subject = 'Public Holidays Table Update'
 body = '<p style="font-family: Calibri; font-size: 14px">'
 if newdbnumber ==0:
     body = body+ 'No new Public Holidays found'
 else:
     body = body+ str(newdbnumber) +' New Public Holidays found . Refer to table below: ' + '<br>' + df_insert.to_html() + ' </br>'
 mail.HTMLBody = body
 mail.Send()
        
if __name__=='__main__':
      main()   
