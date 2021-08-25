# -*- coding: utf-8 -*-
"""
Created on Sat May 15 16:13:10 2021

@author: joetan
"""


#pyinstaller --add-data "C:\Users\joetan\Anaconda\Lib\site-packages\pandas\io\formats\templates\html.tpl;." --onefile "G:\FFMC-Risk\Issuer Exposure\automation\Py_issuer_mapping.py"

from configparser import ConfigParser
import datetime
import win32com.client as win32
import os
import pandas as pd

def create_htmlbody(df):
      htmlbody= '''
            <html>
            <head>
            <style>
                table { 
                    margin-left: auto;
                    margin-right: auto;
                    width: 100%                
                }
                table, th, td {
                    border: 1px solid black;
                    border-collapse: collapse;
                }
                th, td {
                    padding: 1px;
                    text-align: left;
                    font-family: Calibri;
                    font-size: 14px;
                }
                table tbody tr:hover {
                    background-color: #dddddd;
                }
                .wide {
                    width: 65%; 
                }
            </style>
            </head>
                '''
      htmlbody+= '<p style="font-family: Calibri; font-size: 14px"> Hi, <br><br>Please refer to table below or attached file for unmapped Issuers'
      htmlbody += df.to_html(index=False)      

      return htmlbody
def email(dt,body,recipients,filepath):
     
      outlook = win32.Dispatch('outlook.application')
      mail = outlook.CreateItem(0)
      mail.Subject =  'Unmapped Issuers as of ' + dt.strftime("%Y-%m-%d")


      import getpass
      if getpass.getuser() == 'joetan':
            user = 'Joe'
      else:
            user = 'Risk'
      mail.Attachments.Add(filepath)
   
      mail.To = recipients
      mail.HTMLBody = body +'<br>Thanks</br>' + '<br> {} </br>'.format(user)
      
      mail.Display()

def restart_date(configfilepath):

        config = ConfigParser()
        config.read(configfilepath)
        configfile = open(configfilepath,'w')
        config.set('params','date',"")
        config.write(configfile)
        configfile.close() 

def main():
      
      tdy = datetime.datetime.today().date()
      if tdy.isoweekday() in [1,2]:   #Monday and Tuesday use t-4 logic; Remaining use t-2
            delta = 4
      else:
            delta = 2
    
    
      path = os.path.dirname(os.path.realpath(__file__))

 #     path = r'G:\FFMC-Risk\Issuer Exposure\automation'
      dt = datetime.datetime.today()+ datetime.timedelta(days=-delta)  
      config = ConfigParser()
      configfilepath = os.path.join(path,'Issuer_exposure_config.ini')
      config.read(configfilepath)    
      dt_adhoc = config.get('params','date')
      dt_list = dt_adhoc.split('-')
      if len(dt_adhoc)>0:
            dt =  datetime.datetime(int(dt_list[2]),int(dt_list[1]),int(dt_list[0]))
  
  
      
      masterlist=[]
      mapping_file = config.get('params', 'mapping_file')
      barradump_path = config.get('params','barradump_issuer')
      unmapped_path = os.path.join(os.path.split(barradump_path)[0],'unmapped issuers')
      mapping_table = pd.read_excel(mapping_file, header=2)
      recipients = config.get('params','recipients')
      date_barrapath = os.path.join(barradump_path,"Analysis Date - " + dt.strftime('%Y-%m-%d'))
      for port in os.listdir(date_barrapath):
            if os.path.isdir(os.path.join(date_barrapath,port)):
                  port_path = os.path.join(date_barrapath,port)
                  for files in os.listdir(port_path):
                        df = pd.read_excel(os.path.join(port_path, files),header=0,skiprows=[1])
                        print('Reading '+ port + '...............')
                        df.insert(0,'Portfolio',port)
                  masterlist.append(df)  
      masterdf = pd.concat(masterlist)
      
      masterdf['Issuer'] = masterdf['Issuer'].str.rstrip()
      masterdf['Issuer'] = masterdf['Issuer'].str.lower()
      mapping_table['Issuer'] = mapping_table['Issuer'].str.lower()
      joindf = masterdf.merge(mapping_table, on='Issuer', how='left')

      unmapped_df = joindf.loc[(joindf['Obligor Name'].isnull()) & (~joindf['Issuer'].isnull()) & (~joindf['ISIN'].isnull()) & (joindf['Issuer']!= 'proxy issuer')]
      unmapped_df = unmapped_df.loc[(~unmapped_df['Issuer'].str.contains('proxy'))]
   
      unique_df = unmapped_df[['Portfolio','ISIN','Asset Name','Issuer']].drop_duplicates(subset=['Issuer'])
      unique_df = unique_df.sort_values(by=['Portfolio','Issuer'])
      
      unique_df['Issuer'] =  unique_df['Issuer'].str.upper()
      filepath = os.path.join(unmapped_path,dt.strftime("%Y%m%d")+'_unmappedIssuers.xlsx')
      unique_df[unique_df.columns[1:]].to_excel(filepath,index=False)
      htmlbod = create_htmlbody(unique_df[unique_df.columns[1:]])
      
      email(dt,htmlbod,recipients,filepath)
      restart_date(configfilepath)

if __name__=='__main__':
    main()