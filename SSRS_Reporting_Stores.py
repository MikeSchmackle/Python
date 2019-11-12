import pyodbc 

import urllib.parse as ulp

import datetime

import requests, configparser, os 
from requests_ntlm import HttpNtlmAuth

import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 

#Global variables for the script
SSRS_SERVER_NAME    = 'reports1'
TEMP_DIRECTORY = 'C:\\Test_Files'
DB_SERVER = 'Server=reportsvm-dc1;'
DB_NAME = 'Database=WITS2;'

#This function connects to a URL, authenticates, and downloads the resulting exported SSRS report in the file format passes via URL
def get_url(url, username, password, output_file):
    """downlaod SSRS Url using credentials supplied"""
    
    # access the URL
    r = requests.get(url,auth=HttpNtlmAuth(username,password))

    # check request status, output if there was an error
    try:
        r.raise_for_status()
    except Exception as e:
        print('an error occurred downloading the file: %s' % (e))

    # output the request data to a file    
    downloaded_file=open(output_file,'wb')        
    for chunk in r.iter_content(100000):
        downloaded_file.write(chunk)

    return output_file

def modify_url_date(url, new_date, date_parameter="CloseDate"):
    """modify the SSRS URL to replace the data parameter with the new date"""
    # parse URL for query string parts
    o = ulp.urlparse(url)
    query_parts = ulp.parse_qsl(o.query, True)

    # search for date parameter in query string
    for i in range(0,len(query_parts)):
        if query_parts[i][0]==date_parameter:
            query_parts.remove(query_parts[i])

    # reconstruct query string with date removed
    date_string = str(new_date)
    query_string=ulp.urlencode(query_parts, True)

    # The report name will be a query string parameter with empty string value.  remove equal sign.
    query_string = query_string.replace("=&rs","&rs",1) 
    full_path = o.path+"/?"+query_string+"&"+date_parameter+"="+date_string
    
    # get a new URL from parts, adding query string as part of path
    o_new = ulp.ParseResult(o.scheme, o.netloc, full_path 
                            , o.params, '', o.fragment)
    return o_new.geturl()


# The following four lines build a parameter list for the Division report parameter.  Each value needs to be assigned to the Division parameter
#   IE Division=Apparel&Division=Bedding&Division=Consumables
# The thought was to build this dynamically, I could potentially pass in some parameters if need be, and it would dynamically build
# the parameter list and append it to the url that call the SSRS report

Divisions = ['Apparel', 'Bedding', 'Consumables', 'Furniture', 'Hardlines','Home','Misc','Not for retail','Seasonal','Uncategorized']
ParamString = 'Division='

AddParameterPrefix = ['Division=' + sub for sub in Divisions] 
DivisionParameterList ='&'.join(AddParameterPrefix)

#Establish a Connection to the SQL Server and retrieve the # of stores to email a report to.
#This is accomplished by passing a store # to the SSRS report parameter Store when the report is called via URL.

conn = pyodbc.connect('Driver={SQL Server};'+
                      DB_SERVER+
                      DB_NAME+
                      'Trusted_Connection=yes;')

cursor = conn.cursor()
cursor.execute("SELECT Top (2) Venue_ID, Venue FROM TW_List_Venues where Active_YN=1 AND Sales_Channel_ID=2 ORDER BY Venue_ID")

fromaddr = 'mike.schmackle@gmail.com'
DSP      = 'mschmackle@bargainhunt.com' 
toaddr   = 'mschmackle@bargainhunt.com'

for row in cursor:
    url = 'http://' + SSRS_SERVER_NAME + '/ReportServer/Pages/ReportViewer.aspx?%2fRetail%2fInventory+Markdown&rs:Command=Render&Store=' + str(row.Venue_ID) + '&' +DivisionParameterList + '&rs:Format=excel'
    file_name = 'InventoryByMarkDown-' + row.Venue + '.xls'
    output_file = get_url(url, 'mschmackle', 'BargainHunt2019!', file_name)

    # instance of MIMEMultipart 
    msg = MIMEMultipart() 
    
    # storing the senders email address   
    msg['From'] = DSP

    # storing the senders email address   
    msg['ReplyTo'] = DSP
    
    # storing the receivers email address  
    msg['To'] = toaddr  
    
    # storing the subject  
    msg['Subject'] = "Inventory Markdown Report for " + row.Venue
    
    # string to store the body of the mail 
    body = "Please review the attached inventory counts"
    
    # attach the body with the msg instance 
    msg.attach(MIMEText(body, 'plain')) 
    
    # open the file to be sent  

    attachment = open(output_file, "rb") 
    
    # instance of MIMEBase and named as p 
    p = MIMEBase('application', 'octet-stream') 
    
    # To change the payload into encoded form 
    p.set_payload((attachment).read()) 
    
    # encode into base64 
    encoders.encode_base64(p) 
    
    p.add_header('Content-Disposition', "attachment; filename= %s" % file_name) 
    
    # attach the instance 'p' to instance 'msg' 
    msg.attach(p) 
    
    # creates SMTP session 
    s = smtplib.SMTP('smtp.gmail.com', 587) 
    
    # start TLS for security 
    s.starttls() 
    
    # Authentication 
    s.login(fromaddr, "hcjbsqesjzuiejnx") 
    
    # Converts the Multipart msg into a string 
    text = msg.as_string() 
    
    # sending the mail 
    s.sendmail(fromaddr, toaddr, text) 
    
    # terminating the session 
    s.quit() 

    print (url)