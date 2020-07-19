"""
Birthday Wisher
"""

# pandas package is used for accessing the data from the excel file and writing back to it
import pandas as pd
import datetime

# smtplib package is used for sending email 
import smtplib

# Enter your credentials for login
GMAIL_ID = ""
GMAIL_PSWD = ""

# Function to send emails.
def sendemail(to, sub, msg):
    print(f"It's for {to} and subject is {sub} and message is {msg}")
    # Here we used 578 port number for tls connection
    s= smtplib.SMTP('smtp.gmail.com',578)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PSWD)
    s.sendmail(GMAIL_ID, to, f"Subject: {sub} \n\n{msg}")
    s.quit()
    

if __name__ == '__main__':
    # Fetching the data from the excel sheet
    df = pd.read_excel("data.xlsx")
    # Getting today's date in format: 21-08(day-month) and year
    today = datetime.datetime.now().strftime("%d-%m")
    year =  datetime.datetime.now().strftime("%Y")

    # Setting the flag false so that if there is no bday today then we dont need to write back to excel sheet
    flag = False

    # iterrows is an pandas package which gives us two things (index and item of each row)
    for index, item in df.iterrows():
        bday = item['Birthday'].strftime("%d-%m")
        
        # Matching today's date and date in the excel sheet
        # Also check the year..Year parameter is updated once we have sended the email.
        if (today == bday and year == str(item['Year'])):   
            # If all the conditions matches then change the year of dataframe to next year
            # so that we won't keep sending the email 4 or 5 times that day
            df.loc[index, 'Year'] = str(int(year) + 1)
            sendemail(item['Email'],'Wishes to you', item['Dialogue'])
            # Since today is someone's birthday in your list. so update the flag
            # so that we can update the changes to excel sheet
            flag = True
    
    if flag:
        # Here we write back to excel sheet and keeping the index 0
        df.to_excel('data.xlsx', index=False)

"""
Scheduler:
    To schedule the program execution daily at 12 pm.
    Steps:
        1) Search task scheduler in start bar
        2) Go to task library and create a new task
        3) Give name to the task and add description
        4) Set the trigger to execute daily and time as 00:00:00
        5) Set action :
            i) Program execution: set it to the path of python executor i.e., python.exe
            ii) And add file in double quotes in below line "path_of_file/file_name.py"
        6) Save the file
     
"""