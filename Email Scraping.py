##########################################
# Title: Querying Inbox
# Purpose: Read outlook emails with Python
# Author: Thomas Handscomb
##########################################

########################################
# Read outlook emails with Python
########################################

# import modules into session
import pandas as pd
import win32com.client
from tqdm import tqdm # Useful module for displaying a progress bar during long loops

# Define Outlook location
outlook = win32com.client.Dispatch("Outlook.application")
mapi = outlook.GetNamespace("MAPI")

# Find the folder number of the 'Thomas.Handscomb@Schroders.com' meta data folder to start with
for k in range(1, len(mapi.Folders)+1):
    try:
        fol = mapi.Folders.Item(k)
        if fol.name == 'Thomas.Handscomb@Schroders.com':
            folnum = k
            #print(folnum)
            break            
            
    except Exception as e:
        print('Error:' + '(' + str(k) + ')')
        pass

print(folnum)
    
# Once you have determined the above folder number, find the 'Inbox' and 'Sent Items' folders within this
Inboxnum, Sentnum = -1, -1

for l in range(1,30):
    try:
        subfol = mapi.Folders.Item(folnum).Folders.Item(l)
        
        if Inboxnum > 0 and Sentnum > 0:
            break
        
        elif subfol.name =='Inbox':            
                Inboxnum = l
                
        elif subfol.name =='Sent Items':            
                Sentnum = l             
                
    except Exception as e:
        print('Error at loop: %.f' %l)
        pass

print("%0.f, %0.f" %(Inboxnum, Sentnum))

# Once the folder numbers are defined, use these to specify the 'Inbox' and 'Sent' folders
Inbox = mapi.Folders.Item(folnum).Folders.Item(Inboxnum)
Sent = mapi.Folders.Item(folnum).Folders.Item(Sentnum)

# Double check the name
if Inbox.name == 'Inbox' and Sent.name == 'Sent Items':
    print('Inbox and Sent folders assigned correctly')
    pass
else:
    print('An error has occured')

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
## Examine Inbox
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Now that the Inbox and Sent Items folders have been determined,
# create a blank data frame to store email metadata, in this case (date/time sent, sender name, email subject)
Inbox_col_names =  ['Full Date', 'Date', 'Hour', 'Sender', 'Subject']
Inbox_df = pd.DataFrame(columns = Inbox_col_names)
Inbox_df

# Loop through all Inbox.Items (i.e. emails)
# the tqdm wrapper puts a progress bar on the loop
for message in tqdm(Inbox.Items):
    try:
        Inbox_df.loc[len(Inbox_df)] = [message.LastModificationTime.strftime("%Y-%m-%d %H:%M:%S")
        , message.LastModificationTime.strftime("%Y-%m-%d")
        , message.LastModificationTime.strftime("%H")
        , message.Sender
        , message.Subject]
    except:
        pass

# Confirm you are picking up all emails
Inbox_df.groupby(['Date']).size()

# Output data frame to review
Output_filepath = 'C:/Users'

Inbox_df.to_csv(Output_filepath+'/Inbox.csv'
               , encoding = 'utf-8'
               #, mode = 'a'
               , index = False
               , header = True)

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
## Examine my own Sent Items
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Specify Sent Items. This is now defined above
#Sent = mapi.Folders.Item(folnum).Folders.Item(Sentnum)
print(Sent.name)

Outbox_col_names =  ['Full Date', 'Date', 'Hour', 'Recipient', 'Subject']
Outbox_df = pd.DataFrame(columns = Outbox_col_names)
Outbox_df

j=1
m=1

# Double check you are in the correct Sent Folder
i=1
for message in Sent.Items:
    if i<25:
        print(message.LastModificationTime.strftime("%Y-%m-%d %H:%M:%S"))
        print(message.Subject)
        i = i+1
    else:
        #raise
        break

# Build up the Outbox data frame
for message in tqdm(Sent.Items):
    m=+m+1
    try:
        Outbox_df.loc[len(Outbox_df)] = [message.LastModificationTime.strftime("%Y-%m-%d %H:%M:%S")
                , message.LastModificationTime.strftime("%Y-%m-%d")
        , message.LastModificationTime.strftime("%H")
        , message.To
        , message.Subject]
    except Exception as e:
        j=+j+1
        #print('Error:' + str(e))
        pass

print(m)
print(j)

Outbox_df.shape

# Check the dates that are appearing in the Outbox data frame - Convert to a dataframe to sort by values
pd.DataFrame(Outbox_df.groupby(['Date']).size()).sort_values(by = 'Date', ascending = True)

# Another way
Outbox_df[('Date')].unique()

# Output data frame to tbe used in Tableau
Outbox_df.to_csv(Output_filepath+'/Outbox.csv'
               , encoding = 'utf-8'
               #, mode = 'a'
               , index = False
               , header = True)