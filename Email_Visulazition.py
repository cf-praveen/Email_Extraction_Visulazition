import imaplib
import os
import email
import pyzmail
from openpyxl import Workbook
from email.utils import getaddresses
from email.header import decode_header
import win32com.client
import sys
import numpy as np
from pandas import Series, DataFrame
import pandas as pd
import matplotlib
import plotly.express as px
import matplotlib.pyplot as plt
from collections import Counter
from collections import OrderedDict
import datetime
import plotly.io as pio

email_user = input("Please enter the company Name: ")


if email_user == 'Emvia':
    email_user = 'submissions@emvia.com'
    email_pass = "Emvia@2020"
    mail = imaplib.IMAP4_SSL("pop.ionos.com",993)
    mail.login(email_user, email_pass)
    mail.select()
    print('Login to Emvia succesful')
elif email_user == 'Skolix':
    email_user = 'submissions@emvia.com'
    email_pass = "Emvia@2020"
    print('Login to Skolix succesful')
elif email_user == 'Praveen':
    email_user = 'submissions@emvia.com'
    email_pass = "Emvia@2020"
    print('Login to Praveen succesfull')
else:
     print('Wrong!, Please check the company Name')
     
def loginm():
    #email_user = "submissions@emvia.com"
    #email_pass = "Emvia@2020"
    #email_user = input("Email ID :")
    #email_pass = input("Password :")
    mail = imaplib.IMAP4_SSL("pop.ionos.com",993)
    mail.login(email_user, email_pass)
    mail.select()   
    print("Login successful....")
    return mail
    

def email_attributes(message):
    #print("Extracting email attributes....")
    from_addresses = []
    to_address = []
    subjects = []
    dates = []
    days = []
    #cc_recipents = []
   # bcc_recipents = []
    months = []
    years = []
    times = []
    Body = []
    Seen_Unseen =[]
    sent_received = []
    unsub_links = []
    category = "None" #categorize_emails(imapobj, UIDs)
    email_data = []
    
      
    if message.get_address('from')[1] == email_user:
        sent_received.append("sent")
    else:
        sent_received.append("Received")
    subject = message.get_subject("Subjects")
    #d1 = [line for line in message]
    msg = email.message_from_string(str(message), policy=email.policy.default)
    body = msg.get_body(('plain',))
    
    if body:
        body = body.get_content()
        #print("Sample Body :",body) 
    #to_rcpt = msg.get_all('to', [])
    #print(to_rcpt)
    cc_rcpt = msg.get_all('cc', [])
    #print("CC Recipent:",cc_rcpt)
    bcc_rcpt = msg.get_all('bcc', [])
    #print("bcc Recepient:",bcc_rcpt)
        
   
    
   # msgseen = mail.store(num,'+FLAGS', '\\Seen')
    #print(msgseen)
    
    #if mail.store(num,'+FLAGS', '\\Seen') == "Seen":
       # msgseen = "yes"
        #print("yes")
    #else:
       # msgseen="No"
        #print("No")
   
    
   # print("Body",msg.get("Body"))

    #msg=email.message_from_string(message[0][1])
    #while msg.is_multipart():
        #msg = msg.get_payload(0)
        #content = msg.get_payload(decode=True)
        #print(content)
    
    #print(subject.split())
    full_date = message.get_decoded_header('date')
    from_addresses.append (message.get_address('from'))
    to_address.append(message.get_addresses('To'))
    
    
    #msg1 =  pyzmail.PyzMessage.factory(message[b'BODY[]'])
    #print(message)
    #print(full_date)
    day = full_date.split()[0].strip(',')
    date = full_date.split()[1]
    month = full_date.split()[2]
    year = full_date.split()[3]
    time2 = full_date.split()[4]
    bcc_recipents = bcc_rcpt
    cc_recipents= cc_rcpt
    subject = subject
    #Seen_Unseen = msgseen
    #Answer = msganswer
    #print(Seen_Unseen)
    Body.append(body)
    days.append(day)
    dates.append(date)
    months.append(month)
    years.append(year)
    times.append(time2)
    subjects.append(subject)
    #Seen_Unseen.append(msgseen)
    #print('day:',day,'date:',date,'month:',month,'year:',year,'time2:',time2)
    
    email_data.extend([dates, months, years, days, times,Body, from_addresses, subjects,sent_received, category, unsub_links,to_address,cc_recipents,bcc_recipents])
    #print('email Data..',email_data)
    return email_data
    
#If you want to select folders you can select folders here
for i in mail.list()[1]:
    #print(i)
    data1 = i.decode().split(' "/" ')[1]
    print (data1)
    #print(data1[0].split('\\'))
    
import csv
mail=loginm()
email_data = []
header_added = True
for i in mail.list()[1]:
    data1 = i.decode().split(' "/" ')[1]
    print(data1)
    mail.select(data1)
    typ, data = mail.search(None, 'ALL')
    mail_ids = data[0]
    #print('total ids',len(mail_ids))
    id_list = mail_ids.split()
    for num in id_list:
        typ, data = mail.fetch(num,'(RFC822)')
        raw_email = data[0][1]
        message = pyzmail.PyzMessage.factory(raw_email)
        email_data.append(email_attributes(message))
        #mail.close()
    #msgseen= mail.store(num,'+FLAGS', '\\Seen')
    #if msgseen =="\\Seen":
        #msgseen = "yes"
        #print("Seen:yes")
            #print("yes")
    #else:
        #msgseen="No"
        #print("Seen:No")
        
    #msganswer = mail.store(num,'+FLAGS', '\\Answered')  
    #if msganswer =='\\Answered':
        #msganswer ="yes"
    #else:
       # msganswer ="No"
        
    #for num in :
        #mail.store(num,'+FLAGS','\\DELETED')
    
    #print("No")
        cwd = os.getcwd()
        #print(cwd)
        # data =email_data[1:1000]
    print(len(email_data))
    
    with open('sample.csv','w',newline='',encoding='utf-8') as f:
        w = csv.writer(f)
                   # dates, months, years, days, times,Body, from_addresses, subjects, sent_received,Seen_Unseen, category, unsub_links,to_address
        w.writerow(['Date', 'Month','Year', 'Day', 'Time','Body', 'From(Sender)',  'Subject','From(Email ID)','Sent/Received','Category',  'To_Address','cc_recipents','bcc_recipents'])
        w.writerows(email_data)
        #print("extraction from email is over you can proceed the next step")
        #time.sleep(60)

data = pd.read_csv("C:\\Users\\mypc\\Desktop\\Email Submission Code\\sample.csv",encoding="ISO-8859-1")

#removing comas,quotes, slash and \n from thje column 
data['Date'] = data['Date'].str.replace(r"[(\"\',)]", '')
data['Date'] = data['Date'].replace(r'\\n',' ', regex=True)
data['Date'] = data['Date'].str.strip('[]').astype(int)

#removing comas,quotes, slash and \n from thje column 
data['Month'] = data['Month'].str.replace(r"[(\"\',)]", '')
data['Month'] = data['Month'].replace(r'\\n',' ', regex=True)
data['Month'] = data['Month'].str.strip('[]')

#removing comas,quotes, slash and \n from thje column 
data['Year'] = data['Year'].str.replace(r"[(\"\',)]", '')
data['Year'] = data['Year'].replace(r'\\n',' ', regex=True)
data['Year'] = data['Year'].str.strip('[]').astype(int)

#removing comas,quotes, slash and \n from thje column 
data['Time'] = data['Time'].str.replace(r"[(\"\',)]", '')
data['Time'] = data['Time'].replace(r'\\n',' ', regex=True)
data['Time'] = data['Time'].str.strip('[]')

#removing comas,quotes, slash and \n from thje column 
data['Body'] = data['Body'].str.replace(r"[(\"\',)]", '')
data['Body'] = data['Body'].replace(r'\\n','', regex=True)
data['Body'] = data['Body'].str.strip('[]')
data['Body'] = data['Body'].str.replace("None", " ")
#remove website link
data['Body']  = data['Body'] .str.replace('http\S+|www.\S+', '', case=False)
# removing everything except alphabets
data['Body'] =data['Body'] .str.replace("[^a-zA-Z#]", " ")

#removing comas,quotes, slash and \n from thje column 
data['From(Sender)'] = data['From(Sender)'].str.replace(r"[(\"\',)]", '')
data['From(Sender)'] = data['From(Sender)'].replace(r'\\n','', regex=True)
data['From(Sender)'] = data['From(Sender)'].str.strip('[]')

#change Column name 
data=data.rename(columns = {'From(Sender)':'From_Sender'})

#removing comas,quotes, slash and \n from thje column 
#data['Subject'] = data['Subject'].str.replace(r"[(\"\',)]", '')
data['Subject'] = data['Subject'].replace(r'\\n','', regex=True)
data['Subject'] = data['Subject'].str.strip('[]')
# removing everything except alphabets
#data['Subject'] =data['Subject'] .str.replace("[^a-zA-Z#]", " ")


#removing comas,quotes, slash and \n from thje column 
data['From(Email ID)'] = data['From(Email ID)'].str.replace(r"[(\"\',)]", '')
data['From(Email ID)'] = data['From(Email ID)'].replace(r'\\n','', regex=True)
data['From(Email ID)'] = data['From(Email ID)'].str.strip('[]')

#removing comas,quotes, slash and \n from thje column 
data['Day'] = data['Day'].str.replace(r"[(\"\',)]", '')
data['Day'] = data['Day'].replace(r'\\n','', regex=True)
data['Day'] = data['Day'].str.strip('[]')

#removing comas,quotes, slash and \n from thje column 
data['To_Address'] = data['To_Address'].str.replace(r"[(\"\',)]", '')
data['To_Address'] = data['To_Address'].replace(r'\\n','', regex=True)
data['To_Address'] = data['To_Address'].str.strip('[]')

#removing comas,quotes, slash and \n from thje column 
data['cc_recipents'] = data['cc_recipents'].str.replace(r"[(\"\',)]", '')
data['cc_recipents'] = data['cc_recipents'].replace(r'\\n','', regex=True)
data['cc_recipents'] = data['cc_recipents'].str.strip('[]')

#removing comas,quotes, slash and \n from thje column 
data['bcc_recipents'] = data['bcc_recipents'].str.replace(r"[(\"\',)]", '')
data['bcc_recipents'] = data['bcc_recipents'].replace(r'\\n','', regex=True)
data['bcc_recipents'] = data['bcc_recipents'].str.strip('[]')

#removing comas,quotes, slash and \n from thje column 
data['Category'] = data['Category'].str.replace(r"[(\"\',)]", '')
data['Category'] = data['Category'].replace(r'\\n','', regex=True)
data['Category'] = data['Category'].str.strip('[]')

#change character to numeric
data['Monthnumeric'] = pd.Categorical(data.Month)
data['Monthnumeric'] = data.Monthnumeric.cat.codes

#change character to numeric
data['Daynumeric'] = pd.Categorical(data.Day)
data['Daynumeric'] = data.Daynumeric.cat.codes

#change character to numeric
data['FromSendernumeric'] = pd.Categorical(data.From_Sender)
data['FromSendernumeric'] = data.FromSendernumeric.cat.codes

#change Column name 
data=data.rename(columns = {'Day':'Days'})

#change Column name 
data=data.rename(columns = {'Monthtonumber':'Month'})

#change Column name 
data=data.rename(columns = {'Date':'Day'})

look_up = {'Jan': 1,'Feb': 2,'Mar': 3,'Apr':4,'May':5,'Jun':6,'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}

data['Month'] = data['Month'].apply(lambda x: look_up[x])

data['FullDate'] = pd.to_datetime(data[['Year', 'Month', 'Day']])

look_up = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}

data['Monthchar'] = data['Month'].apply(lambda x: look_up[x])

# Weekly traffic
# sort by day of the week
data['Days'] = pd.Categorical(data['Days'], categories= ['Mon','Tue','Wed','Thu','Fri','Sat', 'Sun'],ordered=True)
count_sorted_by_day = data['Days'].value_counts().sort_index()

plt.figure(1)
count_sorted_by_day.plot(marker = 'o', color = 'blueviolet', linewidth = 2, ylim = [0,500])
plt.title('Weekly Email Traffic', fontweight = 'bold' ,fontsize = 14)
plt.ylabel("Received Email Count", fontweight = 'bold', labelpad = 15)
plt.grid()

# Hourly traffic
# splitting only the hour portion in the time column
# sort by hour of the day - using sort_index for numeric sort
received = data[data['Sent/Received'] == 'Received']
hour = data['Time'].str.split(':').str[0] + ':00'
count_sorted_by_hour = hour.value_counts().sort_index()

plt.figure()
fig = plt.figure()
count_sorted_by_hour.plot(marker = 'o', color = 'green')
plt.title('Hourly Email Traffic', fontsize = 14, fontweight = 'bold')
plt.ylabel("Received Email Count", fontweight = 'bold', labelpad = 15)
plt.xlabel("Hour of the Day", fontweight = 'bold', labelpad = 15)
plt.xticks(range(len(count_sorted_by_hour.index)), count_sorted_by_hour.index)
plt.xticks(rotation=90)
plt.grid()
#plt.write_html(fig , file='Hourly_Email_Traffic.html', auto_open=False)
#plt.savefig('Hourly_Email_Traffic.html')

sender_top_20 =  data['From_Sender'].value_counts().nlargest(20)
sender_top_20_count = sender_top_20.values
sender_top_20_names = sender_top_20.index.tolist()

plt.figure()
#fig = plt.figure()
plt.barh(sender_top_20_names, sender_top_20_count, color = 'forestgreen', ec = 'black', linewidth = 1.0)
plt.gca().invert_yaxis()
plt.title('Top 20 Senders', fontsize = 14 ,fontweight = 'bold')
plt.xlabel('Received Email Count', fontweight = 'bold')

for i, v in enumerate(sender_top_20_count):
    plt.text(v + 3, i + .25, str(v), color='blue', fontweight='bold')
plt.figure(figsize=(20,10))
plt.tight_layout()
#pio.write_html(fig, file='Top_From_senders.html', auto_open=False)

# count the number of words in the subject
data['Subject Word Count'] = data['Subject'].str.split(' ').str.len()

#plt.figure(figsize=(20,10))
plt.hist(data['Subject Word Count'], bins=15, color = 'slategray', ec = 'black')
plt.axis()
plt.xlabel('Word Count', fontweight = 'bold')
#plt.figure()
plt.ylabel('No. of Emails', fontweight = 'bold')
plt.title('Subject Word Count Histogram', fontsize = 14, fontweight = 'bold')

# count the number of words in the subject
data['Body Word Count'] = data['Body'].str.split(' ').str.len()

plt.figure()
plt.hist(data['Body Word Count'], bins=15, color = 'slategray', ec = 'black')
plt.axis()
plt.xlabel('Word Count', fontweight = 'bold')
plt.ylabel('No. of Emails', fontweight = 'bold')
plt.title('Body Word Count Histogram', fontsize = 14, fontweight = 'bold')

# split the subject line into words and store them as a list
word_list_2d = data['Subject'].str.split(' ').fillna('none').tolist()
word_list_1d = [word for list in word_list_2d for word in list]

# treat all words as lower case
word_list_1d = [word.lower() for word in word_list_1d]

# exclude common words and words with three or lesser letters
exclude_list = ['this', 'that', 'your', 'with', 'from']
word_list_1d = [word for word in word_list_1d if word not in exclude_list and len(word)>4]

# extract common words in subject lines and their frequencies of occurrence
common_words_map = Counter(word_list_1d).most_common(10)
common_words = [pair[0] for pair in common_words_map]
frequency = [pair[1] for pair in common_words_map]

plt.figure()
fig= plt.figure()
plt.barh(common_words, frequency, color = 'lightcoral', ec = 'black', linewidth = 1.25)
plt.gca().invert_yaxis()
plt.title('Most Common Words in Subjects', fontsize = 14 ,fontweight = 'bold')
y = 0.15
for i in range(len(frequency)):
    if len(str(frequency[i])) == 3:
        x = frequency[i] - 14
    else:
        x = frequency[i] - 10
    plt.text(x,y,frequency[i], fontsize = 13,fontweight = 'bold')
    y = y + 1
plt.xticks([0,200])
plt.xlabel('Occurrences', fontweight = 'bold', labelpad=-5)
#pio.write_html(fig, file='Most_frequent_words_in_Subject.html', auto_open=False)

# split the subject line into words and store them as a list
word_list_2d = data['Body'].str.split(' ').fillna('none').tolist()
word_list_1d = [word for list in word_list_2d for word in list]

# treat all words as lower case
word_list_1d = [word.lower() for word in word_list_1d]

# exclude common words and words with three or lesser letters
exclude_list = ['this', 'that', 'your', 'with', 'from']
word_list_1d = [word for word in word_list_1d if word not in exclude_list and len(word)>3]

# extract common words in subject lines and their frequencies of occurrence
common_words_map = Counter(word_list_1d).most_common(10)
common_words = [pair[0] for pair in common_words_map]
frequency = [pair[1] for pair in common_words_map]

plt.figure()
fig= plt.figure()
plt.barh(common_words, frequency, color = 'lightcoral', ec = 'black', linewidth = 1.25)
plt.gca().invert_yaxis()
plt.title('Most Common Words in Body', fontsize = 14 ,fontweight = 'bold')
y = 0.15
for i in range(len(frequency)):
    if len(str(frequency[i])) == 3:
        x = frequency[i] - 14
    else:
        x = frequency[i] - 10
    plt.text(x,y,frequency[i], fontsize = 10,fontweight = 'bold')
    y = y + 1
plt.xticks([0,200])
plt.xlabel('Occurrences', fontweight = 'bold', labelpad=-5)
#pio.write_html(fig, file='Most_frequent_words_in_body.html', auto_open=False)

import nltk
nltk.download('punkt')
#lines = data['Subject'].apply(word_tokenize)

import nltk
nltk.download('averaged_perceptron_tagger')

import textblob
from textblob import TextBlob
def get_adjectives(text):
    blob = TextBlob(text)
    return [ word for (word,tag) in blob.tags if  tag=='VB' or tag == 'NN']

data['Pos_Tag'] = data['Subject'].apply(get_adjectives)

data['Pos_Tag'] = data['Pos_Tag'].apply(lambda x: ' '.join([w for w in x]))

import plotly
import plotly.graph_objs as go
from plotly.offline import init_notebook_mode, iplot

def get_adjectives(text):
    blob = TextBlob(text)
    return [ word for (word,tag) in blob.tags if  tag=='VB' or tag == 'NN']

data['Body_Pos_Tag'] = data['Body'].apply(get_adjectives)

data['Body_Pos_Tag'] = data['Body_Pos_Tag'].apply(lambda x: ' '.join([w for w in x]))

#pip install chart-studio

import chart_studio.plotly as py
import plotly.graph_objs as go
import cufflinks as cf
cf.go_offline()
cf.set_config_file(offline=False, world_readable=True)


fig = px.bar(data, x="From_Sender", y="Monthnumeric", color='Monthchar', barmode='group',height=400)
fig.show(figsize=(20,10))
pio.write_html(fig, file='From_Email_tracking_by_months.html', auto_open=False)

#From sender based on Day
fig = px.bar(data, x="From_Sender", y="Daynumeric", color='Days', barmode='group',height=400)
fig.show()
pio.write_html(fig, file='From_Email_tracking_by_days.html', auto_open=False)

#To sender based on Day
fig = px.bar(data, x="To_Address", y="Daynumeric", color='Days', barmode='group',height=400)
fig.show()
pio.write_html(fig, file='To_Address_tracking_by_days.html', auto_open=False)

#To sender based on Month
fig = px.bar(data, x="To_Address", y="Monthnumeric", color='Monthchar', barmode='group',height=800)
fig.show()
pio.write_html(fig, file='To_Address_tracking_by_months_and_days.html', auto_open=False)

df = data[["Days","Subject","Monthchar","Daynumeric","Monthnumeric"]]


import nltk
nltk.download('stopwords')
from nltk.corpus import stopwords
stop_words = stopwords.words('english')

# function to remove stopwords
def remove_stopwords(rev):
    rev_new = " ".join([i for i in rev if i not in stop_words])
    return rev_new
exclude_list = ['this', 'that', 'your', 'with', 'from']

# remove short words (length < 3)
df['Subject'] = df['Subject'].apply(lambda x: ' '.join([w for w in x.split() if w not in exclude_list and len(w)>3]))

# remove stopwords from the text
reviews = [remove_stopwords(r.split()) for r in df['Subject']]
#print(reviews)

# make entire text lowercase
reviews = [r.lower() for r in reviews]
#print(reviews)

reviews_3 = []
for i in range(len(reviews)):
    reviews_3.append(''.join(reviews[i]))

df['Subject_content'] = reviews_3
df['Subject_content'] = df['Subject_content'].str.replace(r"[(\"\',)]", '')
df['Subject_content'] =df['Subject_content'] .str.replace("[^a-zA-Z#]", " ")

fig = px.bar(df, x="Subject_content", y="Daynumeric", color='Days', barmode='group',height=800)
fig.show()

fig = px.bar(df, x="Subject_content", y="Monthnumeric", color='Monthchar', barmode='group', height=800)
fig.show()

new_df = data[["Days","Body","Monthchar","Daynumeric","Monthnumeric"]]

from collections import OrderedDict
# function to remove stopwords
def remove_stopwords(rev):
    rev_new = " ".join([i for i in rev if i not in stop_words])
    return rev_new


# make entire text lowercase
new_df['Body'] = [r.lower() for r in new_df['Body']]
#print(new_df['Body'] )

exclude_list = ['this', 'that', 'your', 'with', 'from']

# remove short words (length < 3)
new_df['Body'] = new_df['Body'].apply(lambda x: ' '.join([w for w in x.split() if w not in exclude_list and len(w)>3]))

#remove repeated words
new_df['Body']= (new_df['Body'].str.split().apply(lambda x: OrderedDict.fromkeys(x).keys()).str.join(' '))




# remove stopwords from the text
reviews = [remove_stopwords(r.split()) for r in new_df['Body']]
print(len(reviews))

import spacy
nlp = spacy.load('en_core_web_sm', disable=['parser', 'ner'])

def lemmatization(texts, tags=['NOUN', 'ADJ']):
    # filter noun and adjective
    output = []
    for sent in texts:
        doc = nlp(" ".join(sent)) 
        output.append([token.lemma_ for token in doc if token.pos_ in tags])
    return output
 
tokenized_reviews = pd.Series(reviews).apply(lambda x: x.split())
print(len(tokenized_reviews))

reviews_2 = lemmatization(tokenized_reviews)
print(len(reviews_2) )# print lemmatized review

reviews_3 = []
for i in range(len(reviews_2)):
    reviews_3.append(' '.join(reviews_2[i]))
    #print(reviews_3)
    #print(reviews_2)

new_df['noduplicates_Body'] = reviews_3

fig = px.bar(new_df, x="Body", y="Daynumeric", color='Days', barmode='group',height=800)
fig.show()

fig = go.Figure(data=[
    go.Bar(name='Day', x=data['From_Sender'], y=data['Daynumeric']),
    go.Bar(name='Month', x=data['From_Sender'], y=data['Monthnumeric'])])
# Change the bar mode
fig.update_layout(barmode='group')
fig.show()

fig = go.Figure(data=[
    go.Bar(name='Day', x=data['To_Address'], y=data['Daynumeric']),
    go.Bar(name='Month', x=data['To_Address'], y=data['Monthnumeric'])])
# Change the bar mode
fig.update_layout(barmode='group')
fig.update_layout(showlegend=True)
fig.show()

df1 = data.groupby(["FullDate","From_Sender", "Days","cc_recipents"])["Day"].count().reset_index(name="count")

import plotly.graph_objs as go
fig = px.bar(df1, x='FullDate', y='count',title="Eamil_tracking_by_months_and_days",color='Days')
fig.update_xaxes(
    rangeslider_visible=True,
    rangeselector=dict(
        buttons=list([
            dict(count=1, label="1m", step="month", stepmode="backward"),
            dict(count=6, label="6m", step="month", stepmode="backward"),
            dict(count=1, label="YTD", step="year", stepmode="todate"),
            dict(count=1, label="1y", step="year", stepmode="backward"),
            dict(step="all")])))
fig.show()
pio.write_html(fig, file='Email_tracking_by_days.html', auto_open=False)

import plotly.graph_objs as go
fig = px.bar(df1, x='FullDate', y='count',title="From_Sender_tracking_by_months_and_days",color='From_Sender')
fig.update_layout(
    autosize=True,
    width=1100,
    height=500,)
fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')
fig.update_xaxes(
    rangeslider_visible=True,
    rangeselector=dict(
        buttons=list([
            dict(count=1, label="1day", step="day", stepmode="backward"),
            dict(count=1, label="1m", step="month", stepmode="backward"),
            dict(count=6, label="6m", step="month", stepmode="backward"),
            dict(count=1, label="YTD", step="year", stepmode="todate"),
            dict(count=1, label="1y", step="year", stepmode="backward"),
            dict(step="all")])))
fig.show()
pio.write_html(fig, file='From_Sender_tracking_by_months_and_days.html', auto_open=False)

df_cc_recipents = data.groupby(["FullDate", "Days","cc_recipents","Monthchar","From_Sender"])["Days"].count().reset_index(name="count")


import re
def extract_email_ID(string):
    email = re.findall(r'<(.+?)>',string)
    if not  email:
        email = list(filter(lambda y: '@' in y ,string.split()))
    return email[0] if email else False

df_cc_recipents['cc_recipents'] = df_cc_recipents['cc_recipents'].apply(lambda x:extract_email_ID(x))

df_cc_recipents['From_Sender']=df_cc_recipents['From_Sender'].apply(lambda x:extract_email_ID(x))

import os
import plotly.graph_objs as go
fig = px.bar(df_cc_recipents, x='FullDate', y='count',title="CC_Adress_tracking_by_months_and_days",color='From_Sender',text ='cc_recipents')
fig.update_layout(
    autosize=True,
    width=1100,
    height=500,)
fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')
fig.update_xaxes(
    rangeslider_visible=True,
    rangeselector=dict(
        buttons=list([
            dict(count=1, label="1day", step="day", stepmode="backward"),
            dict(count=1, label="1m", step="month", stepmode="backward"),
            dict(count=6, label="6m", step="month", stepmode="backward"),
            dict(count=1, label="YTD", step="year", stepmode="todate"),
            dict(count=1, label="1y", step="year", stepmode="backward"),
            dict(step="all")])))
fig.show()
cwd = os.getcwd()
print(cwd)
pio.write_html(fig, file='CC_Adress_tracking_by_months_and_days.html', auto_open=False)
#False Represent No Email_Id