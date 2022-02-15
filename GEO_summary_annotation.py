#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#Explanation: This code aims to scrape data from web for a series of provided GEO (Gene expression omnibus) ids followed by
#             storing all the meta data for each id in a relational db created by this code and lastly annotating the 
#             summaries of each id on the fly using "becas api" which annotates text and PubMed abstracts 
#             with biomedical concepts and then find out those GEO ids which are related to disease samples using annotation.


# Modules imported to achieve the above mentioned objective:

import requests
import json
from bs4 import BeautifulSoup
import sqlalchemy
from sqlalchemy import create_engine, MetaData, Table, Column, Integer, String
from sqlalchemy import text


#Module for scraping data from GEO web server and storing.

def fetch_geo(GEO_ids,attr): #scraping data from GEO database
    id=GEO_ids
    attr=attr
    
    url="https://www.ncbi.nlm.nih.gov/geo/query/acc.cgi?acc="+str(id)
    page=requests.get(url).text
    soup = BeautifulSoup(page, 'lxml')
    data_table=soup.find("table",attrs={"cellpadding":"2","cellspacing":"0","width":"600"})
    data=data_table.find_all("td")
    
    attr['ID']=id
    for i in range(len(data)):
        if data[i].text in fields:
            fields[data[i].text]=data[i+1].text
    return(attr)


#Module for annotating Summaries of each GEO id using Becas api.

def becas_api(data): #fetching keywords by using becas api
    data=data
    payload={"text":data['Summary'],"echo":""}
    URL='http://bioinformatics.ua.pt/becas/api/text/annotate?email=<example@gmail.com>&tool=<test>'
    headers={'Content-Type':'application/json','Accept':'text/plain'}
    response=requests.post(URL,headers=headers,data=json.dumps(payload)) #getting annotation of summary from becas api
    
    val=response.status_code
    if val == 200:
        res_dict=json.loads(response.text)
        keyword=[]
        for k in range(len(res_dict['entities'])):
            annotation=res_dict['entities'][k].split("|")
            keyword.append(annotation[0])
        keyword=",".join(keyword)
        data['Annotation']=keyword
        return data
    else:
        print ("ERROR")

#Module for creating a user defined database and table in the same database in local system/server.

def create_db(data): #creating db, table,inserting data and executing select query
    info=data
    engine = create_engine('sqlite:///GSE_annotation.db')   #creating database 
    meta=MetaData()
    
    data_information= Table('data_information', meta,
                       Column('ID',String,primary_key= True),Column('Title',String),
                       Column('Organism',String),Column('Experiment type',String),
                       Column('Summary',String),Column('Overall design',String),
                       Column('Contributor(s)',String),Column('Citation(s)',String),
                       Column('Annotation',String))         #creating table
    
    meta.create_all(engine)
    conn=engine.connect()
    conn.execute(data_information.insert(),info) #inserting data in table
    
    #Fetching those GEO ids which are related to disease samples.
    text_query=text("SELECT ID FROM data_information WHERE Annotation LIKE '%disease%' ") #making select query on table
    result = conn.execute(text_query).fetchall() #executing select query on table.
    return result   


if __name__ == "__main__":

    GEO=['GSE63312','GSE78224','GSE74018','GSE50734','GSE114644','GSE60477','GSE53599','GSE80582','GSE109493','GSE35200']
    fields={"ID":"","Title":"","Organism":"","Experiment type":"","Summary":"","Overall design":"","Contributor(s)":"","Citation(s)":"","Annotation":""}
    
    db_data=[]
    for geo_id in GEO:
        data=fetch_geo(geo_id,fields)
        print (data)
        info=becas_api(data)
        print (info)
        db_data.append(info.copy())
    
  
    output=create_db(db_data) #storing data returned from database in output variable
    
    if len(output) != 0:       #Printing data returned from database
        for row in output:
            row="".join(row)
            print (row)
    else:
        print ("There are no ID's present in database for desired keyword")
    

