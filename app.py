#!/usr/bin/env python
# coding: utf-8

# In[28]:


from flask import Flask, render_template, request , send_file
from werkzeug.utils import secure_filename
import pandas as pd
from docxtpl import DocxTemplate
import datetime as dt
import json


# In[32]:


app = Flask(__name__)
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        filename = secure_filename(file.filename)
        file.save("static/" + filename)
        df = pd.read_excel("static/" + filename)
        
        cusList = [x for x in df.Source.unique() if str(x) != 'nan']
        invoiceList = []
        for i in cusList:
            df1 = df.loc[df['Source'] == i]
            df2 = df1[["Student's_Name", 'Subject','Hourly_Rate_MM','Hourly_Rate_parents','Total_Lesson_fee_parents','Total_lesson_fee_MM','Others']]
            df3 = df2[df2["Student's_Name"].notna()]
            df4 = df3.loc[df3['Total_Lesson_fee_parents']!=0]
            df4.columns = ['sName', 'service', 'aRate', 'pRate','pAmt','aAmt','other']

            #define Customer
            cusName = i
            invoice_num = cusName+dt.datetime.now().strftime("%Y%b")

            # create a document object
            doc = DocxTemplate("static/" +"InvoiceTemplate.docx")

            # create context dictionary
            context = {
                "date": dt.datetime.now().strftime("%d-%b-%Y"),
                "bill_to": cusName,
                "invoice_num": invoice_num
            }
            result = df4.to_json(orient="records")
            parsed = json.loads(result) 

            context['content'] = parsed
            context['pTotal'] = df4.pAmt.sum()+df4.other.sum()
            context['aTotal'] = df4.aAmt.sum()
            # render context into the document object
            doc.render(context)
            doc.save(f'static/{invoice_num}.docx')
            invoiceList.append(f'{invoice_num}.docx')
            
        return(render_template("download.html", result=invoiceList))
    else:
        return(render_template("index.html", result="pending"))
    
    
@app.route('/download', methods=['GET', 'POST'])
def download():
    if request.method == 'POST':
        invoiceName = request.form.get("invoice")
        return send_file(f'static/{invoiceName}', as_attachment=True)
        



if __name__ == "__main__":
    app.run()






