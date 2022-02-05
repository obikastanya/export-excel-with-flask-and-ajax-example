import base64
from flask import Flask
from flask import render_template
import json
import xlsxwriter
from io import BytesIO

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/download_excel_api')
def downloadExcelApi():
    apiResponse = createApiResponse()
    return apiResponse


def createApiResponse():
    bufferFile = writeBufferExcelFile()
    stringFile = parseBufferToString(bufferFile)
    response = {'status': True, 'data': [stringFile], 'message': 'Data Is Success'}
    return response


def writeBufferExcelFile():
    buffer = BytesIO()
    workbook = xlsxwriter.Workbook(buffer)
    worksheet = workbook.add_worksheet()
    jsonProductData= loadData()

    dataHeader=["SKU", "Product","Type","Price","UPC","Shipping","Description", "Manufacturer","Model"]
    headerStyle=workbook.add_format(createHeadStyle())
    worksheet.write_row(0,0,dataHeader,headerStyle)

    for rowIndex,product in enumerate(jsonProductData):
        productValues=list(product.values())
        format=workbook.add_format(createDataStyle())
        if rowIndex%2==1:
            format=workbook.add_format(createDataStyle('#e2efd9'))
        dataToBeWritten=restuctDataBeforeWritten(productValues)
        worksheet.write_row(rowIndex+1, 0,dataToBeWritten, format)
    worksheet.set_column(1, 8, 27)

    workbook.close()
    buffer.seek(0)
    return buffer


def parseBufferToString(buffer):
    binaryFile = buffer.read()
    unicodeBase64File = base64.b64encode(binaryFile).decode('UTF-8')
    return unicodeBase64File


def loadData():
    productData=open('data.json')
    jsonProductData=json.load(productData)
    return jsonProductData

def restuctDataBeforeWritten(productValues):
    dataToBeWritten=[
        productValues[0],
        productValues[1],
        productValues[2], 
        productValues[3], 
        productValues[4],
        productValues[6],
        productValues[7],
        productValues[8],
        productValues[9]
    ]
    return dataToBeWritten

def createDataStyle(bgColor='#FFFFFF'):
    dataStyle={
        'border': 1,
        'fg_color': bgColor
    }
    return dataStyle

def createHeadStyle():
    headStyle={
       'border': 1,
       'font_size':'12',
       'bold':True,
        'fg_color': '#00b050'
    }
    return headStyle
     
        
if __name__ == '__main__':
    app.run(debug=True)