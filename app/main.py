__import__('pysqlite3')
import sys
sys.modules['sqlite3'] = sys.modules.pop('pysqlite3')
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv
from controllers import chatbot_router, metalprice_router
from controllers import excel_dashboard_router
import os
from starlette.responses import FileResponse
import json
import pandas as pd
import xlwings as xw
import re

app = FastAPI()

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Or specify your frontend URL
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

load_dotenv()

# Load OpenAI API Key
os.environ['OPENAI_API_KEY'] = os.getenv("OPENAI_API_KEY")


# Include routers
app.include_router(chatbot_router, prefix="/chatbot", tags=["Chatbot"])
app.include_router(metalprice_router, prefix="/metalprice", tags=["MetalPrice"])
app.include_router(excel_dashboard_router, prefix="/excel-dashboard", tags=["ExcelDashboard"])

# EXCEL FILE

# # Helper Functions to clean and format date-time strings
# def clean_datetime_format(datetime_str):
#     return re.sub(r'\.\d+', '', datetime_str)

# def remove_time_from_date(datetime_str):
#     return datetime_str.split('T')[0]

# def remove_space(datetime_str):
#     return datetime_str.split(' ')[0]

# # Function to create the Excel report
# def create_excel_report(data):
#     df = pd.DataFrame(data)
    
#     # Remove time components from date columns
#     df['rfqStartDate'] = df['rfqStartDate'].apply(remove_time_from_date).apply(remove_space)
#     df['rfqDeadline'] = df['rfqDeadline'].apply(clean_datetime_format)
#     df['rfqConfigureDate'] = df['rfqConfigureDate'].apply(clean_datetime_format)
    
#     # Remove unused columns
#     df = df.drop(['rfqParameterID', 'role', 'rqDescription', 'rfqEndDate', 'rfqStatus', 'rfqLastInvoicePrice', 'rfqTargetPrice'], axis=1)
    
#     # Rename columns
#     df.rename(columns={
#         'rfqid': 'Event ID', 
#         'rfqSubject': 'RFQ Subject',
#         'rfqConfiguredBy': 'Configured By',
#         'rfqStartDate': 'Start Date',
#         'rfqConfigureDate': 'Configure Date',
#         'rfqDeadline': 'RFQ Deadline',
#         'rfqItemCode': 'Item Code',
#         'rfqShortName': 'Short Name',
#         'quantity': 'Quantity',
#         'uom': 'UOM',
#         'currencyName': 'Currency',
#         'vendorName': 'Vendor',
#         'rfqValueAsLastInvoicePrice': 'Last Invoice Price (LIP)',
#         'rfqValueAsTargetPrice': 'Target Price Value',
#         'invitedVendors': 'Invited Vendors',
#         'participatedVendors': 'Participated Vendors'
#     }, inplace=True)

#     df['Start Date'] = pd.to_datetime(df['Start Date'], errors='coerce').dt.date
#     df['RFQ Deadline'] = pd.to_datetime(df['RFQ Deadline'], errors='coerce').dt.date
#     df['Configure Date'] = pd.to_datetime(df['Configure Date'], errors='coerce').dt.date

#     # Create a new Excel workbook using xlwings
#     app = xw.App(visible=False)
#     wb = app.books.add()
#     sheet = wb.sheets[0]
#     sheet.name = 'RFQ_Summary'

#     sheet.range('A1').options(index=False, header=True).value = df

#     table_range = sheet.range('A1').expand()
#     table = sheet.tables.add(source=table_range, name='RFQTable')
#     table.table_style = 'TableStyleMedium9'

#     file_path = 'RFQ_Summary_21.xlsm'
#     wb.save(file_path)
#     wb.close()
#     app.quit()

#     return file_path

# @app.get("/generate-excel")
# async def generate_excel_report():
#     # Load data from the JSON file or request (replace with actual input)
#     with open('data.json', 'r') as file:
#         data = json.load(file)

#     # Call the function to create the Excel report
#     file_path = create_excel_report(data)

#     # Path to the existing workbook
#     vba_code = """ 
#     Sub CreatePivotTableAndChart()
#         Dim wsData As Worksheet
#         Dim wsPivot As Worksheet
#         Dim ptCache As PivotCache
#         Dim pt1 As PivotTable
#         Dim pt2 As PivotTable
#         Dim chartObj1 As ChartObject
#         Dim chartObj2 As ChartObject
#         Dim dataRange As Range

#         On Error GoTo ErrorHandler

#         ' Set the worksheets
#         Set wsData = ThisWorkbook.Sheets("RFQ_Summary")
#         Set wsPivot = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
#         wsPivot.Name = "PivotTable"
        
#         ' Define the data range
#         Set dataRange = wsData.Range("A1").CurrentRegion

#         ' Create the Pivot Cache
#         Set ptCache = ThisWorkbook.PivotCaches.Create( _
#             SourceType:=xlDatabase, _
#             SourceData:=dataRange)

#         ' Create the first Pivot Table
#         Set pt1 = wsPivot.PivotTables.Add( _
#             PivotCache:=ptCache, _
#             TableDestination:=wsPivot.Range("A3"), _
#             TableName:="PivotTable1")

#         ' Configure the first Pivot Table
#         With pt1
#             ' Add Start Date as Row Field
#             With .PivotFields("Start Date")
#                 .Orientation = xlRowField
#                 .Position = 1
#             End With
            
#             ' Add Count of Start Date as Data Field
#             With .PivotFields("Start Date")
#                 .Orientation = xlDataField
#                 .Function = xlCount
#                 .Position = 1
#                 .NumberFormat = "0" ' Optional: Format the count as a number without decimal places
#                 .Caption = "Events"
#             End With
#         End With

#         ' Create a Line Chart from the first Pivot Table
#         Set chartObj = wsPivot.ChartObjects.Add( _
#         Left:=wsPivot.Cells(9, 10).Left, _
#         Width:=wsPivot.Range("J9:Q9").Width, _
#         Top:=wsPivot.Cells(9, 10).Top, _
#         Height:=wsPivot.Range("J9:J24").Height)
#         With chartObj.Chart
#             .SetSourceData pt1.TableRange2
#             .ChartType = xlArea
#             .HasTitle = True
#             .ChartTitle.Text = "Total Events"
#         End With

#         ' Create the second Pivot Table next to the first one (starting at column G)
#         Set pt2 = wsPivot.PivotTables.Add( _
#             PivotCache:=ptCache, _
#             TableDestination:=wsPivot.Range("C3"), _
#             TableName:="PivotTable2")

#         ' Configure the second Pivot Table
#         With pt2
#             ' Add Vendor as Row Field
#             With .PivotFields("Configured By")
#                 .Orientation = xlRowField
#                 .Position = 1
#             End With
            
#             ' Add Sum of Quantity as Data Field
#             With .PivotFields("Start Date")
#                 .Orientation = xlDataField
#                 .Function = xlCount
#                 .Position = 1
#                 .NumberFormat = "0" ' Optional: Format the sum as a number with commas
#                 .Caption = "Events"
#             End With
#         End With
        
#         ' Create a Donut Chart from the second Pivot Table
#         Set chartObj2 = wsPivot.ChartObjects.Add( _
#             Left:=wsPivot.Cells(9, 18).Left, _
#             Width:=wsPivot.Range("R9:Y9").Width, _
#             Top:=wsPivot.Cells(9, 18).Top, _
#             Height:=wsPivot.Range("R9:R24").Height)
#         With chartObj2.Chart
#             .SetSourceData pt2.TableRange2
#             .ChartType = xlDoughnut
#             .HasTitle = True
#             .ChartTitle.Text = "Configured By"
#         End With
        
#         Exit Sub

#     ErrorHandler:
#         MsgBox "Error encountered: " & Err.Description, vbExclamation

# End Sub

# Sub Addslicer()

#     Dim ws As Worksheet
#     Dim pt As PivotTable
#     Dim sc1 As slicerCache
#     Dim sc2 As slicerCache
#     Dim sl1 As slicer
#     Dim sl2 As slicer

#     Set ws = Worksheets("PivotTable")
#     Set pt = ws.PivotTables("PivotTable1")

#     On Error Resume Next
#     ThisWorkbook.SlicerCaches("ConfiguredSlicerCache").Delete
#     ThisWorkbook.SlicerCaches("DateSlicerCaches").Delete
#     On Error GoTo 0

#     Set sc1 = ThisWorkbook.SlicerCaches.Add2( _
#         pt, _
#         "Configured By", _
#         "ConfiguredSlicerCache", _
#         XlSlicerCacheType.xlSlicer)
        
#     Set sl1 = sc1.Slicers.Add(ws, , "ConfiguredSlicer", "Choose Configured By", ws.Cells(9, 7).Top, ws.Cells(9, 7).Left, ws.Range("G9:I9").Width, ws.Range("G9:G24").Height)
        
#     Set sc2 = ThisWorkbook.SlicerCaches.Add2( _
#         pt, _
#         "Start Date", _
#         "DateSlicerCache", _
#         XlSlicerCacheType.xlTimeline)
        
#     Set sl2 = sc2.Slicers.Add( _
#         ws, , _
#         "DateSlicer", _
#         "Select Date Range", ws.Cells(1, 7).Top, ws.Cells(1, 7).Left, ws.Range("G1:Y1").Width, ws.Range("G1:G8").Height)

# End Sub

# Sub ConnectSlicerToMultiplePivots()

#     Dim ws As Worksheet
#     Dim pt As PivotTable
#     Dim sc1 As slicerCache
#     Dim sc2 As slicerCache

#     Set ws = Worksheets("PivotTable")
#     Set pt = ws.PivotTables("PivotTable2")
#     Set sc1 = ThisWorkbook.SlicerCaches("ConfiguredSlicerCache")
#     Set sc2 = ThisWorkbook.SlicerCaches("DateSlicerCache")

#     sc1.PivotTables.AddPivotTable pt
#     sc2.PivotTables.AddPivotTable pt

# End Sub
#     """
    
#     wb = xw.Book(file_path)
#     wb.api.VBProject.VBComponents.Add(1).CodeModule.AddFromString(vba_code)

#     macro1 = wb.macro("Module1.CreatePivotTableAndChart")
#     macro2 = wb.macro("Module1.Addslicer")
#     macro3 = wb.macro("Module1.ConnectSlicerToMultiplePivots")

#     macro1()
#     macro2()
#     macro3()

#     # Save and close the workbook
#     wb.save()
#     wb.close()

#     # Return the file path as a response
#     return FileResponse(path=file_path, filename="RFQ_Summary_21.xlsm")

# # Root endpoint
@app.get("/")
async def read_root():
    return {"message": "Welcome to the FastAPI app"}























