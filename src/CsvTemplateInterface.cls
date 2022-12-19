VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvTemplateInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Members for the interface
Public dbTableName As String
Public csvSheetName As String
Public paramSheetName As String
Public dbConnString As String
Public sqlQuery As String
Public hasReferenceTable As Boolean
Public KEYNAME As String
Public VALUENAME As String

'Methods - Declare here for the interface; then implement in the classes
Public Sub PrintInfo()
Public Sub ExtractData(ByRef csvWorkbook As Workbook, Optional ByRef dict As Dictionary)
