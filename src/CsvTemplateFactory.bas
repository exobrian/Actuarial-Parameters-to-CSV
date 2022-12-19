Attribute VB_Name = "csvTemplateFactory"
Option Explicit
'Global variables we'll be using for every table.
Global effDate As String
Global expDate As String

Sub main()
    Dim classObject As Object
    Dim csvTable As String
    Dim effDateValue As Date
    Dim csvInterface As CsvTemplateInterface
    Dim csvWorkbook As Workbook
    Dim dict As Dictionary
    
    'turn these off before running to conserve resources
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    'test variable
    csvTable = "ManualRate"
    effDateValue = DateValue(ThisWorkbook.Sheets("Summary").UsedRange.Find(what:="Effective", lookat:=xlWhole, LookIn:=xlValues).Offset(0, 1).Value)
    effDate = CStr(Format(DateValue(effDateValue), "YYYY-MM-DD"))
    expDate = CStr(Format(DateAdd("yyyy", 1, effDateValue), "YYYY-MM-DD"))
    
    'Reflection to get object type
    If csvTable = "AssessmentFees" Then
        Set classObject = New AssessmentFeesTable
    ElseIf csvTable = "Olf" Then
        Set classObject = New OlfTable
    ElseIf csvTable = "ManualRate" Then
        Set classObject = New ManualRateTable
    Else
        MsgBox "Please enter a valid table as input.", Title:="Error: Input table not allowed"
        Exit Sub
    End If
        
    'Need to check if table has foreign keys in db. If not, we don't need to fetch data and map.
    Set csvInterface = classObject
    Set csvWorkbook = createCsvFile(csvInterface)
    
    'Create dictionary using sql mapping table
    If csvInterface.hasReferenceTable Then
        Set dict = CreateDict(csvInterface, csvWorkbook)
        csvInterface.ExtractData csvWorkbook, dict
    Else
        csvInterface.ExtractData csvWorkbook
    End If
    
    
    Call saveCsv(csvWorkbook, csvInterface.csvSheetName)
    csvWorkbook.Close False
        
    'Turn these global settings back on
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
End Sub
