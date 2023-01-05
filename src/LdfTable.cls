VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LdfTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements CsvTemplateInterface

Private m_dbTableName As String
Private m_csvSheetName As String
Private m_dbConnString As String
Private m_sqlQuery As String
Private m_paramSheetName As String
Private m_hasReferenceTable As Boolean
Private m_KEYNAME As String
Private m_VALUENAME As String

Private Property Let CsvTemplateInterface_paramSheetName(ByVal paramSheetName As String)
    m_paramSheetName = paramSheetName
End Property

Private Property Get CsvTemplateInterface_paramSheetName() As String
    CsvTemplateInterface_paramSheetName = m_paramSheetName
End Property

Private Property Let CsvTemplateInterface_dbTableName(ByVal dbTableName As String)
    m_dbTableName = dbTableName
End Property

Private Property Get CsvTemplateInterface_dbTableName() As String
    CsvTemplateInterface_dbTableName = m_dbTableName
End Property

Private Property Let CsvTemplateInterface_csvSheetName(ByVal csvSheetName As String)
    m_csvSheetName = csvSheetName
End Property

Private Property Get CsvTemplateInterface_csvSheetName() As String
    CsvTemplateInterface_csvSheetName = m_csvSheetName
End Property

Private Property Let CsvTemplateInterface_dbConnString(ByVal dbConnString As String)
    m_dbConnString = dbConnString
End Property

Private Property Get CsvTemplateInterface_dbConnString() As String
    CsvTemplateInterface_dbConnString = m_dbConnString
End Property

Private Property Let CsvTemplateInterface_sqlQuery(ByVal sqlQuery As String)
    m_sqlQuery = sqlQuery
End Property

Private Property Get CsvTemplateInterface_sqlQuery() As String
    CsvTemplateInterface_sqlQuery = m_sqlQuery
End Property

Private Property Let CsvTemplateInterface_hasReferenceTable(ByVal hasReferenceTable As Boolean)
    m_hasReferenceTable = hasReferenceTable
End Property

Private Property Get CsvTemplateInterface_hasReferenceTable() As Boolean
    CsvTemplateInterface_hasReferenceTable = m_hasReferenceTable
End Property

Private Property Let CsvTemplateInterface_KEYNAME(ByVal KEYNAME As String)
    m_KEYNAME = KEYNAME
End Property

Private Property Get CsvTemplateInterface_KEYNAME() As String
    CsvTemplateInterface_KEYNAME = m_KEYNAME
End Property

Private Property Let CsvTemplateInterface_VALUENAME(ByVal VALUENAME As String)
    m_VALUENAME = VALUENAME
End Property

Private Property Get CsvTemplateInterface_VALUENAME() As String
    CsvTemplateInterface_VALUENAME = m_VALUENAME
End Property

'============================================================================================================
'Initializing members for this class
'============================================================================================================

Private Sub Class_Initialize()
    
    'These are specific to the class
    m_paramSheetName = "LDF"
    m_dbTableName = "Ldf"
    m_csvSheetName = "Ldf_Table"
    m_KEYNAME = ""
    m_VALUENAME = ""
    m_hasReferenceTable = False
    
    'This query may be more complicated depending on the table pulled
    m_sqlQuery = stringFormat("SELECT * FROM dbo.{0}", m_dbTableName)
    m_dbConnString = paramsDbConnectString(m_dbTableName)
    
End Sub

Private Sub CsvTemplateInterface_ExtractData(ByRef csvWorkbook As Workbook, Optional ByRef dict As Dictionary)
    'This method should extract the data from the pricing parameter file and paste it into our csv file.
    'Eventually we'll need to clean the table and normalize it as well to match the formatting of the SQL table.
    'Current idea is to search the sheet for the keyword or column name for our data. Note that some do not have column names yet.
    Dim keyStone, tableColumnRange, stateKey, ldfTypeKey, ageStartKey, ageRange, ldfValues As Variant
    Dim searchKey As String
    Dim stateStartRow, stateEndRow, ageStartCol, ageEndCol, currentRow, i As Integer
    
    With ThisWorkbook.Sheets(m_paramSheetName)
        'This should be a unique key to search that contains a known value in the table header. Be careful with this.
        'Not a fan of all this searching and use of entire row.. but there are blank columns in between headers we can't ignore
        searchKey = "State"
        Set keyStone = .UsedRange.Find(what:=searchKey, LookIn:=xlValues, lookat:=xlWhole)
        Set tableColumnRange = keyStone.EntireRow
        Set stateKey = tableColumnRange.Find(what:="State", LookIn:=xlValues, lookat:=xlWhole)
        Set ldfTypeKey = tableColumnRange.Find(what:="Pd or Inc", LookIn:=xlValues, lookat:=xlWhole)
        Set ageStartKey = tableColumnRange.Find(what:="1", LookIn:=xlValues, lookat:=xlWhole)
        
        'set start/end for our loop
        stateStartRow = stateKey.Row + 1
        stateEndRow = stateKey.End(xlDown).Row
        ageStartCol = ageStartKey.Column
        ageEndCol = ageStartKey.End(xlToRight).Column
        
        'Use this to get the age values
        Set ageRange = .Range(ageStartKey.Address & ":" & ageStartKey.End(xlToRight).Address)
                
        'LDF should have State, PaidOrIncurred, AccidentAge, Value
        csvWorkbook.Sheets(m_csvSheetName).Range("A1") = "State"
        csvWorkbook.Sheets(m_csvSheetName).Range("B1") = "PaidOrIncurred"
        csvWorkbook.Sheets(m_csvSheetName).Range("C1") = "AccidentAge"
        csvWorkbook.Sheets(m_csvSheetName).Range("D1") = "Value"
        
        'Main Loop
        With csvWorkbook.Sheets(m_csvSheetName)
            currentRow = 2
            For i = stateStartRow To stateEndRow
                'Grab each state's ldf values and set here
                Set ldfValues = ThisWorkbook.Sheets(m_paramSheetName).Range(ThisWorkbook.Sheets(m_paramSheetName).Cells(i, ageStartCol), ThisWorkbook.Sheets(m_paramSheetName).Cells(i, ageEndCol))
                
                .Range(.Cells(currentRow, 1), .Cells(currentRow + ageRange.Count - 1, 1)).Value = ThisWorkbook.Sheets(m_paramSheetName).Cells(i, stateKey.Column).Value
                .Range(.Cells(currentRow, 2), .Cells(currentRow + ageRange.Count - 1, 2)).Value = ThisWorkbook.Sheets(m_paramSheetName).Cells(i, ldfTypeKey.Column).Value
                .Range(.Cells(currentRow, 3), .Cells(currentRow + ageRange.Count - 1, 3)).Value = Application.WorksheetFunction.Transpose(ageRange.Value)
                .Range(.Cells(currentRow, 4), .Cells(currentRow + ageRange.Count - 1, 4)).Value = Application.WorksheetFunction.Transpose(ldfValues.Value)
                currentRow = currentRow + ageRange.Count
            Next i
        End With
        
    End With
End Sub


'============================================================================================================
'Testing tools
'============================================================================================================
Private Sub CsvTemplateInterface_PrintInfo()
    Debug.Print CsvTemplateInterface_dbTableName
    Debug.Print CsvTemplateInterface_csvSheetName
    Debug.Print CsvTemplateInterface_dbConnString
    Debug.Print CsvTemplateInterface_sqlQuery
End Sub
