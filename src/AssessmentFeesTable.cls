VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AssessmentFeesTable"
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
    m_paramSheetName = "GC"
    m_dbTableName = "AssessmentFees_Mapping"
    m_csvSheetName = "AssessmentFees_Table"
    m_KEYNAME = "AssessmentFeeDescription"
    m_VALUENAME = "Id"
    m_hasReferenceTable = True
    
    'This query may be more complicated depending on the table pulled
    m_sqlQuery = stringFormat("SELECT MAX([Id]) as [Id], [AssessmentFeeDescription]" & _
        " FROM [dbo].{0}" & _
        " GROUP BY [AssessmentFeeDescription]" & _
        " ORDER BY [AssessmentFeeDescription] ASC" _
        , m_dbTableName)
    m_dbConnString = paramsDbConnectString(m_dbTableName)
    
End Sub


'============================================================================================================
'Methods for this class
'============================================================================================================

Private Sub CsvTemplateInterface_ExtractData(ByRef csvWorkbook As Workbook, Optional ByRef dict As Dictionary)
    'This method should extract the data from the pricing parameter file and paste it into our csv file.
    'Eventually we'll need to clean the table and normalize it as well to match the formatting of the SQL table.
    'Current idea is to search the sheet for the keyword or column name for our data. Note that some do not have column names yet.
    Dim keyStone, cell As Variant
    Dim searchKey As String
    Dim startRow, endRow, stateCol, rateCol As Integer
    
    'Top left column header for range to copy
    searchKey = "Assessment Fees Description"
    Set keyStone = ThisWorkbook.Sheets(m_paramSheetName).UsedRange.Find(what:=searchKey, LookIn:=xlValues, lookat:=xlPart)
    
    startRow = keyStone.Row + 1
    endRow = keyStone.End(xlDown).Row
    stateCol = keyStone.EntireRow.Find(what:="State", LookIn:=xlValues, lookat:=xlPart).Column
    
    'Annoyingly, new value column in the pricing params for this table does not have a heading. For now hardcode, but later should put a unique header like NewValue or something
    rateCol = stateCol + 3
    
    'Assessment fee csv table should have State, AssessmentFeeId, Rate
    csvWorkbook.Sheets(m_csvSheetName).Range("A1") = "State"
    csvWorkbook.Sheets(m_csvSheetName).Range("B1") = "AssessmentFeeId"
    csvWorkbook.Sheets(m_csvSheetName).Range("C1") = "Rate"
    
    With ThisWorkbook.Sheets(m_paramSheetName)
        .Range(.Cells(startRow, stateCol), .Cells(endRow, stateCol)).Copy
        csvWorkbook.Sheets(m_csvSheetName).Range("A2").PasteSpecial xlValues
        
        .Range(.Cells(startRow, rateCol), .Cells(endRow, rateCol)).Copy
        csvWorkbook.Sheets(m_csvSheetName).Range("C2").PasteSpecial xlValues
        
        If m_hasReferenceTable Then
            i = 2
            For Each cell In .Range(keyStone.Offset(1, 0), keyStone.End(xlDown))
                csvWorkbook.Sheets(m_csvSheetName).Cells(i, 2).Value = dict(cell.Value)
                i = i + 1
            Next cell
        End If
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
