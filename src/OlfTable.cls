VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OlfTable"
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
    m_paramSheetName = "OLF"
    m_dbTableName = "Olf_Types"
    m_csvSheetName = "Olf_Table"
    m_hasReferenceTable = True 'SQL db has a reference table (OLF_Types) but it's not exact to the pricing params wkbk. We'll hardcode for now.
    m_KEYNAME = "description"
    m_VALUENAME = "type"
    
    'This query may be more complicated depending on the table pulled
    m_sqlQuery = stringFormat("SELECT * FROM dbo.{0}", m_dbTableName)
    m_dbConnString = paramsDbConnectString(m_dbTableName)
    
End Sub

'============================================================================================================
'Methods for this class
'============================================================================================================

Private Sub CsvTemplateInterface_ExtractData(ByRef csvWorkbook As Workbook, Optional ByRef dict As Dictionary)
    'This method should extract the data from the pricing parameter file and paste it into our csv file.
    'Eventually we'll need to clean the table and normalize it as well to match the formatting of the SQL table.
    'Current idea is to search the sheet for the keyword or column name for our data. Note that some do not have column names yet.
    Dim keyStone As Variant
    Dim searchKey As String
    Dim startRow, stateCol, rateCol As Integer
    Dim stateRange, accMonthRange, olfValueRange As Range
    
    'Olf csv should have AccidentMonth, State, OlfType, Olf, OlfVersion as column headers
    csvWorkbook.Sheets(m_csvSheetName).Range("A1") = "State"
    csvWorkbook.Sheets(m_csvSheetName).Range("B1") = "AccidentMonth"
    csvWorkbook.Sheets(m_csvSheetName).Range("C1") = "OlfType"
    csvWorkbook.Sheets(m_csvSheetName).Range("D1") = "Olf"
    csvWorkbook.Sheets(m_csvSheetName).Range("E1") = "OlfVersion"
    csvWorkbook.Sheets(m_csvSheetName).Range("F1") = "OlfAad"
    
    'State vector range: Assumes AK is first and OTHER is last.
    Set stateKeyStone = ThisWorkbook.Sheets(m_paramSheetName).UsedRange.Find(what:="AK", lookat:=xlWhole, LookIn:=xlValues)
    Set stateKeyStone2 = ThisWorkbook.Sheets(m_paramSheetName).UsedRange.Find(what:="OTHER", lookat:=xlWhole, LookIn:=xlValues)
    Set stateRange = ThisWorkbook.Sheets(m_paramSheetName).Range(stateKeyStone, stateKeyStone2)
    
    'Getting metadata locations for later
    aadRow = ThisWorkbook.Sheets(m_paramSheetName).UsedRange.Find(what:="AAD", lookat:=xlWhole, LookIn:=xlValues).Row
    versionRow = ThisWorkbook.Sheets(m_paramSheetName).UsedRange.Find(what:="Version", lookat:=xlWhole, LookIn:=xlValues).Row
    
    'Assumes accident month starts with 1/1/1900
    Set accMonthKeyStone = ThisWorkbook.Sheets(m_paramSheetName).Range("A:A").Find(what:="1/1/1900", lookat:=xlWhole, LookIn:=xlValues)
    Set accMonthKeyStone2 = accMonthKeyStone.End(xlDown)
    Set accMonthRange = ThisWorkbook.Sheets(m_paramSheetName).Range(accMonthKeyStone, accMonthKeyStone2)
    
    With ThisWorkbook.Sheets(m_paramSheetName)
        'start and end rows for accident months; note these are the same for every table
        startRow = accMonthRange.Item(1).Row
        endRow = accMonthRange.Item(accMonthRange.Count).Row
        currRow = 2 'starting row for pasting into the csv file
        
        'retrieve foreign keys from db dictionary
        'Note: These keys are messed up right now since the db does not correspond exactly to the pricing parameters workbook.
        'Hardcoding here but in the future, we may want to change the pricing parameters to match.
        'i = foreign key for table in db
        Dim olfKeys(0 To 4) As String
        olfKeys(0) = "Combined Indemnity & Medical"     'TABLE 1
        olfKeys(1) = "Indemnity Frequency Benefit Level Trend Factors (Specific to CA & FL Only)"   'TABLE 2
        olfKeys(2) = "Frequency Residual Trend Factors, excluding Med-only claims, On-level for exposure"   'TABLE 3
        olfKeys(3) = "Severity Residual Trend Factors, excluding Med-only claims, On-level for Benefit & Wage"  'TABLE 4
        olfKeys(4) = "Statewide Average Weekly Wage Factors"    'TABLE 5
        
        rowOffset = 0
        
        'Note: Tables start at 1 in db, but here we index at i = 0.
        For i = 0 To (UBound(olfKeys) - LBound(olfKeys))
            'Offset the starting row by the olfKeys above to navigate to olf tables.
            'olfKeys are one row above the 1st accident months.
            rowOffset = .UsedRange.Find(what:=olfKeys(i), lookat:=xlWhole, LookIn:=xlValues).Row + 1 - startRow
            
            For Each st In stateRange
                stColumn = st.Column
                stAad = st.Offset(aadRow - st.Row).Value    'to be pushed to db later. For now, just store this.
                stVersion = st.Offset(versionRow - st.Row).Value
                
                'Olf values for state
                Set olfValueRange = .Range(.Cells(startRow + rowOffset, stColumn), .Cells(endRow + rowOffset, stColumn))
                csvWorkbook.Sheets(m_csvSheetName).Range(CStr("A" & currRow), CStr("A" & currRow + olfValueRange.Count - 1)) = st.Value
                csvWorkbook.Sheets(m_csvSheetName).Range(CStr("B" & currRow), CStr("B" & currRow + olfValueRange.Count - 1)) = accMonthRange.Value
                csvWorkbook.Sheets(m_csvSheetName).Range(CStr("C" & currRow), CStr("C" & currRow + olfValueRange.Count - 1)) = i + 1
                csvWorkbook.Sheets(m_csvSheetName).Range(CStr("D" & currRow), CStr("D" & currRow + olfValueRange.Count - 1)) = olfValueRange.Value
                csvWorkbook.Sheets(m_csvSheetName).Range(CStr("E" & currRow), CStr("E" & currRow + olfValueRange.Count - 1)) = stVersion
                csvWorkbook.Sheets(m_csvSheetName).Range(CStr("F" & currRow), CStr("F" & currRow + olfValueRange.Count - 1)) = stAad
                
                'Need to format this different due to the way we pasted in a range
                csvWorkbook.Sheets(m_csvSheetName).Range(CStr("B" & currRow), CStr("B" & currRow + olfValueRange.Count - 1)).NumberFormat = "yyyy-mm-dd;@"
                csvWorkbook.Sheets(m_csvSheetName).Range(CStr("F" & currRow), CStr("F" & currRow + olfValueRange.Count - 1)).NumberFormat = "yyyy-mm-dd;@"
                
                'Need to format this in order to prevent truncation before saving as csv
                csvWorkbook.Sheets(m_csvSheetName).Range(CStr("D" & currRow), CStr("D" & currRow + olfValueRange.Count - 1)).NumberFormat = "0.0000000000000000;@"
                
                currRow = currRow + olfValueRange.Count
            Next st
        Next i
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

