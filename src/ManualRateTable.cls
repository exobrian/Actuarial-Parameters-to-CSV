VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManualRateTable"
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
    m_paramSheetName = "Rates I"
    m_dbTableName = "ManualRate"
    m_csvSheetName = "ManualRate_Table"
    m_KEYNAME = "CompanyTierKey"
    m_VALUENAME = "CompanyTierValue"
    m_hasReferenceTable = True
    
    'This query may be more complicated depending on the table pulled; Pull duplicate CompanyTier column to create a dictionary unto itself for CreateDict arguments
    m_sqlQuery = stringFormat("SELECT DISTINCT CompanyTier AS CompanyTierKey, CompanyTier AS CompanyTierValue FROM dbo.{0} ORDER BY CompanyTier ASC", m_dbTableName)
    m_dbConnString = paramsDbConnectString(m_dbTableName)
    
End Sub

Private Sub CsvTemplateInterface_ExtractData(ByRef csvWorkbook As Workbook, Optional ByRef dict As Dictionary)
    'This method should extract the data from the pricing parameter file and paste it into our csv file.
    'Eventually we'll need to clean the table and normalize it as well to match the formatting of the SQL table.
    'Current idea is to search the sheet for the keyword or column name for our data. Note that some do not have column names yet.
    Dim keyStone As Variant
    Dim searchKey As String
    Dim startRow As Integer
    Dim stateCol As Integer
    Dim rateCol As Integer
    Dim keyCell As Object
    Dim stateDict(), tempDict As Object
    
    With ThisWorkbook.Sheets(m_paramSheetName)
        'if autofilter is not on, be sure to turn it on
        If Not .AutoFilterMode Then
            tableColumnRange.AutoFilter
        End If
                    
        'This should be a unique key to search that contains a known value in the table header. Be careful with this.
        searchKey = "Class Code"
        Set keyStone = .UsedRange.Find(what:=searchKey, LookIn:=xlValues, lookat:=xlWhole)
        
        'Need these after clearing filter
        .AutoFilter.ShowAllData
        startRow = keyStone.Row + 1
        endRow = keyStone.End(xlDown).Row
        
        'setting range for us to find column by name. Note that "Min Premium" is annoyingly not in the same cell and can't be searched as one string
        Set tableColumnRange = .Range(keyStone.End(xlToLeft), keyStone.End(xlToRight))
        Set stateKey = tableColumnRange.Find(what:="State", LookIn:=xlValues, lookat:=xlWhole)
        Set colorKey = tableColumnRange.Find(what:="Color", LookIn:=xlValues, lookat:=xlWhole)
        Set minPremKey = tableColumnRange.Find(what:="Premium", LookIn:=xlValues, lookat:=xlWhole)
        
        'Getting a dictionary of state names contained in this table for iterating
        Set tempDict = CreateObject("scripting.dictionary")
        stateDict = .Range(ThisWorkbook.Sheets(m_paramSheetName).Cells(startRow, stateKey.Column), ThisWorkbook.Sheets(m_paramSheetName).Cells(endRow, stateKey.Column)).Value
        For stateRow = 1 To UBound(stateDict)
            tempDict(stateDict(stateRow, 1)) = Empty
        Next stateRow
        stateDict = WorksheetFunction.Transpose(tempDict.Keys())
                
        'ManualRate should have State, ClassCode, ClassColor, MinimumPremium, CompanyTier, ManualRate
        csvWorkbook.Sheets(m_csvSheetName).Range("A1") = "State"
        csvWorkbook.Sheets(m_csvSheetName).Range("B1") = "ClassCode"
        csvWorkbook.Sheets(m_csvSheetName).Range("C1") = "ClassColor"
        csvWorkbook.Sheets(m_csvSheetName).Range("D1") = "MinimumPremium"
        csvWorkbook.Sheets(m_csvSheetName).Range("E1") = "CompanyTier"
        csvWorkbook.Sheets(m_csvSheetName).Range("F1") = "ManualRate"
        
        
        'This part is messy at best. The idea is to iterate through the columns in the rate table using the db unique CompanyTier string as keys.
        'Then filter each column for each state to check if the state has any rates for this combination. Excel has no easy way of checking for nonempty tables, so we use a hack.
        'Eventually, we'll want to get the keys from a mapping table. For now, distinct values from preexisting uploads is good enough.
        currentRow = 2
        For Each cotierKey In dict.Keys()
            Set keyCell = tableColumnRange.Find(what:=dict(cotierKey), LookIn:=xlValues, lookat:=xlWhole)
            keyCell.AutoFilter Field:=Application.WorksheetFunction.Match(dict(cotierKey), tableColumnRange, 0), Criteria1:="<>"
            For Each st In stateDict
                stateKey.AutoFilter Field:=Application.WorksheetFunction.Match("State", tableColumnRange, 0), Criteria1:=st
                If (stateKey.End(xlDown).End(xlDown).End(xlUp).Value <> "State") Then
                    currRowCount = .Range(.Cells(startRow, keyCell.Column), .Cells(endRow, keyCell.Column)).SpecialCells(xlCellTypeVisible).Count
                    
                    'State
                    csvWorkbook.Sheets(m_csvSheetName).Range(csvWorkbook.Sheets(m_csvSheetName).Cells(currentRow, 1), csvWorkbook.Sheets(m_csvSheetName).Cells(currentRow + currRowCount - 1, 1)) = _
                        st
                    'Class Code
                    csvWorkbook.Sheets(m_csvSheetName).Range(csvWorkbook.Sheets(m_csvSheetName).Cells(currentRow, 2), csvWorkbook.Sheets(m_csvSheetName).Cells(currentRow + currRowCount - 1, 2)) = _
                        .Range(.Cells(startRow, keyStone.Column), .Cells(endRow, keyStone.Column)).SpecialCells(xlCellTypeVisible).Value
                    'Class Color
                    csvWorkbook.Sheets(m_csvSheetName).Range(csvWorkbook.Sheets(m_csvSheetName).Cells(currentRow, 3), csvWorkbook.Sheets(m_csvSheetName).Cells(currentRow + currRowCount - 1, 3)) = _
                        .Range(.Cells(startRow, colorKey.Column), .Cells(endRow, colorKey.Column)).SpecialCells(xlCellTypeVisible).Value
                    'Minimum Premium
                    csvWorkbook.Sheets(m_csvSheetName).Range(csvWorkbook.Sheets(m_csvSheetName).Cells(currentRow, 4), csvWorkbook.Sheets(m_csvSheetName).Cells(currentRow + currRowCount - 1, 4)) = _
                        .Range(.Cells(startRow, minPremKey.Column), .Cells(endRow, minPremKey.Column)).SpecialCells(xlCellTypeVisible).Value
                    'Company Tier
                    csvWorkbook.Sheets(m_csvSheetName).Range(csvWorkbook.Sheets(m_csvSheetName).Cells(currentRow, 5), csvWorkbook.Sheets(m_csvSheetName).Cells(currentRow + currRowCount - 1, 5)) = _
                        cotierKey
                    'Manual Rate
                    csvWorkbook.Sheets(m_csvSheetName).Range(csvWorkbook.Sheets(m_csvSheetName).Cells(currentRow, 6), csvWorkbook.Sheets(m_csvSheetName).Cells(currentRow + currRowCount - 1, 6)) = _
                        .Range(.Cells(startRow, keyCell.Column), .Cells(endRow, keyCell.Column)).SpecialCells(xlCellTypeVisible).Value
                    csvWorkbook.Sheets(m_csvSheetName).Range(csvWorkbook.Sheets(m_csvSheetName).Cells(currentRow, 6), csvWorkbook.Sheets(m_csvSheetName).Cells(currentRow + currRowCount - 1, 6)).NumberFormat = "0.00;@"
                                        
                    currentRow = currentRow + currRowCount
                End If
            Next st
            
            'reset objects and filter
            .AutoFilter.ShowAllData
            Set keyCell = Nothing
            
        Next cotierKey
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

