VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IndustryLerTemperingFactorTable"
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
    m_paramSheetName = "LER"
    m_dbTableName = "Ler_Icw"
    m_csvSheetName = "IndustryLerTf_Table"
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
    Dim keyStone, tableColumnRange, stateKey, limitTypeKey, lerTfStartKey, hgRange, lerTfValues As Variant
    Dim searchKey As String
    Dim stateStartRow, stateEndRow, lerTfStartCol, lerTfEndCol, currentRow, i As Integer
    
    With ThisWorkbook.Sheets(m_paramSheetName)
        'This should be a unique key to search that contains a known value in the table header. Be careful with this.
        'NOTE: There is an empty column in between the LERs and the tempering factors. We need this in order for this code to work.
        searchKey = "State"
        Set keyStone = .UsedRange.Find(what:=searchKey, LookIn:=xlValues, lookat:=xlWhole)
        Set tableColumnRange = keyStone.EntireRow
        Set stateKey = tableColumnRange.Find(what:="State", LookIn:=xlValues, lookat:=xlWhole)
        Set limitKey = tableColumnRange.Find(what:="Limit", LookIn:=xlValues, lookat:=xlWhole)
        
        'Ler set is first, then the tempering factors follow. I hate this part, but the spreadsheet does not have unique headers
        Set lerTfStartKey = tableColumnRange.Find(what:="1", LookIn:=xlValues, lookat:=xlWhole).End(xlToRight).End(xlToRight)
        
        'set start/end for our loop.
        stateStartRow = stateKey.Row + 1
        stateEndRow = stateKey.End(xlDown).Row
        lerTfStartCol = lerTfStartKey.Column
        lerTfEndCol = lerTfStartKey.End(xlToRight).Column
        
        'Use this to get the hazard group values
        Set hgRange = .Range(lerTfStartKey.Address & ":" & lerTfStartKey.End(xlToRight).Address)
        
        'LerTemperingFactors should have State, LossLimit, HazardGroup, IndustryLer
        'Note: Need to change this if/when IndustryLer column name in db is changed
        csvWorkbook.Sheets(m_csvSheetName).Range("A1") = "State"
        csvWorkbook.Sheets(m_csvSheetName).Range("B1") = "LossLimit"
        csvWorkbook.Sheets(m_csvSheetName).Range("C1") = "HazardGroup"
        csvWorkbook.Sheets(m_csvSheetName).Range("D1") = "IndustryLer"
        
        'Main Loop
        With csvWorkbook.Sheets(m_csvSheetName)
            currentRow = 2
            For i = stateStartRow To stateEndRow
                'Grab each state's ler tempering factor values and set here
                Set lerTfValues = ThisWorkbook.Sheets(m_paramSheetName).Range(ThisWorkbook.Sheets(m_paramSheetName).Cells(i, lerTfStartCol), ThisWorkbook.Sheets(m_paramSheetName).Cells(i, lerTfEndCol))
                
                .Range(.Cells(currentRow, 1), .Cells(currentRow + hgRange.Count - 1, 1)).Value = ThisWorkbook.Sheets(m_paramSheetName).Cells(i, stateKey.Column).Value
                .Range(.Cells(currentRow, 2), .Cells(currentRow + hgRange.Count - 1, 2)).Value = ThisWorkbook.Sheets(m_paramSheetName).Cells(i, limitKey.Column).Value
                .Range(.Cells(currentRow, 3), .Cells(currentRow + hgRange.Count - 1, 3)).Value = Application.WorksheetFunction.Transpose(hgRange.Value)
                .Range(.Cells(currentRow, 4), .Cells(currentRow + hgRange.Count - 1, 4)).Value = Application.WorksheetFunction.Transpose(lerTfValues.Value)
                currentRow = currentRow + hgRange.Count
            Next i
        End With
        
        'Some Ler Tempering Factors may have many sig figs.. converting to text here to make sure we capture it.
        csvWorkbook.Sheets(m_csvSheetName).Range("D:D").NumberFormat = "0.0000000000000000;@"
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
