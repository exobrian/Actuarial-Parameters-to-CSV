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

Private Sub CsvTemplateInterface_PrintInfo()
    Debug.Print CsvTemplateInterface_dbTableName
    Debug.Print CsvTemplateInterface_csvSheetName
End Sub

'initialize members for this class
Private Sub Class_Initialize()
    m_dbTableName = "dimOlfTypes"
    m_csvSheetName = "OlfTypes"
End Sub

Private Sub dbConnect()
    Dim connString As String
    connString = paramsDbConnectString(m_dbTableName)
End Sub
