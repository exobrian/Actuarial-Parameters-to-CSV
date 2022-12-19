Attribute VB_Name = "Util"
'Helper Functions

Public Function stringFormat(ByVal mask As String, ParamArray tokens()) As String
    'pass parameters into a string
    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        mask = Replace(mask, "{" & i & "}", tokens(i))
    Next
    stringFormat = mask
 
End Function

Public Function paramsDbConnectString(ByVal dbTableName As String) As String
    'Uses global sql connection parameters for main server and db to create a connection string
    'Table depends on object
    
    sqlDriver = "SQL Server"    'had weird issues with other drivers.. try to use this one only.
    sqlServer = "DC1BISQLDEV02"
    sqlDatabase = "RatingPlatformParameters"
    sqlTable = dbTableName
    
    paramsDbConnectString = stringFormat("Driver={{0}};Server={1};Database={2};Trusted_Connection=yes", sqlDriver, sqlServer, sqlDatabase)
    
End Function

Function createCsvFile(csvInterface) As Workbook
    'Simply creates template with the object's csvSheetName. It also fetches the db table using the object's sqlQuery if the object's hasReferenceTable is True
    Dim csvWorkbook As Workbook
    Set csvWorkbook = Workbooks.Add
    
    csvWorkbook.Sheets(1).Name = csvInterface.csvSheetName
    
    'Most tables don't need to lookup a reference table
    If csvInterface.hasReferenceTable Then
        Dim conn As New ADODB.Connection
        Dim records As New ADODB.Recordset
        
        'Fetch data from db
        conn.Open csvInterface.dbConnString
        records.Open csvInterface.sqlQuery, conn, adOpenStatic, adLockReadOnly
        
        csvWorkbook.Sheets.Add.Name = csvInterface.dbTableName
        csvWorkbook.Sheets(csvInterface.dbTableName).Range("A2").CopyFromRecordset records
        
        'Sql Table Column Names
        For i = 0 To records.Fields.Count - 1
            csvWorkbook.Sheets(csvInterface.dbTableName).Cells(1, i + 1) = records.Fields(i).Name
        Next i
    End If
    
    Set createCsvFile = csvWorkbook
End Function

Public Function CreateDict(ByRef csvInterface As Object, ByRef csvWorkbook As Workbook) As Dictionary
    'Create dictionary using table's KEYNAME which is usually a string description in the pricing parameters wkbk and mapping the
    'foreign id from the db as its value
    
    Dim dict As Object
    Dim startRow As Integer, endRow As Integer
    
    Set dict = CreateObject("Scripting.Dictionary")
        
    With csvWorkbook.Sheets(csvInterface.dbTableName)
        'first row is always column header
        startRow = 2
        endRow = .Cells(2, 1).End(xlDown).Row
        keyColumn = .Cells(1, 1).EntireRow.Find(what:=csvInterface.KEYNAME, LookIn:=xlValues, lookat:=xlWhole).Column
        valueColumn = .Cells(1, 1).EntireRow.Find(what:=csvInterface.VALUENAME, LookIn:=xlValues, lookat:=xlWhole).Column
        
        For i = startRow To endRow
            dict.Add Key:=.Cells(i, keyColumn).Value, Item:=.Cells(i, valueColumn).Value
        Next i
    End With
    
    Set CreateDict = dict
End Function

Sub saveCsv(ByRef csvWorkbook As Workbook, ByVal csvSheetName As String)
    'This sub cleans the data file and strips all sheets other than the main table to be uploaded.
    'It will also create the CSV data path in the current directory if it does not exist
    'Finally it'll save the csv file with the appropriate version number
    Dim fso As New FileSystemObject
    Dim savePath As String
    
    savePath = ThisWorkbook.Path & "\CSV\"
    If Not fso.FolderExists(savePath) Then
        fso.CreateFolder savePath
    End If
    
    'passing effective date and expiration date for prefilling UI with default values to save user some time entering in
    With csvWorkbook.Sheets(csvSheetName)
        firstRow = 2
        lastRow = .Range("A1").End(xlDown).Row
        
        .Range("A1").End(xlToRight).Offset(0, 1) = "EffectiveDate"
        .Range(.Range("A1").End(xlToRight).Offset(1, 0), .Range("A1").End(xlToRight).Offset(lastRow - 1, 0)) = effDate
        .Range(.Range("A1").End(xlToRight).Offset(1, 0), .Range("A1").End(xlToRight).Offset(lastRow - 1, 0)).NumberFormat = "yyyy-mm-dd;@"
        
        .Range("A1").End(xlToRight).Offset(0, 1) = "ExpirationDate"
        .Range(.Range("A1").End(xlToRight).Offset(1, 0), .Range("A1").End(xlToRight).Offset(lastRow - 1, 0)) = expDate
        .Range(.Range("A1").End(xlToRight).Offset(1, 0), .Range("A1").End(xlToRight).Offset(lastRow - 1, 0)).NumberFormat = "yyyy-mm-dd;@"
    End With
    
    fileNumber = getCountOfFiles(savePath, csvSheetName)
    csvWorkbook.Sheets(csvSheetName).SaveAs savePath & csvSheetName & "_" & fileNumber, xlCSV
End Sub

Function getCountOfFiles(ByVal savePath As String, ByVal csvSheetName As String) As Integer
    'counts number of files with csvSheetName pattern
    
    Dim fso As New FileSystemObject
    Dim searchString As String
        
    searchString = Dir(savePath & "*" & csvSheetName & "*")
    
    i = 0
    Do While Len(searchString) > 0
        searchString = Dir
        i = i + 1
    Loop
    
    getCountOfFiles = i
End Function
