Attribute VB_Name = "TestingModule"
Sub CsvTemplateInterface_ExtractData()
    'This method should extract the data from the pricing parameter file and paste it into our csv file.
    'Eventually we'll need to clean the table and normalize it as well to match the formatting of the SQL table.
    'Current idea is to search the sheet for the keyword or column name for our data. Note that some do not have column names yet.
    Dim keyStone As Variant
    Dim searchKey As String
    Dim startRow As Integer
    Dim stateCol As Integer
    Dim rateCol As Integer
    
    m_paramSheetName = "Rates I"
    
    Set csvWorkbook = ThisWorkbook
    
    'Top left column header for range to copy
    searchKey = "Class Code"
    Set keyStone = ThisWorkbook.Sheets(m_paramSheetName).UsedRange.Find(what:=searchKey, LookIn:=xlValues, lookat:=xlWhole)
    
    startRow = keyStone.Row + 1
    endRow = keyStone.End(xlDown).Row
    
    'Column numbers: consider using a 2d array to map these later.
    Set tableColumnRange = ThisWorkbook.Sheets(m_paramSheetName).Range(keyStone.End(xlToLeft), keyStone.End(xlToRight))
    stateCol = tableColumnRange.Find(what:="State", LookIn:=xlValues, lookat:=xlWhole).Column
    classCol = tableColumnRange.Find(what:="Class Code", LookIn:=xlValues, lookat:=xlWhole).Column
    concatCol = tableColumnRange.Find(what:="State", LookIn:=xlValues, lookat:=xlWhole).Column
    colorCol = tableColumnRange.Find(what:="Color", LookIn:=xlValues, lookat:=xlWhole).Column
    minPremCol = tableColumnRange.Find(what:="Premium", LookIn:=xlValues, lookat:=xlWhole).Column
    icw1Col = tableColumnRange.Find(what:="ICW1", LookIn:=xlValues, lookat:=xlWhole).Column
    icw2Col = tableColumnRange.Find(what:="ICW2", LookIn:=xlValues, lookat:=xlWhole).Column
    icw3Col = tableColumnRange.Find(what:="ICW3", LookIn:=xlValues, lookat:=xlWhole).Column
    icw4Col = tableColumnRange.Find(what:="ICW4", LookIn:=xlValues, lookat:=xlWhole).Column
    icw5Col = tableColumnRange.Find(what:="ICW5", LookIn:=xlValues, lookat:=xlWhole).Column
    icw6Col = tableColumnRange.Find(what:="ICW6", LookIn:=xlValues, lookat:=xlWhole).Column
    icw7Col = tableColumnRange.Find(what:="ICW7", LookIn:=xlValues, lookat:=xlWhole).Column
    icw8Col = tableColumnRange.Find(what:="ICW8", LookIn:=xlValues, lookat:=xlWhole).Column
    icw9Col = tableColumnRange.Find(what:="ICW9", LookIn:=xlValues, lookat:=xlWhole).Column
    eic1Col = tableColumnRange.Find(what:="EIC1", LookIn:=xlValues, lookat:=xlWhole).Column
        
    'Assessment fee csv table should have State, AssessmentFeeId, Rate
    csvWorkbook.Sheets(m_csvSheetName).Range("A1") = "State"
    csvWorkbook.Sheets(m_csvSheetName).Range("B1") = "AssessmentFeeId"
    csvWorkbook.Sheets(m_csvSheetName).Range("C1") = "Rate"
    
    With ThisWorkbook.Sheets(m_paramSheetName)
        .Range(.Cells(startRow, stateCol), .Cells(endRow, stateCol)).Copy
        csvWorkbook.Sheets(m_csvSheetName).Range("A2").PasteSpecial xlValues
        
        .Range(.Cells(startRow, rateCol), .Cells(endRow, rateCol)).Copy
        csvWorkbook.Sheets(m_csvSheetName).Range("C2").PasteSpecial xlValues
        
    End With
End Sub
