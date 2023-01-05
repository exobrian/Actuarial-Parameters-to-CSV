Attribute VB_Name = "TestingModule"
Sub CsvTemplateInterface_ExtractData()
    Dim test As Dictionary
        Set test = columnsToDict(Selection)
        For Each Key In test.Keys()
            Debug.Print Key
        Next Key
End Sub
