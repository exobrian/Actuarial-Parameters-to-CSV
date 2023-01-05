Attribute VB_Name = "Injector"
Public Sub GetParamFiles()
    Dim fso As New FileSystemObject
    
    mainPath = "\\newfs5\EXCEL\ACTUARIA\WCDEV\LDP_On-Level Analysis\"
    fileString = "pricing parameters"
    
    'Note: Expecting folder format of "True_Up_YYYYMMDD_Pricing"
    currFolder = Dir(mainPath & "*_Pricing", vbDirectory)
    
    While Len(currFolder) > 0
        'Grab the date from folder path with patter YYYYMMDD.
        currDate = Left(Right(currFolder, 16), 8)
        
        'Main loop: some param files may not have the right tables.. we'll skip and log these
        savePath = ThisWorkbook
        On Error GoTo Skip
        If Not fso.FolderExists(savePath) Then
            fso.CreateFolder savePath
        End If
        
        
        
Skip:
        If Err.Number > 0 Then
            Debug.Print "Error in " & currDate; ""
            fso.DeleteFolder savePath
            Err.Clear
        End If
        
        'next folder
        currFolder = Dir()
    Wend
End Sub

Public Sub ExportModules()
    'This code will export all modules and class files in this workbook into the path below
    'Make sure Tools > References > Microsoft Visual Basic For Applications Extensibility 5.3 and Microsoft Scripting Runtime enabled
    Dim cmpComponent As VBIDE.VBComponent
    
    'Save path for macro files
    'srcExportPath = "\\sancfs001\Actuarial$\Staff\Brian Tran\Code Samples\Macros\ParametersToCsv\src\"
    srcExportPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(Application.ThisWorkbook.path) & "\src\"
    
    Set srcWorkbook = Application.Workbooks(ThisWorkbook.Name)
    'Set srcWorkbook = ThisWorkbook
    For Each cmpComponent In srcWorkbook.VBProject.vbcomponents
        boolExport = True
        srcFileName = cmpComponent.Name
        
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                srcFileName = srcFileName & ".cls"
            Case vbext_ct_StdModule
                srcFileName = srcFileName & ".bas"
            Case Else
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                boolExport = False
        End Select
        
        If boolExport Then
            cmpComponent.Export srcExportPath & srcFileName
        End If
        
    Next cmpComponent
End Sub
