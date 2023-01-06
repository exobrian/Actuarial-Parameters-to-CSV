Attribute VB_Name = "Injector"
Public Sub processParamFiles()
    'This code loops through original set of parameter files and injects all the vba code from thisworkbook into it
    'It will also inject references as well, then call the main subroutine
    Dim fso As New FileSystemObject
    Dim folderList As New Collection
    
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    
    mainPath = "\\newfs5\EXCEL\ACTUARIA\WCDEV\LDP_On-Level Analysis\"
    fileString = "pricing parameters"
    
    'Need to create array of paths for subfolders first since Dir is static and cannot be used again once
    'Note: Expecting folder format of "True_Up_YYYYMMDD_Pricing"
    currFolder = Dir(mainPath & "*_Pricing", vbDirectory)
    While Len(currFolder)
        folderList.Add currFolder
        currFolder = Dir
    Wend
    
    For Each currFolder In folderList
        'Grab the date from folder path with patter YYYYMMDD.
        currDate = Left(Right(currFolder, 16), 8)
        
        'Main loop: some param files may not have the right tables.. we'll skip and log these
        savePath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(Application.ThisWorkbook.path) & "\" & currDate & "\"
        On Error GoTo Skip
        If Not fso.FolderExists(savePath) Then
            fso.CreateFolder savePath
        End If
        
        currPath = mainPath & currFolder & "\"
        paramFileName = getNewest(CStr(currPath), "pricing parameters")
        
        Set paramFile = Workbooks.Open(currPath & paramFileName, ReadOnly:=True)
        paramFileBaseName = Mid(paramFileName, 1, InStrRev(paramFileName, ".") - 1)
        If paramFile.MultiUserEditing Then
            'Note: if file is shared, we need to unshare it, close then reopen to inject the vba
            Debug.Print "Warning: The workbook " & paramFileName & " is a shared file."
            paramFile.SaveAs FileName:=savePath & paramFileBaseName & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled, accessmode:=xlExclusive
            paramFile.Close (False)
            Set paramFile = Workbooks.Open(savePath & paramFileBaseName & ".xlsm", ReadOnly:=False)
        Else
            'save with accessmode as is
            paramFile.SaveAs FileName:=savePath & paramFileBaseName & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
        End If
        
        Call importModules(paramFile)
        Call injectReferences(paramFile)
        Run "'" & paramFileBaseName & ".xlsm'!main", "Olf"
        Run "'" & paramFileBaseName & ".xlsm'!main", "ManualRate"
        Run "'" & paramFileBaseName & ".xlsm'!main", "Ldf"
        
Skip:
        If Err.Number > 0 Then
            Debug.Print "Error in " & currDate; ""
            'fso.DeleteFolder savePath
            Err.Clear
        End If
                
        paramFile.Close (True)
        Set paramFile = Nothing
    Next currFolder
    
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
End Sub

Public Sub exportModules()
    'This code will export all modules and class files in this workbook into the path below
    'Make sure Tools > References > Microsoft Visual Basic For Applications Extensibility 5.3 and Microsoft Scripting Runtime enabled
    Dim cmpComponent As VBIDE.VBComponent
    
    'Save path for macro files
    srcExportPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(Application.ThisWorkbook.path) & "\src\"
    
    Set srcWorkbook = Application.Workbooks(ThisWorkbook.Name)
    For Each cmpComponent In srcWorkbook.VBProject.VBComponents
        boolExport = True
        srcFileName = cmpComponent.Name
        
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                srcFileName = srcFileName & ".cls"
            Case vbext_ct_StdModule
                srcFileName = srcFileName & ".bas"
            Case Else
                ' This is a worksheet or workbook object so don't export.
                boolExport = False
        End Select
        
        If boolExport Then
            cmpComponent.Export srcExportPath & srcFileName
        End If
        
    Next cmpComponent
End Sub

Public Sub importModules(ByRef wkbkTarget As Variant)
    'This code will import all modules/class files from the specified path into the targeted workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim cmpComponents As VBIDE.VBComponents
    
    srcPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(Application.ThisWorkbook.path) & "\src\"
    
    Set cmpComponents = wkbkTarget.VBProject.VBComponents
    Set objFSO = New Scripting.FileSystemObject
    
    For Each objFile In objFSO.GetFolder(srcPath).Files
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.path
        End If
    Next objFile

End Sub

Public Sub injectReferences(ByRef wkbkTarget As Variant)
    'This codes clones the current workbooks' references and injects them into the target workbook
    
    Dim strGUIDs As Dictionary, strGUID As Variant, theRef As Variant, i As Long
    
    'Set to continue in case of error
    On Error Resume Next
    
    'Get dictionary of references we'll need from parent workbook
    Set strGUIDs = getRefGUIDs
     
'     'Remove any missing references. Not really needed but leaving here in case of other bugs later
'    For i = wkbkTarget.VBProject.References.Count To 1 Step -1
'        Set theRef = wkbkTarget.VBProject.References.Item(i)
'        If theRef.IsBroken = True Then
'            wkbkTarget.VBProject.References.Remove theRef
'        End If
'    Next i
     
     'Clear any errors so that error trapping for GUID additions can be evaluated
    Err.Clear
    
    i = 0
     'Add the reference
    For Each strGUIDKey In strGUIDs.Keys()
        wkbkTarget.VBProject.References.AddFromGuid _
        GUID:=strGUIDKey, Major:=0, Minor:=0
        
         'If an error was encountered, inform the user
        Select Case Err.Number
        Case Is = 32813
             'Reference already in use.  No action necessary
             Debug.Print "Item " & i, strGUIDKey, strGUIDs.Item(strGUIDKey), "Reference in use"
        Case Is = vbNullString
             'Reference added without issue
             Debug.Print "Item " & i, strGUIDKey, strGUIDs.Item(strGUIDKey), "Reference added"
        Case Else
             'An unknown error was encountered, so alert the user
            MsgBox "A problem was encountered trying to" & vbNewLine _
            & "add or remove a reference in this file" & vbNewLine & "Please check the " _
            & "references in your VBA project!", vbCritical + vbOKOnly, "Error!"
        End Select
        
        i = i + 1
        On Error GoTo -1
    Next strGUIDKey
    Err.Clear
End Sub


Sub xlfVBEListReferences()
    ' This code prints out all the references currently enabled in this workbook. Used for debugging.
    ' Requires References :: Microsoft Visual Basic for Applications Extensibility 5.3 (GUID {0002E157-0000-0000-C000-000000000046})
    ' C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
    Dim oRef As VBIDE.Reference   ' Item
    Dim oRefs As VBIDE.References ' Collection
    Dim i As Integer
     
    Set oRefs = Application.VBE.ActiveVBProject.References
 
    Debug.Print "Print Time: " & Time & " :: Item - Name and Description"
        For Each oRef In oRefs
            i = i + 1
            Debug.Print "Item " & i, oRef.Name, oRef.Description
        Next oRef
    Debug.Print vbNewLine
 
    i = 0
    Debug.Print "Print Time: " & Time & " :: Item - Full Path"
        For Each oRef In oRefs
            i = i + 1
            Debug.Print "Item " & i, oRef.FullPath
        Next oRef
    Debug.Print vbNewLine
 
    i = 0
    ' List the Globally Unique Identifier (GUID) for each library referenced in the current project
    Debug.Print "Print Time: " & Time & " :: Item - GUID"
        For Each oRef In oRefs
            i = i + 1
            Debug.Print "Item " & i, oRef.GUID
        Next oRef
    Debug.Print vbNewLine
 
End Sub

Function getRefGUIDs() As Dictionary
    'Here are the references I use. Some are necessary, some may not be.
    'This will need to be run on a workbook that already has the references enabled.
    ' Requires References :: Microsoft Visual Basic for Applications Extensibility 5.3 (GUID {0002E157-0000-0000-C000-000000000046})
    ' C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
    
    'Visual basic for Applications
    'Microsoft Excel 16.0 Object Library
    'OLE Automation
    'Microsoft Office 16.0 Object Library
    'Microsoft ActiveX Data Objects 6.1 Library
    'Microsoft Scripting Runtime
    'Microsoft visual Basic for Applications Extensibility 5.3

    'This code returns all the references currently enabled in this workbook
    
    Dim oRef As VBIDE.Reference   ' Item
    Dim oRefs As VBIDE.References ' Collection
    Dim dict As Object
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    'Super important! Make sure thisworkbook has the correct references enabled. Otherwise the incorrect references will get cloned
    Set oRefs = ThisWorkbook.VBProject.References
    
    ' List the Globally Unique Identifier (GUID) for each library referenced in the current project
    For Each oRef In oRefs
        dict.Add Key:=CStr(oRef.GUID), Item:=CStr(oRef.Description)
    Next oRef
    
    Set getRefGUIDs = dict

End Function
