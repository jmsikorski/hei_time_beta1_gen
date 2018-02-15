Attribute VB_Name = "ExportVisualBasicCode"
' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Sub a()
Application.EnableEvents = True
End Sub
Public Sub callExportVBA()
    ExportVBA ThisWorkbook.name
End Sub

Public Sub ExportVBA(Optional xlFile As String)
    Dim ans As Integer
    Dim mFile As String
    If Environ$("username") <> "jsikorski" Then
    End If
    ans = MsgBox("Export?", vbYesNo, ThisWorkbook.name)
    If ans <> vbYes Then
      Exit Sub
    End If
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24

    Dim VBComponent As VBComponent
    Dim count As Integer
    Dim path As String
    Dim dir_main As String
    Dim extension As String
    Dim FSO As New FileSystemObject
    Dim dirs() As String
    Dim directory As Variant
    Dim xFileDir
    If xlFile <> "" Then
        ans = MsgBox("Publish?", vbYesNo, ThisWorkbook.name)
        If Left(xlFile, 13) = "Time Card Gen" Then
            xFileDir = "Generator"
            mFile = "hei_time_beta1_gen"
        ElseIf Left(xlFile, 8) = "Launcher" Then
            xFileDir = "Installer"
            mFile = "hei_time_beta1_inst"
        Else
            ans = vbNo
        End If
    Else
        ans = vbNo
    End If
    If ans = vbYes Then
        ReDim dirs(3)
        dirs(3) = "C:\Users\jsikorski\Desktop\Time Card BETA\BETA 1\GitHub\"
        If xFileDir = "Generator" Then
            mFile = "hei_time_beta1_gen"
            dirs(3) = dirs(3) & mFile
        ElseIf xFileDir = "Installer" Then
            mFile = "hei_time_beta1_inst"
            dirs(3) = dirs(3) & mFile
        Else
            ReDim dirs(1)
        End If
        dirs(2) = "C:\Users\jsikorski\Desktop\Time Card BETA\BETA 1\Time Card Project\Code\" & xFileDir
    Else
        ReDim dirs(1)
    End If
    dirs(1) = "C:\Users\jsikorski\Documents\VBAProjectFiles\ALL VBA CODE\jms_hei_vba"
    dirs(0) = "C:\Users\jsikorski\Documents\VBAProjectFiles\ALL VBA CODE\jms_hei_vba\" & ThisWorkbook.name & "_code"
    
    count = 0

    For Each directory In dirs
        If Not FSO.FolderExists(directory) Then
            Call FSO.CreateFolder(directory)
        Else
            If directory <> dirs(1) Then clearFolder (directory)
        End If
    Next
    Set FSO = Nothing
    Dim cnt As Integer
    For Each directory In dirs
        cnt = 0
        For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
            cnt = cnt + VBComponent.CodeModule.CountOfLines
            If directory = dirs(0) Then
                Exit For
            End If
            Select Case VBComponent.Type
                Case ClassModule, Document
                    extension = ".cls"
                Case Form
                    extension = ".frm"
                Case Module
                    extension = ".bas"
                Case Else
                    extension = ".txt"
            End Select


            On Error Resume Next
            Err.Clear
            If directory = dirs(1) Then
                If Left(VBComponent.name, 12) = "ThisWorkbook" Or _
                  Left(VBComponent.name, 5) = "Sheet" Then
                    path = dirs(0) & "\" & VBComponent.name & extension
                Else
                    path = dirs(1) & "\" & VBComponent.name & extension
                End If
            Else
                path = directory & "\" & VBComponent.name & extension
            End If
            Call VBComponent.Export(path)
            
            If Err.Number <> 0 Then
                 Call MsgBox("Failed to export " & VBComponent.name & " to " & path, vbCritical)
            Else
                count = count + 1
                Debug.Print "Exported " & Left$(VBComponent.name & ":" & space(Padding), Padding) & path
            End If

            On Error GoTo 0
        Next
    Next
    MsgBox cnt & " lines"
    Application.StatusBar = "Successfully exported " & CStr(count) & " VBA files to " & dir_main
    Application.StatusBar = False
    If ans = vbYes Then
        On Error Resume Next
        Dim commitMessage As String
        commitMessage = InputBox("Commit Message: ", "MESSAGE", Left(ThisWorkbook.name, Len(ThisWorkbook.name) - 5))
        Application.Run "'VBA_IDE.xlsm'!gitCommit", commitMessage, dirs(3)
        On Error GoTo 0
        ans = MsgBox("Release?", vbYesNo, ThisWorkbook.name)
        If ans = vbYes Then
            Zip_All_Files_in_Folder Left(dirs(2), Len(dirs(2)) - Len(xFileDir)), "C:\Users\jsikorski\Helix Electric Inc\TeslaTimeCard - Documents\Time Card Files\Data"
        End If
    End If
End Sub

Public Sub Zip_All_Files_in_Folder(Optional FolderName As String, Optional DefPath)
    Dim FileNameZip
    Dim strDate As String
    Dim oApp As Shell

    If DefPath = vbNullString Then
        DefPath = Application.DefaultFilePath
    End If
    If Right(DefPath, 1) <> "\" Then
        DefPath = DefPath & "\"
    End If

    If FolderName = vbNullString Then
        FolderName = "C:\Users\jsikorski\Desktop\Time Card Project"
    End If

    strDate = Format(Now, "mm.dd.yy.nnssAM")
    FileNameZip = DefPath & "Time Card Project" & ".zip"

    'Create empty Zip File
    NewZip (FileNameZip)
    Set oApp = New Shell
    'Copy the files to the compressed folder
    oApp.Namespace(FileNameZip).CopyHere oApp.Namespace(FolderName).Items

    'Keep script waiting until Compressing is done
    On Error Resume Next
    Do Until oApp.Namespace(FileNameZip).Items.count = _
       oApp.Namespace(FolderName).Items.count
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    On Error GoTo 0

    Debug.Print "You find the zipfile here: " & FileNameZip
End Sub

Public Sub NewZip(sPath)
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub

Public Function bIsBookOpen(ByRef szBookName As String) As Boolean
' Rob Bovey
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function


Public Function Split97(sStr As Variant, sdelim As String) As Variant
'Tom Ogilvy
    Split97 = Evaluate("{""" & _
                       Application.Substitute(sStr, sdelim, """,""") & """}")
End Function

Public Function clearFolder(xFolder As String) As Integer
    Dim FSO As New FileSystemObject
    Set FSO = New FileSystemObject
retry:
    On Error GoTo close_file
    Dim xFile As File
    If Not FSO.FolderExists(xFolder) Then
        Call FSO.CreateFolder(xFolder)
        clearFolder = 1
        GoTo clean_up
    End If
    
    For Each xFile In FSO.GetFolder(xFolder).Files
        xFile.Delete
    Next
    clearFolder = 1
    GoTo clean_up
close_file:
    Err.Clear
    Dim ans As Integer
    ans = MsgBox("Unable to remove " & xFile.name, vbAbortRetryIgnore + vbCritical, "ERROR!")
    If ans = vbRetry Then
        Resume
    ElseIf ans = vbAbort Then
        clearFolder = -2
    Else
        Resume Next
'        MsgBox "Unable to recover, the file will now close", vbCritical & vbOKOnly, "CRITICAL ERROR!"
'        Set FSO = Nothing
'        ThisWorkbook.Close , False
    End If
clean_up:
    Set FSO = Nothing

End Function

Public Sub importDataFile()
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    Dim xFolder As String
    Dim xFile As Object
    Dim objShell As WshShell
    Set objShell = New WshShell
    xFolder = ThisWorkbook.path
    
    For Each xFile In FSO.GetFolder(xFolder).Files
        If FSO.GetExtensionName(xFile.name) = "xlsx" Or FSO.GetExtensionName(xFile.name) = "xlsm" Then
            FSO.CopyFile xFile, ThisWorkbook.Worksheets(1).Range("aPath") & "/" & xFile.name
        End If
    Next
End Sub
Public Sub rebuildFile(rFile As Integer)
Attribute rebuildFile.VB_ProcData.VB_Invoke_Func = "S\n14"
    Dim xlFile As String
    Dim rng As Range
    Dim cFolder As String
    Dim templateName As String
    templateName = " TEMPLATE.xlsm"
    Set objShell = New WshShell
    cFolder = objShell.SpecialFolders("Desktop")
    
    Select Case rFile
        Case 1 ' Rebuild Master File
            xlFile = ThisWorkbook.Worksheets(1).Range("aFile")
            xlFile = Left(xlFile, Len(xlFile) - 5) & templateName
            cFolder = ThisWorkbook.path & "\Generator"
        Case 2 ' Rebuild Builder File
'            xlFile =
'            cFolder = cFolder & "\Time Card Project\Builder"
        Case 3 ' Rebuild Installer File
'            xlFile =
'            cFolder = cFolder & "\Time Card Project\Installer"
        Case Else
            Debug.Print "ERROR REBUILDING FILE"
            Exit Sub
    End Select
    Application.EnableEvents = False
    Workbooks.Open ThisWorkbook.path & "\" & xlFile
    Application.EnableEvents = True
    Workbooks(xlFile).Activate
    ImportModules cFolder
    For Each rng In ThisWorkbook.Worksheets("BUILD").Range("A1", ThisWorkbook.Worksheets("BUILD").Range("A1").End(xlDown))
        AddReference rng.Value, rng.Offset(0, 1).Value
    Next
    Application.DisplayAlerts = False
    Dim newFile As String
    newFile = ThisWorkbook.Worksheets(1).Range("aPath") & "\" & ThisWorkbook.Worksheets(1).Range("aFile")
    ActiveWorkbook.SaveAs newFile
    Application.DisplayAlerts = True

End Sub

Function GetFolder() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .title = "Select code Folder"
        .AllowMultiSelect = False
        .InitialFileName = FolderWithVBAProjectFiles
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Sub Unzip1()
    Dim FSO As Object
    Dim oApp As Object
    Dim Fname As Variant
    Dim FileNameFolder As Variant
    Dim DefPath As String
    Dim strDate As String

    Fname = Application.GetOpenFilename(filefilter:="Zip Files (*.zip), *.zip", MultiSelect:=False)
    If Fname = False Then
        'Do nothing
    Else
        'Root folder for the new folder.
        'You can also use DefPath = "C:\Users\Ron\test\"
        'DefPath = Application.DefaultFilePath
        DefPath = "C:"
        If Right(DefPath, 1) <> "\" Then
            DefPath = DefPath & "\"
        End If

        'Create the folder name
        strDate = Format(Now, " dd-mm-yy h-mm-ss")
        FileNameFolder = DefPath & "MyUnzipFolder " & strDate & "\"

        'Make the normal folder in DefPath
        MkDir FileNameFolder

        'Extract the files into the newly created folder
        Set oApp = CreateObject("Shell.Application")

        oApp.Namespace(FileNameFolder).CopyHere oApp.Namespace(Fname).Items

        'If you want to extract only one file you can use this:
        'oApp.Namespace(FileNameFolder).CopyHere _
         'oApp.Namespace(Fname).items.Item("test.txt")

        MsgBox "You find the files here: " & FileNameFolder

        On Error Resume Next
        Set FSO = CreateObject("scripting.filesystemobject")
        FSO.DeleteFolder Environ("Temp") & "\Temporary Directory*", True
    End If
End Sub

Public Sub ImportModules(Optional codeFolder As String)
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents
    Dim objShell As WshShell
    
    If codeFolder = "" Then
        Set objShell = New WshShell
        codeFolder = objShell.SpecialFolders("Desktop")
        codeFolder = codeFolder & "\Time Card Project\installer"
    End If
    
    If ActiveWorkbook.name = ThisWorkbook.name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles(codeFolder) = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.name
    
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    
'    If wkbTarget.VBProject.Protection = 1 Then
'    MsgBox "The VBA in this workbook is protected," & _
'        "not possible to Import the code"
'    Exit Sub
'    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = codeFolder & "\"
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms
    Dim cnt As Integer
    cnt = 0
    Set cmpComponents = wkbTarget.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.name) = "cls") Then
            If Left(objFile.name, 12) = "ThisWorkbook" Then
                With ActiveWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule
                    .DeleteLines StartLine:=1, count:=.CountOfLines
                    .AddFromFile objFile.path
                    .DeleteLines StartLine:=1, count:=4
                End With
            ElseIf Left(objFile.name, 5) = "Sheet" Then
                On Error Resume Next
                With ActiveWorkbook.VBProject.VBComponents(Left(objFile.name, Len(objFile.name) - 4)).CodeModule
                    .DeleteLines StartLine:=1, count:=.CountOfLines
                    .AddFromFile objFile.path
                    .DeleteLines StartLine:=1, count:=4
                End With
            Else
                cmpComponents.Import objFile.path
                On Error GoTo 0
            End If
        
        ElseIf (objFSO.GetExtensionName(objFile.name) = "frm") Then
            cmpComponents.Import objFile.path
            ThisWorkbook.VBProject.VBComponents(objFile).CodeModule.DeleteLines StartLine:=16, count:=1
        ElseIf (objFSO.GetExtensionName(objFile.name) = "bas") Then
            If objFile.name <> "main_module.bas" Then
                cmpComponents.Import objFile.path
            End If
        End If
        cnt = cnt + 1
    Next objFile
    Set wkbTarget = Nothing
    Set objShell = Nothing
    Set objFSO = Nothing
    Set cmpComponents = Nothing
End Sub

Function FolderWithVBAProjectFiles(Optional xFolder As String) As String
    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String
    
    If xFolder = "" Then
        Set WshShell = CreateObject("WScript.Shell")
        Set FSO = CreateObject("scripting.filesystemobject")
    
        SpecialPath = WshShell.SpecialFolders("DEsktop")
    
        If Right(SpecialPath, 1) <> "\" Then
            SpecialPath = SpecialPath & "\"
        End If
        xFolder = SpecialPath & "Time Card Project - JASON\hei_time\Time Card Gen BETA.xlsm_VBA"
    End If
    
    Set FSO = New FileSystemObject
    If FSO.FolderExists(Left(xFolder, Len(xFolder) - 1)) = False Then
        On Error Resume Next
        MkDir xFolder
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(xFolder) = True Then
        FolderWithVBAProjectFiles = xFolder
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    Set FSO = Nothing
    Set WshShell = Nothing
End Function

Function DeleteVBAModulesAndUserForms(Optional saveModule As String)
        Dim vbProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        
        If saveModule = "" Then
            saveModule = "main_module"
        End If
        Set vbProj = ActiveWorkbook.VBProject
        For Each VBComp In vbProj.VBComponents
            If VBComp.name <> saveModule Then
                If VBComp.Type = vbext_ct_Document Then
                    'Thisworkbook or worksheet module
                    'We do nothing
                Else
                    vbProj.VBComponents.Remove VBComp
                End If
            End If
        Next VBComp
        Set vbProj = Nothing
End Function

Sub AddReference(rName As String, rLoc As String)
    Dim VBAEditor As VBIDE.VBE
    Dim vbProj As VBIDE.VBProject
    Dim chkRef As VBIDE.Reference
    Dim BoolExists As Boolean

    Set VBAEditor = Application.VBE
    Set vbProj = ActiveWorkbook.VBProject

    '~~> Check if rName is already added
    For Each chkRef In vbProj.References
        If chkRef.Description = rName Then
            BoolExists = True
            GoTo CleanUp
        End If
    Next

    vbProj.References.AddFromFile rLoc

CleanUp:
    Set vbProj = Nothing
    Set VBAEditor = Nothing
End Sub
