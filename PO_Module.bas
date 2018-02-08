Attribute VB_Name = "PO_Module"

Public Sub send_req()
    Dim xSht As Worksheet
    Dim xOutlookObj As Object
    Dim xEmailObj As Object ' Outlook.MailItem
    Dim send_to As String
    Dim xYesorNo As Integer
    Set xSht = ActiveSheet
'GET DEFAULT EMAIL SIGNATURE
    On Error Resume Next
    Dim signature As String
    signature = Environ("appdata") & "\Microsoft\Signatures\"
    If Dir(signature, vbDirectory) <> vbNullString Then
        signature = signature & Dir$(signature & "*.txt")
    Else:
        signature = ""
    End If
    signature = CreateObject("Scripting.FileSystemObject").GetFile(signature).OpenAsTextStream(1, -2).ReadAll
    On Error GoTo 0
    Set xOutlookObj = CreateObject("Outlook.Application")
    Set xEmailObj = xOutlookObj.CreateItem(olMailItem)
    With xEmailObj
        .To = ThisWorkbook.Worksheets("INSTRUCTIONS").Range("email_to")
        .CC = ThisWorkbook.Worksheets("INSTRUCTIONS").Range("email_cc")
        .Subject = job & " PO " & xSht.name
        
        xYesorNo = MsgBox("Preview E-mail?", vbYesNoCancel + vbQuestion, "Send?")
        If xYesorNo = vbYes Then
            .display
        ElseIf xYesorNo = vbCancel Then
            MsgBox "PDF created, but not submitted", vbInformation
        Else
            .Body = "Hello," & vbNewLine & vbNewLine & "Attached is a PO Request for " & job & vbNewLine & vbNewLine & signature
            .Attachments.Add xFolder
            .Send
        End If
    End With
End Sub
