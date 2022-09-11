Attribute VB_Name = "Module1"
Sub ExportEmailToMsg()
' Export E-Mails from an Outlook folder to msg-files
' By Lvnou, 2022
' Inspired by:
'       https://stackoverflow.com/questions/57379087/save-outlook-email-to-my-internal-drive-as-msg-file
'       https://www.extendoffice.com/documents/outlook/5034-outlook-save-multiple-emails-as-msg.html

Dim OlApp As Outlook.Application
Set OlApp = New Outlook.Application
Dim Mailobject As Object
Dim Email As String
Dim NS As NameSpace
Dim Folder As MAPIFolder
Set OlApp = CreateObject("Outlook.Application")

Dim xlObj As Object
Set xlObj = CreateObject("Excel.Application")

Dim fso As Object
Dim fldrname As String
Dim fldrpath As String

Dim mail_num_total As Integer
Dim mail_subj As String
Dim mail_time As Date
Dim mail_name As String

Const invalid_chars As String = "\/:*?<>|[]"""
Const n_chars_filename As Integer = 64

Debug.Print "Set up export"

Set NS = ThisOutlookSession.Session

' Display select folder dialog
Set Folder = NS.PickFolder
Set fso = CreateObject("Scripting.FileSystemObject")

' Choose destination folder
With xlObj.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then
            fldrpath = .SelectedItems(1)
        End If
End With

If fldrpath = "" Then
    Exit Sub
End If


' read all emails from mail items
i = 1
i_failed = 0

Debug.Print "Start export"

mail_num_total = Folder.Items.Count

For Each Mailobject In Folder.Items
    On Error GoTo ErrIter
    
    mail_subj = Mailobject.Subject
    mail_time = Mailobject.ReceivedTime
    mail_name = Format(mail_time, "yyyymmdd", vbUseSystemDayOfWeek, _
          vbUseSystem) & Format(mail_time, "-hhnnss", _
          vbUseSystemDayOfWeek, vbUseSystem) & "-" & mail_subj
    
    ' remove invalid chars from file name
    For ctr = 1 To Len(invalid_chars)
        mail_name = Replace(mail_name, Mid(invalid_chars, ctr, 1), "")
    Next
    
    mail_name = Left$(mail_name, n_chars_filename - 4) & ".msg"
    
    Debug.Print "[Item " & i & "/" & mail_num_total & "]     " & mail_name
    
    Mailobject.SaveAs fldrpath & "\" & mail_name, olMSG
    GoTo NextIteration
    
ErrIter:
    Debug.Print "[Item " & i & "/" & mail_num_total & "]     EXPORT FAILED: " & mail_subj
    i_failed = i_failed + 1
    Resume NextIteration

NextIteration:
    i = i + 1
Next

Debug.Print "Export finished"
Debug.Print "Export failed for "; i_failed & " items"

Set OlApp = Nothing
Set Mailobject = Nothing


End Sub

