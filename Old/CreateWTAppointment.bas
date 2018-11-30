Attribute VB_Name = "CreateWTAppointment"
'This macro creates an appointment without attendees
Public Sub CreateWalkthroughAppointment()
    Dim Ns As Outlook.NameSpace
    Dim objItems As Outlook.Items
    Dim objApp As Outlook.Application
    Dim objAppt As Outlook.AppointmentItem
    Dim objFolder As Outlook.Folder
    Dim Item As Object ' works with any outlook item
    Dim objMail As Outlook.MailItem
    Dim objRecipients As Outlook.Recipients
    Dim objRecipient As Outlook.Recipient
    Dim attendee As Outlook.Recipient
    
    Dim strAddress As String
    Dim myCounter As Integer
    Dim arrUserName() As String
    Dim strUserName As String
    Dim strCcEmails As String
    
    On Error Resume Next
    
    Set objApp = Application
    Set Ns = Application.GetNamespace("MAPI")
    Ns.Logon
    
    'Select active mail item
    Select Case TypeName(objApp.ActiveWindow)
    Case "Explorer"
        Set Item = objApp.ActiveExplorer.Selection.Item(1)
        If (TypeOf Item Is Outlook.MailItem) Then
            Set objMail = Item
        End If
    Case "Inspector"
        Set Item = objApp.ActiveInspector.CurrentItem
        If (TypeOf Item Is Outlook.MailItem) Then
            Set objMail = Item
        End If
    End Select
    
    'Find mail recipients
    Set objRecipients = objMail.Recipients
    
'Edit mail body
    'Body string for editing
    Dim strBody As String
    Dim strSubject As String
    Dim strDate As String
    Dim strTime As String
    Dim dtStart As Date
    'Starting position of email template
    Dim intStart As Integer
    
    'Default
    dtStart = Now
       
    'Remove tabs
    strBody = objMail.Body
    strBody = Replace(strBody, Chr(9) + Chr(13), "")
    strBody = Replace(strBody, Chr(9), Chr(58) + Chr(32))
    'adjust to capture first character and Remove non-table info
    intStart = (Len(strBody)) - (InStr(strBody, "Scheduled Walkthrough Request")) + 1
    strBody = Right(strBody, intStart)
    'adjust to capture first character and Remove footer
    intStart = InStr(strBody, "Thank you very much") - 1
    If intStart > 0 Then
        strBody = Left(strBody, intStart)
    End If
    'Find start of date and capture
    intStart = InStr(strBody, "Scheduled Date") + 18
    If intStart > 0 Then
        strDate = Mid(strBody, intStart, 10)
    End If
    'MsgBox (strDate)
    'Find start of time and capture
    intStart = InStr(strBody, "Scheduled Time") + 23
    If intStart > 0 Then
        strTime = Mid(strBody, intStart, 5)
    End If
    'MsgBox (strTime)
    'Set appointment date/time
    If IsDate(strDate) And IsDate(strTime) Then
        dtStart = CDate(strDate + " " + strTime)
    End If
    'Remove extra characters from subject line
    strSubject = objMail.Subject
    intStart = InStr(strSubject, "Scheduled WT Request")
    If intStart > 0 Then
        strSubject = Mid(strSubject, intStart)
    End If
    
    
    '**************************************************************************************
    'Find shared inbox folder
    Set objFolder = GetFolderPath("LCree Walkthrough Team\Calendar")
    If Not objFolder Is Nothing Then
        Set objItems = objFolder.Items
        Set objAppt = objItems.Add
        If objAppt Is Nothing Then
            Set objAppt = objApp.CreateItem(olAppointmentItem)
        End If
        'Add current user as "location"
        arrUserName = Split(Ns.CurrentUser.Name, ",")
        strUserName = Trim(arrUserName(1)) + " " + Trim(arrUserName(0))
        'Create appointment
        With objAppt
            .Subject = strSubject
            .Attachments.Add objMail
'           .MeetingStatus = olMeeting
'           .Body = objMail.Body
            .Body = strBody
            .Start = dtStart
'           .End = Now
            .Location = strUserName
            .Duration = "15"
            .ReminderSet = True
            .BusyStatus = olBusy
            .ReminderMinutesBeforeStart = "0"
            '.OptionalAttendees = "lcree@west.com"
'        .Save
        .Display 'show to add notes
    End With
    Else
        MsgBox ("Folder not found!")
    
    End If
    
    Set objAppt = Nothing
    'Set objMail = Nothing
    Set objFolder = Nothing
    'Set objOwner = Nothing
    Ns.Logoff
    Set Ns = Nothing
    
    Set objApp = Nothing
End Sub

'Function courtesy of http://www.slipstick.com/developer/working-vba-nondefault-outlook-folders/
Function GetFolderPath(ByVal FolderPath As String) As Outlook.Folder
    Dim oFolder As Outlook.Folder
    Dim FoldersArray As Variant
    Dim i As Integer
         
    On Error GoTo GetFolderPath_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set oFolder = Application.Session.Folders.Item(FoldersArray(0))
    If Not oFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = oFolder.Folders
            Set oFolder = SubFolders.Item(FoldersArray(i))
            If oFolder Is Nothing Then
                Set GetFolderPath = Nothing
            End If
        Next
    End If
    'Return the oFolder
    Set GetFolderPath = oFolder
    Exit Function
         
GetFolderPath_Error:
    Set GetFolderPath = Nothing
    Exit Function
End Function
