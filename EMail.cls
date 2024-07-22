VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMailClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This module contains the Microsoft Outlook related procedures.
Option Explicit
Private WithEvents EMail As Outlook.MailItem          'Contains a reference to a Microsoft Outlook e-mail message.
Attribute EMail.VB_VarHelpID = -1
Private WithEvents MSOutlook As Outlook.Application   'Contains a reference to Microsoft Outlook.
Attribute MSOutlook.VB_VarHelpID = -1

'This procedure manages the Microsoft Outlook status information.
Private Function OutlookAlreadyActive(Optional NewOutlookAlreadyActive As Variant) As Boolean
On Error GoTo ErrorTrap
Static CurrentOutlookAlreadyActive As Boolean

   If Not IsMissing(NewOutlookAlreadyActive) Then CurrentOutlookAlreadyActive = CBool(NewOutlookAlreadyActive)

EndRoutine:
   OutlookAlreadyActive = CurrentOutlookAlreadyActive
   Exit Function

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure adds the specified exported query results to an e-mail.
Public Sub AddQueryResults(Optional ExportPath As Variant = vbNullString, Optional ExportPaths As Collection = Nothing)
On Error GoTo ErrorTrap

   If Not (EMail Is Nothing Or MSOutlook Is Nothing) Then
      If ExportPaths Is Nothing Then
         EMail.Attachments.Add ExportPath
      Else
         For Each ExportPath In ExportPaths
            If Not ExportPath = vbNullString Then EMail.Attachments.Add ExportPath
         Next ExportPath
      End If
   
      If Settings().ExportAutoSend Then EMail.Send
   End If

EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub


'This procedure initializes this module.
Private Sub Class_Initialize()
On Error GoTo ErrorTrap
   OutlookAlreadyActive NewOutlookAlreadyActive:=False

   Set MSOutlook = New Outlook.Application
   
   If Not MSOutlook Is Nothing Then
      Set EMail = MSOutlook.CreateItem(olMailItem)
      EMail.GetInspector.Activate
   End If

EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub


'This procedure is executed when this module is closed.
Private Sub Class_Terminate()
On Error GoTo ErrorTrap
   Set EMail = Nothing
   Set MSOutlook = Nothing
EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub


'This procedure is executed when a new e-mail is opened.
Private Sub EMail_Open(Cancel As Boolean)
On Error GoTo ErrorTrap
   With Settings()
      If Not EMail Is Nothing Then
         EMail.Body = ReplaceSymbols(.EMailText)
         EMail.CC = .ExportCCRecipient
         EMail.SentOnBehalfOfName = .ExportSender
         EMail.Subject = ReplaceSymbols(.ExportSubject)
         EMail.To = .ExportRecipient
      End If
   End With
EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure is executed when an e-mail is closed.
Private Sub EMail_Unload()
On Error GoTo ErrorTrap
   If Not (Settings().QueryAutoClose Or OutlookAlreadyActive()) Then
      If Not MSOutlook Is Nothing Then
         MSOutlook.GetNamespace("MAPI").Logoff
         MSOutlook.Quit
      End If
   End If
EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub


'This procedure is executed when Microsoft Outlook is started.
Private Sub MSOutlook_Startup()
On Error GoTo ErrorTrap
   OutlookAlreadyActive NewOutlookAlreadyActive:=True
EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub


