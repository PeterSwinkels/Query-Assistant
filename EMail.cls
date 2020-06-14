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
'Deze module bevat de Microsoft Outlook gerelateerde procedures.
Option Explicit
Private WithEvents EMail As Outlook.MailItem          'Bevat een verwijzing naar een Microsoft Outlook e-mail bericht.
Attribute EMail.VB_VarHelpID = -1
Private WithEvents MSOutlook As Outlook.Application   'Bevat een verwijzing naar Microsoft Outlook.
Attribute MSOutlook.VB_VarHelpID = -1

'Deze procedure beheert de Microsoft Outlook status informatie.
Private Function OutlookReedsActief(Optional NieuweOutlookReedsActief As Variant) As Boolean
On Error GoTo Fout
Static HuidigeOutlookReedsActief As Boolean

   If Not IsMissing(NieuweOutlookReedsActief) Then HuidigeOutlookReedsActief = CBool(NieuweOutlookReedsActief)

EindeProcedure:
   OutlookReedsActief = HuidigeOutlookReedsActief
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function


'Deze procedure voegt de opgegeven geëxporteerde query resultaten toe aan een e-mail.
Public Sub VoegQueryResultatenToe(Optional ExportPad As Variant = vbNullString, Optional ExportPaden As Collection = Nothing)
On Error GoTo Fout

   If Not (EMail Is Nothing Or MSOutlook Is Nothing) Then
      If ExportPaden Is Nothing Then
         EMail.Attachments.Add ExportPad
      Else
         For Each ExportPad In ExportPaden
            If Not ExportPad = vbNullString Then EMail.Attachments.Add ExportPad
         Next ExportPad
      End If
   
      If Instellingen().ExportAutoVerzenden Then EMail.Send
   End If

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure stelt deze module in.
Private Sub Class_Initialize()
On Error GoTo Fout
   OutlookReedsActief NieuweOutlookReedsActief:=False

   Set MSOutlook = New Outlook.Application
   
   If Not MSOutlook Is Nothing Then
      Set EMail = MSOutlook.CreateItem(olMailItem)
      EMail.GetInspector.Activate
   End If

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure wordt uitgevoerd wanneer deze module wordt afgesloten.
Private Sub Class_Terminate()
On Error GoTo Fout
   Set EMail = Nothing
   Set MSOutlook = Nothing
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure wordt uitgevoerd wanneer een nieuwe e-mail wordt geopend.
Private Sub EMail_Open(Cancel As Boolean)
On Error GoTo Fout
   With Instellingen()
      If Not EMail Is Nothing Then
         EMail.Body = VervangSymbolen(.EMailTekst)
         EMail.CC = .ExportCCOntvanger
         EMail.SentOnBehalfOfName = .ExportAfzender
         EMail.Subject = VervangSymbolen(.ExportOnderwerp)
         EMail.To = .ExportOntvanger
      End If
   End With
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure wordt uitgevoerd wanneer een e-mail wordt afgesloten.
Private Sub EMail_Unload()
On Error GoTo Fout
   If Not (Instellingen().QueryAutoSluiten Or OutlookReedsActief()) Then
      If Not MSOutlook Is Nothing Then
         MSOutlook.GetNamespace("MAPI").Logoff
         MSOutlook.Quit
      End If
   End If
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure wordt uitgevoerd wanneer Microsoft Outlook wordt gestart.
Private Sub MSOutlook_Startup()
On Error GoTo Fout
   OutlookReedsActief NieuweOutlookReedsActief:=True
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


