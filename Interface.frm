VERSION 5.00
Begin VB.Form InterfaceVenster 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9270
   Icon            =   "Interface.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   545
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   618
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox StatusVeld 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   528
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7530
      Width           =   9015
   End
   Begin VB.Frame ExportFrame 
      Caption         =   "Export"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   4440
      TabIndex        =   19
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton ResultaatExporterenKnop 
         Caption         =   "Resultaat &Exporteren"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CheckBox OpenResultaatNaExportVeld 
         Caption         =   "&Open resultaat na export."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   4452
      End
      Begin VB.TextBox ExportPadVeld 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   3975
      End
      Begin VB.CommandButton ExportPadSelecterenKnop 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         Picture         =   "Interface.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox MaakEMailMetExportBijgevoegdVeld 
         Caption         =   "Maak e-&mail met export bijgevoegd."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   4452
      End
      Begin VB.CheckBox AutomatischResultaatExporterenVeld 
         Caption         =   "&Automatisch resultaat exporteren na query."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   4452
      End
      Begin VB.Label ExporteerResultaatNaarLabel 
         Caption         =   "Exporteer resultaat naar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   3972
      End
   End
   Begin VB.Frame QueryFrame 
      Caption         =   "Query"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton QueryUitvoerenKnop 
         Caption         =   "Query &Uitvoeren"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox QueryPadVeld 
         Height          =   285
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   0
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton QuerySelecterenKnop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         Picture         =   "Interface.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton QueryOpenenKnop 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         Picture         =   "Interface.frx":098E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.Frame ParametersFrame 
         Caption         =   "Parameters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   3975
         Begin VB.VScrollBar ParameterFrameSchuifBalk 
            Height          =   1215
            Left            =   3720
            Max             =   0
            TabIndex        =   4
            Top             =   120
            Width           =   255
         End
         Begin VB.PictureBox ParameterVeldHouder 
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   120
            ScaleHeight     =   65
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   241
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   240
            Width           =   3615
            Begin VB.TextBox ParameterVelden 
               Height          =   285
               Index           =   0
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   3
               Top             =   0
               Width           =   2055
            End
            Begin VB.Label ParameterLabel 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Parameter:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Index           =   0
               Left            =   48
               TabIndex        =   18
               Top             =   0
               Width           =   1176
            End
         End
      End
      Begin VB.Label QueryLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Query:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame ResultaatFrame 
      Caption         =   "Resultaat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   9015
      Begin VB.TextBox QueryResultaatVeld 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   12
         Top             =   240
         Width           =   8775
      End
   End
   Begin VB.Menu ProgrammaHoofdMenu 
      Caption         =   "&Programma"
      Begin VB.Menu InformatieMenu 
         Caption         =   "&Informatie"
         Shortcut        =   ^I
      End
      Begin VB.Menu SluitenMenu 
         Caption         =   "&Sluiten"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "InterfaceVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het interfacevenster van dit programma.
Option Explicit


'Deze procedure geeft opdracht om het queryresultaat te exporteren.
Private Sub GeefExportOpdracht()
On Error GoTo Fout
Dim EMail As EMailClass
Dim ExportPad As String

   ExportPad = ExportPadVeld.Text
   If ExportPad = vbNullString Then
      MsgBox "Geen export pad opgegeven.", vbExclamation
   ElseIf Me.Visible Then
      Screen.MousePointer = vbHourglass
      ResultaatExporterenKnop.Enabled = False
      ToonStatus "Bezig met het exporteren van het queryresultaat..." & vbCrLf
   
      ExportPad = BestandsSysteem().GetAbsolutePathName(VerwijderAanhalingsTekens(Trim$(VervangSymbolen(ExportPad))))

      If BestandsSysteem().FolderExists(BestandsSysteem().GetParentFolderName(ExportPad)) Then
         If ExporteerResultaat(ExportPad) Then
            If BestandsSysteem().FileExists(ExportPad) Then
               If OpenResultaatNaExportVeld.Value = vbChecked Then
                  ToonStatus "De export wordt automatisch geopend..." & vbCrLf
                  ControleerOpAPIFout ShellExecuteA(CLng(0), "open", ExportPad, vbNullString, vbNullString, SW_SHOWNORMAL)
               End If
               If MaakEMailMetExportBijgevoegdVeld.Value = vbChecked Then
                  ToonStatus "Bezig met het maken van de e-mail met de export..." & vbCrLf
                  Set EMail = New EMailClass
                  EMail.VoegQueryResultatenToe ExportPad
                  Set EMail = Nothing
               End If
            End If
            ToonStatus "Exporteren gereed." & vbCrLf
         Else
            ToonStatus "Export afgebroken." & vbCrLf
         End If
      Else
         MsgBox "Ongeldig export pad." & vbCr & "Huidig pad: " & CurDir$(), vbExclamation
         ToonStatus "Ongeldig export pad." & vbCrLf
      End If
   End If

EindeProcedure:
   ResultaatExporterenKnop.Enabled = True
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om de geselecteerde query met de opgegeven parameters uit te voeren.
Private Sub GeefQueryOpdracht()
On Error GoTo Fout
   If Not Query().Code = vbNullString Then
      QueryUitvoerenKnop.Enabled = False
      If ParametersGeldig(ParameterVelden) Then
         Screen.MousePointer = vbHourglass
         ToonStatus "Bezig met het uitvoeren van de query..." & vbCrLf
   
         QueryResultaten , ResultatenVerwijderen:=True
         VoerQueryUit Query().Code
   
         If VerbindingGeopend(Verbinding()) Then
            ToonQueryResultaat QueryResultaatVeld, ResultaatIndex:=0

            If Verbinding().Errors.Count = 0 Then
               If AutomatischResultaatExporterenVeld.Value = vbChecked Then GeefExportOpdracht
            Else
               ToonStatus FoutenLijstTekst(Verbinding().Errors)
            End If
   
            Verbinding , , Reset:=True
         End If
      End If
   End If
EindeProcedure:
   QueryUitvoerenKnop.Enabled = ((Not (Query().Pad = vbNullString)) And VerbindingGeopend(Verbinding()))
   Screen.MousePointer = vbDefault

   If (Instellingen().QueryAutoSluiten) Or (Not VerwerkSessieLijst() = vbNullString) Then Unload Me
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure past dit venster aan de geselecteerde query aan.
Private Sub PasVensterAan()
On Error GoTo Fout
Dim EersteParameter As Long
Dim LaatsteParameter As Long
Dim ParameterIndex As Long
Dim VeldIndex As Long
Dim ZichtbareVelden As Long

   QueryParameters , , , EersteParameter, LaatsteParameter

   If Not (EersteParameter = GEEN_PARAMETER And LaatsteParameter = GEEN_PARAMETER) Then
      ParameterIndex = EersteParameter
      ZichtbareVelden = 0
      VeldIndex = ParameterVelden.LBound
      Do While ParameterIndex <= LaatsteParameter
         If VeldIndex > ParameterVelden.UBound Then
            Load ParameterLabel(VeldIndex)
            Load ParameterVelden(VeldIndex)
            ParameterLabel(VeldIndex).Top = (ZichtbareVelden * (ParameterLabel(VeldIndex).Height * 1.75))
            ParameterVelden(VeldIndex).Top = ParameterLabel(VeldIndex).Top
         End If

         With QueryParameters(, ParameterIndex)
            ParameterLabel(VeldIndex).Caption = .ParameterNaam & ":"
            ParameterLabel(VeldIndex).Enabled = True
            ParameterLabel(VeldIndex).ToolTipText = Left$(ParameterLabel(VeldIndex).Caption, Len(ParameterLabel(VeldIndex).Caption) - 1)
            ParameterLabel(VeldIndex).Visible = .VeldIsZichtbaar

            ParameterVelden(VeldIndex).Enabled = True
            ParameterVelden(VeldIndex).Locked = (.Masker = vbNullString)
            ParameterVelden(VeldIndex).MaxLength = Len(.Masker)
            ParameterVelden(VeldIndex).TabIndex = (QueryOpenenKnop.TabIndex + 1) + VeldIndex
            ParameterVelden(VeldIndex).Text = .StandaardWaarde & Mid$(.Masker, Len(.StandaardWaarde) + 1)
            If Not Trim$(.Commentaar) = vbNullString Then ParameterVelden(VeldIndex).ToolTipText = .Commentaar
            ParameterVelden(VeldIndex).Visible = .VeldIsZichtbaar
            If ParameterVelden(VeldIndex).Visible Then
               ParametersFrame.Enabled = True
               ZichtbareVelden = ZichtbareVelden + 1
            End If
         End With

         ParameterIndex = ParameterIndex + 1
         VeldIndex = VeldIndex + 1
      Loop

      ParameterFrameSchuifBalk.Enabled = True
      ParameterFrameSchuifBalk.Max = ZichtbareVelden
      ParameterFrameSchuifBalk.Value = 0

      For VeldIndex = ParameterVelden.LBound To ParameterVelden.UBound
         If ParameterVelden(VeldIndex).Visible Then
            ParameterVelden(VeldIndex).SetFocus
            Exit For
         End If
      Next VeldIndex
   End If

   QueryUitvoerenKnop.Enabled = ((Not (Query().Pad = vbNullString)) And VerbindingGeopend(Verbinding()))
   If QueryUitvoerenKnop.Enabled And ZichtbareVelden = 0 Then QueryUitvoerenKnop.SetFocus

   If Not Query().Pad = vbNullString Then
      Me.Caption = App.Title & " " & ProgrammaVersie() & " - " & Query().Pad
      ToonStatus "Query: " & Query().Pad & vbCrLf
   End If

   QueryPadVeld.Text = VerwijderAanhalingsTekens(Query().Pad)
   QueryPadVeld.SelStart = 0
   If Not QueryPadVeld.Text = vbNullString Then QueryPadVeld.SelStart = Len(QueryPadVeld.Text) - 1

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure reset dit venster.
Private Sub ResetVenster()
On Error GoTo Fout
Dim Index As Long

   AutomatischResultaatExporterenVeld.Enabled = True
   ExporteerResultaatNaarLabel.Enabled = True
   ExportFrame.Enabled = True
   ExportPadSelecterenKnop.Enabled = True
   MaakEMailMetExportBijgevoegdVeld.Enabled = True
   OpenResultaatNaExportVeld.Enabled = True
   ParameterFrameSchuifBalk.Enabled = False
   ParametersFrame.Enabled = False
   QueryLabel.Enabled = True
   QueryOpenenKnop.Enabled = True
   QueryPadVeld.Enabled = True
   QueryResultaatVeld.Enabled = True
   QuerySelecterenKnop.Enabled = True
   QueryUitvoerenKnop.Enabled = False
   ResultaatExporterenKnop.Enabled = True
   ResultaatFrame.Enabled = True

   QueryResultaatVeld.Text = vbNullString

   AutomatischResultaatExporterenVeld.ToolTipText = "Als dit veld is aangevinkt, dan wordt het queryresultaat naar het opgegeven pad geëxporteerd."
   ExportPadSelecterenKnop.ToolTipText = "Klik hier om een venster te openen om naar een map voor het export bestand te bladeren."
   ExportPadVeld.ToolTipText = "Hier kan het pad waar het queryresultaat naar wordt geëxporteerd opgegeven worden."
   MaakEMailMetExportBijgevoegdVeld.ToolTipText = "Als dit veld is aangevinkt, dan wordt een e-mail met het geëxporteerde queryresultaat gemaakt."
   OpenResultaatNaExportVeld.ToolTipText = "Als dit veld is aangevinkt, dan wordt het geëxporteerde resultaat geopend."
   QueryOpenenKnop.ToolTipText = "Klik hier om het opgegeven query bestand te openen."
   QueryPadVeld.ToolTipText = "Hier kan het pad van een querybestand worden opgegeven."
   QueryResultaatVeld.ToolTipText = "Hier wordt het queryresultaat weergegeven. Druk op de toetsen Control + Page Up of Page Down om te bladeren tussen meerdere queryresultaten."
   QuerySelecterenKnop.ToolTipText = "Klik hier om een venster te openen om naar een query bestand te bladeren."
   QueryUitvoerenKnop.ToolTipText = "Klik hier om de query met de opgegeven parameters uit te voeren."
   ResultaatExporterenKnop.ToolTipText = "Klik hier om het queryresultaat te exporteren naar het opgegeven pad."
   StatusVeld.ToolTipText = "Hier wordt de status informatie weergegeven. Klik met de rechtermuisknop in de tekst voor opties."

   For Index = ParameterVelden.LBound To ParameterVelden.UBound
      ParameterLabel(Index).Caption = "Parameter:"
      ParameterLabel(Index).Enabled = False
      ParameterLabel(Index).ToolTipText = ParameterLabel(Index).Caption
      ParameterVelden(Index).Enabled = False
      ParameterVelden(Index).Text = vbNullString
      ParameterVelden(Index).ToolTipText = "Voer hier een waarde in voor de parameter."
      If Index > ParameterVelden.LBound Then
         Unload ParameterLabel(Index)
         Unload ParameterVelden(Index)
      End If
   Next Index

   Me.Caption = App.Title & " " & ProgrammaVersie()
EindeProcedure:
   If Instellingen().BatchInteractief Then ZetVensterInBatchModus
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht het opgegeven/eerder geladen querybestand te tonen.
Private Sub ToonQuery(Optional Pad As String = vbNullString)
On Error GoTo Fout
   If Not BatchModusActief() Then
      If Pad = vbNullString Then
         QueryParameters Query().Code
      Else
         QueryParameters Query(VerwijderAanhalingsTekens(Pad)).Code
      End If
   End If

   QueryResultaten , ResultatenVerwijderen:=True
   ResetVenster

   PasVensterAan

   If Instellingen().QueryAutoUitvoeren Then GeefQueryOpdracht

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure verschuift de schuifbalk zodat het opgegeven parameterveld zichtbaar wordt.
Private Sub VerschuifBalk(VeldIndex As Long)
On Error GoTo Fout
Dim Index As Long
Dim Rij As Long

   Rij = 0
   For Index = ParameterVelden.LBound To VeldIndex
      If ParameterVelden(Index).Visible Then Rij = Rij + 1
   Next Index

EindeProcedure:
   ParameterFrameSchuifBalk.Value = Rij - 1
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure zet dit venster in interactieve batchmodus.
Private Sub ZetVensterInBatchModus()
On Error GoTo Fout
   AutomatischResultaatExporterenVeld.Enabled = False
   ExporteerResultaatNaarLabel.Enabled = False
   ExportFrame.Enabled = False
   ExportPadSelecterenKnop.Enabled = False
   MaakEMailMetExportBijgevoegdVeld.Enabled = False
   OpenResultaatNaExportVeld.Enabled = False
   QueryLabel.Enabled = False
   QueryOpenenKnop.Enabled = False
   QueryPadVeld.Enabled = False
   QueryResultaatVeld.Enabled = False
   QuerySelecterenKnop.Enabled = False
   ResultaatExporterenKnop.Enabled = False
   ResultaatFrame.Enabled = False

   ExportPadVeld.Text = vbNullString
   QueryUitvoerenKnop.Caption = "Batch &Uitvoeren"
   QueryUitvoerenKnop.ToolTipText = "Klik hier om de batch met de opgegeven parameters uit te voeren."
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om de gebruiker te verzoeken een export pad op te geven.
Private Sub ExportPadSelecterenKnop_Click()
On Error GoTo Fout
   ExportPadVeld.Text = VraagExportPad(ExportPadVeld.Text)
   ExportPadVeld.SelStart = 0
   If Not ExportPadVeld.Text = vbNullString Then ExportPadVeld.SelStart = Len(ExportPadVeld.Text) - 1
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om een eventuele bij het starten van dit programma geladen query te tonen.
Private Sub Form_Activate()
On Error GoTo Fout
   If BatchModusActief() Then
      PasVensterAan
   Else
      If Not Trim$(Command$()) = vbNullString Then ToonStatus "Opdrachtregel: " & Command$() & vbCrLf
      If Not VerwerkSessieLijst() = vbNullString Then ToonStatus "Sessie lijst: " & VerwerkSessieLijst() & vbCrLf
      ToonVerbindingsstatus
      If Not Query().Pad = vbNullString Then ToonQuery Query().Pad
   End If

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure stelt dit venster in wanneer het wordt geopend.
Private Sub Form_Load()
On Error GoTo Fout
   If Not BatchModusActief() And Not OpdrachtRegelParameters().QueryPad = vbNullString Then Query OpdrachtRegelParameters().QueryPad

   ResetVenster
   ToonStatus , NieuwVeld:=StatusVeld

   With Instellingen()
      ExportPadVeld.Text = .ExportStandaardPad

      AutomatischResultaatExporterenVeld.Value = vbUnchecked
      MaakEMailMetExportBijgevoegdVeld.Value = vbUnchecked
      OpenResultaatNaExportVeld.Value = vbUnchecked

      If .ExportAutoOpenen Then OpenResultaatNaExportVeld.Value = vbChecked
      If Not .ExportStandaardPad = vbNullString Then AutomatischResultaatExporterenVeld.Value = vbChecked
      If Not (.ExportOntvanger = vbNullString And .ExportCCOntvanger = vbNullString) Then MaakEMailMetExportBijgevoegdVeld.Value = vbChecked
   End With
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure sluit dit programma na bevestiging van de gebruiker.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo Fout
Dim Keuze As Long

   With Instellingen()
      If Not .QueryAutoSluiten Then
         If UnloadMode = vbFormControlMenu Then
            If (InteractieveBatchAfbreken() And InteractieveBatchModusActief()) Or Not BatchModusActief() Then
               Keuze = MsgBox("Dit programma sluiten?", vbQuestion Or vbYesNo Or vbDefaultButton2)
               Select Case Keuze
                  Case vbNo
                     Cancel = CInt(True)
                  Case vbYes
                     If Not VerwerkSessieLijst() = vbNullString Then SessiesAfbreken NieuweSessiesAfbreken:=True
                     Cancel = CInt(False)
               End Select
            End If
         End If
      End If
   End With
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om programmainformatie te tonen.
Private Sub InformatieMenu_Click()
On Error GoTo Fout
   ToonProgrammaInformatie
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure verplaatst de parametervelden wanneer de knop op de schuifbalk verschoven wordt.
Private Sub ParameterFrameSchuifBalk_Change()
On Error GoTo Fout
Dim Index As Long
Dim Rij As Long

   Rij = 0
   For Index = ParameterVelden.LBound To ParameterVelden.UBound
      ParameterLabel(Index).Top = (Rij - ParameterFrameSchuifBalk.Value) * (ParameterLabel(Index).Height * 1.75)
      ParameterVelden(Index).Top = ParameterLabel(Index).Top
      If Not QueryParameters(, Index).ParameterNaam = vbNullString Then Rij = Rij + 1
   Next Index
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure selecteert de inhoud van het geactiveerde parameterveld.
Private Sub ParameterVelden_GotFocus(Index As Integer)
On Error GoTo Fout
   With ParameterVelden(Index)
      If .Top - .Height < 0 Or .Top > ParameterVeldHouder.ScaleHeight Then VerschuifBalk CLng(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure filtert de toetsaanslagen van de gebruiker in een parameterveld.
Private Sub ParameterVelden_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Fout

   If Not Shift = Asc(vbNullChar) Then
      If Not ((KeyCode = vbKeyF4) And (Shift And vbAltMask) = vbAltMask) Then
         KeyCode = Asc(vbNullChar)
         Shift = Asc(vbNullChar)
      End If
   End If

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure verwerkt de invoer van de gebruiker in een parameterveld.
Private Sub ParameterVelden_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Fout
Dim CursorPositie As Long
Dim MaskerTeken As String
Dim Teken As String
Dim Tekst As String

   If Not KeyAscii = vbMenuAccelCtrlC Then
      With ParameterVelden(Index)
         Teken = vbNullString
         Select Case KeyAscii
            Case vbKeyBack
               If .SelStart > 0 Then
                  CursorPositie = .SelStart
                  Teken = Mid$(QueryParameters(, CLng(Index)).Masker, CursorPositie, 1)
               End If
            Case vbMenuAccelCtrlV
               If .SelLength = 0 Then
                  .Text = Clipboard.GetText(vbCFText)
               Else
                  .SelText = Clipboard.GetText(vbCFText)
               End If
            Case vbMenuAccelCtrlX
               If .SelLength = 0 Then
                  Clipboard.SetText .Text, vbCFText
                  .Text = QueryParameters(, CLng(Index)).Masker
               Else
                  Clipboard.SetText .SelText, vbCFText
                  .SelText = QueryParameters(, CLng(Index)).Masker
               End If
            Case Else
               Teken = UCase$(Chr$(KeyAscii))
               CursorPositie = .SelStart + 1
   
               MaskerTeken = Mid$(QueryParameters(, CLng(Index)).Masker, CursorPositie, 1)
               If Not ParameterMaskerTekenGeldig(Teken, MaskerTeken) = vbNullString Then Teken = vbNullString
         End Select
   
         If CursorPositie > 0 And CursorPositie <= Len(QueryParameters(, CLng(Index)).Masker) Then
            Tekst = .Text & Mid$(QueryParameters(, CLng(Index)).Masker, Len(.Text) + 1)
            Mid$(Tekst, CursorPositie, 1) = Teken
            .Text = Tekst
            If KeyAscii = vbKeyBack Then .SelStart = CursorPositie - 1 Else .SelStart = CursorPositie
         End If
   
         KeyAscii = Asc(vbNullChar)
      End With
   End If
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure maakt het parameterveld leeg wanneer de gebruiker de delete knop in drukt.
Private Sub ParameterVelden_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Fout
   If KeyCode = vbKeyDelete Then
      With QueryParameters(, CLng(Index))
         ParameterVelden(Index).Text = .StandaardWaarde & Mid$(.Masker, Len(.StandaardWaarde) + 1)
      End With
   End If
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure selecteert de inhoud van het geactiveerde parameterveld.
Private Sub ParameterVelden_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Fout
   ParameterVelden(Index).SelStart = 0
   ParameterVelden(Index).SelLength = Len(ParameterVelden(Index).Text)
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure opent de eerste van een of meer bestanden die in het querypadveld gesleept worden.
Private Sub QueryPadVeld_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Fout
   If Data.Files.Count > 0 Then ToonQuery Data.Files.Item(1)
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om de query op het opgegeven pad te openen.
Private Sub QueryOpenenKnop_Click()
On Error GoTo Fout
   If Not QueryPadVeld.Text = vbNullString Then ToonQuery QueryPadVeld.Text
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure verwerkt de toetsaanslagen van de gebruiker in het queryresultaat veld.
Private Sub QueryResultaatVeld_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Fout
Dim EersteResultaat As Long
Dim LaatsteResultaat As Long
Static ResultaatIndex As Long

   If (Shift And vbCtrlMask) = vbCtrlMask Then
      QueryResultaten , , , EersteResultaat, LaatsteResultaat

      Select Case KeyCode
         Case vbKeyPageUp
            If ResultaatIndex > EersteResultaat Then ResultaatIndex = ResultaatIndex - 1
         Case vbKeyPageDown
            If ResultaatIndex < LaatsteResultaat Then ResultaatIndex = ResultaatIndex + 1
      End Select

      ToonQueryResultaat QueryResultaatVeld, ResultaatIndex
   End If
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure geeft de opdracht om de gebruiker te verzoeken een query te selecteren.
Private Sub QuerySelecterenKnop_Click()
On Error GoTo Fout
Dim QueryPad As String

   QueryPad = VraagQueryPad()
   If Not QueryPad = vbNullString Then ToonQuery QueryPad
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om de geselecteerde query met de opgegeven parameters uit te voeren.
Private Sub QueryUitvoerenKnop_Click()
On Error GoTo Fout
   If Instellingen().BatchInteractief Then
      If ParametersGeldig(ParameterVelden) Then
         InteractieveBatchAfbreken BatchAfbreken:=False
         Me.Enabled = False
         QueryUitvoerenKnop.Enabled = False
      End If
   Else
      GeefQueryOpdracht
   End If
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om het queryresultaat te exporteren.
Private Sub ResultaatExporterenKnop_Click()
On Error GoTo Fout
   GeefExportOpdracht
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure sluit dit venster.
Private Sub SluitenMenu_Click()
On Error GoTo Fout
   Unload Me
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

