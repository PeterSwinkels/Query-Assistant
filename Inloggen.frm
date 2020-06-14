VERSION 5.00
Begin VB.Form InloggenVenster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inloggen"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   ControlBox      =   0   'False
   Icon            =   "Inloggen.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5.563
   ScaleMode       =   4  'Character
   ScaleWidth      =   33.125
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton AnnulerenKnop 
      Cancel          =   -1  'True
      Caption         =   "&Annuleren"
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
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton InloggenKnop 
      Caption         =   "&Inloggen"
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
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox WachtwoordVeld 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   255
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox GebruikerVeld 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      MaxLength       =   255
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label WachtwoordLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Wachtwoord:"
      Enabled         =   0   'False
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
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   1332
   End
   Begin VB.Label GebruikerLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Gebruiker:"
      Enabled         =   0   'False
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
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   1332
   End
End
Attribute VB_Name = "InloggenVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het inlogvenster.
Option Explicit

'Deze procedure sluit dit venster.
Private Sub AnnulerenKnop_Click()
On Error GoTo Fout
   Unload Me
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure stelt dit venster in wanneer het wordt geopend.
Private Sub Form_Load()
On Error GoTo Fout
   Me.Left = (Screen.Width / 2) - (Me.Width / 2)
   Me.Top = (Screen.Height / 3) - (Me.Height / 2)

   GebruikerLabel.Enabled = (InStr(UCase$(Instellingen().VerbindingsInformatie), GEBRUIKER_VARIABEL) > 0)
   WachtwoordLabel.Enabled = (InStr(UCase$(Instellingen().VerbindingsInformatie), WACHTWOORD_VARIABEL) > 0)
   GebruikerVeld.Enabled = GebruikerLabel.Enabled
   WachtwoordVeld.Enabled = WachtwoordLabel.Enabled

   AnnulerenKnop.ToolTipText = "Klik hier om het inloggen af te breken en het programma te beëindigen."
   GebruikerVeld.ToolTipText = "Voer hier een gebruikersnaam in, als deze vereist is voor de verbinding met de database."
   InloggenKnop.ToolTipText = Instellingen().Bestand
   WachtwoordVeld.ToolTipText = "Voer hier een wachtwoord in, als deze vereist is voor de verbinding met de database."

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure sluit dit venster.
Private Sub InloggenKnop_Click()
On Error GoTo Fout
   Verbinding VerwerkInlogGegevens(GebruikerVeld.Text, WachtwoordVeld.Text, Instellingen().VerbindingsInformatie)
EindeProcedure:
   If VerbindingGeopend(Verbinding()) Then
      Unload Me
   Else
      If GebruikerVeld.Enabled Then
         GebruikerVeld.SetFocus
      ElseIf WachtwoordVeld.Enabled Then
         WachtwoordVeld.SetFocus
      End If
   End If
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume EindeProcedure
End Sub

