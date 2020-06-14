VERSION 5.00
Begin VB.Form LogonWindow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Logon"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   ControlBox      =   0   'False
   Icon            =   "Logon.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5.563
   ScaleMode       =   4  'Character
   ScaleWidth      =   33.125
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
   Begin VB.CommandButton LogonButton 
      Caption         =   "&Logon"
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
   Begin VB.TextBox PasswordBox 
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
   Begin VB.TextBox UserBox 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      MaxLength       =   255
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label PasswordLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
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
   Begin VB.Label UserLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "User:"
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
Attribute VB_Name = "LogonWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains the logon window.
Option Explicit

'This procedure closes this window.
Private Sub CancelButton_Click()
On Error GoTo ErrorTrap
   Unload Me
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure initializes this window when it is opened.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   Me.Left = (Screen.Width / 2) - (Me.Width / 2)
   Me.Top = (Screen.Height / 3) - (Me.Height / 2)

   UserLabel.Enabled = (InStr(UCase$(Settings().ConnectionInformation), USER_VARIABLE) > 0)
   PasswordLabel.Enabled = (InStr(UCase$(Settings().ConnectionInformation), PASSWORD_VARIABLE) > 0)
   UserBox.Enabled = UserLabel.Enabled
   PasswordBox.Enabled = PasswordLabel.Enabled
   
   CancelButton.ToolTipText = "Click here to cancel logon and quit this program."
   UserBox.ToolTipText = "Specify a user here, if this is required to connect with the database."
   LogonButton.ToolTipText = Settings().FileName
   PasswordBox.ToolTipText = "Specify a password here, if this is required to connect with the database."
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure closes this window.
Private Sub LogonButton_Click()
On Error GoTo ErrorTrap
   Connection ProcessLogonInformation(UserBox.Text, PasswordBox.Text, Settings().ConnectionInformation)
EndRoutine:
   If ConnectionOpened(Connection()) Then
      Unload Me
   Else
      If UserBox.Enabled Then
         UserBox.SetFocus
      ElseIf PasswordBox.Enabled Then
         PasswordBox.SetFocus
      End If
   End If
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume EndRoutine
End Sub

