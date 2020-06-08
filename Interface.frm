VERSION 5.00
Begin VB.Form InterfaceWindow 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8172
   ClientLeft      =   48
   ClientTop       =   612
   ClientWidth     =   9264
   Icon            =   "Interface.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   681
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   772
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox StatusBox 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   7.2
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
         Size            =   7.8
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
      Begin VB.CommandButton ExportResultButton 
         Caption         =   "Export &Result"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
      Begin VB.CheckBox OpenResultAfterExportBox 
         Caption         =   "&Open result after export."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
      Begin VB.TextBox ExportPathBox 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   3975
      End
      Begin VB.CommandButton SelectExportPathButton 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
      Begin VB.CheckBox CreateEMailWithExportAttachedBox 
         Caption         =   "Create e-&mail with export attached."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
      Begin VB.CheckBox AutomaticallyExportResultBox 
         Caption         =   "&Automatically export result after query."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
      Begin VB.Label ExportResultToLabel 
         Caption         =   "Export result to:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Size            =   7.8
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
      Begin VB.CommandButton ExecuteQueryButton 
         Caption         =   "&Execute Query"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
      Begin VB.TextBox QueryPathBox 
         Height          =   285
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   0
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton SelectQueryButton 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
      Begin VB.CommandButton OpenQueryButton 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
            Size            =   7.8
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
         Begin VB.VScrollBar ParameterFrameScrollBar 
            Height          =   1215
            Left            =   3720
            Max             =   0
            TabIndex        =   4
            Top             =   120
            Width           =   255
         End
         Begin VB.PictureBox ParameterBoxContainer 
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   120
            ScaleHeight     =   81
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   301
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   240
            Width           =   3615
            Begin VB.TextBox ParameterBoxes 
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
                  Size            =   7.8
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
            Size            =   7.8
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
   Begin VB.Frame ResultFrame 
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Begin VB.TextBox QueryResultBox 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   7.2
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
   Begin VB.Menu ProgramMainMenu 
      Caption         =   "&Program"
      Begin VB.Menu InformationMenu 
         Caption         =   "&Information"
         Shortcut        =   ^I
      End
      Begin VB.Menu CloseMenu 
         Caption         =   "&Close"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's interface window.
Option Explicit

'This procedure adjusts the scrollbar so that the specified parameter box becomes visible.
Private Sub AdjustScrollBar(BoxIndex As Long)
On Error GoTo ErrorTrap
Dim Index As Long
Dim Row As Long

   Row = 0
   For Index = ParameterBoxes.LBound To BoxIndex
      If ParameterBoxes(Index).Visible Then Row = Row + 1
   Next Index

EndRoutine:
   ParameterFrameScrollBar.Value = Row - 1
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub


'This procedure adjusts this window to the selected query.
Private Sub AdjustWindow()
On Error GoTo ErrorTrap
Dim BoxIndex As Long
Dim FirstParameter As Long
Dim LastParameter As Long
Dim ParameterIndex As Long
Dim VisibleBoxes As Long

   QueryParameters , , , FirstParameter, LastParameter
   
   If Not (FirstParameter = NO_PARAMETER And LastParameter = NO_PARAMETER) Then
      ParameterIndex = FirstParameter
      VisibleBoxes = 0
      BoxIndex = ParameterBoxes.LBound
      Do While ParameterIndex <= LastParameter
         If BoxIndex > ParameterBoxes.UBound Then
            Load ParameterLabel(BoxIndex)
            Load ParameterBoxes(BoxIndex)
            ParameterLabel(BoxIndex).Top = (VisibleBoxes * (ParameterLabel(BoxIndex).Height * 1.75))
            ParameterBoxes(BoxIndex).Top = ParameterLabel(BoxIndex).Top
         End If

         With QueryParameters(, ParameterIndex)
            ParameterLabel(BoxIndex).Caption = .ParameterName & ":"
            ParameterLabel(BoxIndex).Enabled = True
            ParameterLabel(BoxIndex).ToolTipText = Left$(ParameterLabel(BoxIndex).Caption, Len(ParameterLabel(BoxIndex).Caption) - 1)
            ParameterLabel(BoxIndex).Visible = .InputBoxIsVisible
         
            ParameterBoxes(BoxIndex).Enabled = True
            ParameterBoxes(BoxIndex).Locked = (.Mask = vbNullString)
            ParameterBoxes(BoxIndex).MaxLength = Len(.Mask)
            ParameterBoxes(BoxIndex).TabIndex = (OpenQueryButton.TabIndex + 1) + BoxIndex
            ParameterBoxes(BoxIndex).Text = .DefaultValue & Mid$(.FixedMask, Len(.DefaultValue) + 1)
            If Not Trim$(.Comments) = vbNullString Then ParameterBoxes(BoxIndex).ToolTipText = .Comments
            ParameterBoxes(BoxIndex).Visible = .InputBoxIsVisible
            If ParameterBoxes(BoxIndex).Visible Then
               ParametersFrame.Enabled = True
               VisibleBoxes = VisibleBoxes + 1
            End If
         End With

         ParameterIndex = ParameterIndex + 1
         BoxIndex = BoxIndex + 1
      Loop

      ParameterFrameScrollBar.Enabled = True
      ParameterFrameScrollBar.Max = VisibleBoxes
      ParameterFrameScrollBar.Value = 0
         
      For BoxIndex = ParameterBoxes.LBound To ParameterBoxes.UBound
         If ParameterBoxes(BoxIndex).Visible Then
            ParameterBoxes(BoxIndex).SetFocus
            Exit For
         End If
      Next BoxIndex
   End If
    
   ExecuteQueryButton.Enabled = ((Not (Query().Path = vbNullString)) And ConnectionOpened(Connection()))
   If ExecuteQueryButton.Enabled And VisibleBoxes = 0 Then ExecuteQueryButton.SetFocus
     
   If Not Query().Path = vbNullString Then
      Me.Caption = App.Title & " " & ProgramVersion() & " - " & Query().Path
      DisplayStatus "Query: " & Query().Path & vbCrLf
   End If
   
   QueryPathBox.Text = Unquote(Query().Path)
   QueryPathBox.SelStart = 0
   If Not QueryPathBox.Text = vbNullString Then QueryPathBox.SelStart = Len(QueryPathBox.Text) - 1
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub
'This procedure gives the command to display the specified/already loaded query.
Private Sub DisplayQuery(Optional Path As String = vbNullString)
On Error GoTo ErrorTrap
   If Not BatchModeActive() Then
      If Path = vbNullString Then
         QueryParameters Query().Code
      Else
         QueryParameters Query(Unquote(Path)).Code
      End If
   End If
   
   QueryResults , RemoveResults:=True
   ResetWindow

   AdjustWindow
   
   If Settings().QueryAutoExecute Then GiveQueryCommand
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub


'This procedure gives the command to export the query result.
Private Sub GiveExportCommand()
On Error GoTo ErrorTrap
Dim EMail As EMailClass
Dim ExportPath As String

   ExportPath = ExportPathBox.Text
   If ExportPath = vbNullString Then
      MsgBox "No export path specified.", vbExclamation
   ElseIf Me.Visible Then
      Screen.MousePointer = vbHourglass
      ExportResultButton.Enabled = False
      DisplayStatus "Busy exporting the query result..." & vbCrLf
      
      ExportPath = FileSystemO().GetAbsolutePathName(Unquote(Trim$(ReplaceSymbols(ExportPath))))
      
      If FileSystemO().FolderExists(FileSystemO().GetParentFolderName(ExportPath)) Then
         If ExportResult(ExportPath) Then
            If FileSystemO().FileExists(ExportPath) Then
               If OpenResultAfterExportBox.Value = vbChecked Then
                  DisplayStatus "The export will be opened automatically..." & vbCrLf
                  CheckForAPIError ShellExecuteA(CLng(0), "open", ExportPath, vbNullString, vbNullString, SW_SHOWNORMAL)
               End If
               If CreateEMailWithExportAttachedBox.Value = vbChecked Then
                  DisplayStatus "Busy creating the e-mail containing the export..." & vbCrLf
                  Set EMail = New EMailClass
                  EMail.AddQueryResults ExportPath
                  Set EMail = Nothing
               End If
            End If
            DisplayStatus "Finished export." & vbCrLf
         Else
            DisplayStatus "Export canceled." & vbCrLf
         End If
      Else
         MsgBox "Invalid export path." & vbCr & "Current path: " & CurDir$(), vbExclamation
         DisplayStatus "Invalid export path." & vbCrLf
      End If
   End If
   
EndRoutine:
   ExportResultButton.Enabled = True
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to excute the selected query with the specified parameters.
Private Sub GiveQueryCommand()
On Error GoTo ErrorTrap
   If Not Query().Code = vbNullString Then
      ExecuteQueryButton.Enabled = False
      If ParametersValid(ParameterBoxes) Then
         Screen.MousePointer = vbHourglass
         DisplayStatus "Busy executing the query..." & vbCrLf
         
         QueryResults , RemoveResults:=True
         ExecuteQuery Query().Code
   
         If ConnectionOpened(Connection()) Then
            DisplayQueryResult QueryResultBox, ResultIndex:=0

            If Connection().Errors.Count = 0 Then
               If AutomaticallyExportResultBox.Value = vbChecked Then GiveExportCommand
            Else
               DisplayStatus ErrorListText(Connection().Errors)
            End If

            Connection , , Reset:=True
         End If
      End If
   End If
EndRoutine:
   ExecuteQueryButton.Enabled = ((Not (Query().Path = vbNullString)) And ConnectionOpened(Connection()))
   Screen.MousePointer = vbDefault
   
   If (Settings().QueryAutoClose) Or (Not ProcessSessionList() = vbNullString) Then Unload Me
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure sets this window to interactive batch mode.
Private Sub PutWindowInBatchMode()
On Error GoTo ErrorTrap
   AutomaticallyExportResultBox.Enabled = False
   CreateEMailWithExportAttachedBox.Enabled = False
   ExportFrame.Enabled = False
   ExportResultButton.Enabled = False
   ExportResultToLabel.Enabled = False
   OpenResultAfterExportBox.Enabled = False
   OpenQueryButton.Enabled = False
   QueryLabel.Enabled = False
   QueryPathBox.Enabled = False
   QueryResultBox.Enabled = False
   ResultFrame.Enabled = False
   SelectExportPathButton.Enabled = False
   SelectQueryButton.Enabled = False
   
   ExportPathBox.Text = vbNullString
   ExecuteQueryButton.Caption = "&Execute Batch"
   ExecuteQueryButton.ToolTipText = "Click here to execute the batch with the specified parameters."
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub


'This procedure resets this window.
Private Sub ResetWindow()
On Error GoTo ErrorTrap
Dim Index As Long
   
   AutomaticallyExportResultBox.Enabled = True
   CreateEMailWithExportAttachedBox.Enabled = True
   ExecuteQueryButton.Enabled = False
   ExportFrame.Enabled = True
   ExportResultButton.Enabled = True
   ExportResultToLabel.Enabled = True
   OpenQueryButton.Enabled = True
   OpenResultAfterExportBox.Enabled = True
   ParameterFrameScrollBar.Enabled = False
   ParametersFrame.Enabled = False
   QueryLabel.Enabled = True
   QueryPathBox.Enabled = True
   QueryResultBox.Enabled = True
   ResultFrame.Enabled = True
   SelectExportPathButton.Enabled = True
   SelectQueryButton.Enabled = True
   
   QueryResultBox.Text = vbNullString
   
   AutomaticallyExportResultBox.ToolTipText = "If this box is checked, the query result will be exported to the specified path."
   CreateEMailWithExportAttachedBox.ToolTipText = "If this box is checked, an e-mail with the exported query result will be created."
   ExecuteQueryButton.ToolTipText = "Click here to execute the query with the specified parameters."
   ExportPathBox.ToolTipText = "Here the path to which the query result will exported to can be specified."
   ExportResultButton.ToolTipText = "Click here to export the query result to the specified path."
   OpenQueryButton.ToolTipText = "Click here to the open the specified query file."
   OpenResultAfterExportBox.ToolTipText = "If this box is checked, the exported query result will be opened."
   QueryPathBox.ToolTipText = "Here a query file's path can be specified."
   QueryResultBox.ToolTipText = "The query result is displayed here. Press the Control + Page Up or Page Down keys to browse between multiple query results."
   SelectExportPathButton.ToolTipText = "Click here to open a window to browse to a folder for the export file."
   SelectQueryButton.ToolTipText = "Click here to open a window to browse to a query file."
   StatusBox.ToolTipText = "The status information is displayed here. Right click inside the text for options."
   
   For Index = ParameterBoxes.LBound To ParameterBoxes.UBound
      ParameterLabel(Index).Caption = "Parameter:"
      ParameterLabel(Index).Enabled = False
      ParameterLabel(Index).ToolTipText = ParameterLabel(Index).Caption
      ParameterBoxes(Index).Enabled = False
      ParameterBoxes(Index).Text = vbNullString
      ParameterBoxes(Index).ToolTipText = "Enter a value for the parameter here."
      If Index > ParameterBoxes.LBound Then
         Unload ParameterLabel(Index)
         Unload ParameterBoxes(Index)
      End If
   Next Index
   
   Me.Caption = App.Title & " " & ProgramVersion()
EndRoutine:
   If Settings().BatchInteractive Then PutWindowInBatchMode
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure closes this window.
Private Sub CloseMenu_Click()
On Error GoTo ErrorTrap
   Unload Me
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to execute the selected query with the specified parameters.
Private Sub ExecuteQueryButton_Click()
On Error GoTo ErrorTrap
   If Settings().BatchInteractive Then
      If ParametersValid(ParameterBoxes) Then
         AbortInteractiveBatch AbortBatch:=False
         Me.Enabled = False
         ExecuteQueryButton.Enabled = False
      End If
   Else
      GiveQueryCommand
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub


'This procedure gives the command to the export the query result.
Private Sub ExportResultButton_Click()
On Error GoTo ErrorTrap
   GiveExportCommand
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub



'This procedure gives the command to diplay any query that has been loaded while starting this program.
Private Sub Form_Activate()
On Error GoTo ErrorTrap
   If BatchModeActive() Then
      AdjustWindow
   Else
      If Not Trim$(Command$()) = vbNullString Then DisplayStatus "Command line: " & Command$() & vbCrLf
      If Not ProcessSessionList() = vbNullString Then DisplayStatus "Session list: " & ProcessSessionList() & vbCrLf
      DisplayConnectionStatus
      If Not Query().Path = vbNullString Then DisplayQuery Query().Path
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure initializes this window when it is opened.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   If Not BatchModeActive() And Not CommandLineArguments().QueryPath = vbNullString Then Query CommandLineArguments().QueryPath

   ResetWindow
   DisplayStatus , NewBox:=StatusBox

   With Settings()
      ExportPathBox.Text = .ExportDefaultPath
      
      AutomaticallyExportResultBox.Value = vbUnchecked
      CreateEMailWithExportAttachedBox.Value = vbUnchecked
      OpenResultAfterExportBox.Value = vbUnchecked
      
      If .ExportAutoOpen Then OpenResultAfterExportBox.Value = vbChecked
      If Not .ExportDefaultPath = vbNullString Then AutomaticallyExportResultBox.Value = vbChecked
      If Not (.ExportRecipient = vbNullString And .ExportCCRecipient = vbNullString) Then CreateEMailWithExportAttachedBox.Value = vbChecked
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure closes this window after confirmation from the user.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrorTrap
Dim Choice As Long
   
   With Settings()
      If Not .QueryAutoClose Then
         If UnloadMode = vbFormControlMenu Then
            If (AbortInteractiveBatch() And InteractiveBatchModeActive()) Or Not BatchModeActive() Then
                Choice = MsgBox("Close this program?", vbQuestion Or vbYesNo Or vbDefaultButton2)
               Select Case Choice
                  Case vbNo
                     Cancel = CInt(True)
                  Case vbYes
                     If Not ProcessSessionList() = vbNullString Then AbortSessions NewAbortSessions:=True
                     Cancel = CInt(False)
               End Select
            End If
         End If
      End If
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to display information about this program.
Private Sub InformationMenu_Click()
On Error GoTo ErrorTrap
   DisplayProgramInformation
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to open the query at the specified path.
Private Sub OpenQueryButton_Click()
On Error GoTo ErrorTrap
   If Not QueryPathBox.Text = vbNullString Then DisplayQuery QueryPathBox.Text
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub


'This procedure selects the contents of the activated parameter box.
Private Sub ParameterBoxes_GotFocus(Index As Integer)
On Error GoTo ErrorTrap
   With ParameterBoxes(Index)
      If .Top - .Height < 0 Or .Top > ParameterBoxContainer.ScaleHeight Then AdjustScrollBar CLng(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub


'This procedure filters the user's keystrokes in a parameter box.
Private Sub ParameterBoxes_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap

   If Not Shift = Asc(vbNullChar) Then
      If Not ((KeyCode = vbKeyF4) And (Shift And vbAltMask) = vbAltMask) Then
         KeyCode = Asc(vbNullChar)
         Shift = Asc(vbNullChar)
      End If
   End If
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure handles the user's input in a parameter box.
Private Sub ParameterBoxes_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ErrorTrap
Dim Character As String
Dim CursorPosition As Long
Dim FixedInputCharacter As String
Dim MaskCharacter As String
Dim Text As String

   If Not KeyAscii = vbMenuAccelCtrlC Then
      With ParameterBoxes(Index)
         Character = vbNullString
         Select Case KeyAscii
            Case vbKeyBack
               If .SelStart > 0 Then
                  CursorPosition = .SelStart
                  Character = Mid$(QueryParameters(, CLng(Index)).FixedMask, CursorPosition, 1)
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
                  .Text = QueryParameters(, CLng(Index)).FixedMask
               Else
                  Clipboard.SetText .SelText, vbCFText
                  .SelText = QueryParameters(, CLng(Index)).FixedMask
               End If
            Case Else
               Character = UCase$(Chr$(KeyAscii))
               CursorPosition = .SelStart + 1
               
               With QueryParameters(, CLng(Index))
                  FixedInputCharacter = Mid$(.FixedInput, CursorPosition, 1)
                  MaskCharacter = Mid$(.Mask, CursorPosition, 1)
               End With
   
               If Not ParameterMaskCharacterValid(Character, MaskCharacter, FixedInputCharacter) = vbNullString Then Character = vbNullString
         End Select
      
         If CursorPosition > 0 And CursorPosition <= Len(QueryParameters(, CLng(Index)).Mask) Then
            Text = .Text & Mid$(QueryParameters(, CLng(Index)).FixedMask, Len(.Text) + 1)
            Mid$(Text, CursorPosition, 1) = Character
            .Text = Text
            If KeyAscii = vbKeyBack Then .SelStart = CursorPosition - 1 Else .SelStart = CursorPosition
         End If
       
         KeyAscii = Asc(vbNullChar)
      End With
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure clears the parameter box when the user presses the delete button.
Private Sub ParameterBoxes_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap
   If KeyCode = vbKeyDelete Then
      With QueryParameters(, CLng(Index))
         ParameterBoxes(Index).Text = .DefaultValue & Mid$(.FixedMask, Len(.DefaultValue) + 1)
      End With
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure selects the activated parameter box' contents.
Private Sub ParameterBoxes_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorTrap
   ParameterBoxes(Index).SelStart = 0
   ParameterBoxes(Index).SelLength = Len(ParameterBoxes(Index).Text)
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure moves the parameter boxes when the button on the scrollbar is moved.
Private Sub ParameterFrameScrollBar_Change()
On Error GoTo ErrorTrap
Dim Index As Long
Dim Row As Long
   
   Row = 0
   For Index = ParameterBoxes.LBound To ParameterBoxes.UBound
      ParameterLabel(Index).Top = (Row - ParameterFrameScrollBar.Value) * (ParameterLabel(Index).Height * 1.75)
      ParameterBoxes(Index).Top = ParameterLabel(Index).Top
      If Not QueryParameters(, Index).ParameterName = vbNullString Then Row = Row + 1
   Next Index
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub



'This procedure opens the first of one or more files that have been dragged into the query path box.
Private Sub QueryPathBox_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorTrap
   If Data.Files.Count > 0 Then DisplayQuery Data.Files.Item(1)
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure processes the user's keystrokes in the query result box.
Private Sub QueryResultBox_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap
Dim FirstResult As Long
Dim LastResult As Long
Static ResultIndex As Long

   If (Shift And vbCtrlMask) = vbCtrlMask Then
      QueryResults , , , FirstResult, LastResult

      Select Case KeyCode
         Case vbKeyPageUp
            If ResultIndex > FirstResult Then ResultIndex = ResultIndex - 1
         Case vbKeyPageDown
            If ResultIndex < LastResult Then ResultIndex = ResultIndex + 1
      End Select

      DisplayQueryResult QueryResultBox, ResultIndex
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to request the user to select a query.
Private Sub SelectQueryButton_Click()
On Error GoTo ErrorTrap
Dim QueryPath As String

   QueryPath = RequestQueryPath()
   If Not QueryPath = vbNullString Then DisplayQuery QueryPath
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure gives the command to request the user to specify an export path.
Private Sub SelectExportPathButton_Click()
On Error GoTo ErrorTrap
   ExportPathBox.Text = RequestExportPath(ExportPathBox.Text)
   ExportPathBox.SelStart = 0
   If Not ExportPathBox.Text = vbNullString Then ExportPathBox.SelStart = Len(ExportPathBox.Text) - 1
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(DoNotAskForChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub


