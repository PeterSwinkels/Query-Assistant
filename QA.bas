Attribute VB_Name = "QAModule"
'This module contains this program's main procedures.
Option Explicit

'The Microsoft Windows API constants, functions, and structures used by this program.
Private Type OPENFILENAME
   lStructSize As Long
   hwndOwner As Long
   hInstance As Long
   lpstrFilter As String
   lpstrCustomFilter As String
   nMaxCustomFilter As Long
   nFilterIndex As Long
   lpstrFile As String
   nMaxFile As Long
   lpstrFileTitle As String
   nMaxFileTitle As Long
   lpstrInitialDir As String
   lpstrTitle As String
   flags As Long
   nFileOffset As Integer
   nFileExtension As Integer
   lpstrDefExt As String
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type

Public Const SW_SHOWNORMAL As Long = 1
Private Const ERROR_SUCCESS As Long = 0
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY As Long = &H2000&
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Private Const MAX_STRING As Long = 65535
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_FILEMUSTEXIST  As Long = &H1000&
Private Const OFN_HIDEREADONLY As Long = &H4&
Private Const OFN_LONGNAMES As Long = &H200000
Private Const OFN_NOCHANGEDIR As Long = &H8&
Private Const OFN_PATHMUSTEXIST  As Long = &H800&

Public Declare Function ShellExecuteA Lib "Shell32.dll" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetOpenFileNameA Lib "Comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileNameA Lib "Comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SafeArrayGetDim Lib "Oleaut32.dll" (ByRef saArray() As Any) As Long
Private Declare Function SetCurrentDirectoryA Lib "Kernel32.dll" (ByVal lpPathName As String) As Long
Private Declare Function WaitMessage Lib "User32.dll" () As Long

'The constants, enumerations, and structures used by this program.

'Contains a list of the parameter definition elements.
Private Enum ParameterDefinitionList
   NameElement
   MaskElement
   FixedElement
   DefaultValueElement
   PropertiesElement
   CommentsElement
End Enum

'This structure defines this program's settings.
Public Type SettingsStructure
   BatchInteractive As Boolean        'Indicates whether the user must first specify any parameters before a batch can be executed.
   BatchQueryPath As String           'Contains the path and/or filename without indexes of the query's in a batch to be executed.
   BatchRange As String               'Contains the indexes of the first and last query in a batch to be executed.
   ConnectionInformation As String    'Contains the information required for a connection with a database.
   EMailText As String                'Contains the text of the e-mail with the exported results.
   ExportAutoOpen As Boolean          'Indicates whether an export is opened automatically after having been exported.
   ExportAutoOverwrite As Boolean     'Indicates whether a file is automatically overwritten while exporting the query results.
   ExportAutoSend As Boolean          'Indicates whether the e-mail with the exported results is sent automatically.
   ExportCCRecipient As String        'Contains the e-mail address to which the exported result's copy is sent.
   ExportDefaultPath As String        'Contains the default path for the export of query results.
   ExportPadColumn As Boolean         'Indicates whether the data in a column should be padded with spaces.
   ExportRecipient As String          'Contains the e-mail address to which the exported query results are sent.
   ExportSender As String             'Contains the e-mail containing the exported results' sender.
   ExportSubject As String            'Contains the e-mail with the exported results' subject.
   FileName As String                 'Contains the path and/or filename of the program's settings file.
   QueryAutoClose As Boolean          'Indicates whether this program after executing a query and export is closed automatically.
   QueryAutoExecute As Boolean        'Indicates whether a query is automatically executed after having been loaded.
   QueryRecordSets As Boolean         'Indicates whether the database can return more than one recordset as the result of a query.
   QueryTimeout As Long               'Contains the number of seconds the program will wait for the query result after the command to execute the query has been given.
   PreviewColumnWidth As Long         'Contains the maximum column width used to display the query result in the preview window.
   PreviewLines As Long               'Contains the maximum number of lines that is displayed of the query result in the preview window.
End Type

'This structure defines the parameter information for the selected query.
Public Type QueryParameterStructure
   Comments As String            'The comments for the parameter.
   DefaultValue As String        'The default value for the parameter.
   FixedInput As String          'The parameter definition's fixed input.
   FixedMask As String           'The input mask merged with the fixed input.
   InputBoxIsVisible As Boolean  'Indicates whether the parameter's input box is visible.
   Length As Long                'The parameter definition's length.
   LengthIsVariable As Boolean   'Indicates whether variable length input is allowed.
   Mask As String                'The parameter's input mask.
   ParameterName As String       'The parameter's name.
   Position As Long              'The parameter's position relative to the previous parameter's position.
   Properties As String          'The parameter defintion's properties.
   Value As String               'The user's input.
End Type

'This structure defines any command line arguments specified when starting this program.
Public Type CommandLineArgumentsStructure
   SettingsPath As String     'Contains the specified settings path.
   QueryPath As String        'Contains the specified query path.
   SessionPath As String      'Containsthe  specified session list path.
   Processed As Boolean       'Indicates whether the command line arguments were processed without errors.
End Type

'This structure defines a query.
Public Type QueryStructure
   Code As String             'A query's code.
   Path As String             'A query file's path.
   Opened As Boolean          'Indicates whether a query could be opened.
End Type

'This structure defines a query's result.
Public Type QueryResultStructure
   ColumnWidth() As Long     'Indicates the information's maximum width in bytes for each column.
   RightAligned() As Boolean 'Indicates whether the information will right aligned when displayed.
   Table() As String         'Contains the information retrieved from a database by a query.
End Type

Public Const NO_PARAMETER As Long = -1                      'Stands for "no parameter".
Public Const PASSWORD_VARIABLE As String = "$$PASSWORD$$"   'This variable indicates the password's position when pressent in the connection information.
Public Const USER_VARIABLE As String = "$$USER$$"           'This variable indicates the user name's position when pressent in the connection information.
Private Const ARGUMENT_CHARACTER As String = "?"             'Delimits the command line arguments.
Private Const ASCII_A As Long = 65                           'The ASCII value for the  "A" character.
Private Const ASCII_Z As Long = 90                           'The ASCII value for the  "Z" character.
Private Const COMMENT_CHARACTER As String = "#"              'Indicates that a line in a settings file is a comment.
Private Const CONNECTION_PARAMETER_CHARACTER As String = ";" 'Delimits the connection information parameters.
Private Const DEFINITION_CHARACTERS As String = "$$"         'Indicates where a parameter definition in a query begins and ends.
Private Const ELEMENT_CHARACTER As String = ":"              'Delimits the parameter definition elements.
Private Const EXCEL_MAXIMUM_COLUMN_NUMBER As Long = 255      'The maximum number of columns supported by Microsoft Excel.
Private Const MASK_DIGIT As String = "#"                     'Indicates in a mask that a number is expected as input.
Private Const MASK_FIXED As String = " "                     'Indicates in a mask that the character is fixed input.
Private Const MASK_UPPERCASE As String = "_"                 'Indicates in a mask that a capital letter is expected as input.
Private Const NO_LETTER As Long = 64                         'Stands for "no letter". (The ASCII character that precedes the "A" character.)
Private Const NO_MAXIMUM As Long = -1                        'Stands for "no maximum column width or maximum number of lines in a preview".
Private Const NO_RESULT As Long = -1                         'Stands for "no query result".
Private Const NOT_FIXED As String = "*"                      'Stands for "not a fixed input character".
Private Const PROPERTY_HIDDEN As String = "H"                'Indicates that a parameter's input box is hidden.
Private Const PROPERTY_VARIABLE_LENGTH As String = "V"       'When present at the start of a mask this character indicates that the input length is variable.
Private Const UNKNOWN_NUMBER As Long = -1                    'Stands for "unknown number for the specified dimension in the specified array".
Private Const SECTION_NAME_END As String = "]"               'Indicates the end of a section name in a settings file.
Private Const SECTION_NAME_START As String = "["             'Indicates the start of a section name in a settings file.
Private Const SQL_COMMENT_BLOCK_START As String = "/*"       'Stands for the start of a SQL commentblock.
Private Const SQL_COMMENT_BLOCK_END As String = "*/"         'Stands for the end of a SQL commentblock.
Private Const SQL_COMMENT_LINE_START As String = "--"        'Stands for the start of a SQL comment line.
Private Const SQL_COMMENT_LINE_END As String = vbNullString  'Stands for the end of a SQL comment line.
Private Const STRING_CHARACTERS As String = "'"""            'Stands for the characters that indicate where a string starts and ends.
Private Const SYMBOL_CHARACTER As String = "*"               'Indicates where a symbol starts and ends.
Private Const VALUE_CHARACTER As String = "="                'Delimits a settings parameter's value and name.


'This procedure indicates whether an interactive batch should be aborted.
Public Function AbortInteractiveBatch(Optional AbortBatch As Variant) As Boolean
On Error GoTo ErrorTrap
Static Abort As Boolean
   
   If Not IsMissing(AbortBatch) Then Abort = CBool(AbortBatch)
   
EndRoutine:
   AbortInteractiveBatch = Abort
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure indicates whether a series of sessions has been aborted.
Public Function AbortSessions(Optional NewAbortSessions As Variant) As Boolean
On Error GoTo ErrorTrap
Static CurrentAbortSessions As Boolean
   
   If Not IsMissing(NewAbortSessions) Then CurrentAbortSessions = CBool(NewAbortSessions)

EndRoutine:
   AbortSessions = CurrentAbortSessions
   Exit Function

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure indicates whether the batch mode is active.
Public Function BatchModeActive() As Boolean
On Error GoTo ErrorTrap
EndRoutine:
   With Settings()
      BatchModeActive = Not (.BatchRange = vbNullString Or .BatchQueryPath = vbNullString)
   End With
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure checks whether an error has occurred during the most recent API function call.
Public Function CheckForAPIError(TerugGestuurd As Long, Optional ExtraInformation As String = Empty) As Long
Dim Description As String
Dim ErrorCode As Long
Dim Message As String
Dim Length As Long

   ErrorCode = Err.LastDllError
   Err.Clear
   On Error GoTo ErrorTrap

   If Not ErrorCode = ERROR_SUCCESS Then
      Description = String$(MAX_STRING, vbNullChar)
      Length = FormatMessageA(FORMAT_MESSAGE_ARGUMENT_ARRAY Or FORMAT_MESSAGE_FROM_SYSTEM, CLng(0), ErrorCode, CLng(0), Description, Len(Description), StrPtr(StrConv(ExtraInformation, vbFromUnicode)))
      If Length = 0 Then
         Description = "No description."
      Else
         Description = Left$(Description, Length - 1)
      End If
     
      Message = "API Error code: " & CStr(ErrorCode) & vbCr
      Message = Message & Description
      If Not Right$(Message, 1) = vbCr Then Message = Message & vbCr
      Message = Message & "Return value: " & CStr(TerugGestuurd)
      MsgBox Message, vbExclamation
   End If
EndRoutine:
   CheckForAPIError = TerugGestuurd
   Exit Function

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure closes all windows that might be open.
Public Sub CloseAllWindows()
On Error GoTo ErrorTrap
Dim Window As Form
   
   For Each Window In Forms
      Unload Window
   Next Window
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure closes the workbook at the specified path if it is open in Microsoft Excel.
Private Sub CloseExcelWorkBook(Path As String)
On Error GoTo ErrorTrap
Dim WorkBookO As Excel.Workbook

   Set WorkBookO = GetObject(Path)
   WorkBookO.Close SaveChanges:=False
   Set WorkBookO = Nothing
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False, TypePath:="Path: ", Path:=Path) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub


'This procedure manages the current session's command line arguments.
Public Function CommandLineArguments(Optional SessionParameters As String = vbNullString) As CommandLineArgumentsStructure
On Error GoTo ErrorTrap
Dim Extensions As Collection
Dim Message As String
Dim Parameter As Variant
Dim Parameters() As String
Dim Position As Long
Static CurrentCommandLineArguments As CommandLineArgumentsStructure

   With CurrentCommandLineArguments
      .Processed = True

      If Not SessionParameters = vbNullString Then
         ItemIsUnique Extensions, , ResetList:=True

         Position = InStr(SessionParameters, ARGUMENT_CHARACTER & ARGUMENT_CHARACTER)
         If Position > 0 Then
            .SettingsPath = Mid$(SessionParameters, Position + Len(ARGUMENT_CHARACTER))
         Else
            Parameters = Split(SessionParameters, ARGUMENT_CHARACTER)

            For Each Parameter In Parameters
               If Not Trim$(Parameter) = vbNullString Then
                  Parameter = Unquote(CStr(Parameter))
   
                  If ItemIsUnique(Extensions, "." & LCase$(FileSystemO().GetExtensionName(CStr(Parameter)))) Then
                     Select Case "." & LCase$(FileSystemO().GetExtensionName(CStr(Parameter)))
                        Case ".ini"
                           .SettingsPath = Parameter
                        Case ".lst"
                           .SessionPath = Parameter
                        Case ".qa"
                           .QueryPath = Parameter
                        Case Else
                           If Not Trim$(Parameter) = vbNullString Then
                              Message = "Unrecognized command line argument: """ & Parameter & """."
                              If ProcessSessionList() = vbNullString Then
                                 MsgBox Message, vbExclamation
                              Else
                                 Message = Message & vbCr & "Session list: """ & ProcessSessionList() & """."
                                 If MsgBox(Message, vbExclamation Or vbOKCancel) = vbCancel Then AbortSessions NewAbortSessions:=True
                              End If
                              .Processed = False
                           End If
                     End Select
                  Else
                     Message = "Only one settings file and/or query can be specified at a time."
                     If Not ProcessSessionList() = vbNullString Then Message = Message & vbCr & "Session list: """ & ProcessSessionList() & """."
                     MsgBox Message, vbExclamation
                     .Processed = False
                  End If
               End If
            Next Parameter
         End If
      End If
   End With

EndRoutine:
   CommandLineArguments = CurrentCommandLineArguments
   Exit Function

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function
'This procedure manages the connection with a database.
Public Function Connection(Optional ConnectionInformation As String = vbNullString, Optional CloseConnection As Boolean = False, Optional Reset As Boolean = False) As Adodb.Connection
On Error GoTo ErrorTrap
Static DatabaseConnection As New Adodb.Connection
   
   If Not DatabaseConnection Is Nothing Then
      If Reset Then
         DatabaseConnection.Errors.Clear
      ElseIf Not ConnectionInformation = vbNullString Then
         If Not FormatConnectionInformation(ConnectionInformation) = vbNullString Then DatabaseConnection.Open ConnectionInformation
      ElseIf CloseConnection Then
         If ConnectionOpened(DatabaseConnection) Then
            DatabaseConnection.Close
            Set DatabaseConnection = Nothing
         End If
      End If
   End If
   
EndRoutine:
   Set Connection = DatabaseConnection
   Exit Function
  
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure checks whether the specified connection is open.
Public Function ConnectionOpened(ConnectionO As Adodb.Connection) As Boolean
On Error GoTo ErrorTrap
Dim Opened As Boolean

   If Not ConnectionO Is Nothing Then Opened = (ConnectionO.State = adStateOpen)

EndRoutine:
   ConnectionOpened = Opened
   Exit Function
   
ErrorTrap:
   Opened = False
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure returns this program's default settings.
Private Function DefaultSettings() As SettingsStructure
On Error GoTo ErrorTrap
Dim ProgramSettings As SettingsStructure
   
   With ProgramSettings
      .BatchInteractive = False
      .BatchRange = vbNullString
      .BatchQueryPath = vbNullString
      .ConnectionInformation = vbNullString
      .EMailText = vbNullString
      .ExportAutoOpen = False
      .ExportAutoOverwrite = False
      .ExportAutoSend = False
      .ExportCCRecipient = vbNullString
      .ExportDefaultPath = ".\Export.xls"
      .ExportPadColumn = False
      .ExportRecipient = vbNullString
      .ExportSender = vbNullString
      .ExportSubject = vbNullString
      .FileName = "Qa.ini"
      .PreviewColumnWidth = -1
      .PreviewLines = 10
      .QueryAutoClose = False
      .QueryAutoExecute = False
      .QueryRecordSets = False
      .QueryTimeout = 10
   End With
   
EndRoutine:
   DefaultSettings = ProgramSettings
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function
'This procedure displays the connection status.
Public Sub DisplayConnectionStatus()
On Error GoTo ErrorTrap
   If ConnectionOpened(Connection()) Then
      DisplayStatus "Connected to the database. - Settings: " & Settings().FileName & vbCrLf
   Else
      DisplayStatus "There is no connection to a database." & vbCrLf
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub




'This program displays information about this program.
Public Sub DisplayProgramInformation()
On Error GoTo ErrorTrap
   With App
      MsgBox .Comments, vbInformation, .Title & " " & ProgramVersion() & " - " & "by: " & .CompanyName & ", " & "2009-2016"
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub


'This procedure displays the specified query result.
Public Sub DisplayQueryResult(QueryResultBox As TextBox, ResultIndex As Long)
On Error GoTo ErrorTrap
Dim ResultText As String

   DisplayStatus "Busy creating the query result preview..." & vbCrLf
   ResultText = QueryResultText(QueryResults(, , ResultIndex))
   QueryResultBox.Text = ResultText
   If InterfaceWindow.Visible And Len(QueryResultBox.Text) < Len(ResultText) Then MsgBox "The query result cannot be fully displayed.", vbInformation
EndRoutine:
   DisplayStatus StatusAfterQuery(ResultIndex) & vbCrLf
   Exit Sub

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub



'This procedure displays the specified text in the specified box.
Public Sub DisplayStatus(Optional Text As String = vbNullString, Optional NewBox As TextBox = Nothing)
On Error GoTo ErrorTrap
Dim PreviousLength As Long
Static Box As TextBox

   If Not NewBox Is Nothing Then Set Box = NewBox

   If Not Box Is Nothing Then
      With Box
         PreviousLength = Len(.Text)
         .Text = .Text & Text
         If Len(.Text) < PreviousLength + Len(Text) Then .Text = Text
         .SelStart = Len(.Text) - Len(Text)
         .SelLength = 0
      End With
   End If

   DoEvents
EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure converts the specified error list to text.
Public Function ErrorListText(List As Adodb.Errors) As String
On Error GoTo ErrorTrap
Dim ErrorO As Adodb.Error
Dim Text As String

   Text = vbNullString
   If List.Count = 1 Then Text = Text & "1 error " Else Text = Text & CStr(List.Count) & " errors"
   Text = Text & " occurred while executing the query:" & vbCrLf
   Text = Text & Pad("Native", 11)
   Text = Text & Pad("Code", 11)
   Text = Text & Pad("Source", 36)
   Text = Text & Pad("SQL state", 11)
   Text = Text & "Description" & vbCrLf
   For Each ErrorO In List
      With ErrorO
         Text = Text & Pad(CStr(.NativeError), 10, PadLeft:=True) & " "
         Text = Text & Pad(CStr(.Number), 10, PadLeft:=True) & " "
         Text = Text & Pad(.Source, 35) & " "
         Text = Text & Pad(.SqlState, 10, PadLeft:=True) & " "
         Text = Text & .Description & vbCrLf
      End With
   Next ErrorO

EndRoutine:
   ErrorListText = Text
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure returns the Microsoft Excel column id for the specified column number.
Private Function ExcelColumnId(ByVal Column As Long) As String
On Error GoTo ErrorTrap
Dim ColumnId As String
Dim Letter1 As Long
Dim Letter2 As Long
     
   ColumnId = vbNullString
   If Column > EXCEL_MAXIMUM_COLUMN_NUMBER Then
      ExcelColumnId = vbNullString
      Exit Function
   End If
   
   For Letter1 = NO_LETTER To ASCII_Z
      For Letter2 = ASCII_A To ASCII_Z
         If Column = 0 Then
            If Letter1 = NO_LETTER Then
               ColumnId = Chr$(Letter2)
            Else
               ColumnId = Chr$(Letter1) & Chr$(Letter2)
            End If
            ExcelColumnId = ColumnId
            Exit Function
         End If
         Column = Column - 1
      Next Letter2
   Next Letter1
   
EndRoutine:
   ExcelColumnId = ColumnId
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure executes a quuery batch.
Private Sub ExecuteBatch()
On Error GoTo ErrorTrap
Dim EMail As EMailClass
Dim ErrorInformation As String
Dim ExportPath As String
Dim ExportPaths As New Collection
Dim ExportExecuted As Boolean
Dim FirstParameter As Long
Dim FirstQuery As Long
Dim Index As Long
Dim LastParameter As Long
Dim LastQuery As Long
Dim Position As Long
Dim QueryIndex As Long
Dim QueryPath As String
Dim QueryPathExtension As String

   DisplayConnectionStatus

   With Settings()
      Position = InStr(.BatchRange, "-")
      If Not Position = 0 Then
         FirstQuery = CLng(Val(Trim$(Left$(.BatchRange, Position - 1))))
         LastQuery = CLng(Val(Trim$(Mid$(.BatchRange, Position + 1))))
         QueryPathExtension = "." & FileSystemO().GetExtensionName(.BatchQueryPath)
   
         If CStr(FirstQuery) = Trim$(Left$(.BatchRange, Position - 1)) And CStr(LastQuery) = Trim$(Mid$(.BatchRange, Position + 1)) And FirstQuery <= LastQuery Then
            For QueryIndex = FirstQuery To LastQuery
               QueryPath = Unquote(Left$(.BatchQueryPath, Len(.BatchQueryPath) - Len(QueryPathExtension)) & CStr(QueryIndex) & QueryPathExtension)
   
               If Query(QueryPath).Opened Then
                  QueryParameters Query().Code
   
                  If .BatchInteractive And QueryIndex = FirstQuery Then
                     AbortInteractiveBatch AbortBatch:=True
                     InterfaceWindow.Show

                     If Not Trim$(Command$()) = vbNullString Then DisplayStatus "Command line: " & Command$() & vbCrLf
                     If Not ProcessSessionList() = vbNullString Then DisplayStatus "Session list: " & ProcessSessionList() & vbCrLf
                     DisplayStatus "Query: " & QueryPath & vbCrLf

                     Do While DoEvents() > 0 And InterfaceWindow.Enabled: CheckForAPIError WaitMessage(): Loop
                     If AbortInteractiveBatch() Then Exit Sub
   
                     Screen.MousePointer = vbHourglass
                     InteractiveBatchParameters , , Remove:=True
                     QueryParameters , , , FirstParameter, LastParameter
                     For Index = FirstParameter To LastParameter
                        InteractiveBatchParameters , QueryParameters(, Index).Value
                     Next Index
                  Else
                     If QueryIndex = FirstQuery Then
                        If Not Trim$(Command$()) = vbNullString Then DisplayStatus "Command line: " & Command$() & vbCrLf
                        If Not ProcessSessionList() = vbNullString Then DisplayStatus "Session list: " & ProcessSessionList() & vbCrLf
                     End If

                     DisplayStatus "Query: " & QueryPath & vbCrLf
   
                     QueryParameters , , , FirstParameter, LastParameter
                     For Index = FirstParameter To LastParameter
                        With QueryParameters(, Index)
                           QueryParameters , Index, .DefaultValue & Mid$(.Mask, Len(.DefaultValue) + 1)
                           If Not (.Comments = vbNullString And .FixedInput = vbNullString And .FixedMask = vbNullString And .Mask = vbNullString And .ParameterName = vbNullString And .Properties = vbNullString) Then ParameterSymbolError "Found ignored elements in batch query.", Index
                        End With
                     Next Index
                  End If
                  
                  If InvalidParameterInput(ErrorInformation) = NO_PARAMETER Then
                     DisplayStatus "Busy executing the query..." & vbCrLf
                     QueryResults , RemoveResults:=True
                     ExecuteQuery Query().Code
   
                     If ConnectionOpened(Connection()) Then
                        If Connection().Errors.Count = 0 Then
                           DisplayStatus StatusAfterQuery(ResultIndex:=0) & vbCrLf
                           If Not .ExportDefaultPath = vbNullString Then
                              DisplayStatus "Busy exporting the query result..." & vbCrLf
                              ExportPath = FileSystemO().GetAbsolutePathName(Unquote(Trim$(ReplaceSymbols(.ExportDefaultPath))))
   
                              If FileSystemO().FolderExists(FileSystemO().GetParentFolderName(ExportPath)) Then
                                 ExportPaths.Add ExportPath
                                 ExportExecuted = ExportResult(ExportPath)
                                 If ExportExecuted Then
                                    If FileSystemO().FileExists(ExportPath) And .ExportAutoOpen Then
                                       DisplayStatus "The export will be opened automatically..." & vbCrLf
                                       CheckForAPIError ShellExecuteA(CLng(0), "open", ExportPath, vbNullString, vbNullString, SW_SHOWNORMAL)
                                    End If
            
                                    DisplayStatus "Finished exporting." & vbCrLf
                                 Else
                                    DisplayStatus "Export canceled." & vbCrLf
                                 End If
                              Else
                                 MsgBox "Invalid export path." & vbCr & "Current path: " & CurDir$(), vbExclamation
                                 DisplayStatus "Invalid export path." & vbCrLf
                              End If
                           Else
                              DisplayStatus ErrorListText(Connection().Errors)
                           End If
                        End If
                        
                        Connection , , Reset:=True
                     End If
                  Else
                     ParameterSymbolError "Invalid parameter input: " & ErrorInformation
                  End If
               End If
            Next QueryIndex
            
            If (Not .ExportDefaultPath = vbNullString) And ExportExecuted Then
               If Not (.ExportRecipient = vbNullString And .ExportCCRecipient = vbNullString) Then
                  DisplayStatus "Busy creating the e-mail containing the export..." & vbCrLf
                  Set EMail = New EMailClass
                  EMail.AddQueryResults , ExportPaths
                  Set EMail = Nothing
               End If
            End If
         Else
            MsgBox "Invalid query batch range: """ & .BatchRange & """.", vbExclamation
         End If
      End If
   End With
   
EndRoutine:
   Screen.MousePointer = vbDefault
   Unload InterfaceWindow
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub



'This procedure executes a query on a database.
Public Sub ExecuteQuery(QueryCode As String)
On Error GoTo ErrorTrap
Dim CommandO As New Adodb.Command
Dim Result As Adodb.Recordset
Dim QueryPath As String

   QueryPath = Query().Path

   If ConnectionOpened(Connection()) Then
      Set CommandO.ActiveConnection = Connection()

      If Not CommandO Is Nothing Then
         CommandO.CommandText = FillInParameters(QueryCode)
         CommandO.CommandText = RemoveFormatting(CommandO.CommandText, SQL_COMMENT_LINE_START, SQL_COMMENT_LINE_END, STRING_CHARACTERS)
         CommandO.CommandText = RemoveFormatting(CommandO.CommandText, SQL_COMMENT_BLOCK_START, SQL_COMMENT_BLOCK_END, STRING_CHARACTERS)
         CommandO.CommandTimeout = Settings().QueryTimeout
         CommandO.CommandType = adCmdText

         Set Result = CommandO.Execute
      End If

      Do While ConnectionOpened(Result.ActiveConnection)
         QueryResults Result

         If Settings().QueryRecordSets Then Set Result = Result.NextRecordset Else Exit Do
      Loop
   End If
   
EndRoutine:
   DisplayStatus "Executed query: " & vbCrLf & CommandO.CommandText & vbCrLf

   If Not Result Is Nothing Then
      If ConnectionOpened(Result.ActiveConnection) Then Result.Close
   End If
   
   Set CommandO = Nothing
   Set Result = Nothing
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False, TypePath:="Query: ", Path:=QueryPath) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub
'This procedure exacutes a session using the specified parameters.
Private Sub ExecuteSession(SessionParameters As String)
On Error GoTo ErrorTrap
Static MostRecentConnectionInformation As String

   With CommandLineArguments(SessionParameters)
      If .Processed Then
         If .SettingsPath = vbNullString Then
            Settings FileSystemO().BuildPath(App.Path, DefaultSettings().FileName)
         Else
            Settings .SettingsPath
         End If
         
         With Settings()
            If Not .ConnectionInformation = MostRecentConnectionInformation Then
               Connection , CloseConnection:=True
               If InStr(UCase$(.ConnectionInformation), USER_VARIABLE) > 0 Or InStr(UCase$(.ConnectionInformation), PASSWORD_VARIABLE) > 0 Then
                  LogonWindow.Show vbModal
               Else
                  Connection .ConnectionInformation
               End If
            End If
      
            MostRecentConnectionInformation = .ConnectionInformation
         End With

         If ConnectionOpened(Connection()) Then
            If BatchModeActive() Then
               ExecuteBatch
            Else
               InterfaceWindow.Show
               Do While DoEvents() > 0: CheckForAPIError WaitMessage(): Loop
            End If
        End If
     End If
   End With

EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub




'This procedure exports the query result to a text file.
Private Function ExportAsText(ExportPath As String) As Boolean
On Error GoTo ErrorTrap
Dim Column As Long
Dim ExportAborted As Boolean
Dim FileHandle As Long
Dim FirstResult As Long
Dim LastResult As Long
Dim ResultIndex As Long
Dim Row As Long

   ExportAborted = False
   QueryResults , , , FirstResult, LastResult

   FileHandle = FreeFile()
   Open ExportPath For Output Lock Read Write As FileHandle
      For ResultIndex = FirstResult To LastResult
         With QueryResults(, , ResultIndex)
            If Not CheckForAPIError(SafeArrayGetDim(.Table())) = 0 Then
               If Not LastResult = FirstResult Then Print #FileHandle, "[Result: #" & CStr((ResultIndex - FirstResult) + 1) & "]"
               For Row = LBound(.Table(), 1) To UBound(.Table(), 1)
                  For Column = LBound(.Table(), 2) To UBound(.Table(), 2)
                     If Settings().ExportPadColumn Then
                        Print #FileHandle, Pad(.Table(Row, Column), .ColumnWidth(Column), .RightAligned(Column)) & " ";
                     Else
                        Print #FileHandle, .Table(Row, Column); vbTab;
                     End If
                  Next Column
                  Print #FileHandle,
               Next Row
            End If
         End With
      Next ResultIndex
EndRoutine:
   Close FileHandle

   ExportAsText = ExportAborted
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False, TypePath:="Export path: ", Path:=ExportPath) = vbIgnore Then
      ExportAborted = True
      Resume EndRoutine
   End If
   If HandleError() = vbRetry Then Resume
End Function

'This procedure exports the query result.
Public Function ExportResult(ExportPath As String) As Boolean
On Error GoTo ErrorTrap
Dim ExportAborted As Boolean
Dim FileType As String

   ExportAborted = False
   FileType = "." & LCase$(Trim$(FileSystemO().GetExtensionName(ExportPath)))

   If FileSystemO().FileExists(ExportPath) Then
      If Not Settings().ExportAutoOverwrite Then
         If MsgBox("The file """ & ExportPath & """ already exists. Overwrite?", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbNo Then ExportAborted = True
      End If

      If Not ExportAborted Then
         If FileType = ".xls" Or FileType = ".xlsx" Then CloseExcelWorkBook ExportPath
         Kill ExportPath
      End If
   End If

   If Not ExportAborted Then
      Select Case FileType
         Case ".xls"
            ExportAborted = ExportToExcel(ExportPath, xlWorkbookNormal)
         Case ".xlsx"
            ExportAborted = ExportToExcel(ExportPath, xlWorkbookDefault)
         Case Else
            ExportAborted = ExportAsText(ExportPath)
      End Select
   End If
   
EndRoutine:
   ExportResult = Not ExportAborted
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then
      ExportAborted = True
      Resume EndRoutine
   End If
   If HandleError() = vbRetry Then Resume
End Function


'This procedure exports the query result to a Microsoft Excel workbook.
Private Function ExportToExcel(ExportPath As String, ExcelFormat As Long) As Boolean
On Error GoTo ErrorTrap
Dim Column As Long
Dim ColumnId As String
Dim ExportAborted As Boolean
Dim FirstResult As Long
Dim LastResult As Long
Dim Message As String
Dim MSExcel As New Excel.Application
Dim ResultIndex As Long
Dim WorkBookO As Excel.Workbook
Dim WorkSheetO As Excel.Worksheet

   ExportAborted = False
   QueryResults , , , FirstResult, LastResult

   MSExcel.DisplayAlerts = False
   MSExcel.Interactive = False
   MSExcel.ScreenUpdating = False
   MSExcel.Workbooks.Add

   Set WorkBookO = MSExcel.Workbooks.Item(1)
   WorkBookO.Activate

   Do Until WorkBookO.Worksheets.Count <= 1
      WorkBookO.Worksheets.Item(WorkBookO.Worksheets.Count).Delete
   Loop

   Do Until WorkBookO.Worksheets.Count >= Abs(LastResult - FirstResult) + 1
      WorkBookO.Worksheets.Add
   Loop

   For ResultIndex = FirstResult To LastResult
      With QueryResults(, , ResultIndex)
         If Not CheckForAPIError(SafeArrayGetDim(.Table())) = 0 Then
            If NumberOfItems(.Table, Dimension:=2) > EXCEL_MAXIMUM_COLUMN_NUMBER Then
               Message = "The query result contains too many columns to export these to Microsoft Excel." & vbCr
               Message = Message & "The maximum allowed number of columns is: " & CStr(EXCEL_MAXIMUM_COLUMN_NUMBER)
               MsgBox Message, vbExclamation
            Else
               Set WorkSheetO = WorkBookO.Worksheets.Item((ResultIndex - FirstResult) + 1)
               WorkSheetO.Activate
               If Not LastResult = FirstResult Then WorkSheetO.Name = "Result " & CStr((ResultIndex - FirstResult) + 1)
   
               WorkSheetO.Range("A1:" & ExcelColumnId(NumberOfItems(.Table(), Dimension:=2)) & CStr(NumberOfItems(.Table(), Dimension:=1) + 1)).Value = .Table()
               For Column = LBound(.Table(), 2) To UBound(.Table(), 2)
                  ColumnId = ExcelColumnId(Column)
                  WorkSheetO.Range(ColumnId & "1:" & ColumnId & "1").Font.Bold = True
                  If .RightAligned(Column) Then WorkSheetO.Range(ColumnId & "1:" & ColumnId & CStr(NumberOfItems(.Table(), Dimension:=1) + 1)).HorizontalAlignment = xlRight
               Next Column
               WorkSheetO.Range("A:" & ExcelColumnId(NumberOfItems(.Table(), Dimension:=2))).Columns.AutoFit
            End If
         End If
      End With

      If ResultIndex = LastResult Then
         WorkBookO.Worksheets.Item(1).Activate
         WorkBookO.SaveAs ExportPath, ExcelFormat
         WorkBookO.Close
      End If
   Next ResultIndex

EndRoutine:
   If Not MSExcel Is Nothing Then
      MSExcel.Quit
      MSExcel.DisplayAlerts = True
      MSExcel.Interactive = True
      MSExcel.ScreenUpdating = True
   End If

   Set MSExcel = Nothing
   Set WorkSheetO = Nothing
   Set WorkBookO = Nothing

   ExportToExcel = ExportAborted
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False, TypePath:="Export path: ", Path:=ExportPath) = vbIgnore Then
      ExportAborted = True
      Resume EndRoutine
   End If
   If HandleError() = vbRetry Then Resume
End Function

'This procedure manages the file system related functions.
Public Function FileSystemO() As FileSystemObject
On Error GoTo ErrorTrap
Static CurrentFileSystem As New FileSystemObject
EndRoutine:
   Set FileSystemO = CurrentFileSystem
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure fills the specified query code with the parameter input.
Private Function FillInParameters(QueryCode As String) As String
On Error GoTo ErrorTrap
Dim FirstParameter As Long
Dim Index As Long
Dim LastParameter As Long
Dim Position As Long
Dim QueryWithoutParameters As String
Dim QueryWithParameters As String

   QueryParameters , , , FirstParameter, LastParameter
   If FirstParameter = NO_PARAMETER And LastParameter = NO_PARAMETER Then
      QueryWithParameters = QueryCode
   Else
      Index = FirstParameter
      QueryWithParameters = vbNullString
      QueryWithoutParameters = QueryCode
      Do Until Index > LastParameter
         SessionParameters , QueryParameters(, Index).Value

         Position = QueryParameters(, Index).Position
         QueryWithParameters = QueryWithParameters & Left$(QueryWithoutParameters, Position - 1)
         QueryWithParameters = QueryWithParameters & ReplaceSymbols(QueryParameters(, Index).Value)
         QueryWithoutParameters = Mid$(QueryWithoutParameters, Position + QueryParameters(, Index).Length)
         Index = Index + 1
      Loop
      QueryWithParameters = QueryWithParameters & QueryWithoutParameters
   End If
EndRoutine:
   FillInParameters = QueryWithParameters
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure checks the specified connection information and formats it.
Private Function FormatConnectionInformation(ConnectionInformation As String) As String
On Error GoTo ErrorTrap
Dim Character As String
Dim CurrentStringCharacter As String
Dim FormattedConnectionInformation As String
Dim Parameter As String
Dim ParameterName As String
Dim ParameterNames As Collection
Dim ParameterStart As Long
Dim Position As Long
Dim Value As String

   FormattedConnectionInformation = vbNullString
  
   If Not Trim$(ConnectionInformation) = vbNullString Then
      CurrentStringCharacter = vbNullString
      ItemIsUnique ParameterNames, , ResetList:=True
      Position = 1
      ParameterStart = Position
      If Not Right$(Trim$(ConnectionInformation), Len(CONNECTION_PARAMETER_CHARACTER)) = CONNECTION_PARAMETER_CHARACTER Then ConnectionInformation = ConnectionInformation & CONNECTION_PARAMETER_CHARACTER
      Do Until Position > Len(ConnectionInformation)
         Character = Mid$(ConnectionInformation, Position, 1)
         If InStr(STRING_CHARACTERS, Character) > 0 Then
            If CurrentStringCharacter = vbNullString Then
               CurrentStringCharacter = Character
            ElseIf Character = CurrentStringCharacter Then
               CurrentStringCharacter = vbNullString
            End If
         ElseIf Character = CONNECTION_PARAMETER_CHARACTER Then
            If CurrentStringCharacter = vbNullString Then
               Parameter = Mid$(ConnectionInformation, ParameterStart, Position - ParameterStart)

               If InStr(Parameter, VALUE_CHARACTER) = 0 Then
                  FormattedConnectionInformation = vbNullString
                  MsgBox "Invalid parameter present in connection information: """ & Parameter & """. Expected character: " & VALUE_CHARACTER, vbExclamation
                  Exit Do
               End If

               Value = ReadParameter(Parameter, ParameterName)

               If Not ItemIsUnique(ParameterNames, ParameterName) Then
                  FormattedConnectionInformation = vbNullString
                  MsgBox "Parameter present multiple times in connection information: """ & Parameter & """.", vbExclamation
                  Exit Do
               End If
               ParameterStart = Position + 1

               FormattedConnectionInformation = FormattedConnectionInformation & ParameterName & VALUE_CHARACTER & Trim$(Value) & CONNECTION_PARAMETER_CHARACTER
            End If
         End If
   
         Position = Position + 1
      Loop
   
      If Not CurrentStringCharacter = vbNullString Then
         FormattedConnectionInformation = vbNullString
         MsgBox "Unclosed string in connection information. Expected character: " & CurrentStringCharacter, vbExclamation
      End If
   End If

EndRoutine:
   FormatConnectionInformation = FormattedConnectionInformation
   Exit Function

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure formats the specified error description.
Private Function FormatErrorDescription(ErrorDescription As String) As String
On Error Resume Next
Dim Description As String

   Description = Trim$(ErrorDescription)
   Do
      Select Case Right$(Description, 1)
         Case vbCr, vbLf
            Description = Trim$(Left$(Description, Len(Description) - 1))
         Case Else
            Exit Do
      End Select
   Loop
   If Not Right$(Description, 1) = "." Then Description = Description & "."

FormatErrorDescription = Description
End Function

'This procedure generates an input mask that is merged with the fixed input for the specified parameter.
Private Function GenerateFixedMask(Index As Long) As String
On Error GoTo ErrorTrap
Dim Character As String
Dim FixedMask As String
Dim Position As Long

With QueryParameters(, Index)
   FixedMask = vbNullString
   For Position = 1 To Len(.FixedInput)
      Character = Mid$(.FixedInput, Position, 1)
      If Character = NOT_FIXED Then
         FixedMask = FixedMask & Mid$(.Mask, Position, 1)
      Else
         FixedMask = FixedMask & Character
      End If
   Next Position
End With

EndRoutine:
   GenerateFixedMask = FixedMask
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure handles any errors that occur.
Public Function HandleError(Optional ReturnPreviousChoice As Boolean = True, Optional TypePath As String = vbNullString, Optional Path As String = vbNullString, Optional ExtraInformation As String = vbNullString) As Long
Dim ErrorCode As Long
Dim ErrorDescription As String
Dim Message As String
Dim Source As String
Static Choice As Long
   
   Source = Err.Source
   ErrorCode = Err.Number
   ErrorDescription = Err.Description
   Err.Clear

   On Error Resume Next

   If Not ReturnPreviousChoice Then
      Message = FormatErrorDescription(ErrorDescription) & vbCr
      Message = Message & "Error code: " & CStr(ErrorCode)
      If Not Source = vbNullString Then Message = Message & vbCr & "Source: " & Source
      If Not (TypePath = vbNullString Or Path = vbNullString) Then Message = Message & vbCr & TypePath & FileSystemO().GetAbsolutePathName(Path)
      If Not ExtraInformation = vbNullString Then Message = Message & vbCr & ExtraInformation

      Choice = MsgBox(Message, vbExclamation Or vbAbortRetryIgnore Or vbDefaultButton2)
   End If

   HandleError = Choice

   If Choice = vbAbort Then End
End Function

'This procedure indicates whether the interactive batch mode is active.
Public Function InteractiveBatchModeActive() As Boolean
On Error GoTo ErrorTrap
EndRoutine:
   InteractiveBatchModeActive = Settings().BatchInteractive And BatchModeActive()
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure manages the interactive batch parameters.
Private Function InteractiveBatchParameters(Optional Index As Long = 0, Optional NewParameter As Variant, Optional Remove As Boolean = False) As String
On Error GoTo ErrorTrap
Static Parameters As New Collection
   
   If Not IsMissing(NewParameter) Then
      Parameters.Add CStr(NewParameter)
   ElseIf Remove Then
      Set Parameters = New Collection
   End If
   
EndRoutine:
   If Parameters.Count = 0 Then
      InteractiveBatchParameters = vbNullString
   Else
      InteractiveBatchParameters = Parameters(Index + 1)
   End If
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure checks the query parameter input and returns the index of any incorrectly filled in inputbox and an error description.
Private Function InvalidParameterInput(Optional ByRef ErrorInformation As String = vbNullString) As Long
On Error GoTo ErrorTrap
Dim FirstParameter As Long
Dim Index As Long
Dim InvalidBox As Long
Dim LastParameter As Long
Dim Length As Long
Dim Position As Long

   QueryParameters , , , FirstParameter, LastParameter
   InvalidBox = NO_PARAMETER
   
   For Index = FirstParameter To LastParameter
      With QueryParameters(, Index)
         If .Mask = vbNullString Then
            Length = Len(.Value)
         Else
            If .LengthIsVariable Then Length = ParameterInputLength(Index) Else Length = Len(.Mask)
            For Position = 1 To Length
               ErrorInformation = ParameterMaskCharacterValid(Mid$(.Value, Position, 1), Mid$(.Mask, Position, 1), Mid$(.FixedInput, Position, 1))
               If Not ErrorInformation = vbNullString Then
                  ErrorInformation = vbCr & """" & ErrorInformation & """." & vbCr & "Character position: " & CStr(Position)
                  InvalidBox = Index
                  Exit For
               End If
            Next Position
         End If

         If Not InvalidBox = NO_PARAMETER Then Exit For
         QueryParameters , Index, Left$(.Value, Length)
      End With
   Next Index
   
   If Not InvalidBox = NO_PARAMETER Then
      For Index = FirstParameter To LastParameter
         QueryParameters , Index, vbNullString
      Next Index
   End If
   
EndRoutine:
   InvalidParameterInput = InvalidBox
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure indicates whether the specified datatype should be left aligned.
Private Function IsLeftAligned(DataType As Long) As Boolean
On Error GoTo ErrorTrap
Dim LeftAligned As Boolean
Dim TypeIndex As Long

   LeftAligned = False
   For TypeIndex = LBound(LeftAlignedDataTypes()) To UBound(LeftAlignedDataTypes())
      If DataType = LeftAlignedDataTypes(TypeIndex) Then
         LeftAligned = True
         Exit For
      End If
   Next TypeIndex
EndRoutine:
   IsLeftAligned = LeftAligned
   Exit Function

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function
'This procedure indicates whether the specified line of text contains a settings section name.
Private Function IsSettingsSection(Line As String) As Boolean
On Error GoTo ErrorTrap
EndRoutine:
   IsSettingsSection = (Left$(Trim$(Line), 1) = SECTION_NAME_START And Right$(Trim$(Line), 1) = SECTION_NAME_END)
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure adds the specified item if it does not yet appear on the specified list.
Private Function ItemIsUnique(ByRef List As Collection, Optional Item As Variant, Optional ResetList As Boolean = False) As Boolean
On Error GoTo ErrorTrap
Dim Index As Long
Dim Unique As Boolean

   Unique = True

   If ResetList Then
      Set List = New Collection
   ElseIf Not IsMissing(Item) Then
      For Index = 1 To List.Count
         If List(Index) = Item Then
            Unique = False
            Exit For
         End If
      Next Index

      If Unique Then List.Add Item
   End If

EndRoutine:
   ItemIsUnique = Unique
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure returns a list of database datatypes that are left aligned.
Private Function LeftAlignedDataTypes() As Variant
On Error GoTo ErrorTrap
EndRoutine:
   LeftAlignedDataTypes = Array(adBSTR, adChar, adDBDate, adDBTime, adDBTimeStamp, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar)
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure loads this program's settings.
Private Function LoadSettings(SettingsPath As String) As SettingsStructure
On Error GoTo ErrorTrap
Dim Abort As Boolean
Dim ConnectionInformation As String
Dim FileHandle As Long
Dim Line As String
Dim MostRecentValidSection As String
Dim ParameterName As String
Dim ProcessedParameters As New Collection
Dim ProcessedSections As New Collection
Dim ProgramSettings As SettingsStructure
Dim Section As String

   Abort = False
   ItemIsUnique ProcessedParameters, , ResetList:=True
   ItemIsUnique ProcessedSections, , ResetList:=True
   ProgramSettings = DefaultSettings()
   MostRecentValidSection = vbNullString
   Section = vbNullString
   
   With ProgramSettings
      .FileName = SettingsPath
      FileHandle = FreeFile()
      Open .FileName For Input Lock Read Write As FileHandle
         Do Until EOF(FileHandle) Or Abort
            Line Input #FileHandle, Line
            
            If Not Left$(Trim$(Line), 1) = COMMENT_CHARACTER Then
               If IsSettingsSection(Line) Then
                  Line = Trim$(Line)
                  MostRecentValidSection = Section
                  Section = UCase$(Mid$(Line, Len(SECTION_NAME_START) + 1, Len(Line) - (Len(SECTION_NAME_START) + Len(SECTION_NAME_END))))
                  If Not ItemIsUnique(ProcessedSections, Section) Then
                     If SettingsError("Multiple instances of the same section have been found.", SettingsPath, Section, Line) = vbCancel Then Abort = True
                  End If
                  ItemIsUnique ProcessedParameters, , ResetList:=True
               Else
                  Select Case Section
                     Case "BATCH", "EXPORT", "PREVIEW", "QUERY"
                        If Not Trim$(Line) = vbNullString Then
                           ReadParameter Line, ParameterName
                           If Not ItemIsUnique(ProcessedParameters, ParameterName) Then
                              If SettingsError("Multiple instances of the parameter section have been found.", SettingsPath, Section, Line) = vbCancel Then Abort = True
                           End If
                        End If
                  End Select
               End If

               Select Case Section
                  Case "BATCH"
                     If Not (IsSettingsSection(Line) Or Trim$(Line) = vbNullString) Then
                        If Not ProcessBatchSettings(Line, Section, ProgramSettings) Then Abort = True
                     End If
                  Case "CONNECTION"
                     If Not (IsSettingsSection(Line) Or Trim$(Line) = vbNullString) Then .ConnectionInformation = .ConnectionInformation & Trim$(Line)
                  Case "EMAILTEXT"
                     If Not IsSettingsSection(Line) Then .EMailText = .EMailText & Line & vbCrLf
                  Case "EXPORT"
                     If Not (IsSettingsSection(Line) Or Trim$(Line) = vbNullString) Then
                        If Not ProcessExportSettings(Line, Section, ProgramSettings) Then Abort = True
                     End If
                  Case "PREVIEW"
                     If Not (IsSettingsSection(Line) Or Trim$(Line) = vbNullString) Then
                       If Not ProcessPreviewSettings(Line, Section, ProgramSettings) Then Abort = True
                     End If
                  Case "QUERY"
                     If Not (IsSettingsSection(Line) Or Trim$(Line) = vbNullString) Then
                       If Not ProcessQuerySettings(Line, Section, ProgramSettings) Then Abort = True
                     End If
                  Case Else
                     If Not Trim$(Line) = vbNullString Then
                        If IsSettingsSection(Line) Then
                           Section = MostRecentValidSection
                           If SettingsError("Unrecognized section.", SettingsPath, Section, Line) = vbCancel Then Abort = True
                        Else
                           If SettingsError("Unrecognized parameter.", SettingsPath, Section, Line) = vbCancel Then Abort = True
                        End If
                     End If
               End Select
            End If
         Loop
      Close FileHandle

      If Trim$(.ConnectionInformation) = vbNullString And Not Abort Then
         ConnectionInformation = Trim$(RequestConnectionInformation())
         If Not ConnectionInformation = vbNullString Then
            .ConnectionInformation = ConnectionInformation
            SaveSettings SettingsPath, ProgramSettings, "The settings have been written to:"
         End If
      End If

      .ConnectionInformation = FormatConnectionInformation(.ConnectionInformation)
   End With

EndRoutine:
   Close FileHandle
     
   LoadSettings = ProgramSettings
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False, TypePath:="Settings file: ", Path:=SettingsPath) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function
'This procedure is executed when this program is started.
Private Sub Main()
On Error GoTo ErrorTrap
Dim SettingsPath As String
   CheckForAPIError SetCurrentDirectoryA(Left$(App.Path, InStr(App.Path, ":")))
   ChDir App.Path

   With CommandLineArguments(Command$())
      If .Processed Then
         If Left$(Trim$(.SettingsPath), Len(ARGUMENT_CHARACTER)) = ARGUMENT_CHARACTER Then
            SettingsPath = Unquote(Mid$(Trim$(.SettingsPath), Len(ARGUMENT_CHARACTER) + 1))
            If SettingsPath = vbNullString Then
               MsgBox "Cannot save the settings. No target file specified.", vbExclamation
            Else
               SaveSettings SettingsPath, DefaultSettings(), "The default settings have been written to:"
            End If
         ElseIf Not .SessionPath = vbNullString Then
            SessionParameters , , Remove:=True
            ProcessSessionList .SessionPath
         Else
            SessionParameters , , Remove:=True
            ExecuteSession Command$()
         End If
      End If
   End With

EndRoutine:
   Connection , CloseConnection:=True
   CloseAllWindows
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure checks whether the specified parameter's input mask contains invalid characters.
Private Function MaskValid(Index As Long) As Boolean
On Error GoTo ErrorTrap
Dim Character As String
Dim IsValid As Boolean
Dim FixedInput As String
Dim Mask As String
Dim Position As Long

   IsValid = True
   FixedInput = QueryParameters(, Index).FixedInput
   Mask = QueryParameters(, Index).Mask
   For Position = 1 To Len(Mask)
      Character = Mid$(Mask, Position, 1)
      If Character = MASK_FIXED Then
         If Mid$(FixedInput, Position, 1) = NOT_FIXED Then
            ParameterSymbolError "Fixed character indicated in mask but not in fixed characters at: #" & CStr(Position) & ".", Index
            IsValid = False
            Exit For
         End If
      ElseIf Character = MASK_DIGIT Or Character = MASK_UPPERCASE Then
         If Not Mid$(FixedInput, Position, 1) = NOT_FIXED Then
            ParameterSymbolError "Fixed character indicated in fixed characters but not in mask at: #" & CStr(Position) & ".", Index
            IsValid = False
            Exit For
         End If
      Else
         ParameterSymbolError "Invalid mask character """ & Character & """ at: #" & CStr(Position) & ".", Index
         IsValid = False
         Exit For
      End If
   Next Position

EndRoutine:
   MaskValid = IsValid
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function



'This procedure returns the number of items minus one for the specified dimension in the specified array.
Private Function NumberOfItems(ArrayV As Variant, Optional Dimension As Long = 1) As Long
On Error GoTo ErrorTrap
Dim Count As Long

   Count = UBound(ArrayV, Dimension) - LBound(ArrayV, Dimension)
EndRoutine:
   NumberOfItems = Count
   Exit Function

ErrorTrap:
   Count = UNKNOWN_NUMBER
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure pads the specified text with the number spaces specified.
Private Function Pad(Text As String, Width As Long, Optional PadLeft As Boolean = False) As String
On Error GoTo ErrorTrap
Dim PaddedText As String

   PaddedText = Text
   If Len(PaddedText) > Width Then PaddedText = Left$(PaddedText, Width)
   If PadLeft Then
      PaddedText = Space$(Width - Len(PaddedText)) & PaddedText
   Else
      PaddedText = PaddedText & Space$(Width - Len(PaddedText))
   End If

EndRoutine:
   Pad = PaddedText
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure returns the input's length for the specified query parameter.
Private Function ParameterInputLength(Index As Long) As Long
On Error GoTo ErrorTrap
Dim Length As Long
Dim Position As Long

   Length = 0
   With QueryParameters(, Index)
      For Position = 1 To Len(.Value)
         If Not Mid$(.Value, Position, 1) = Mid$(.Mask, Position, 1) Then Length = Position
      Next Position
   End With

EndRoutine:
   ParameterInputLength = Length
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure compares the specified character with the specified query parameter mask mask character.
Public Function ParameterMaskCharacterValid(Character As String, MaskCharacter As String, FixedInputCharacter As String) As String
On Error GoTo ErrorTrap
Dim Result As String
   
   Result = vbNullString

   If FixedInputCharacter = NOT_FIXED Then
      Select Case MaskCharacter
         Case MASK_UPPERCASE
            If Not (Character >= "A" And Character <= "Z") Then Result = "Uppercase character expected."
         Case MASK_DIGIT
            If Not (Character >= "0" And Character <= "9") Then Result = "Number expected."
      End Select
   Else
      If Not Character = FixedInputCharacter Then Result = "Fixed mask character """ & FixedInputCharacter & """ expected."
   End If
   
EndRoutine:
   ParameterMaskCharacterValid = Result
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function




'This procedure checks whether the parameters specified by the user are valid and returns the result.
Public Function ParametersValid(ParameterBoxes As Object) As Boolean
On Error GoTo ErrorTrap
Dim ErrorInformation As String
Dim Index As Long
Dim Valid As Boolean

   For Index = ParameterBoxes.LBound To ParameterBoxes.UBound
      QueryParameters , Index, ParameterBoxes(Index).Text
   Next Index
   
   Index = InvalidParameterInput(ErrorInformation)
   Valid = (Index = NO_PARAMETER)
   
   If Not Valid Then
      With ParameterBoxes(Index)
         If .Visible Then
            ErrorInformation = "This input box has not been correctly or fully filled in:" & ErrorInformation
         Else
            ErrorInformation = "Invisible parameter #" & CStr(Index - ParameterBoxes.LBound) & "  has not been correctly or fully filled in:" & ErrorInformation
         End If
         MsgBox ErrorInformation, vbExclamation
         If .Visible Then .SetFocus
      End With
   End If
   
EndRoutine:
   ParametersValid = Valid
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure displays parameter and/or symbol related error messages.
Private Sub ParameterSymbolError(Message As String, Optional Index As Long = NO_PARAMETER)
On Error GoTo ErrorTrap
Dim FirstParameter As Long

   QueryParameters , , , FirstParameter

   If Not Index = NO_PARAMETER Then
      Message = Message & vbCr & "Parameter definition: #" & CStr((Index - FirstParameter) + 1)
      With QueryParameters(, Index)
         If Not .ParameterName = vbNullString Then Message = Message & vbCr & "Name: """ & .ParameterName & """"
         If Not .Value = vbNullString Then Message = Message & vbCr & "Input: """ & .Value & """"
         If Not .DefaultValue = vbNullString Then Message = Message & vbCr & "Default value: """ & .DefaultValue & """"
         If Not .Mask = vbNullString Then Message = Message & vbCr & "Mask: """ & .Mask & """"
         If Not .FixedInput = vbNullString Then Message = Message & vbCr & "Fixed input: """ & .FixedInput & """"
      End With
   End If
   If Not Query().Path = vbNullString Then Message = Message & vbCr & "Query: " & Query().Path
   MsgBox Message, vbExclamation
EndRoutine:
Exit Sub

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub




'This procedure processes the batch settings.
Private Function ProcessBatchSettings(Line As String, Section As String, ByRef BatchSettings As SettingsStructure) As Boolean
On Error GoTo ErrorTrap
Dim ParameterName As String
Dim Processed As Boolean
Dim Value As String

   ParameterName = vbNullString
   Processed = True
   Value = ReadParameter(Line, ParameterName)
   
   With BatchSettings
      Select Case ParameterName
         Case "interactive"
            .BatchInteractive = CBool(Value)
         Case "querypath"
            .BatchQueryPath = Value
         Case "range"
            .BatchRange = Value
         Case Else
            If SettingsError("Unrecognized parameter.", BatchSettings.FileName, Section, Line) = vbCancel Then Processed = False
      End Select
   End With
EndRoutine:
   ProcessBatchSettings = Processed
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False, TypePath:="Settings file: ", Path:=BatchSettings.FileName, ExtraInformation:="Section: " & Section & vbCr & "Line: " & Line) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure processes the export settings.
Private Function ProcessExportSettings(Line As String, Section As String, ByRef ExportSettings As SettingsStructure) As Boolean
On Error GoTo ErrorTrap
Dim ParameterName As String
Dim Processed As Boolean
Dim Value As String
   
   ParameterName = vbNullString
   Processed = True
   Value = ReadParameter(Line, ParameterName)
   
   With ExportSettings
      Select Case ParameterName
         Case "autoopen"
            .ExportAutoOpen = CBool(Value)
         Case "autooverwrite"
            .ExportAutoOverwrite = CBool(Value)
         Case "autosend"
            .ExportAutoSend = CBool(Value)
         Case "ccrecipient"
            .ExportCCRecipient = Value
         Case "defaultpath"
            .ExportDefaultPath = Value
         Case "padcolumn"
            .ExportPadColumn = CBool(Value)
         Case "recipient"
            .ExportRecipient = Value
         Case "sender"
            .ExportSender = Value
         Case "subject"
            .ExportSubject = Value
         Case Else
            If SettingsError("Unrecognized parameter.", ExportSettings.FileName, Section, Line) = vbCancel Then Processed = False
      End Select
   End With
EndRoutine:
   ProcessExportSettings = Processed
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False, TypePath:="Settings file: ", Path:=ExportSettings.FileName, ExtraInformation:="Section: " & Section & vbCr & "Line: " & Line) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure returns the connection information containing the specified logon information.
Public Function ProcessLogonInformation(User As String, Password As String, ConnectionInformation As String) As String
On Error GoTo ErrorTrap
Dim LeftPart As String
Dim Position As Long
Dim ProcessedLogonInformation As String
Dim RightPart As String

   ProcessedLogonInformation = ConnectionInformation

   Position = InStr(UCase$(ProcessedLogonInformation), USER_VARIABLE)
   If Position > 0 Then
      LeftPart = Left$(ProcessedLogonInformation, Position - 1)
      RightPart = Mid$(ProcessedLogonInformation, Position + Len(USER_VARIABLE))
      ProcessedLogonInformation = LeftPart & User & RightPart
   End If

   Position = InStr(UCase$(ProcessedLogonInformation), PASSWORD_VARIABLE)
   If Position > 0 Then
      LeftPart = Left$(ProcessedLogonInformation, Position - 1)
      RightPart = Mid$(ProcessedLogonInformation, Position + Len(PASSWORD_VARIABLE))
      ProcessedLogonInformation = LeftPart & Password & RightPart
   End If
   
EndRoutine:
   ProcessLogonInformation = ProcessedLogonInformation
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure processes the preview settings.
Private Function ProcessPreviewSettings(Line As String, Section As String, ByRef PreviewSettings As SettingsStructure) As Boolean
On Error GoTo ErrorTrap
Dim ParameterName As String
Dim Processed As Boolean
Dim Value As String
   
   ParameterName = vbNullString
   Processed = True
   Value = ReadParameter(Line, ParameterName)
   
   With PreviewSettings
      Select Case ParameterName
         Case "columnwidth"
            .PreviewColumnWidth = CLng(Value)
         Case "rows"
            .PreviewLines = CLng(Value)
         Case Else
            If SettingsError("Unrecognized parameter.", PreviewSettings.FileName, Section, Line) = vbCancel Then Processed = False
      End Select
   End With
EndRoutine:
   ProcessPreviewSettings = Processed
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False, TypePath:="Settings file: ", Path:=PreviewSettings.FileName, ExtraInformation:="Section: " & Section & vbCr & "Line: " & Line) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure processes the query settings.
Private Function ProcessQuerySettings(Line As String, Section As String, ByRef QuerySettings As SettingsStructure) As Boolean
On Error GoTo ErrorTrap
Dim ParameterName As String
Dim Processed As Boolean
Dim Value As String

   ParameterName = vbNullString
   Processed = True
   Value = ReadParameter(Line, ParameterName)
   
   With QuerySettings
      Select Case ParameterName
         Case "autoclose"
            .QueryAutoClose = CBool(Value)
         Case "autoexecute"
            .QueryAutoExecute = CBool(Value)
         Case "recordsets"
            .QueryRecordSets = CBool(Value)
         Case "timeout"
            .QueryTimeout = CLng(Value)
         Case Else
            If SettingsError("Unrecognized parameter.", QuerySettings.FileName, Section, Line) = vbCancel Then Processed = False
      End Select
   End With
EndRoutine:
   ProcessQuerySettings = Processed
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False, TypePath:="Settings file: ", Path:=QuerySettings.FileName, ExtraInformation:="Section: " & Section & vbCr & "Line: " & Line) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure processes the specified session list.
Public Function ProcessSessionList(Optional SessionListPath As String = vbNullString) As String
On Error GoTo ErrorTrap
Dim FileHandle As Long
Dim SessionParameters As String
Static CurrentSessionListPath As String

   If Not SessionListPath = vbNullString Then
      AbortSessions NewAbortSessions:=False
      FileHandle = FreeFile()
      CurrentSessionListPath = SessionListPath
      Open CurrentSessionListPath For Input Lock Read Write As FileHandle
         Do Until EOF(FileHandle) Or AbortSessions()
            Line Input #FileHandle, SessionParameters
            If Not Trim$(SessionParameters) = vbNullString Then ExecuteSession SessionParameters
         Loop
      Close FileHandle
   End If

EndRoutine:
   ProcessSessionList = CurrentSessionListPath
   Exit Function

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False, TypePath:="Session list: ", Path:=CurrentSessionListPath) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function



'This procedure returns the value represented by the specified symbol.
Private Function ProcessSymbol(Symbol As String) As String
On Error GoTo ErrorTrap
Dim IsNumber As Boolean
Dim Message As String
Dim SymbolArgument As String
Dim Value As String

   On Error GoTo IsNotANumber
   IsNumber = CStr(CLng(Val(Symbol))) = Symbol
   On Error GoTo ErrorTrap

   If IsNumber Then
      If CLng(Val(Symbol)) = 0 Then
         Value = FileSystemO().GetBaseName(Query().Path)
      Else
         Value = QueryParameters(, CLng(Val(Symbol)) - 1).Value
      End If
   Else
      SymbolArgument = Mid$(Symbol, 2)
      Symbol = Left$(Symbol, 1)

      Select Case Symbol
         Case "D"
            Value = Format$(Day(Date), "00") & Format$(Month(Date), "00") & CStr(Year(Date))
         Case "b"
            If CStr(CLng(Val(SymbolArgument))) = SymbolArgument Then Value = InteractiveBatchParameters(CLng(Val(SymbolArgument)))
         Case "c"
            If CStr(CLng(Val(SymbolArgument))) = SymbolArgument Then Value = ChrW$(CLng(Val(SymbolArgument)))
         Case "d"
            Value = Format$(Day(Date), "00")
         Case "e"
            Value = Environ$(SymbolArgument)
         Case "m"
            Value = Format$(Month(Date), "00")
         Case "s"
            If CStr(CLng(Val(SymbolArgument))) = SymbolArgument Then Value = SessionParameters(CLng(Val(SymbolArgument)))
         Case "y"
            Value = Format$(Year(Date), "0000")
         Case Else
            If Not Symbol = vbNullString Then ParameterSymbolError "Symbol """ & Symbol & """ is unknown. It will be ignored."
      End Select
   End If
   
EndRoutine:
   ProcessSymbol = Value
   Exit Function

ErrorTrap:
   Message = "Symbol """ & Symbol & """ has caused the following error: " & vbCr
   Message = Message & Err.Description & "." & vbCr
   Message = Message & "Error code: " & Err.Number
   ParameterSymbolError Message
   Resume EndRoutine

IsNotANumber:
   IsNumber = False
   Resume Next
End Function


'This procedure returns this program's version.
Public Function ProgramVersion() As String
On Error GoTo ErrorTrap
EndRoutine:
   With App
      ProgramVersion = "v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision)
   End With
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure checks whether the specified parameter's properties contain invalid characters.
Private Function PropertiesValid(Index As Long) As Boolean
On Error GoTo ErrorTrap
Dim Character As String
Dim IsValid As Boolean
Dim Position As Long
Dim Properties As String

   IsValid = True
   Properties = QueryParameters(, Index).Properties
   For Position = 1 To Len(Properties)
      Character = Mid$(Properties, Position, 1)
      If Not (Character = PROPERTY_HIDDEN Or Character = PROPERTY_VARIABLE_LENGTH) Then
         ParameterSymbolError "Invalid property """ & Character & """ at: #" & CStr(Position) & ".", Index
         IsValid = False
         Exit For
      End If
   Next Position

EndRoutine:
   PropertiesValid = IsValid
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function




'This procedure loads the specified query or returns a previously loaded query.
Public Function Query(Optional QueryPath As String = vbNullString) As QueryStructure
On Error GoTo ErrorTrap
Dim FileHandle As Long
Static CurrentQuery As QueryStructure

   With CurrentQuery
      .Opened = False
   
      If Not QueryPath = vbNullString Then
         FileHandle = FreeFile()
         Open QueryPath For Input Lock Read Write As FileHandle: Close FileHandle
         
         FileHandle = FreeFile()
         Open QueryPath For Binary Lock Read Write As FileHandle
            .Code = Input$(LOF(FileHandle), FileHandle)
         Close FileHandle
   
         .Path = QueryPath
         .Opened = True
      End If
   End With
   
EndRoutine:
   Query = CurrentQuery
   Exit Function
   
ErrorTrap:
   CurrentQuery.Opened = False
   
   If HandleError(ReturnPreviousChoice:=False, TypePath:="Query path: ", Path:=QueryPath) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure searches the specified query for parameter definitions or returns a parameter definition found earlier.
Public Function QueryParameters(Optional QueryCode As String = vbNullString, Optional ParameterIndex As Long = 0, Optional Value As Variant, Optional ByRef FirstParameter As Long = 0, Optional ByRef LastParameter As Long = 0) As QueryParameterStructure
On Error GoTo ErrorTrap
Dim Definition As String
Dim DefinitionEnd As Long
Dim DefinitionStart As Long
Dim Elements() As String
Dim UnprocessedCode As String
Static Parameters() As QueryParameterStructure
  
   If Not IsMissing(Value) Then
      If Not CheckForAPIError(SafeArrayGetDim(Parameters)) = 0 Then Parameters(ParameterIndex).Value = CStr(Value)
   ElseIf Not QueryCode = vbNullString Then
      Erase Parameters()

      UnprocessedCode = QueryCode
      Do
         DefinitionStart = InStr(UnprocessedCode, DEFINITION_CHARACTERS)
         If DefinitionStart > 0 Then
            DefinitionEnd = InStr(DefinitionStart + Len(DEFINITION_CHARACTERS), UnprocessedCode, DEFINITION_CHARACTERS)
            If DefinitionEnd > 0 Then
               If CheckForAPIError(SafeArrayGetDim(Parameters())) = 0 Then
                  ReDim Parameters(0 To 0) As QueryParameterStructure
               Else
                  ReDim Preserve Parameters(LBound(Parameters()) To UBound(Parameters()) + 1) As QueryParameterStructure
               End If

               Definition = Mid$(UnprocessedCode, DefinitionStart + Len(DEFINITION_CHARACTERS), (DefinitionEnd - DefinitionStart) - Len(DEFINITION_CHARACTERS))
               UnprocessedCode = Mid$(UnprocessedCode, DefinitionEnd + Len(DEFINITION_CHARACTERS))

               With Parameters(UBound(Parameters()))
                  .Length = Len(DEFINITION_CHARACTERS & Definition & DEFINITION_CHARACTERS)
                  .Position = DefinitionStart

                  Elements = Split(Definition, ELEMENT_CHARACTER)
                  If NumberOfItems(Elements) > Abs(CommentsElement - NameElement) Then ParameterSymbolError "To many elements, these will be ignored.", UBound(Parameters())
                  ReDim Preserve Elements(NameElement To CommentsElement) As String

                  .ParameterName = Elements(NameElement)

                  .Mask = Elements(MaskElement)

                  .FixedInput = Elements(FixedElement)
                  If Not .Mask = vbNullString Then
                     If .FixedInput = vbNullString Then
                        .FixedInput = String$(Len(.Mask), NOT_FIXED)
                     Else
                        If Not Len(.FixedInput) = Len(.Mask) Then ParameterSymbolError ("The fixed input must be the same length as the mask. Any surplus characters will be removed."), UBound(Parameters())
                     End If
                  End If

                  .DefaultValue = ReplaceSymbols(Elements(DefaultValueElement))
                  If Not .Mask = vbNullString Then If Len(.DefaultValue) > Len(.Mask) Then ParameterSymbolError "The default value is longer than the mask. The surplus characters will be removed.", UBound(Parameters())

                  .Properties = Elements(PropertiesElement)

                  .Comments = Elements(CommentsElement)
                  .Value = .DefaultValue

                  .FixedMask = GenerateFixedMask(UBound(Parameters()))
                  .InputBoxIsVisible = Not (InStr(.Properties, PROPERTY_HIDDEN) > 0)
                  .LengthIsVariable = (InStr(.Properties, PROPERTY_VARIABLE_LENGTH) > 0)

                  MaskValid UBound(Parameters())
                  PropertiesValid UBound(Parameters())
               End With
            Else
               ParameterSymbolError "No end of parameter definition marker. This definition will be ignored.", UBound(Parameters())
               Exit Do
            End If
         Else
            Exit Do
         End If
      Loop
   End If
   
EndRoutine:
   FirstParameter = NO_PARAMETER
   LastParameter = NO_PARAMETER
   If Not CheckForAPIError(SafeArrayGetDim(Parameters())) = 0 Then
      FirstParameter = LBound(Parameters())
      LastParameter = UBound(Parameters())
      QueryParameters = Parameters(ParameterIndex)
   End If
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function
'This procedure handles any query result read errors that occur.
Public Function QueryResultReadError(Optional Row As Long = 0, Optional Column As Long = 0, Optional ColumnName As String = vbNullString, Optional ReturnPreviousChoice As Boolean = True) As Long
Dim ErrorCode As Long
Dim ErrorDescription As String
Dim Message As String
Dim Source As String
Static Choice As Long

   Source = Err.Source
   ErrorCode = Err.Number
   ErrorDescription = Err.Description
   Err.Clear

   On Error Resume Next

   If Not ReturnPreviousChoice Then
      Message = "An error has occurred while reading the query result." & vbCr
      Message = Message & "Row: " & CStr(Row) & vbCr
      Message = Message & "Column: " & CStr(Column) & vbCr
      Message = Message & "Column name: " & CStr(ColumnName) & vbCr
      Message = Message & "Description: " & FormatErrorDescription(ErrorDescription) & vbCr
      Message = Message & "Error code: " & CStr(ErrorCode)
      If Not Source = vbNullString Then Message = Message & vbCr & "Source: " & Source

      Choice = MsgBox(Message, vbExclamation Or vbAbortRetryIgnore Or vbDefaultButton2)
   End If

   QueryResultReadError = Choice
End Function


'This procedure manages the query results.
Public Function QueryResults(Optional NewQueryResult As Adodb.Recordset = Nothing, Optional RemoveResults As Boolean = False, Optional ResultIndex As Long = 0, Optional ByRef FirstResult As Long = 0, Optional ByRef LastResult As Long = 0) As QueryResultStructure
On Error GoTo ErrorTrap
Dim Column As Long
Dim Row As Long
Dim TemporaryTable() As String
Static Results() As QueryResultStructure

   If Not NewQueryResult Is Nothing Then
      With NewQueryResult
         If Not .BOF Then
            If CheckForAPIError(SafeArrayGetDim(Results())) = 0 Then
               ReDim Results(0 To 0) As QueryResultStructure
            Else
               ReDim Preserve Results(LBound(Results()) To UBound(Results()) + 1) As QueryResultStructure
            End If

            With Results((UBound(Results())))
               If CheckForAPIError(SafeArrayGetDim(.ColumnWidth())) = 0 Then ReDim .ColumnWidth(0 To 0) As Long
               If CheckForAPIError(SafeArrayGetDim(.RightAligned())) = 0 Then ReDim .RightAligned(0 To 0) As Boolean
               If CheckForAPIError(SafeArrayGetDim(.Table())) = 0 Then ReDim .Table(0 To 0, 0 To 0) As String
            End With

            Row = 0
            ReDim Results(UBound(Results())).ColumnWidth(0 To .Fields.Count - 1) As Long
            ReDim Results(UBound(Results())).RightAligned(0 To .Fields.Count - 1) As Boolean
            ReDim TemporaryTable(0 To .Fields.Count - 1, 0 To Row) As String
            For Column = 0 To .Fields.Count - 1
               TemporaryTable(Column, Row) = Trim$(.Fields.Item(Column).Name)
               Results(UBound(Results())).ColumnWidth(Column) = Len(TemporaryTable(Column, Row))
               Results(UBound(Results())).RightAligned(Column) = Not IsLeftAligned(.Fields.Item(Column).Type)
            Next Column
            Row = Row + 1

            On Error GoTo ReadError
            ReDim Preserve TemporaryTable(LBound(TemporaryTable(), 1) To .Fields.Count - 1, LBound(TemporaryTable(), 2) To Row) As String
            Do Until .EOF
               For Column = 0 To .Fields.Count - 1
                  If Not IsNull(.Fields.Item(Column).Value) Then
                     TemporaryTable(Column, Row) = Trim$(.Fields.Item(Column).Value)
                     If Len(TemporaryTable(Column, Row)) > Results(UBound(Results())).ColumnWidth(Column) Then Results(UBound(Results())).ColumnWidth(Column) = Len(TemporaryTable(Column, Row)) + 1
                  End If
NextValue:
               Next Column
               .MoveNext
               Row = Row + 1
               ReDim Preserve TemporaryTable(LBound(TemporaryTable(), 1) To .Fields.Count - 1, LBound(TemporaryTable(), 2) To Row) As String
            Loop
            On Error GoTo 0
EndReading:

            ReDim Results(UBound(Results())).Table(0 To Row, 0 To .Fields.Count - 1) As String
            For Row = LBound(Results(UBound(Results())).Table(), 1) To UBound(Results(UBound(Results())).Table(), 1) - 1
               For Column = LBound(Results(UBound(Results())).Table(), 2) To UBound(Results(UBound(Results())).Table(), 2)
                  Results(UBound(Results())).Table(Row, Column) = TemporaryTable(Column, Row)
               Next Column
            Next Row
         End If
      End With
   ElseIf RemoveResults Then
      Erase Results()
   End If

EndRoutine:
   FirstResult = NO_RESULT
   LastResult = NO_RESULT
   If Not CheckForAPIError(SafeArrayGetDim(Results())) = 0 Then
      FirstResult = LBound(Results())
      LastResult = UBound(Results())
      QueryResults = Results(ResultIndex)
   End If
   Exit Function

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
Exit Function

ReadError:
   If QueryResultReadError(Row, Column, TemporaryTable(Column, 0), ReturnPreviousChoice:=False) = vbAbort Then Resume EndReading
   If QueryResultReadError() = vbIgnore Then Resume NextValue
   If QueryResultReadError() = vbRetry Then Resume
End Function



'This procedure returns the query result as text.
Public Function QueryResultText(Result As QueryResultStructure) As String
On Error GoTo ErrorTrap
Dim Column As Long
Dim LastLine As Long
Dim ResultText As String
Dim Row As Long
Dim Text As String
Dim Width As Long

   With Result
      If Not CheckForAPIError(SafeArrayGetDim(.Table)) = 0 Then
         If Settings().PreviewLines = NO_MAXIMUM Or Settings().PreviewLines > NumberOfItems(.Table(), Dimension:=1) Then
            LastLine = NumberOfItems(.Table(), Dimension:=1)
         Else
            LastLine = Settings().PreviewLines - 1
         End If
         
         ResultText = vbNullString
         For Row = LBound(.Table(), 1) To LastLine
            For Column = LBound(.Table(), 2) To UBound(.Table(), 2)
               Width = .ColumnWidth(Column)
               Text = .Table(Row, Column)
               Text = Replace(Text, vbCr, " ")
               Text = Replace(Text, vbLf, " ")
               Text = Replace(Text, vbTab, " ")
               If Not Settings().PreviewColumnWidth = NO_MAXIMUM Then
                  If .ColumnWidth(Column) > Settings().PreviewColumnWidth Then
                     Width = Settings().PreviewColumnWidth
                     Text = Left$(Text, Settings().PreviewColumnWidth)
                  End If
               End If
   
               ResultText = ResultText & Pad(Text, Width, .RightAligned(Column)) & " "
            Next Column
            ResultText = ResultText & vbCrLf
            If DoEvents() = 0 Then Exit For
         Next Row
      End If
   End With
   
EndRoutine:
   QueryResultText = ResultText
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure returns the settings parameter's value and name contained in the specified line of text.
Private Function ReadParameter(Line As String, ByRef ParameterName As String) As String
On Error GoTo ErrorTrap
Dim Position As Long
Dim Value As String

   ParameterName = vbNullString
   Value = vbNullString
   Position = InStr(Line, VALUE_CHARACTER)
   If Position > 0 Then
      ParameterName = LCase$(Trim$(Left$(Line, Position - 1)))
      Value = Trim$(Mid$(Line, Position + 1))
   End If
   
EndRoutine:
   ReadParameter = Value
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procecure procedure removes any formatting from the specified querycode.
Private Function RemoveFormatting(QueryCode As String, CommentStart As String, CommentEnd As String, StringCharacters As String) As String
On Error GoTo ErrorTrap
Dim Character As String
Dim CurrentStringCharacter As String
Dim InComment As Boolean
Dim Index As Long
Dim QueryWithoutFormatting As String
   
   Character = vbNullString
   CurrentStringCharacter = vbNullString
   InComment = False
   Index = 1
   QueryWithoutFormatting = vbNullString
   Do Until Index > Len(QueryCode)
      Character = Mid$(QueryCode, Index, 1)

      If InComment Then
         If CommentEnd = vbNullString Then
            If Mid$(QueryCode, Index, 1) = vbCr Or Mid$(QueryCode, Index, 1) = vbLf Then
               CurrentStringCharacter = vbNullString
               InComment = False
               Character = " "
            End If
         Else
            If Mid$(QueryCode, Index, Len(CommentEnd)) = CommentEnd Then
               CurrentStringCharacter = vbNullString
               InComment = False
               Index = Index + (Len(CommentEnd) - 1)
               Character = " "
            End If
         End If
      Else
         If InStr(STRING_CHARACTERS, Mid$(QueryCode, Index, 1)) > 0 Then
            If CurrentStringCharacter = vbNullString Then
               CurrentStringCharacter = Character
            ElseIf Character = CurrentStringCharacter Then
               CurrentStringCharacter = vbNullString
            End If
         ElseIf Mid$(QueryCode, Index, Len(CommentStart)) = CommentStart Then
            If CurrentStringCharacter = vbNullString Then InComment = True
         End If
      End If

      If Not InComment Then
         If CurrentStringCharacter = vbNullString Then
            If Mid$(QueryCode, Index, 1) = vbCr Or Mid$(QueryCode, Index, 1) = vbLf Then Character = " "

            If InStr(vbTab & " ", Character) > 0 Then
               Character = " "
               If Right$(QueryWithoutFormatting, 1) = " " Then Character = vbNullString
            End If
         End If

         QueryWithoutFormatting = QueryWithoutFormatting & Character
      End If

      Index = Index + 1
   Loop
   
EndRoutine:
   RemoveFormatting = QueryWithoutFormatting
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'Deze procecure vervangt de symbolen in de opgegeven tekst met de tekst waar ze voor staan.
Public Function ReplaceSymbols(Text As String) As String
On Error GoTo ErrorTrap
Dim Symbol As String
Dim SymbolEnd As Long
Dim SymbolStart As Long
Dim TextWithoutSymbols As String
Dim TextWithSymbols As String
      
   TextWithSymbols = Text
   TextWithoutSymbols = vbNullString
   Do
      SymbolStart = InStr(TextWithSymbols, SYMBOL_CHARACTER)
      If SymbolStart = 0 Then
         TextWithoutSymbols = TextWithoutSymbols & TextWithSymbols
         Exit Do
      Else
         SymbolEnd = InStr(SymbolStart + 1, TextWithSymbols, SYMBOL_CHARACTER)
         If SymbolEnd = 0 Then
            TextWithoutSymbols = TextWithoutSymbols & TextWithSymbols
            Exit Do
         Else
            TextWithoutSymbols = TextWithoutSymbols & Left$(TextWithSymbols, SymbolStart - 1)
            Symbol = Mid$(TextWithSymbols, SymbolStart + 1, (SymbolEnd - SymbolStart) - 1)
            TextWithSymbols = Mid$(TextWithSymbols, SymbolEnd + 1)
            
            If Symbol = vbNullString Then
               ParameterSymbolError "An empty symbol has been found. It will be ignored."
            Else
               TextWithoutSymbols = TextWithoutSymbols & ProcessSymbol(Symbol)
           End If
         End If
      End If
   Loop
   
EndRoutine:
   ReplaceSymbols = TextWithoutSymbols
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function

'This procedure request the user to specify the information for a connection with a database.
Private Function RequestConnectionInformation() As String
On Error GoTo ErrorTrap
Dim ConnectionInformation As String

   Do While Trim$(ConnectionInformation) = vbNullString
      ConnectionInformation = InputBox$("Information to connect with a database:")
      If StrPtr(ConnectionInformation) = 0 Then
         Exit Do
      ElseIf Trim$(ConnectionInformation) = vbNullString Then
         MsgBox "This information is required.", vbExclamation
      End If
   Loop

EndRoutine:
   RequestConnectionInformation = Trim$(ConnectionInformation)
   Exit Function

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function
'This procedure opens a dialog window with which the user can browse to the path for the query result to be exported.
Public Function RequestExportPath(CurrentExportPath As String) As String
On Error GoTo ErrorTrap
Dim ExportPathDialog As OPENFILENAME
Dim NewExportPath As String

   NewExportPath = CurrentExportPath

   With ExportPathDialog
      .hInstance = CLng(0)
      .hwndOwner = CLng(0)
      .lCustData = CLng(0)
      .lpfnHook = CLng(0)
      .lpstrCustomFilter = vbNullString
      .lpstrDefExt = vbNullString
      .lpstrFile = String$(MAX_STRING, vbNullChar) & vbNullChar
      .lpstrFileTitle = String$(MAX_STRING, vbNullChar) & vbNullChar
      .lpTemplateName = vbNullString
      .lStructSize = Len(ExportPathDialog)
      .nFileExtension = CLng(0)
      .nFileOffset = CLng(0)
      .nFilterIndex = CLng(1)
      .nMaxCustomFilter = CLng(0)
      .nMaxFile = Len(.lpstrFile)
      .nMaxFileTitle = Len(.lpstrFileTitle)
   
      .flags = OFN_EXPLORER
      .flags = .flags Or OFN_HIDEREADONLY
      .flags = .flags Or OFN_LONGNAMES
      .flags = .flags Or OFN_NOCHANGEDIR
      .flags = .flags Or OFN_PATHMUSTEXIST
      .lpstrTitle = "Export the query result to:" & vbNullChar
      .lpstrFilter = "Text file (*.txt)" & vbNullChar & "*.txt" & vbNullChar
      .lpstrFilter = .lpstrFilter & "Microsoft Excel file (*.xls)" & vbNullChar & "*.xls" & vbNullChar
      .lpstrFilter = .lpstrFilter & "Microsoft Excel 2007 file (*.xlsx)" & vbNullChar & "*.xlsx" & vbNullChar
      .lpstrFilter = .lpstrFilter & vbNullChar
      .lpstrInitialDir = App.Path & vbNullChar
   
      If CBool(CheckForAPIError(GetSaveFileNameA(ExportPathDialog))) Then NewExportPath = Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
   End With
EndRoutine:
   RequestExportPath = NewExportPath
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure opens a dialog window with which the user can browse to a query file.
Public Function RequestQueryPath() As String
On Error GoTo ErrorTrap
Dim QueryPath As String
Dim QueryPathDialog As OPENFILENAME
   
   QueryPath = vbNullString

   With QueryPathDialog
      .hInstance = CLng(0)
      .hwndOwner = CLng(0)
      .lCustData = CLng(0)
      .lpfnHook = CLng(0)
      .lpstrCustomFilter = vbNullString
      .lpstrDefExt = vbNullString
      .lpstrFile = String$(MAX_STRING, vbNullChar) & vbNullChar
      .lpstrFileTitle = String$(MAX_STRING, vbNullChar) & vbNullChar
      .lpTemplateName = vbNullString
      .lStructSize = Len(QueryPathDialog)
      .nFileExtension = CLng(0)
      .nFileOffset = CLng(0)
      .nFilterIndex = CLng(1)
      .nMaxCustomFilter = CLng(0)
      .nMaxFile = Len(.lpstrFile)
      .nMaxFileTitle = Len(.lpstrFileTitle)
   
      .flags = OFN_EXPLORER
      .flags = .flags Or OFN_FILEMUSTEXIST
      .flags = .flags Or OFN_HIDEREADONLY
      .flags = .flags Or OFN_LONGNAMES
      .flags = .flags Or OFN_NOCHANGEDIR
      .flags = .flags Or OFN_PATHMUSTEXIST
      .lpstrTitle = "Select a query:" & vbNullChar
      .lpstrFilter = "Query Assistant query files (*.qa)" & vbNullChar & "*.qa" & vbNullChar
      .lpstrFilter = .lpstrFilter & vbNullChar
   
      .lpstrInitialDir = App.Path & vbNullChar
      If CBool(CheckForAPIError(GetOpenFileNameA(QueryPathDialog))) Then QueryPath = Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
   End With

EndRoutine:
   RequestQueryPath = QueryPath
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure saves this program's settings.
Private Sub SaveSettings(SettingsPath As String, SettingsToBeSaved As SettingsStructure, Message As String)
On Error GoTo ErrorTrap
Dim FileHandle As Long

   FileHandle = FreeFile()
   Open SettingsPath For Output Lock Read Write As FileHandle
      With SettingsToBeSaved
         Print #FileHandle, SECTION_NAME_START & "BATCH" & SECTION_NAME_END
         Print #FileHandle, "Interactive" & VALUE_CHARACTER & CStr(.BatchInteractive)
         Print #FileHandle, "QueryPath" & VALUE_CHARACTER & .BatchQueryPath
         Print #FileHandle, "Range" & VALUE_CHARACTER & .BatchRange
         Print #FileHandle,

         Print #FileHandle, SECTION_NAME_START & "CONNECTION" & SECTION_NAME_END
         Print #FileHandle, .ConnectionInformation
         Print #FileHandle,

         Print #FileHandle, SECTION_NAME_START & "EMAILTEXT" & SECTION_NAME_END
         Print #FileHandle, .EMailText
         Print #FileHandle,

         Print #FileHandle, SECTION_NAME_START & "EXPORT" & SECTION_NAME_END
         Print #FileHandle, "AutoOpen" & VALUE_CHARACTER & CStr(.ExportAutoOpen)
         Print #FileHandle, "AutoOverwrite" & VALUE_CHARACTER & CStr(.ExportAutoOverwrite)
         Print #FileHandle, "AutoSend" & VALUE_CHARACTER & CStr(.ExportAutoSend)
         Print #FileHandle, "CCRecipient" & VALUE_CHARACTER & .ExportCCRecipient
         Print #FileHandle, "DefaultPath" & VALUE_CHARACTER & .ExportDefaultPath
         Print #FileHandle, "PadColumn" & VALUE_CHARACTER & CStr(.ExportPadColumn)
         Print #FileHandle, "Recipient" & VALUE_CHARACTER & .ExportRecipient
         Print #FileHandle, "Sender" & VALUE_CHARACTER & .ExportSender
         Print #FileHandle, "Subject" & VALUE_CHARACTER & .ExportSubject
         Print #FileHandle,

         Print #FileHandle, SECTION_NAME_START & "PREVIEW" & SECTION_NAME_END
         Print #FileHandle, "ColumnWidth" & VALUE_CHARACTER & CStr(.PreviewColumnWidth)
         Print #FileHandle, "Rows" & VALUE_CHARACTER & CStr(.PreviewLines)
         Print #FileHandle,

         Print #FileHandle, SECTION_NAME_START & "QUERY" & SECTION_NAME_END
         Print #FileHandle, "AutoClose" & VALUE_CHARACTER & CStr(.QueryAutoClose)
         Print #FileHandle, "AutoExecute" & VALUE_CHARACTER & CStr(.QueryAutoExecute)
         Print #FileHandle, "Recordsets" & VALUE_CHARACTER & CStr(.QueryRecordSets)
         Print #FileHandle, "Timeout" & VALUE_CHARACTER & CStr(.QueryTimeout)
      End With
   Close FileHandle

   MsgBox Message & vbCr & SettingsPath, vbInformation

EndRoutine:
   Close FileHandle
   Exit Sub
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False, TypePath:="Settings file: ", Path:=SettingsPath) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Sub

'This procedure manages the session parameters.
Private Function SessionParameters(Optional Index As Long = 0, Optional NewParameter As Variant, Optional Remove As Boolean = False) As String
On Error GoTo ErrorTrap
Static Parameters As New Collection

   If Not IsMissing(NewParameter) Then
      Parameters.Add CStr(NewParameter)
   ElseIf Remove Then
      Set Parameters = New Collection
   End If

EndRoutine::
   If Parameters.Count = 0 Then
      SessionParameters = vbNullString
   Else
      SessionParameters = Parameters(Index + 1)
   End If
   Exit Function

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function




'This procedure returns this program's settings.
Public Function Settings(Optional SettingsPath As String = vbNullString) As SettingsStructure
On Error GoTo ErrorTrap
Dim Message As String
Static ProgramSettings As SettingsStructure

   If Not SettingsPath = vbNullString Then
      If FileSystemO().FileExists(SettingsPath) Then
         ProgramSettings = LoadSettings(SettingsPath)
      Else
         Message = "Cannot find settings file." & vbCr
         Message = Message & "Settings file: " & SettingsPath & vbCr
         Message = Message & "Generate this file?" & vbCr
         Message = Message & "Current path: " & CurDir$()
         If MsgBox(Message, vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes Then
            SaveSettings SettingsPath, DefaultSettings(), "The default settings have been written to:"
            ProgramSettings = LoadSettings(SettingsPath)
         End If
      End If
   End If

EndRoutine:
   Settings = ProgramSettings
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure displays settings file related errors.
Private Function SettingsError(Message As String, Optional SettingsPath As String = vbNullString, Optional Section As String = vbNullString, Optional Line As String = vbNullString, Optional Fatal As Boolean = False) As Long
On Error GoTo ErrorTrap
Dim Choice As Long
Dim Style As Long

   If Not Section = vbNullString Then Message = Message & vbCr & "Section: " & Section
   If Not Line = vbNullString Then Message = Message & vbCr & "Line: " & """" & Line & """"
   If Not SettingsPath = vbNullString Then Message = Message & vbCr & "Settings file: " & SettingsPath

   Style = vbExclamation
   If Not Fatal Then Style = Style Or vbOKCancel Or vbDefaultButton1
   Choice = MsgBox(Message, Style)

EndRoutine:
   SettingsError = Choice
   Exit Function

ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure returns the query result's status after a query has been executed.
Public Function StatusAfterQuery(ResultIndex As Long) As String
On Error GoTo ErrorTrap
Dim ColumnCount As Long
Dim FirstResult As Long
Dim LastResult As Long
Dim ResultCount As Long
Dim RowCount As Long
Dim Status As String

   With QueryResults(, , ResultIndex)
      ColumnCount = 0
      ResultCount = 0
      RowCount = 0
      If Not CheckForAPIError(SafeArrayGetDim(.Table)) = 0 Then
         ColumnCount = NumberOfItems(.Table(), Dimension:=2) + 1
         If NumberOfItems(.Table(), Dimension:=1) = 0 Then ColumnCount = 0
         RowCount = NumberOfItems(.Table(), Dimension:=1)

         QueryResults , , , FirstResult, LastResult
         ResultCount = Abs(LastResult - FirstResult) + 1
      End If

      Status = "Query executed: " & CStr(RowCount)
      If RowCount = 1 Then Status = Status & " row" Else Status = Status & " rows"
      Status = Status & " and " & CStr(ColumnCount)
      If ColumnCount = 1 Then Status = Status & " column." Else Status = Status & " columns."

      If ResultCount > 1 Then Status = Status & " Result " & CStr((ResultIndex - FirstResult) + 1) & " of " & CStr(ResultCount) & "."

      If Settings().PreviewLines >= 0 Then
         Status = Status & " Preview limit: " & CStr(Settings().PreviewLines)
         If Settings().PreviewLines = 1 Then Status = Status & " row." Else Status = Status & " rows."
      End If
   End With

EndRoutine:
   StatusAfterQuery = Status
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function


'This procedure removes any leading/trailing quote from the specified path.
Public Function Unquote(Path As String) As String
On Error GoTo ErrorTrap
Dim UnquotedPath As String
   
   UnquotedPath = Path
   If Left$(UnquotedPath, 1) = """" Then UnquotedPath = Mid$(UnquotedPath, 2)
   If Right$(UnquotedPath, 1) = """" Then UnquotedPath = Left$(UnquotedPath, Len(UnquotedPath) - 1)
   
EndRoutine:
   Unquote = UnquotedPath
   Exit Function
   
ErrorTrap:
   If HandleError(ReturnPreviousChoice:=False) = vbIgnore Then Resume EndRoutine
   If HandleError() = vbRetry Then Resume
End Function

