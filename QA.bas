Attribute VB_Name = "QAModule"
'Deze module bevat de hoofdprocedures van dit programma.
Option Explicit

'De door dit programma gebruikte Microsoft Windows API constanten, functies en structuren.
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
Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwBerichtId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetOpenFileNameA Lib "Comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileNameA Lib "Comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SafeArrayGetDim Lib "Oleaut32.dll" (ByRef saArray() As Any) As Long
Private Declare Function SetCurrentDirectoryA Lib "Kernel32.dll" (ByVal lpPathName As String) As Long
Private Declare Function WaitMessage Lib "User32.dll" () As Long

'De door dit programma gebruikte constanten, definities, en opsommingen.

'Bevat een opsomming van de parameter definitie elementen.
Private Enum ParameterDefinitieOpsomming
   NaamElement
   MaskerElement
   StandaardWaardeElement
   CommentaarElement
End Enum

'Bevat de definities voor de instellingen van dit programma.
Public Type InstellingenDefinitie
   BatchBereik As String              'Bevat de volgnummers van de eerste en de laatste query in een uit te voeren batch.
   BatchInteractief As Boolean        'Geeft aan of de gebruiker eerst parameters moet invoeren voordat een batch uitgevoerd kan worden.
   BatchQueryPad As String            'Bevat het pad en/of de bestandsnaam zonder volgnummers van de query's in een uit te voeren batch.
   Bestand As String                  'Bevat het pad en/of de bestandsnaam van het programmainstellingenbestand.
   EMailTekst As String               'Bevat de tekst van de e-mail met de geëxporteerde resultaten.
   ExportAfzender As String           'Bevat de naam van de afzender van de e-mail met de geëxporteerde resultaten.
   ExportAutoOpenen As Boolean        'Geeft aan of een export automatisch na het exporteren geopend wordt.
   ExportAutoOverschrijven As Boolean 'Geeft aan of een bestand automatisch overschreven wordt bij het exporteren van de queryresultaataten.
   ExportAutoVerzenden As Boolean     'Geeft aan of de e-mail met de geëxporteerde resultaten automatisch verzonden wordt.
   ExportCCOntvanger As String        'Bevat het e-mail adres van de ontvanger van het kopie van de e-mail met de geëxporteerde resultaten.
   ExportKolomAanvullen As Boolean    'Geeft aan of de data in een kolom moet worden aangevuld met spaties.
   ExportOnderwerp As String          'Bevat het onderwerp van de e-mail met de geëxporteerde resultaten.
   ExportOntvanger As String          'Bevat het e-mail adres van de ontvanger van de e-mail met de geëxporteerde resultaten.
   ExportStandaardPad As String       'Bevat het standaardpad voor het exporteren van queryresultaataten.
   QueryAutoSluiten As Boolean        'Geeft aan of dit programma na het uitvoeren van een query en een eventuele export automatisch afgesloten wordt.
   QueryAutoUitvoeren As Boolean      'Geeft aan of een query automatisch uitgevoerd wordt na het laden.
   QueryRecordSets As Boolean         'Geeft aan of er meer dan een recordset kan worden teruggestuurd door de database als het resultaat van een query.
   QueryTimeout As Long               'Bevat het aantal seconden dat het programma wacht op het queryresultaat nadat opdracht is gegeven de query uit te voeren.
   VoorbeeldKolomBreedte As Long      'Bevat de maximale kolombreedte die gebruikt wordt om het queryresultaat te tonen in het voorbeeld venster.
   VoorbeeldRegels As Long            'Bevat het maximum aantal regels dat van het queryresultaat wordt getoond in het voorbeeld venster.
   VerbindingsInformatie As String    'Bevat de voor de verbinding met een database noodzakelijke gegevens.
End Type

'Bevat de definities voor de opdrachtregelparameters die eventueel zijn opgegeven bij het starten van dit programma.
Public Type OpdrachtRegelParametersDefinitie
   InstellingenPad As String   'Bevat het opgegeven instellingenpad.
   QueryPad As String          'Bevat het opgegeven querypad.
   SessiesPad As String        'Bevat het opgegeven sessielijstpad.
   Verwerkt As Boolean         'Geeft aan of de opdrachtregelparameters zonder fouten zijn verwerkt.
End Type

'Bevat de definities voor een query.
Public Type QueryDefinitie
   Code As String              'De code van een query.
   Pad As String               'Het pad van een query bestand.
   Geopend As Boolean          'Geeft aan of het query bestand kon worden geopend.
End Type

'Bevat de definities voor de parameter gegevens van de geselecteerde query.
Public Type QueryParameterDefinitie
   Commentaar As String        'Het commentaar bij de parameter.
   Invoer As String            'De invoer van de gebruiker.
   Lengte As Long              'De lengte van de parameterdefinitie.
   LengteIsVariabel As Boolean 'Geeft aan of de lengte van de invoer variabel is.
   Masker As String            'Het invoer masker van de parameter.
   ParameterNaam As String     'De naam van de parameter.
   Positie As Long             'De positie relatief ten op zichte van de vorige definitie.
   StandaardWaarde As String   'De standaardwaarde van de parameter.
   VeldIsZichtbaar As Boolean  'Geeft aan of het invoer veld zichtbaar is.
End Type

'Bevat de definities voor het resultaat van een query.
Public Type QueryResultaatDefinitie
   KolomBreedte() As Long       'Geeft per kolom de maximale breedte in bytes van de gegevens aan.
   RechtsUitlijnen() As Boolean 'Geeft per kolom aan of de gegevens rechtsuitgelijnd worden bij weergave.
   Tabel() As String            'Bevat de door een query opgevraagde gegevens uit een database.
End Type

Public Const GEBRUIKER_VARIABEL As String = "$$GEBRUIKER$$"       'Indien aanwezig in de verbindingsinformatie geeft dit variabel de positie van de gebruikersnaam aan.
Public Const GEEN_PARAMETER As Long = -1                          'Staat voor "geen parameter".
Public Const WACHTWOORD_VARIABEL As String = "$$WACHTWOORD$$"     'Indien aanwezig in de verbindingsinformatie geeft dit variabel de positie van het wachtwoord aan.
Private Const ASCII_A As Long = 65                                 'De ASCII waarde voor het teken "A".
Private Const ASCII_Z As Long = 90                                 'De ASCII waarde voor het teken "Z".
Private Const COMMENTAAR_TEKEN As String = "#"                     'Geeft aan dat een regel in een instellingenbestand commentaar is.
Private Const DEFINITIE_TEKENS As String = "$$"                    'Geeft het begin en het einde van een parameterdefinitie binnen een query aan.
Private Const ELEMENT_TEKEN As String = ":"                        'Scheidt de parameter definitie elementen van elkaar.
Private Const EXCEL_MAXIMUM_AANTAL_KOLOMMEN As Long = 255          'Het maximale aantal door Microsoft Excel ondersteunde kolommen.
Private Const GEEN_LETTER As Long = 64                             'Staat voor "geen letter". (De ASCII waarde die voor het "A" teken komt.)
Private Const GEEN_MAXIMUM As Long = -1                            'Staat voor "geen maximale kolom breedte of maximum aantal regels in voorbeeld".
Private Const GEEN_RESULTAAT As Long = -1                          'Staat voor "geen queryresultaat".
Private Const MASKER_CIJFER As String = "#"                        'Geeft in een masker aan dat er een cijfer als invoer wordt verwacht.
Private Const MASKER_HOOFDLETTER As String = "_"                   'Geeft in een masker aan dat er een hoofdletter als invoer wordt verwacht.
Private Const ONBEKEND_AANTAL As Long = -1                         'Staat voor "onbekend aantal voor de opgegeven dimensie in de opgegeven array".
Private Const PARAMETER_TEKEN As String = "?"                      'Scheidt de opdrachtregelparameters van elkaar.
Private Const SECTIE_NAAM_BEGIN As String = "["                    'Geeft het begin van een sectie naam in een instellingenbestand aan.
Private Const SECTIE_NAAM_EINDE As String = "]"                    'Geeft het einde van een sectie naam in een instellingenbestand aan.
Private Const SQL_COMMENTAAR_BLOK_BEGIN As String = "/*"           'Staat voor het begin van een SQL commentaarblok.
Private Const SQL_COMMENTAAR_BLOK_EINDE As String = "*/"           'Staat voor het einde van een SQL commentaarblok.
Private Const SQL_COMMENTAAR_REGEL_BEGIN As String = "--"          'Staat voor het begin van een SQL commentaarregel.
Private Const SQL_COMMENTAAR_REGEL_EINDE As String = vbNullString  'Staat voor het einde van een SQL commentaarregel.
Private Const SYMBOOL_TEKEN As String = "*"                        'Geeft het begin en het einde van een symbool in een tekst aan.
Private Const TEKENREEKS_TEKENS As String = "'"""                  'Staat voor de tekens die het begin en einde van een tekenreeks aanduiden.
Private Const VARIABELE_LENGTE_TEKEN As String = "*"               'Indien aanwezig aan het begin van een masker geeft dit teken aan dat de invoer lengte variabel is.
Private Const VERBINDING_PARAMETER_TEKEN As String = ";"           'Scheidt de verbindingsinformatieparameters van elkaar.
Private Const WAARDE_TEKEN As String = "="                         'Scheidt de naam en waarde van een instellingenparameter van elkaar.

'Deze procedure stuurt het aantal items min een voor de opgegeven dimensie in de opgegeven array terug.
Private Function AantalItems(ArrayV As Variant, Optional Dimensie As Long = 1) As Long
On Error GoTo Fout
Dim Aantal As Long

   Aantal = UBound(ArrayV, Dimensie) - LBound(ArrayV, Dimensie)
EindeProcedure:
   AantalItems = Aantal
   Exit Function

Fout:
   Aantal = ONBEKEND_AANTAL
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure geeft aan of de  batchmodus actief is.
Public Function BatchModusActief() As Boolean
On Error GoTo Fout
EindeProcedure:
   With Instellingen()
      BatchModusActief = Not (.BatchBereik = vbNullString Or .BatchQueryPad = vbNullString)
   End With
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure beheert bestandssysteem gerelateerde functies.
Public Function BestandsSysteem() As FileSystemObject
On Error GoTo Fout
Static HuidigBestandSysteem As New FileSystemObject
EindeProcedure:
   Set BestandsSysteem = HuidigBestandSysteem
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function


'Deze procedure bewaart de instellingen van dit programma.
Private Sub BewaarInstellingen(InstellingenPad As String, TeBewarenInstellingen As InstellingenDefinitie, Bericht As String)
On Error GoTo Fout
Dim BestandsHandle As Long

   BestandsHandle = FreeFile()
   Open InstellingenPad For Output Lock Read Write As BestandsHandle
      With TeBewarenInstellingen
         Print #BestandsHandle, SECTIE_NAAM_BEGIN & "BATCH" & SECTIE_NAAM_EINDE
         Print #BestandsHandle, "Bereik" & WAARDE_TEKEN & .BatchBereik
         Print #BestandsHandle, "Interactief" & WAARDE_TEKEN & CStr(.BatchInteractief)
         Print #BestandsHandle, "QueryPad" & WAARDE_TEKEN & .BatchQueryPad
         Print #BestandsHandle,

         Print #BestandsHandle, SECTIE_NAAM_BEGIN & "EMAILTEKST" & SECTIE_NAAM_EINDE
         Print #BestandsHandle, .EMailTekst
         Print #BestandsHandle,

         Print #BestandsHandle, SECTIE_NAAM_BEGIN & "EXPORT" & SECTIE_NAAM_EINDE
         Print #BestandsHandle, "Afzender" & WAARDE_TEKEN & .ExportAfzender
         Print #BestandsHandle, "AutoOpenen" & WAARDE_TEKEN & CStr(.ExportAutoOpenen)
         Print #BestandsHandle, "AutoOverschrijven" & WAARDE_TEKEN & CStr(.ExportAutoOverschrijven)
         Print #BestandsHandle, "AutoVerzenden" & WAARDE_TEKEN & CStr(.ExportAutoVerzenden)
         Print #BestandsHandle, "CCOntvanger" & WAARDE_TEKEN & .ExportCCOntvanger
         Print #BestandsHandle, "KolomAanvullen" & WAARDE_TEKEN & CStr(.ExportKolomAanvullen)
         Print #BestandsHandle, "Onderwerp" & WAARDE_TEKEN & .ExportOnderwerp
         Print #BestandsHandle, "Ontvanger" & WAARDE_TEKEN & .ExportOntvanger
         Print #BestandsHandle, "StandaardPad" & WAARDE_TEKEN & .ExportStandaardPad
         Print #BestandsHandle,

         Print #BestandsHandle, SECTIE_NAAM_BEGIN & "QUERY" & SECTIE_NAAM_EINDE
         Print #BestandsHandle, "AutoSluiten" & WAARDE_TEKEN & CStr(.QueryAutoSluiten)
         Print #BestandsHandle, "AutoUitvoeren" & WAARDE_TEKEN & CStr(.QueryAutoUitvoeren)
         Print #BestandsHandle, "Recordsets" & WAARDE_TEKEN & CStr(.QueryRecordSets)
         Print #BestandsHandle, "Timeout" & WAARDE_TEKEN & CStr(.QueryTimeout)
         Print #BestandsHandle,

         Print #BestandsHandle, SECTIE_NAAM_BEGIN & "VERBINDING" & SECTIE_NAAM_EINDE
         Print #BestandsHandle, .VerbindingsInformatie
         Print #BestandsHandle,

         Print #BestandsHandle, SECTIE_NAAM_BEGIN & "VOORBEELD" & SECTIE_NAAM_EINDE
         Print #BestandsHandle, "KolomBreedte" & WAARDE_TEKEN & CStr(.VoorbeeldKolomBreedte)
         Print #BestandsHandle, "Regels" & WAARDE_TEKEN & CStr(.VoorbeeldRegels)
      End With
   Close BestandsHandle

   MsgBox Bericht & vbCr & InstellingenPad, vbInformation

EindeProcedure:
   Close BestandsHandle
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False, TypePad:="Instellingenbestand: ", Pad:=InstellingenPad) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure controleert of er een fout is opgetreden tijdens de recenste API functie aanroep.
Public Function ControleerOpAPIFout(TerugGestuurd As Long, Optional ExtraInformatie As String = Empty) As Long
Dim Bericht As String
Dim FoutCode As Long
Dim Lengte As Long
Dim Omschrijving As String

   FoutCode = Err.LastDllError
   Err.Clear
   On Error GoTo Fout

   If Not FoutCode = ERROR_SUCCESS Then
      Omschrijving = String$(MAX_STRING, vbNullChar)
      Lengte = FormatMessageA(FORMAT_MESSAGE_ARGUMENT_ARRAY Or FORMAT_MESSAGE_FROM_SYSTEM, CLng(0), FoutCode, CLng(0), Omschrijving, Len(Omschrijving), StrPtr(StrConv(ExtraInformatie, vbFromUnicode)))
      If Lengte = 0 Then
         Omschrijving = "Geen omschrijving."
      Else
         Omschrijving = Left$(Omschrijving, Lengte - 1)
      End If

      Bericht = "API Foutcode: " & CStr(FoutCode) & vbCr
      Bericht = Bericht & Omschrijving
      If Not Right$(Bericht, 1) = vbCr Then Bericht = Bericht & vbCr
      Bericht = Bericht & "Terug gestuurde waarde: " & CStr(TerugGestuurd)
      MsgBox Bericht, vbExclamation
   End If
EindeProcedure:
   ControleerOpAPIFout = TerugGestuurd
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure stuurt de Microsoft Excel kolom id voor het opgegeven kolom nummer terug.
Private Function ExcelKolomId(ByVal Kolom As Long) As String
On Error GoTo Fout
Dim KolomId As String
Dim Letter1 As Long
Dim Letter2 As Long

   KolomId = vbNullString
   If Kolom > EXCEL_MAXIMUM_AANTAL_KOLOMMEN Then
      ExcelKolomId = vbNullString
      Exit Function
   End If

   For Letter1 = GEEN_LETTER To ASCII_Z
      For Letter2 = ASCII_A To ASCII_Z
         If Kolom = 0 Then
            If Letter1 = GEEN_LETTER Then
               KolomId = Chr$(Letter2)
            Else
               KolomId = Chr$(Letter1) & Chr$(Letter2)
            End If
            ExcelKolomId = KolomId
            Exit Function
         End If
         Kolom = Kolom - 1
      Next Letter2
   Next Letter1

EindeProcedure:
   ExcelKolomId = KolomId
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function


'Deze procedure exporteert het queryresultaat naar een tekstbestand.
Private Function ExporteerAlsTekst(ExportPad As String) As Boolean
On Error GoTo Fout
Dim BestandsHandle As Long
Dim EersteResultaat As Long
Dim ExportAfgebroken As Boolean
Dim Kolom As Long
Dim LaatsteResultaat As Long
Dim ResultaatIndex As Long
Dim Rij As Long

   ExportAfgebroken = False
   QueryResultaten , , , EersteResultaat, LaatsteResultaat

   BestandsHandle = FreeFile()
   Open ExportPad For Output Lock Read Write As BestandsHandle
      For ResultaatIndex = EersteResultaat To LaatsteResultaat
         With QueryResultaten(, , ResultaatIndex)
            If Not ControleerOpAPIFout(SafeArrayGetDim(.Tabel())) = 0 Then
               If Not LaatsteResultaat = EersteResultaat Then Print #BestandsHandle, "[Resultaat: #" & CStr((ResultaatIndex - EersteResultaat) + 1) & "]"
               For Rij = LBound(.Tabel(), 1) To UBound(.Tabel(), 1)
                  For Kolom = LBound(.Tabel(), 2) To UBound(.Tabel(), 2)
                     If Instellingen().ExportKolomAanvullen Then
                        Print #BestandsHandle, VulAan(.Tabel(Rij, Kolom), .KolomBreedte(Kolom), .RechtsUitlijnen(Kolom)) & " ";
                     Else
                        Print #BestandsHandle, .Tabel(Rij, Kolom); vbTab;
                     End If
                  Next Kolom
                  Print #BestandsHandle,
               Next Rij
            End If
         End With
      Next ResultaatIndex
EindeProcedure:
   Close BestandsHandle

   ExporteerAlsTekst = ExportAfgebroken
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False, TypePad:="Export pad: ", Pad:=ExportPad) = vbIgnore Then
      ExportAfgebroken = True
      Resume EindeProcedure
   End If
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure exporteert het queryresultaat naar een Microsoft Excel werkmap.
Private Function ExporteerNaarExcel(ExportPad As String, ExcelFormaat As Long) As Boolean
On Error GoTo Fout
Dim Bericht As String
Dim EersteResultaat As Long
Dim ExportAfgebroken As Boolean
Dim Kolom As Long
Dim KolomId As String
Dim LaatsteResultaat As Long
Dim MSExcel As New Excel.Application
Dim ResultaatIndex As Long
Dim WerkBlad As Excel.Worksheet
Dim WerkMap As Excel.Workbook

   ExportAfgebroken = False
   QueryResultaten , , , EersteResultaat, LaatsteResultaat

   MSExcel.DisplayAlerts = False
   MSExcel.Interactive = False
   MSExcel.ScreenUpdating = False
   MSExcel.Workbooks.Add

   Set WerkMap = MSExcel.Workbooks.Item(1)
   WerkMap.Activate

   Do Until WerkMap.Worksheets.Count <= 1
      WerkMap.Worksheets.Item(WerkMap.Worksheets.Count).Delete
   Loop

   Do Until WerkMap.Worksheets.Count >= Abs(LaatsteResultaat - EersteResultaat) + 1
      WerkMap.Worksheets.Add
   Loop

   For ResultaatIndex = EersteResultaat To LaatsteResultaat
      With QueryResultaten(, , ResultaatIndex)
         If Not ControleerOpAPIFout(SafeArrayGetDim(.Tabel())) = 0 Then
            If AantalItems(.Tabel, Dimensie:=2) > EXCEL_MAXIMUM_AANTAL_KOLOMMEN Then
               Bericht = "Het queryresultaat bevat te veel kolommen om deze naar Microsoft Excel te exporteren." & vbCr
               Bericht = Bericht & "Het maximaal toegestane aantal kolommen is: " & CStr(EXCEL_MAXIMUM_AANTAL_KOLOMMEN)
               MsgBox Bericht, vbExclamation
            Else
               Set WerkBlad = WerkMap.Worksheets.Item((ResultaatIndex - EersteResultaat) + 1)
               WerkBlad.Activate
               If Not LaatsteResultaat = EersteResultaat Then WerkBlad.Name = "Resultaat " & CStr((ResultaatIndex - EersteResultaat) + 1)
   
               WerkBlad.Range("A1:" & ExcelKolomId(AantalItems(.Tabel(), Dimensie:=2)) & CStr(AantalItems(.Tabel(), Dimensie:=1) + 1)).Value = .Tabel()
               For Kolom = LBound(.Tabel(), 2) To UBound(.Tabel(), 2)
                  KolomId = ExcelKolomId(Kolom)
                  WerkBlad.Range(KolomId & "1:" & KolomId & "1").Font.Bold = True
                  If .RechtsUitlijnen(Kolom) Then WerkBlad.Range(KolomId & "1:" & KolomId & CStr(AantalItems(.Tabel(), Dimensie:=1) + 1)).HorizontalAlignment = xlRight
               Next Kolom
               WerkBlad.Range("A:" & ExcelKolomId(AantalItems(.Tabel(), Dimensie:=2))).Columns.AutoFit
            End If
         End If
      End With

      If ResultaatIndex = LaatsteResultaat Then
         WerkMap.Worksheets.Item(1).Activate
         WerkMap.SaveAs ExportPad, ExcelFormaat
         WerkMap.Close
      End If
   Next ResultaatIndex

EindeProcedure:
   If Not MSExcel Is Nothing Then
      MSExcel.Quit
      MSExcel.DisplayAlerts = True
      MSExcel.Interactive = True
      MSExcel.ScreenUpdating = True
   End If

   Set MSExcel = Nothing
   Set WerkBlad = Nothing
   Set WerkMap = Nothing

   ExporteerNaarExcel = ExportAfgebroken
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False, TypePad:="Export pad: ", Pad:=ExportPad) = vbIgnore Then
      ExportAfgebroken = True
      Resume EindeProcedure
   End If
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure exporteert het queryresultaat.
Public Function ExporteerResultaat(ExportPad As String) As Boolean
On Error GoTo Fout
Dim BestandsType As String
Dim ExportAfgebroken As Boolean

   ExportAfgebroken = False
   BestandsType = LCase$(Trim$("." & BestandsSysteem().GetExtensionName(ExportPad)))

   If BestandsSysteem().FileExists(ExportPad) Then
      If Not Instellingen().ExportAutoOverschrijven Then
         If MsgBox("Het bestand """ & ExportPad & """ bestaat al. Overschrijven?", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbNo Then ExportAfgebroken = True
      End If

      If Not ExportAfgebroken Then
         If BestandsType = ".xls" Or BestandsType = ".xlsx" Then SluitExcelWerkmap ExportPad
         Kill ExportPad
      End If
   End If

   If Not ExportAfgebroken Then
      Select Case BestandsType
         Case ".xls"
            ExportAfgebroken = ExporteerNaarExcel(ExportPad, xlWorkbookNormal)
         Case ".xlsx"
            ExportAfgebroken = ExporteerNaarExcel(ExportPad, xlWorkbookDefault)
         Case Else
            ExportAfgebroken = ExporteerAlsTekst(ExportPad)
      End Select
   End If

EindeProcedure:
   ExporteerResultaat = Not ExportAfgebroken
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then
      ExportAfgebroken = True
      Resume EindeProcedure
   End If
   If HandelFoutAf() = vbRetry Then Resume
End Function
'Deze procedure zet de opgegeven fouten lijst om naar tekst.
Public Function FoutenLijstTekst(Lijst As Adodb.Errors) As String
On Error GoTo Fout
Dim Fout As Adodb.Error
Dim Tekst As String

   Tekst = "Er "
   If Lijst.Count = 1 Then Tekst = Tekst & "is 1 fout " Else Tekst = Tekst & "zijn " & CStr(Lijst.Count) & " fouten"
   Tekst = Tekst & " opgetreden tijdens het uitvoeren van de query:" & vbCrLf
   Tekst = Tekst & VulAan("Native", 11)
   Tekst = Tekst & VulAan("Code", 11)
   Tekst = Tekst & VulAan("Bron", 36)
   Tekst = Tekst & VulAan("SQL status", 11)
   Tekst = Tekst & "Omschrijving" & vbCrLf
   For Each Fout In Lijst
      With Fout
         Tekst = Tekst & VulAan(CStr(.NativeError), 10, LinksAanvullen:=True) & " "
         Tekst = Tekst & VulAan(CStr(.Number), 10, LinksAanvullen:=True) & " "
         Tekst = Tekst & VulAan(.Source, 35) & " "
         Tekst = Tekst & VulAan(.SqlState, 10, LinksAanvullen:=True) & " "
         Tekst = Tekst & .Description & vbCrLf
      End With
   Next Fout

EindeProcedure:
   FoutenLijstTekst = Tekst
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function


'Deze procedure handelt eventuele fouten af.
Public Function HandelFoutAf(Optional VraagVorigeKeuzeOp As Boolean = True, Optional TypePad As String = vbNullString, Optional Pad As String = vbNullString, Optional ExtraInformatie As String = vbNullString) As Long
Dim Bericht As String
Dim Bron As String
Dim FoutCode As Long
Dim FoutOmschrijving As String
Static Keuze As Long

   Bron = Err.Source
   FoutCode = Err.Number
   FoutOmschrijving = Err.Description
   Err.Clear

   On Error Resume Next

   If Not VraagVorigeKeuzeOp Then
      Bericht = MaakFoutOmschrijvingOp(FoutOmschrijving) & vbCr
      Bericht = Bericht & "Foutcode: " & CStr(FoutCode)
      If Not Bron = vbNullString Then Bericht = Bericht & vbCr & "Bron: " & Bron
      If Not (TypePad = vbNullString Or Pad = vbNullString) Then Bericht = Bericht & vbCr & TypePad & BestandsSysteem().GetAbsolutePathName(Pad)
      If Not ExtraInformatie = vbNullString Then Bericht = Bericht & vbCr & ExtraInformatie

      Keuze = MsgBox(Bericht, vbExclamation Or vbAbortRetryIgnore Or vbDefaultButton2)
   End If

   HandelFoutAf = Keuze

   If Keuze = vbAbort Then End
End Function
'Deze procedure stuurt de instellingen voor dit programma terug.
Public Function Instellingen(Optional InstellingenPad As String = vbNullString) As InstellingenDefinitie
On Error GoTo Fout
Dim Bericht As String
Static ProgrammaInstellingen As InstellingenDefinitie

   If Not InstellingenPad = vbNullString Then
      If BestandsSysteem().FileExists(InstellingenPad) Then
         ProgrammaInstellingen = LaadInstellingen(InstellingenPad)
      Else
         Bericht = "Kan het instellingenbestand niet vinden." & vbCr
         Bericht = Bericht & "Instellingenbestand: " & InstellingenPad & vbCr
         Bericht = Bericht & "Dit bestand genereren?" & vbCr
         Bericht = Bericht & "Huidig pad: " & CurDir$()
         If MsgBox(Bericht, vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes Then
            BewaarInstellingen InstellingenPad, StandaardInstellingen(), "De standaardinstellingen zijn weggeschreven naar:"
            ProgrammaInstellingen = LaadInstellingen(InstellingenPad)
         End If
      End If
   End If

EindeProcedure:
   Instellingen = ProgrammaInstellingen
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure toont instellingsbestand gerelateerde foutmeldingen.
Private Function InstellingenFout(Bericht As String, Optional InstellingenPad As String = vbNullString, Optional Sectie As String = vbNullString, Optional Regel As String = vbNullString, Optional Fataal As Boolean = False) As Long
On Error GoTo Fout
Dim Keuze As Long
Dim Stijl As Long

   If Not Sectie = vbNullString Then Bericht = Bericht & vbCr & "Sectie: " & Sectie
   If Not Regel = vbNullString Then Bericht = Bericht & vbCr & "Regel: " & """" & Regel & """"
   If Not InstellingenPad = vbNullString Then Bericht = Bericht & vbCr & "Instellingenbestand: " & InstellingenPad

   Stijl = vbExclamation
   If Not Fataal Then Stijl = Stijl Or vbOKCancel Or vbDefaultButton1
   Keuze = MsgBox(Bericht, Stijl)

EindeProcedure:
   InstellingenFout = Keuze
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure geeft aan of een interactieve batch moet worden afgebroken.
Public Function InteractieveBatchAfbreken(Optional BatchAfbreken As Variant) As Boolean
On Error GoTo Fout
Static Afbreken As Boolean

   If Not IsMissing(BatchAfbreken) Then Afbreken = CBool(BatchAfbreken)

EindeProcedure:
   InteractieveBatchAfbreken = Afbreken
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure geeft aan of de interactieve batchmodus actief is.
Public Function InteractieveBatchModusActief() As Boolean
On Error GoTo Fout
EindeProcedure:
   InteractieveBatchModusActief = Instellingen().BatchInteractief And BatchModusActief()
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure beheert de interactieve batch parameters.
Private Function InteractieveBatchParameters(Optional Index As Long = 0, Optional NieuweParameter As Variant, Optional Verwijderen As Boolean = False) As String
On Error GoTo Fout
Static Parameters As New Collection

   If Not IsMissing(NieuweParameter) Then
      Parameters.Add CStr(NieuweParameter)
   ElseIf Verwijderen Then
      Set Parameters = New Collection
   End If

EindeProcedure:
   If Parameters.Count = 0 Then
      InteractieveBatchParameters = vbNullString
   Else
      InteractieveBatchParameters = Parameters(Index + 1)
   End If
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function


'Deze procedure geeft aan of de opgegeven regel een instellingen sectie naam bevat.
Private Function IsInstellingsSectie(Regel As String) As Boolean
On Error GoTo Fout
EindeProcedure:
   IsInstellingsSectie = (Left$(Trim$(Regel), 1) = SECTIE_NAAM_BEGIN And Right$(Trim$(Regel), 1) = SECTIE_NAAM_EINDE)
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function


'Deze procedure geeft aan of het opgegeven datatype links uitgelijnd moet worden.
Private Function IsLinksUitgelijnd(DataType As Long) As Boolean
On Error GoTo Fout
Dim LinksUitgelijnd As Boolean
Dim TypeIndex As Long

   LinksUitgelijnd = False
   For TypeIndex = LBound(LinksUitgelijndeDataTypes()) To UBound(LinksUitgelijndeDataTypes())
      If DataType = LinksUitgelijndeDataTypes(TypeIndex) Then
         LinksUitgelijnd = True
         Exit For
      End If
   Next TypeIndex
EindeProcedure:
   IsLinksUitgelijnd = LinksUitgelijnd
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure voegt het opgegeven item toe aan de opgegeven lijst indien deze nog niet voorkomt.
Private Function ItemIsUniek(ByRef Lijst As Collection, Optional Item As Variant, Optional ResetLijst As Boolean = False) As Boolean
On Error GoTo Fout
Dim Index As Long
Dim Uniek As Boolean

   Uniek = True

   If ResetLijst Then
      Set Lijst = New Collection
   ElseIf Not IsMissing(Item) Then
      For Index = 1 To Lijst.Count
         If Lijst(Index) = Item Then
            Uniek = False
            Exit For
         End If
      Next Index

      If Uniek Then Lijst.Add Item
   End If

EindeProcedure:
   ItemIsUniek = Uniek
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure laadt de instellingen voor dit programma.
Private Function LaadInstellingen(InstellingenPad As String) As InstellingenDefinitie
On Error GoTo Fout
Dim Afbreken As Boolean
Dim BestandsHandle As Long
Dim ParameterNaam As String
Dim ProgrammaInstellingen As InstellingenDefinitie
Dim RecensteGeldigeSectie As String
Dim Regel As String
Dim Sectie As String
Dim VerbindingsInformatie As String
Dim VerwerkteParameters As New Collection
Dim VerwerkteSecties As New Collection

   Afbreken = False
   ItemIsUniek VerwerkteParameters, , ResetLijst:=True
   ItemIsUniek VerwerkteSecties, , ResetLijst:=True
   ProgrammaInstellingen = StandaardInstellingen()
   RecensteGeldigeSectie = vbNullString
   Sectie = vbNullString

   With ProgrammaInstellingen
      .Bestand = InstellingenPad
      BestandsHandle = FreeFile()
      Open .Bestand For Input Lock Read Write As BestandsHandle
         Do Until EOF(BestandsHandle) Or Afbreken
            Line Input #BestandsHandle, Regel

            If Not Left$(Trim$(Regel), 1) = COMMENTAAR_TEKEN Then
               If IsInstellingsSectie(Regel) Then
                  Regel = Trim$(Regel)
                  RecensteGeldigeSectie = Sectie
                  Sectie = UCase$(Mid$(Regel, Len(SECTIE_NAAM_BEGIN) + 1, Len(Regel) - (Len(SECTIE_NAAM_BEGIN) + Len(SECTIE_NAAM_EINDE))))
                  If Not ItemIsUniek(VerwerkteSecties, Sectie) Then
                     If InstellingenFout("Sectie is meerdere keren aanwezig.", InstellingenPad, Sectie, Regel) = vbCancel Then Afbreken = True
                  End If
                  ItemIsUniek VerwerkteParameters, , ResetLijst:=True
               Else
                  Select Case Sectie
                     Case "BATCH", "EXPORT", "QUERY", "VOORBEELD"
                        If Not Trim$(Regel) = vbNullString Then
                           LeesParameter Regel, ParameterNaam
                           If Not ItemIsUniek(VerwerkteParameters, ParameterNaam) Then
                              If InstellingenFout("Parameter is meerdere keren aanwezig.", InstellingenPad, Sectie, Regel) = vbCancel Then Afbreken = True
                           End If
                        End If
                  End Select
               End If

               Select Case Sectie
                  Case "BATCH"
                     If Not (IsInstellingsSectie(Regel) Or Trim$(Regel) = vbNullString) Then
                        If Not VerwerkBatchInstellingen(Regel, Sectie, ProgrammaInstellingen) Then Afbreken = True
                     End If
                  Case "EMAILTEKST"
                     If Not IsInstellingsSectie(Regel) Then .EMailTekst = .EMailTekst & Regel & vbCrLf
                  Case "EXPORT"
                     If Not (IsInstellingsSectie(Regel) Or Trim$(Regel) = vbNullString) Then
                        If Not VerwerkExportInstellingen(Regel, Sectie, ProgrammaInstellingen) Then Afbreken = True
                     End If
                  Case "QUERY"
                     If Not (IsInstellingsSectie(Regel) Or Trim$(Regel) = vbNullString) Then
                        If Not VerwerkQueryInstellingen(Regel, Sectie, ProgrammaInstellingen) Then Afbreken = True
                     End If
                  Case "VERBINDING"
                     If Not (IsInstellingsSectie(Regel) Or Trim$(Regel) = vbNullString) Then .VerbindingsInformatie = .VerbindingsInformatie & Trim$(Regel)
                  Case "VOORBEELD"
                     If Not (IsInstellingsSectie(Regel) Or Trim$(Regel) = vbNullString) Then
                        If Not VerwerkVoorbeeldInstellingen(Regel, Sectie, ProgrammaInstellingen) Then Afbreken = True
                     End If
                  Case Else
                     If Not Trim$(Regel) = vbNullString Then
                        If IsInstellingsSectie(Regel) Then
                           Sectie = RecensteGeldigeSectie
                           If InstellingenFout("Niet herkende sectie.", InstellingenPad, Sectie, Regel) = vbCancel Then Afbreken = True
                        Else
                           If InstellingenFout("Niet herkende parameter.", InstellingenPad, Sectie, Regel) = vbCancel Then Afbreken = True
                        End If
                     End If
               End Select
            End If
         Loop
      Close BestandsHandle
      
      If Trim$(.VerbindingsInformatie) = vbNullString And Not Afbreken Then
         VerbindingsInformatie = Trim$(VraagVerbindingsInformatie())
         If Not VerbindingsInformatie = vbNullString Then
            .VerbindingsInformatie = VerbindingsInformatie
            BewaarInstellingen InstellingenPad, ProgrammaInstellingen, "De instellingen zijn weggeschreven naar:"
         End If
      End If

      .VerbindingsInformatie = MaakVerbindingsInformatieOp(.VerbindingsInformatie)
   End With

EindeProcedure:
   Close BestandsHandle

   LaadInstellingen = ProgrammaInstellingen
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False, TypePad:="Instellingenbestand: ", Pad:=InstellingenPad) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function
'Deze procedure stuurt de waarde en de naam van een instellingen parameter in de opgegeven regel terug.
Private Function LeesParameter(Regel As String, ByRef ParameterNaam As String) As String
On Error GoTo Fout
Dim Positie As Long
Dim Waarde As String

   ParameterNaam = vbNullString
   Waarde = vbNullString
   Positie = InStr(Regel, WAARDE_TEKEN)
   If Positie > 0 Then
      ParameterNaam = LCase$(Trim$(Left$(Regel, Positie - 1)))
      Waarde = Trim$(Mid$(Regel, Positie + 1))
   End If

EindeProcedure:
   LeesParameter = Waarde
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure stuurt een lijst van databasedatatypes die linksuitgelijnd worden terug.
Private Function LinksUitgelijndeDataTypes() As Variant
On Error GoTo Fout
EindeProcedure:
   LinksUitgelijndeDataTypes = Array(adBSTR, adChar, adDBDate, adDBTime, adDBTimeStamp, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar)
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function


'Deze procedure maakt de opgegeven fout omschrijving op.
Private Function MaakFoutOmschrijvingOp(FoutOmschrijving As String) As String
On Error Resume Next
Dim Omschrijving As String

   Omschrijving = Trim$(FoutOmschrijving)
   Do
      Select Case Right$(Omschrijving, 1)
         Case vbCr, vbLf
            Omschrijving = Trim$(Left$(Omschrijving, Len(Omschrijving) - 1))
         Case Else
            Exit Do
      End Select
   Loop
   If Not Right$(Omschrijving, 1) = "." Then Omschrijving = Omschrijving & "."

MaakFoutOmschrijvingOp = Omschrijving
End Function

'Deze procedure controleert de opgegeven verbindingsinformatie en maakt deze op.
Private Function MaakVerbindingsInformatieOp(VerbindingsInformatie As String) As String
On Error GoTo Fout
Dim HuidigStringTeken As String
Dim OpgemaakteVerbindingsInformatie As String
Dim Parameter As String
Dim ParameterBegin As Long
Dim ParameterNaam As String
Dim ParameterNamen As Collection
Dim Positie As Long
Dim Teken As String
Dim Waarde As String

   OpgemaakteVerbindingsInformatie = vbNullString
  
   If Not Trim$(VerbindingsInformatie) = vbNullString Then
      HuidigStringTeken = vbNullString
      ItemIsUniek ParameterNamen, , ResetLijst:=True
      Positie = 1
      ParameterBegin = Positie
      If Not Right$(Trim$(VerbindingsInformatie), Len(VERBINDING_PARAMETER_TEKEN)) = VERBINDING_PARAMETER_TEKEN Then VerbindingsInformatie = VerbindingsInformatie & VERBINDING_PARAMETER_TEKEN
      Do Until Positie > Len(VerbindingsInformatie)
         Teken = Mid$(VerbindingsInformatie, Positie, 1)
         If InStr(TEKENREEKS_TEKENS, Teken) > 0 Then
            If HuidigStringTeken = vbNullString Then
               HuidigStringTeken = Teken
            ElseIf Teken = HuidigStringTeken Then
               HuidigStringTeken = vbNullString
            End If
         ElseIf Teken = VERBINDING_PARAMETER_TEKEN Then
            If HuidigStringTeken = vbNullString Then
               Parameter = Mid$(VerbindingsInformatie, ParameterBegin, Positie - ParameterBegin)

               If InStr(Parameter, WAARDE_TEKEN) = 0 Then
                  OpgemaakteVerbindingsInformatie = vbNullString
                  MsgBox "Ongeldige parameter aanwezig in verbindingsinformatie: """ & Parameter & """. Verwacht teken: " & WAARDE_TEKEN, vbExclamation
                  Exit Do
               End If

               Waarde = LeesParameter(Parameter, ParameterNaam)

               If Not ItemIsUniek(ParameterNamen, ParameterNaam) Then
                  OpgemaakteVerbindingsInformatie = vbNullString
                  MsgBox "Parameter meerdere malen aanwezig in verbindingsinformatie: """ & Parameter & """.", vbExclamation
                  Exit Do
               End If
               ParameterBegin = Positie + 1

               OpgemaakteVerbindingsInformatie = OpgemaakteVerbindingsInformatie & ParameterNaam & WAARDE_TEKEN & Trim$(Waarde) & VERBINDING_PARAMETER_TEKEN
            End If
         End If
   
         Positie = Positie + 1
      Loop
   
      If Not HuidigStringTeken = vbNullString Then
         OpgemaakteVerbindingsInformatie = vbNullString
         MsgBox "Niet afgesloten tekenreekswaarde in verbindingsgegevens. Verwacht teken: " & HuidigStringTeken, vbExclamation
      End If
   End If

EindeProcedure:
   MaakVerbindingsInformatieOp = OpgemaakteVerbindingsInformatie
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function




'Deze procedure wordt uitgevoerd wanneer het programma wordt gestart.
Private Sub Main()
On Error GoTo Fout
Dim InstellingenPad As String
   ControleerOpAPIFout SetCurrentDirectoryA(Left$(App.Path, InStr(App.Path, ":")))
   ChDir App.Path

   With OpdrachtRegelParameters(Command$())
      If .Verwerkt Then
         If Left$(Trim$(.InstellingenPad), Len(PARAMETER_TEKEN)) = PARAMETER_TEKEN Then
            InstellingenPad = VerwijderAanhalingsTekens(Mid$(Trim$(.InstellingenPad), Len(PARAMETER_TEKEN) + 1))
            If InstellingenPad = vbNullString Then
               MsgBox "Kan de instellingen niet bewaren. Geen doel bestand opgegeven.", vbExclamation
            Else
               BewaarInstellingen InstellingenPad, StandaardInstellingen(), "De standaardinstellingen zijn weggeschreven naar:"
            End If
         ElseIf Not .SessiesPad = vbNullString Then
            SessieParameters , , Verwijderen:=True
            VerwerkSessieLijst .SessiesPad
         Else
            SessieParameters , , Verwijderen:=True
            VoerSessieUit Command$()
         End If
      End If
   End With

EindeProcedure:
   Verbinding , VerbindingSluiten:=True
   SluitAlleVensters
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub
'Deze procedure controleert de queryparameter invoer en stuurt eventueel de index van een onjuist ingevuld veld en een fout omschrijving terug.
Private Function OngeldigeParameterInvoer(Optional ByRef FoutInformatie As String = vbNullString) As Long
On Error GoTo Fout
Dim EersteParameter As Long
Dim Index As Long
Dim LaatsteParameter As Long
Dim Lengte As Long
Dim OngeldigVeld As Long
Dim Positie As Long

   QueryParameters , , , EersteParameter, LaatsteParameter
   OngeldigVeld = GEEN_PARAMETER

   For Index = EersteParameter To LaatsteParameter
      With QueryParameters(, Index)
         If .Masker = vbNullString Then
            Lengte = Len(.Invoer)
         Else
            If .LengteIsVariabel Then Lengte = ParameterInvoerLengte(Index) Else Lengte = Len(.Masker)
            For Positie = 1 To Lengte
               FoutInformatie = ParameterMaskerTekenGeldig(Mid$(.Invoer, Positie, 1), Mid$(.Masker, Positie, 1))
               If Not FoutInformatie = vbNullString Then
                  FoutInformatie = vbCr & """" & FoutInformatie & """." & vbCr & "Teken positie: " & CStr(Positie)
                  OngeldigVeld = Index
                  Exit For
               End If
            Next Positie
         End If

         If Not OngeldigVeld = GEEN_PARAMETER Then Exit For
         QueryParameters , Index, Left$(.Invoer, Lengte)
      End With
   Next Index

   If Not OngeldigVeld = GEEN_PARAMETER Then
      For Index = EersteParameter To LaatsteParameter
         QueryParameters , Index, vbNullString
      Next Index
   End If

EindeProcedure:
   OngeldigeParameterInvoer = OngeldigVeld
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure beheert de huidige sessie's opdrachtregelparameters.
Public Function OpdrachtRegelParameters(Optional SessieParameters As String = vbNullString) As OpdrachtRegelParametersDefinitie
On Error GoTo Fout
Dim Bericht As String
Dim Extensies As Collection
Dim Parameter As Variant
Dim Parameters() As String
Dim Positie As Long
Static HuidigeOpdrachtRegelParameters As OpdrachtRegelParametersDefinitie

   With HuidigeOpdrachtRegelParameters
      .Verwerkt = True

      If Not SessieParameters = vbNullString Then
         ItemIsUniek Extensies, , ResetLijst:=True

         Positie = InStr(SessieParameters, PARAMETER_TEKEN & PARAMETER_TEKEN)
         If Positie > 0 Then
            .InstellingenPad = Mid$(SessieParameters, Positie + Len(PARAMETER_TEKEN))
         Else
            Parameters = Split(SessieParameters, PARAMETER_TEKEN)

            For Each Parameter In Parameters
               If Not Trim$(Parameter) = vbNullString Then
                  Parameter = VerwijderAanhalingsTekens(CStr(Parameter))
   
                  If ItemIsUniek(Extensies, "." & LCase$(BestandsSysteem().GetExtensionName(CStr(Parameter)))) Then
                     Select Case "." & LCase$(BestandsSysteem().GetExtensionName(CStr(Parameter)))
                        Case ".ini"
                           .InstellingenPad = Parameter
                        Case ".lst"
                           .SessiesPad = Parameter
                        Case ".txt"
                           .QueryPad = Parameter
                        Case Else
                           If Not Trim$(Parameter) = vbNullString Then
                              Bericht = "Niet herkende opdrachtregelparameter: """ & Parameter & """."
                              If VerwerkSessieLijst() = vbNullString Then
                                 MsgBox Bericht, vbExclamation
                              Else
                                 Bericht = Bericht & vbCr & "Sessielijst: """ & VerwerkSessieLijst() & """."
                                 If MsgBox(Bericht, vbExclamation Or vbOKCancel) = vbCancel Then SessiesAfbreken NieuweSessiesAfbreken:=True
                              End If
                              .Verwerkt = False
                           End If
                     End Select
                  Else
                     Bericht = "Er kan maar een instellingenbestand en/of query tegelijk opgegeven worden."
                     If Not VerwerkSessieLijst() = vbNullString Then Bericht = Bericht & vbCr & "Sessielijst: """ & VerwerkSessieLijst() & """."
                     MsgBox Bericht, vbExclamation
                     .Verwerkt = False
                  End If
               End If
            Next Parameter
         End If
      End If
   End With

EindeProcedure:
   OpdrachtRegelParameters = HuidigeOpdrachtRegelParameters
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure controleert de door de gebruiker ingevoerde parameters en stuurt het resultaat terug.
Public Function ParametersGeldig(ParameterVelden As Object) As Boolean
On Error GoTo Fout
Dim FoutInformatie As String
Dim Geldig As Boolean
Dim Index As Long

   For Index = ParameterVelden.LBound To ParameterVelden.UBound
      QueryParameters , Index, ParameterVelden(Index).Text
   Next Index

   Index = OngeldigeParameterInvoer(FoutInformatie)
   Geldig = (Index = GEEN_PARAMETER)

   If Not Geldig Then
      With ParameterVelden(Index)
         If .Visible Then
            FoutInformatie = "Dit veld is niet volledig of onjuist ingevuld:" & FoutInformatie
         Else
            FoutInformatie = "Onzichtbare parameter #" & CStr(Index - ParameterVelden.LBound) & " is niet volledig of onjuist ingevuld:" & FoutInformatie
         End If
         MsgBox FoutInformatie, vbExclamation
         If .Visible Then .SetFocus
      End With
   End If

EindeProcedure:
   ParametersGeldig = Geldig
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure vergelijkt het opgegeven teken met het opgegeven queryparametermaskerteken.
Public Function ParameterMaskerTekenGeldig(Teken As String, MaskerTeken As String) As String
On Error GoTo Fout
Dim Geldig As String

   Geldig = vbNullString

   Select Case MaskerTeken
      Case MASKER_CIJFER
         If Not (Teken >= "0" And Teken <= "9") Then Geldig = "Cijfer verwacht."
      Case MASKER_HOOFDLETTER
         If Not (Teken >= "A" And Teken <= "Z") Then Geldig = "Hoofdletter verwacht."
      Case Else
         If Not Teken = MaskerTeken Then Geldig = "Vast maskerteken """ & MaskerTeken & """ verwacht."
   End Select

EindeProcedure:
   ParameterMaskerTekenGeldig = Geldig
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function


'Deze procedure stuurt de lengte van de invoer voor de opgegeven queryparameter terug.
Private Function ParameterInvoerLengte(Index As Long) As Long
On Error GoTo Fout
Dim Lengte As Long
Dim Positie As Long

   Lengte = 0
   With QueryParameters(, Index)
      For Positie = 1 To Len(.Invoer)
         If Not Mid$(.Invoer, Positie, 1) = Mid$(.Masker, Positie, 1) Then Lengte = Positie
      Next Positie
   End With

EindeProcedure:
   ParameterInvoerLengte = Lengte
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function


'Deze procedure toont parameter en/of symbool gerelateerde foutmeldingen.
Private Sub ParameterSymboolFout(Bericht As String, Optional Index As Long = GEEN_PARAMETER)
On Error GoTo Fout
Dim EersteParameter As Long

   QueryParameters , , , EersteParameter

   If Not Index = GEEN_PARAMETER Then
      Bericht = Bericht & vbCr & "Parameter definitie: #" & CStr((Index - EersteParameter) + 1)
      With QueryParameters(, Index)
         If Not .ParameterNaam = vbNullString Then Bericht = Bericht & vbCr & "Naam: """ & .ParameterNaam & """"
         If Not .Invoer = vbNullString Then Bericht = Bericht & vbCr & "Invoer: """ & .Invoer & """"
         If Not .StandaardWaarde = vbNullString Then Bericht = Bericht & vbCr & "Standaardwaarde: """ & .StandaardWaarde & """"
         If Not .Masker = vbNullString Then Bericht = Bericht & vbCr & "Masker: """ & .Masker & """"
      End With
   End If
   If Not Query().Pad = vbNullString Then Bericht = Bericht & vbCr & "Query: " & Query().Pad
   MsgBox Bericht, vbExclamation
EindeProcedure:
Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub



'Deze procedure stuurt het versienummer van dit programma terug.
Public Function ProgrammaVersie() As String
On Error GoTo Fout
EindeProcedure:
   With App
      ProgrammaVersie = "v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision)
   End With
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure laadt de opgegeven query of stuurt een al geladen query terug.
Public Function Query(Optional QueryPad As String = vbNullString) As QueryDefinitie
On Error GoTo Fout
Dim BestandsHandle As Long
Static HuidigeQuery As QueryDefinitie

   With HuidigeQuery
      .Geopend = False

      If Not QueryPad = vbNullString Then
         BestandsHandle = FreeFile()
         Open QueryPad For Input Lock Read Write As BestandsHandle: Close BestandsHandle

         BestandsHandle = FreeFile()
         Open QueryPad For Binary Lock Read Write As BestandsHandle
            .Code = Input$(LOF(BestandsHandle), BestandsHandle)
         Close BestandsHandle

         .Pad = QueryPad
         .Geopend = True
      End If
   End With

EindeProcedure:
   Query = HuidigeQuery
   Exit Function

Fout:
   HuidigeQuery.Geopend = False

   If HandelFoutAf(VraagVorigeKeuzeOp:=False, TypePad:="Query pad: ", Pad:=QueryPad) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure doorzoekt de opgegeven query op parameterdefinities of stuurt een eerder gevonden parameterdefinitie terug.
Public Function QueryParameters(Optional QueryCode As String = vbNullString, Optional ParameterIndex As Long = 0, Optional Invoer As Variant, Optional ByRef EersteParameter As Long = 0, Optional ByRef LaatsteParameter As Long = 0) As QueryParameterDefinitie
On Error GoTo Fout
Dim Definitie As String
Dim DefinitieBegin As Long
Dim DefinitieEinde As Long
Dim Elementen() As String
Dim OnverwerkteCode As String
Static Parameters() As QueryParameterDefinitie

   If Not IsMissing(Invoer) Then
      If Not ControleerOpAPIFout(SafeArrayGetDim(Parameters)) = 0 Then Parameters(ParameterIndex).Invoer = CStr(Invoer)
   ElseIf Not QueryCode = vbNullString Then
      Erase Parameters()

      OnverwerkteCode = QueryCode
      Do
         DefinitieBegin = InStr(OnverwerkteCode, DEFINITIE_TEKENS)
         If DefinitieBegin > 0 Then
            DefinitieEinde = InStr(DefinitieBegin + Len(DEFINITIE_TEKENS), OnverwerkteCode, DEFINITIE_TEKENS)
            If DefinitieEinde > 0 Then
               If ControleerOpAPIFout(SafeArrayGetDim(Parameters())) = 0 Then
                  ReDim Parameters(0 To 0) As QueryParameterDefinitie
               Else
                  ReDim Preserve Parameters(LBound(Parameters()) To UBound(Parameters()) + 1) As QueryParameterDefinitie
               End If

               Definitie = Mid$(OnverwerkteCode, DefinitieBegin + Len(DEFINITIE_TEKENS), (DefinitieEinde - DefinitieBegin) - Len(DEFINITIE_TEKENS))
               OnverwerkteCode = Mid$(OnverwerkteCode, DefinitieEinde + Len(DEFINITIE_TEKENS))

               With Parameters(UBound(Parameters()))
                  .Lengte = Len(DEFINITIE_TEKENS & Definitie & DEFINITIE_TEKENS)
                  .Positie = DefinitieBegin

                  Elementen = Split(Definitie, ELEMENT_TEKEN)
                  If AantalItems(Elementen) > Abs(CommentaarElement - NaamElement) Then ParameterSymboolFout "Teveel elementen, deze worden genegeerd.", UBound(Parameters())
                  ReDim Preserve Elementen(NaamElement To CommentaarElement) As String

                  .ParameterNaam = Elementen(NaamElement)

                  .VeldIsZichtbaar = Not (.ParameterNaam = vbNullString)

                  .Masker = Elementen(MaskerElement)
                  .LengteIsVariabel = (Left$(.Masker, 1) = VARIABELE_LENGTE_TEKEN)
                  If .LengteIsVariabel Then .Masker = Mid$(.Masker, 2)

                  .StandaardWaarde = VervangSymbolen(Elementen(StandaardWaardeElement))
                  If Not .Masker = vbNullString Then If Len(.StandaardWaarde) > Len(.Masker) Then ParameterSymboolFout "De standaardwaarde is langer dan het masker. De overtollige tekens worden verwijderd.", UBound(Parameters())

                  .Commentaar = Elementen(CommentaarElement)

                  .Invoer = .StandaardWaarde
               End With

            Else
               ParameterSymboolFout "Geen einde markering. Deze wordt genegeerd.", UBound(Parameters())
               Exit Do
            End If
         Else
            Exit Do
         End If
      Loop
   End If

EindeProcedure:
   EersteParameter = GEEN_PARAMETER
   LaatsteParameter = GEEN_PARAMETER
   If Not ControleerOpAPIFout(SafeArrayGetDim(Parameters())) = 0 Then
      EersteParameter = LBound(Parameters())
      LaatsteParameter = UBound(Parameters())
      QueryParameters = Parameters(ParameterIndex)
   End If
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure handelt eventuele queryresultaat lees fouten af.
Public Function QueryResultaatLeesFout(Optional Rij As Long = 0, Optional Kolom As Long = 0, Optional KolomNaam As String = vbNullString, Optional VraagVorigeKeuzeOp As Boolean = True) As Long
Dim Bericht As String
Dim Bron As String
Dim FoutCode As Long
Dim FoutOmschrijving As String
Static Keuze As Long

   Bron = Err.Source
   FoutCode = Err.Number
   FoutOmschrijving = Err.Description
   Err.Clear

   On Error Resume Next

   If Not VraagVorigeKeuzeOp Then
      Bericht = "Er is een fout opgetreden bij het uitlezen van het queryresultaat." & vbCr
      Bericht = Bericht & "Rij: " & CStr(Rij) & vbCr
      Bericht = Bericht & "Kolom: " & CStr(Kolom) & vbCr
      Bericht = Bericht & "Kolom naam: " & CStr(KolomNaam) & vbCr
      Bericht = Bericht & "Omschrijving: " & MaakFoutOmschrijvingOp(FoutOmschrijving) & vbCr
      Bericht = Bericht & "Foutcode: " & CStr(FoutCode)
      If Not Bron = vbNullString Then Bericht = Bericht & vbCr & "Bron: " & Bron

      Keuze = MsgBox(Bericht, vbExclamation Or vbAbortRetryIgnore Or vbDefaultButton2)
   End If

   QueryResultaatLeesFout = Keuze
End Function



'Deze procedure stuurt het queryresultaat terug als tekst.
Public Function QueryResultaatTekst(Resultaat As QueryResultaatDefinitie) As String
On Error GoTo Fout
Dim Breedte As Long
Dim Kolom As Long
Dim LaatsteRegel As Long
Dim ResultaatTekst As String
Dim Rij As Long
Dim Tekst As String

   With Resultaat
      If Not ControleerOpAPIFout(SafeArrayGetDim(.Tabel())) = 0 Then
         If Instellingen().VoorbeeldRegels = GEEN_MAXIMUM Or Instellingen().VoorbeeldRegels > AantalItems(.Tabel(), Dimensie:=1) Then
            LaatsteRegel = AantalItems(.Tabel(), Dimensie:=1)
         Else
            LaatsteRegel = Instellingen().VoorbeeldRegels - 1
         End If
   
         ResultaatTekst = vbNullString
         For Rij = LBound(.Tabel(), 1) To LaatsteRegel
            For Kolom = LBound(.Tabel(), 2) To UBound(.Tabel(), 2)
               Breedte = .KolomBreedte(Kolom)
               Tekst = .Tabel(Rij, Kolom)
               Tekst = Replace(Tekst, vbCr, " ")
               Tekst = Replace(Tekst, vbLf, " ")
               Tekst = Replace(Tekst, vbTab, " ")
               If Not Instellingen().VoorbeeldKolomBreedte = GEEN_MAXIMUM Then
                  If .KolomBreedte(Kolom) > Instellingen().VoorbeeldKolomBreedte Then
                     Breedte = Instellingen().VoorbeeldKolomBreedte
                     Tekst = Left$(Tekst, Instellingen().VoorbeeldKolomBreedte)
                  End If
               End If
   
               ResultaatTekst = ResultaatTekst & VulAan(Tekst, Breedte, .RechtsUitlijnen(Kolom)) & " "
            Next Kolom
            ResultaatTekst = ResultaatTekst & vbCrLf
            If DoEvents() = 0 Then Exit For
         Next Rij
      End If
   End With

EindeProcedure:
   QueryResultaatTekst = ResultaatTekst
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure beheert de queryresulaten.
Public Function QueryResultaten(Optional NieuwQueryResultaat As Adodb.Recordset = Nothing, Optional ResultatenVerwijderen As Boolean = False, Optional ResultaatIndex As Long = 0, Optional ByRef EersteResultaat As Long = 0, Optional ByRef LaatsteResultaat As Long = 0) As QueryResultaatDefinitie
On Error GoTo Fout
Dim Kolom As Long
Dim Rij As Long
Dim TijdelijkeTabel() As String
Static Resultaten() As QueryResultaatDefinitie

   If Not NieuwQueryResultaat Is Nothing Then
      With NieuwQueryResultaat
         If Not .BOF Then
            If ControleerOpAPIFout(SafeArrayGetDim(Resultaten())) = 0 Then
               ReDim Resultaten(0 To 0) As QueryResultaatDefinitie
            Else
               ReDim Preserve Resultaten(LBound(Resultaten()) To UBound(Resultaten()) + 1) As QueryResultaatDefinitie
            End If

            With Resultaten((UBound(Resultaten())))
               If ControleerOpAPIFout(SafeArrayGetDim(.KolomBreedte())) = 0 Then ReDim .KolomBreedte(0 To 0) As Long
               If ControleerOpAPIFout(SafeArrayGetDim(.RechtsUitlijnen())) = 0 Then ReDim .RechtsUitlijnen(0 To 0) As Boolean
               If ControleerOpAPIFout(SafeArrayGetDim(.Tabel())) = 0 Then ReDim .Tabel(0 To 0, 0 To 0) As String
            End With

            Rij = 0
            ReDim Resultaten(UBound(Resultaten())).KolomBreedte(0 To .Fields.Count - 1) As Long
            ReDim Resultaten(UBound(Resultaten())).RechtsUitlijnen(0 To .Fields.Count - 1) As Boolean
            ReDim TijdelijkeTabel(0 To .Fields.Count - 1, 0 To Rij) As String
            For Kolom = 0 To .Fields.Count - 1
               TijdelijkeTabel(Kolom, Rij) = Trim$(.Fields.Item(Kolom).Name)
               Resultaten(UBound(Resultaten())).KolomBreedte(Kolom) = Len(TijdelijkeTabel(Kolom, Rij))
               Resultaten(UBound(Resultaten())).RechtsUitlijnen(Kolom) = Not IsLinksUitgelijnd(.Fields.Item(Kolom).Type)
            Next Kolom
            Rij = Rij + 1

            On Error GoTo LeesFout
            ReDim Preserve TijdelijkeTabel(LBound(TijdelijkeTabel(), 1) To .Fields.Count - 1, LBound(TijdelijkeTabel(), 2) To Rij) As String
            Do Until .EOF
               For Kolom = 0 To .Fields.Count - 1
                  If Not IsNull(.Fields.Item(Kolom).Value) Then
                     TijdelijkeTabel(Kolom, Rij) = Trim$(.Fields.Item(Kolom).Value)
                     If Len(TijdelijkeTabel(Kolom, Rij)) > Resultaten(UBound(Resultaten())).KolomBreedte(Kolom) Then Resultaten(UBound(Resultaten())).KolomBreedte(Kolom) = Len(TijdelijkeTabel(Kolom, Rij)) + 1
                  End If
VolgendeWaarde:
               Next Kolom
               .MoveNext
               Rij = Rij + 1
               ReDim Preserve TijdelijkeTabel(LBound(TijdelijkeTabel(), 1) To .Fields.Count - 1, LBound(TijdelijkeTabel(), 2) To Rij) As String
            Loop
            On Error GoTo 0
EindeUitlezen:

            ReDim Resultaten(UBound(Resultaten())).Tabel(0 To Rij, 0 To .Fields.Count - 1) As String
            For Rij = LBound(Resultaten(UBound(Resultaten())).Tabel(), 1) To UBound(Resultaten(UBound(Resultaten())).Tabel(), 1) - 1
               For Kolom = LBound(Resultaten(UBound(Resultaten())).Tabel(), 2) To UBound(Resultaten(UBound(Resultaten())).Tabel(), 2)
                  Resultaten(UBound(Resultaten())).Tabel(Rij, Kolom) = TijdelijkeTabel(Kolom, Rij)
               Next Kolom
            Next Rij
         End If
      End With
   ElseIf ResultatenVerwijderen Then
      Erase Resultaten()
   End If

EindeProcedure:
   EersteResultaat = GEEN_RESULTAAT
   LaatsteResultaat = GEEN_RESULTAAT
   If Not ControleerOpAPIFout(SafeArrayGetDim(Resultaten())) = 0 Then
      EersteResultaat = LBound(Resultaten())
      LaatsteResultaat = UBound(Resultaten())
      QueryResultaten = Resultaten(ResultaatIndex)
   End If
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
Exit Function

LeesFout:
   If QueryResultaatLeesFout(Rij, Kolom, TijdelijkeTabel(Kolom, 0), VraagVorigeKeuzeOp:=False) = vbAbort Then Resume EindeUitlezen
   If QueryResultaatLeesFout() = vbIgnore Then Resume VolgendeWaarde
   If QueryResultaatLeesFout() = vbRetry Then Resume
End Function

'Deze procedure geeft aan of een reeks sessies is afgebroken.
Public Function SessiesAfbreken(Optional NieuweSessiesAfbreken As Variant) As Boolean
On Error GoTo Fout
Static HuidigeSessiesAfbreken As Boolean
   
   If Not IsMissing(NieuweSessiesAfbreken) Then HuidigeSessiesAfbreken = CBool(NieuweSessiesAfbreken)

EindeProcedure:
   SessiesAfbreken = HuidigeSessiesAfbreken
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure beheert de sessie parameters.
Private Function SessieParameters(Optional Index As Long = 0, Optional NieuweParameter As Variant, Optional Verwijderen As Boolean = False) As String
On Error GoTo Fout
Static Parameters As New Collection

   If Not IsMissing(NieuweParameter) Then
      Parameters.Add CStr(NieuweParameter)
   ElseIf Verwijderen Then
      Set Parameters = New Collection
   End If

EindeProcedure:
   If Parameters.Count = 0 Then
      SessieParameters = vbNullString
   Else
      SessieParameters = Parameters(Index + 1)
   End If
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function




'Deze procedure sluit alle eventueel geopende vensters af.
Public Sub SluitAlleVensters()
On Error GoTo Fout
Dim Venster As Form
   
   For Each Venster In Forms
      Unload Venster
   Next Venster
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure sluit de werkmap met opgegeven pad als deze geopend is in Microsoft Excel.
Private Sub SluitExcelWerkmap(Pad As String)
On Error GoTo Fout
Dim WerkMap As Excel.Workbook

   Set WerkMap = GetObject(Pad)
   WerkMap.Close SaveChanges:=False
   Set WerkMap = Nothing
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False, TypePad:="Pad: ", Pad:=Pad) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure stuurt de standaardinstellingen voor dit programma terug.
Private Function StandaardInstellingen() As InstellingenDefinitie
On Error GoTo Fout
Dim ProgrammaInstellingen As InstellingenDefinitie

   With ProgrammaInstellingen
      .BatchBereik = vbNullString
      .BatchInteractief = False
      .BatchQueryPad = vbNullString
      .Bestand = "Qa.ini"
      .EMailTekst = vbNullString
      .ExportAfzender = vbNullString
      .ExportAutoOpenen = False
      .ExportAutoOverschrijven = False
      .ExportAutoVerzenden = False
      .ExportCCOntvanger = vbNullString
      .ExportKolomAanvullen = False
      .ExportOnderwerp = vbNullString
      .ExportOntvanger = vbNullString
      .ExportStandaardPad = ".\Export.xls"
      .QueryAutoSluiten = False
      .QueryAutoUitvoeren = False
      .QueryRecordSets = False
      .QueryTimeout = 10
      .VerbindingsInformatie = vbNullString
      .VoorbeeldKolomBreedte = -1
      .VoorbeeldRegels = 10
   End With

EindeProcedure:
   StandaardInstellingen = ProgrammaInstellingen
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure stuurt de status van het queryresultaat terug nadat een query is uitgevoerd.
Public Function StatusNaQuery(ResultaatIndex As Long) As String
On Error GoTo Fout
Dim AantalKolommen As Long
Dim AantalResultaten As Long
Dim AantalRijen As Long
Dim EersteResultaat As Long
Dim LaatsteResultaat As Long
Dim Status As String

   With QueryResultaten(, , ResultaatIndex)
      AantalKolommen = 0
      AantalResultaten = 0
      AantalRijen = 0
      If Not ControleerOpAPIFout(SafeArrayGetDim(.Tabel)) = 0 Then
         AantalKolommen = AantalItems(.Tabel(), Dimensie:=2) + 1
         If AantalItems(.Tabel(), Dimensie:=1) = 0 Then AantalKolommen = 0
         AantalRijen = AantalItems(.Tabel(), Dimensie:=1)

         QueryResultaten , , , EersteResultaat, LaatsteResultaat
         AantalResultaten = Abs(LaatsteResultaat - EersteResultaat) + 1
      End If

      Status = "Query uitgevoerd: " & CStr(AantalRijen)
      If AantalRijen = 1 Then Status = Status & " rij" Else Status = Status & " rijen"
      Status = Status & " en " & CStr(AantalKolommen)
      If AantalKolommen = 1 Then Status = Status & " kolom." Else Status = Status & " kolommen."

      If AantalResultaten > 1 Then Status = Status & " Resultaat " & CStr((ResultaatIndex - EersteResultaat) + 1) & " van " & CStr(AantalResultaten) & "."

      If Instellingen().VoorbeeldRegels >= 0 Then
         Status = Status & " Voorbeeld limiet: " & CStr(Instellingen().VoorbeeldRegels)
         If Instellingen().VoorbeeldRegels = 1 Then Status = Status & " regel." Else Status = Status & " regels."
      End If
   End With

EindeProcedure:
   StatusNaQuery = Status
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure toont de programmainformatie.
Public Sub ToonProgrammaInformatie()
On Error GoTo Fout
   With App
      MsgBox .Comments, vbInformation, .Title & " " & ProgrammaVersie() & " - " & "door: " & .CompanyName & ", " & "2009-2016"
   End With
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure toont het opgegeven queryresultaat.
Public Sub ToonQueryResultaat(QueryResultaatVeld As TextBox, ResultaatIndex As Long)
On Error GoTo Fout
Dim ResultaatTekst As String

   ToonStatus "Bezig met maken van voorbeeld weergave voor queryresultaat..." & vbCrLf
   ResultaatTekst = QueryResultaatTekst(QueryResultaten(, , ResultaatIndex))
   QueryResultaatVeld.Text = ResultaatTekst
   If InterfaceVenster.Visible And Len(QueryResultaatVeld.Text) < Len(ResultaatTekst) Then MsgBox "Het queryresultaat kan niet volledig worden weergegeven.", vbInformation
EindeProcedure:
   ToonStatus StatusNaQuery(ResultaatIndex) & vbCrLf
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure toont de opgegeven tekst in het opgegeven veld.
Public Sub ToonStatus(Optional Tekst As String = vbNullString, Optional NieuwVeld As TextBox = Nothing)
On Error GoTo Fout
Dim VorigeLengte As Long
Static Veld As TextBox

   If Not NieuwVeld Is Nothing Then Set Veld = NieuwVeld

   If Not Veld Is Nothing Then
      With Veld
         VorigeLengte = Len(.Text)
         .Text = .Text & Tekst
         If Len(.Text) < VorigeLengte + Len(Tekst) Then .Text = Tekst
         .SelStart = Len(.Text) - Len(Tekst)
         .SelLength = 0
      End With
   End If

   DoEvents
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure toont de verbindingsstatus.
Public Sub ToonVerbindingsstatus()
On Error GoTo Fout
   If VerbindingGeopend(Verbinding()) Then
      ToonStatus "Verbonden met de database. - Instellingen: " & Instellingen().Bestand & vbCrLf
   Else
      ToonStatus "Er is geen verbinding met een database." & vbCrLf
   End If
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub



'Deze procedure beheert de verbinding met een database.
Public Function Verbinding(Optional VerbindingsInformatie As String = vbNullString, Optional VerbindingSluiten As Boolean = False, Optional Reset As Boolean = False) As Adodb.Connection
On Error GoTo Fout
Static DataBaseVerbinding As New Adodb.Connection

   If Not DataBaseVerbinding Is Nothing Then
      If Reset Then
         DataBaseVerbinding.Errors.Clear
      ElseIf Not VerbindingsInformatie = vbNullString Then
         If Not MaakVerbindingsInformatieOp(VerbindingsInformatie) = vbNullString Then DataBaseVerbinding.Open VerbindingsInformatie
      ElseIf VerbindingSluiten Then
         If VerbindingGeopend(DataBaseVerbinding) Then
            DataBaseVerbinding.Close
            Set DataBaseVerbinding = Nothing
         End If
      End If
   End If

EindeProcedure:
   Set Verbinding = DataBaseVerbinding
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure controleert of de opgegeven verbinding geopend is.
Public Function VerbindingGeopend(VerbindingO As Adodb.Connection) As Boolean
On Error GoTo Fout
Dim Geopend As Boolean

   If Not VerbindingO Is Nothing Then Geopend = (VerbindingO.State = adStateOpen)

EindeProcedure:
   VerbindingGeopend = Geopend
   Exit Function

Fout:
   Geopend = False
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procecure vervangt de symbolen in de opgegeven tekst met de tekst waar ze voor staan.
Public Function VervangSymbolen(Tekst As String) As String
On Error GoTo Fout
Dim Symbool As String
Dim SymboolBegin As Long
Dim SymboolEinde As Long
Dim TekstMetSymbolen As String
Dim TekstZonderSymbolen As String

   TekstMetSymbolen = Tekst
   TekstZonderSymbolen = vbNullString
   Do
      SymboolBegin = InStr(TekstMetSymbolen, SYMBOOL_TEKEN)
      If SymboolBegin = 0 Then
         TekstZonderSymbolen = TekstZonderSymbolen & TekstMetSymbolen
         Exit Do
      Else
         SymboolEinde = InStr(SymboolBegin + 1, TekstMetSymbolen, SYMBOOL_TEKEN)
         If SymboolEinde = 0 Then
            TekstZonderSymbolen = TekstZonderSymbolen & TekstMetSymbolen
            Exit Do
         Else
            TekstZonderSymbolen = TekstZonderSymbolen & Left$(TekstMetSymbolen, SymboolBegin - 1)
            Symbool = Mid$(TekstMetSymbolen, SymboolBegin + 1, (SymboolEinde - SymboolBegin) - 1)
            TekstMetSymbolen = Mid$(TekstMetSymbolen, SymboolEinde + 1)

            If Symbool = vbNullString Then
               ParameterSymboolFout "Een leeg symbool is gevonden. Deze wordt genegeerd."
            Else
               TekstZonderSymbolen = TekstZonderSymbolen & VerwerkSymbool(Symbool)
            End If
         End If
      End If
   Loop

EindeProcedure:
   VervangSymbolen = TekstZonderSymbolen
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure verwerkt de batchinstellingen.
Private Function VerwerkBatchInstellingen(Regel As String, Sectie As String, ByRef BatchInstellingen As InstellingenDefinitie) As Boolean
On Error GoTo Fout
Dim ParameterNaam As String
Dim Verwerkt As Boolean
Dim Waarde As String

   ParameterNaam = vbNullString
   Verwerkt = True
   Waarde = LeesParameter(Regel, ParameterNaam)

   With BatchInstellingen
      Select Case ParameterNaam
         Case "bereik"
            .BatchBereik = Waarde
         Case "interactief"
            .BatchInteractief = CBool(Waarde)
         Case "querypad"
            .BatchQueryPad = Waarde
         Case Else
            If InstellingenFout("Niet herkende parameter.", BatchInstellingen.Bestand, Sectie, Regel) = vbCancel Then Verwerkt = False
      End Select
   End With
EindeProcedure:
   VerwerkBatchInstellingen = Verwerkt
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False, TypePad:="Instellingenbestand: ", Pad:=BatchInstellingen.Bestand, ExtraInformatie:="Sectie: " & Sectie & vbCr & "Regel: " & Regel) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure verwerkt de exportinstellingen.
Private Function VerwerkExportInstellingen(Regel As String, Sectie As String, ByRef ExportInstellingen As InstellingenDefinitie) As Boolean
On Error GoTo Fout
Dim ParameterNaam As String
Dim Verwerkt As Boolean
Dim Waarde As String

   ParameterNaam = vbNullString
   Verwerkt = True
   Waarde = LeesParameter(Regel, ParameterNaam)

   With ExportInstellingen
      Select Case ParameterNaam
         Case "afzender"
            .ExportAfzender = Waarde
         Case "autoopenen"
            .ExportAutoOpenen = CBool(Waarde)
         Case "autooverschrijven"
            .ExportAutoOverschrijven = CBool(Waarde)
         Case "autoverzenden"
            .ExportAutoVerzenden = CBool(Waarde)
         Case "ccontvanger"
            .ExportCCOntvanger = Waarde
         Case "kolomaanvullen"
            .ExportKolomAanvullen = CBool(Waarde)
         Case "onderwerp"
            .ExportOnderwerp = Waarde
         Case "ontvanger"
            .ExportOntvanger = Waarde
         Case "standaardpad"
            .ExportStandaardPad = Waarde
         Case Else
            If InstellingenFout("Niet herkende parameter.", ExportInstellingen.Bestand, Sectie, Regel) = vbCancel Then Verwerkt = False
      End Select
   End With
EindeProcedure:
   VerwerkExportInstellingen = Verwerkt
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False, TypePad:="Instellingenbestand: ", Pad:=ExportInstellingen.Bestand, ExtraInformatie:="Sectie: " & Sectie & vbCr & "Regel: " & Regel) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure stuurt de verbindingsinformatie met de opgegeven inloggegevens terug.
Public Function VerwerkInlogGegevens(Gebruiker As String, Wachtwoord As String, VerbindingsInformatie As String) As String
On Error GoTo Fout
Dim LinkerDeel As String
Dim Positie As Long
Dim RechterDeel As String
Dim VerwerkteInlogGegevens As String

   VerwerkteInlogGegevens = VerbindingsInformatie

   Positie = InStr(UCase$(VerwerkteInlogGegevens), GEBRUIKER_VARIABEL)
   If Positie > 0 Then
      LinkerDeel = Left$(VerwerkteInlogGegevens, Positie - 1)
      RechterDeel = Mid$(VerwerkteInlogGegevens, Positie + Len(GEBRUIKER_VARIABEL))
      VerwerkteInlogGegevens = LinkerDeel & Gebruiker & RechterDeel
   End If

   Positie = InStr(UCase$(VerwerkteInlogGegevens), WACHTWOORD_VARIABEL)
   If Positie > 0 Then
      LinkerDeel = Left$(VerwerkteInlogGegevens, Positie - 1)
      RechterDeel = Mid$(VerwerkteInlogGegevens, Positie + Len(WACHTWOORD_VARIABEL))
      VerwerkteInlogGegevens = LinkerDeel & Wachtwoord & RechterDeel
   End If

EindeProcedure:
   VerwerkInlogGegevens = VerwerkteInlogGegevens
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure verwerkt de queryinstellingen.
Private Function VerwerkQueryInstellingen(Regel As String, Sectie As String, ByRef QueryInstellingen As InstellingenDefinitie) As Boolean
On Error GoTo Fout
Dim ParameterNaam As String
Dim Verwerkt As Boolean
Dim Waarde As String

   ParameterNaam = vbNullString
   Verwerkt = True
   Waarde = LeesParameter(Regel, ParameterNaam)

   With QueryInstellingen
      Select Case ParameterNaam
         Case "autosluiten"
            .QueryAutoSluiten = CBool(Waarde)
         Case "autouitvoeren"
            .QueryAutoUitvoeren = CBool(Waarde)
         Case "recordsets"
            .QueryRecordSets = CBool(Waarde)
         Case "timeout"
            .QueryTimeout = CLng(Waarde)
         Case Else
            If InstellingenFout("Niet herkende parameter.", QueryInstellingen.Bestand, Sectie, Regel) = vbCancel Then Verwerkt = False
      End Select
   End With
EindeProcedure:
   VerwerkQueryInstellingen = Verwerkt
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False, TypePad:="Instellingenbestand: ", Pad:=QueryInstellingen.Bestand, ExtraInformatie:="Sectie: " & Sectie & vbCr & "Regel: " & Regel) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure verwerkt de opgegeven sessie lijst.
Public Function VerwerkSessieLijst(Optional SessieLijstPad As String = vbNullString) As String
On Error GoTo Fout
Dim BestandHandle As Long
Dim SessieParameters As String
Static HuidigeSessieLijstPad As String

   If Not SessieLijstPad = vbNullString Then
      SessiesAfbreken NieuweSessiesAfbreken:=False
      BestandHandle = FreeFile()
      HuidigeSessieLijstPad = SessieLijstPad
      Open HuidigeSessieLijstPad For Input Lock Read Write As BestandHandle
         Do Until EOF(BestandHandle) Or SessiesAfbreken()
            Line Input #BestandHandle, SessieParameters
            If Not Trim$(SessieParameters) = vbNullString Then VoerSessieUit SessieParameters
         Loop
      Close BestandHandle
   End If

EindeProcedure:
   VerwerkSessieLijst = HuidigeSessieLijstPad
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False, TypePad:="Sessielijst: ", Pad:=HuidigeSessieLijstPad) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function



'Deze procedure stuurt de door het opgegeven symbool vertegenwoordigde waarde terug.
Private Function VerwerkSymbool(Symbool As String) As String
On Error GoTo Fout
Dim Bericht As String
Dim IsGetal As Boolean
Dim SymboolArgument As String
Dim Waarde As String

   On Error GoTo IsGeenGetal
   IsGetal = CStr(CLng(Val(Symbool))) = Symbool
   On Error GoTo Fout

   If IsGetal Then
      If CLng(Val(Symbool)) = 0 Then
         Waarde = BestandsSysteem().GetBaseName(Query().Pad)
      Else
         Waarde = QueryParameters(, CLng(Val(Symbool)) - 1).Invoer
      End If
   Else
      SymboolArgument = Mid$(Symbool, 2)
      Symbool = Left$(Symbool, 1)

      Select Case Symbool
         Case "D"
            Waarde = Format$(Day(Date), "00") & Format$(Month(Date), "00") & CStr(Year(Date))
         Case "b"
            If CStr(CLng(Val(SymboolArgument))) = SymboolArgument Then Waarde = InteractieveBatchParameters(CLng(Val(SymboolArgument)))
         Case "c"
            If CStr(CLng(Val(SymboolArgument))) = SymboolArgument Then Waarde = ChrW$(CLng(Val(SymboolArgument)))
         Case "d"
            Waarde = Format$(Day(Date), "00")
         Case "e"
            Waarde = Environ$(SymboolArgument)
         Case "j"
            Waarde = Format$(Year(Date), "0000")
         Case "m"
            Waarde = Format$(Month(Date), "00")
         Case "s"
            If CStr(CLng(Val(SymboolArgument))) = SymboolArgument Then Waarde = SessieParameters(CLng(Val(SymboolArgument)))
         Case Else
            If Not Symbool = vbNullString Then ParameterSymboolFout "Symbool """ & Symbool & """ is onbekend. Deze wordt genegeerd."
      End Select
   End If

EindeProcedure:
   VerwerkSymbool = Waarde
   Exit Function

Fout:
   Bericht = "Symbool """ & Symbool & """ veroorzaakt de volgende fout: " & vbCr
   Bericht = Bericht & Err.Description & "." & vbCr
   Bericht = Bericht & "Foutcode: " & Err.Number
   ParameterSymboolFout Bericht
   Resume EindeProcedure

IsGeenGetal:
   IsGetal = False
   Resume Next
End Function

'Deze procedure verwerkt de voorbeeldinstellingen.
Private Function VerwerkVoorbeeldInstellingen(Regel As String, Sectie As String, ByRef VoorbeeldInstellingen As InstellingenDefinitie) As Boolean
On Error GoTo Fout
Dim ParameterNaam As String
Dim Verwerkt As Boolean
Dim Waarde As String

   ParameterNaam = vbNullString
   Verwerkt = True
   Waarde = LeesParameter(Regel, ParameterNaam)

   With VoorbeeldInstellingen
      Select Case ParameterNaam
         Case "kolombreedte"
            .VoorbeeldKolomBreedte = CLng(Waarde)
         Case "regels"
            .VoorbeeldRegels = CLng(Waarde)
         Case Else
            If InstellingenFout("Niet herkende parameter.", VoorbeeldInstellingen.Bestand, Sectie, Regel) = vbCancel Then Verwerkt = False
      End Select
   End With
EindeProcedure:
   VerwerkVoorbeeldInstellingen = Verwerkt
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False, TypePad:="Instellingenbestand: ", Pad:=VoorbeeldInstellingen.Bestand, ExtraInformatie:="Sectie: " & Sectie & vbCr & "Regel: " & Regel) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure verwijdert eventuele aanhalingstekens aan het begin en/of eind van het opgegeven pad.
Public Function VerwijderAanhalingsTekens(Pad As String) As String
On Error GoTo Fout
Dim PadZonderAanhalingsTekens As String

   PadZonderAanhalingsTekens = Pad
   If Left$(PadZonderAanhalingsTekens, 1) = """" Then PadZonderAanhalingsTekens = Mid$(PadZonderAanhalingsTekens, 2)
   If Right$(PadZonderAanhalingsTekens, 1) = """" Then PadZonderAanhalingsTekens = Left$(PadZonderAanhalingsTekens, Len(PadZonderAanhalingsTekens) - 1)

EindeProcedure:
   VerwijderAanhalingsTekens = PadZonderAanhalingsTekens
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procecure verwijdert eventuele opmaak uit de opgegeven querycode.
Private Function VerwijderOpmaak(QueryCode As String, CommentaarBegin As String, CommentaarEinde As String, TekenreeksTekens As String) As String
On Error GoTo Fout
Dim HuidigStringTeken As String
Dim InCommentaar As Boolean
Dim Index As Long
Dim QueryZonderOpmaak As String
Dim Teken As String

   HuidigStringTeken = vbNullString
   InCommentaar = False
   Index = 1
   QueryZonderOpmaak = vbNullString
   Teken = vbNullString
   Do Until Index > Len(QueryCode)
      Teken = Mid$(QueryCode, Index, 1)

      If InCommentaar Then
         If CommentaarEinde = vbNullString Then
            If Mid$(QueryCode, Index, 1) = vbCr Or Mid$(QueryCode, Index, 1) = vbLf Then
               HuidigStringTeken = vbNullString
               InCommentaar = False
               Teken = " "
            End If
         Else
            If Mid$(QueryCode, Index, Len(CommentaarEinde)) = CommentaarEinde Then
               HuidigStringTeken = vbNullString
               InCommentaar = False
               Index = Index + (Len(CommentaarEinde) - 1)
               Teken = " "
            End If
         End If
      Else
         If InStr(TekenreeksTekens, Mid$(QueryCode, Index, 1)) > 0 Then
            If HuidigStringTeken = vbNullString Then
               HuidigStringTeken = Teken
            ElseIf Teken = HuidigStringTeken Then
               HuidigStringTeken = vbNullString
            End If
         ElseIf Mid$(QueryCode, Index, Len(CommentaarBegin)) = CommentaarBegin Then
            If HuidigStringTeken = vbNullString Then InCommentaar = True
         End If
      End If

      If Not InCommentaar Then
         If HuidigStringTeken = vbNullString Then
            If Mid$(QueryCode, Index, 1) = vbCr Or Mid$(QueryCode, Index, 1) = vbLf Then Teken = " "

            If InStr(vbTab & " ", Teken) > 0 Then
               Teken = " "
               If Right$(QueryZonderOpmaak, 1) = " " Then Teken = vbNullString
            End If
         End If

         QueryZonderOpmaak = QueryZonderOpmaak & Teken
      End If

      Index = Index + 1
   Loop

EindeProcedure:
   VerwijderOpmaak = QueryZonderOpmaak
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure voert een querybatch uit.
Private Sub VoerBatchUit()
On Error GoTo Fout
Dim EersteParameter As Long
Dim EersteQuery As Long
Dim EMail As EMailClass
Dim ExportPad As String
Dim ExportPaden As New Collection
Dim ExportUitgevoerd As Boolean
Dim FoutInformatie As String
Dim Index As Long
Dim LaatsteParameter As Long
Dim LaatsteQuery As Long
Dim Positie As Long
Dim QueryIndex As Long
Dim QueryPad As String
Dim QueryPadExtensie As String

   ToonVerbindingsstatus

   With Instellingen()
      Positie = InStr(.BatchBereik, "-")
      If Not Positie = 0 Then
         EersteQuery = CLng(Val(Trim$(Left$(.BatchBereik, Positie - 1))))
         LaatsteQuery = CLng(Val(Trim$(Mid$(.BatchBereik, Positie + 1))))
         QueryPadExtensie = "." & BestandsSysteem().GetExtensionName(.BatchQueryPad)

         If CStr(EersteQuery) = Trim$(Left$(.BatchBereik, Positie - 1)) And CStr(LaatsteQuery) = Trim$(Mid$(.BatchBereik, Positie + 1)) And EersteQuery <= LaatsteQuery Then
            For QueryIndex = EersteQuery To LaatsteQuery
               QueryPad = VerwijderAanhalingsTekens(Left$(.BatchQueryPad, Len(.BatchQueryPad) - Len(QueryPadExtensie)) & CStr(QueryIndex) & QueryPadExtensie)
   
               If Query(QueryPad).Geopend Then
                  QueryParameters Query().Code

                  If .BatchInteractief And QueryIndex = EersteQuery Then
                     InteractieveBatchAfbreken BatchAfbreken:=True
                     InterfaceVenster.Show

                     If Not Trim$(Command$()) = vbNullString Then ToonStatus "Opdrachtregel: " & Command$() & vbCrLf
                     If Not VerwerkSessieLijst() = vbNullString Then ToonStatus "Sessie lijst: " & VerwerkSessieLijst() & vbCrLf
                     ToonStatus "Query: " & QueryPad & vbCrLf

                     Do While DoEvents() > 0 And InterfaceVenster.Enabled: ControleerOpAPIFout WaitMessage(): Loop
                     If InteractieveBatchAfbreken() Then Exit Sub
   
                     Screen.MousePointer = vbHourglass
                     InteractieveBatchParameters , , Verwijderen:=True
                     QueryParameters , , , EersteParameter, LaatsteParameter
                     For Index = EersteParameter To LaatsteParameter
                        InteractieveBatchParameters , QueryParameters(, Index).Invoer
                     Next Index
                  Else
                     If QueryIndex = EersteQuery Then
                        If Not Trim$(Command$()) = vbNullString Then ToonStatus "Opdrachtregel: " & Command$() & vbCrLf
                        If Not VerwerkSessieLijst() = vbNullString Then ToonStatus "Sessie lijst: " & VerwerkSessieLijst() & vbCrLf
                     End If

                     ToonStatus "Query: " & QueryPad & vbCrLf
   
                     QueryParameters , , , EersteParameter, LaatsteParameter
                     For Index = EersteParameter To LaatsteParameter
                        With QueryParameters(, Index)
                           QueryParameters , Index, .StandaardWaarde & Mid$(.Masker, Len(.StandaardWaarde) + 1)
                           If Not (.Commentaar = vbNullString And .Masker = vbNullString And .ParameterNaam = vbNullString) Then ParameterSymboolFout "Genegeerde elementen in batch query gevonden.", Index
                        End With
                     Next Index
                  End If
   
                  If OngeldigeParameterInvoer(FoutInformatie) = GEEN_PARAMETER Then
                     ToonStatus "Bezig met het uitvoeren van de query..." & vbCrLf
                     QueryResultaten , ResultatenVerwijderen:=True
                     VoerQueryUit Query().Code
   
                     If VerbindingGeopend(Verbinding()) Then
                        If Verbinding().Errors.Count = 0 Then
                           ToonStatus StatusNaQuery(ResultaatIndex:=0) & vbCrLf
                           If Not .ExportStandaardPad = vbNullString Then
                              ToonStatus "Bezig met het exporteren van het queryresultaat..." & vbCrLf
                              ExportPad = BestandsSysteem().GetAbsolutePathName(VerwijderAanhalingsTekens(Trim$(VervangSymbolen(.ExportStandaardPad))))
                     
                              If BestandsSysteem().FolderExists(BestandsSysteem().GetParentFolderName(ExportPad)) Then
                                 ExportPaden.Add ExportPad
                                 ExportUitgevoerd = ExporteerResultaat(ExportPad)
                                 If ExportUitgevoerd Then
                                    If BestandsSysteem().FileExists(ExportPad) And .ExportAutoOpenen Then
                                       ToonStatus "De export wordt automatisch geopend..." & vbCrLf
                                       ControleerOpAPIFout ShellExecuteA(CLng(0), "open", ExportPad, vbNullString, vbNullString, SW_SHOWNORMAL)
                                    End If
   
                                    ToonStatus "Exporteren gereed." & vbCrLf
                                 Else
                                    ToonStatus "Export afgebroken." & vbCrLf
                                 End If
                              Else
                                 MsgBox "Ongeldig export pad." & vbCr & "Huidig pad: " & CurDir$(), vbExclamation
                                 ToonStatus "Ongeldig export pad." & vbCrLf
                              End If
                           Else
                              ToonStatus FoutenLijstTekst(Verbinding().Errors)
                           End If
                        End If
                        
                        Verbinding , , Reset:=True
                     End If
                  Else
                     ParameterSymboolFout "Ongeldige parameter invoer: " & FoutInformatie
                  End If
               End If
            Next QueryIndex
   
            If (Not .ExportStandaardPad = vbNullString) And ExportUitgevoerd Then
               If Not (.ExportOntvanger = vbNullString And .ExportCCOntvanger = vbNullString) Then
                  ToonStatus "Bezig met het maken van de e-mail met de export..." & vbCrLf
                  Set EMail = New EMailClass
                  EMail.VoegQueryResultatenToe , ExportPaden
                  Set EMail = Nothing
               End If
            End If
         Else
            MsgBox "Ongeldige querybatchbereik: """ & .BatchBereik & """.", vbExclamation
         End If
      End If
   End With

EindeProcedure:
   Screen.MousePointer = vbDefault
   Unload InterfaceVenster
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure voert een query uit op een database.
Public Sub VoerQueryUit(QueryCode As String)
On Error GoTo Fout
Dim Commando As New Adodb.Command
Dim Resultaat As Adodb.Recordset
Dim QueryPad As String

   QueryPad = Query().Pad

   If VerbindingGeopend(Verbinding()) Then
      Set Commando.ActiveConnection = Verbinding()

      If Not Commando Is Nothing Then
         Commando.CommandText = VulParametersIn(QueryCode)
         Commando.CommandText = VerwijderOpmaak(Commando.CommandText, SQL_COMMENTAAR_REGEL_BEGIN, SQL_COMMENTAAR_REGEL_EINDE, TEKENREEKS_TEKENS)
         Commando.CommandText = VerwijderOpmaak(Commando.CommandText, SQL_COMMENTAAR_BLOK_BEGIN, SQL_COMMENTAAR_BLOK_EINDE, TEKENREEKS_TEKENS)
         Commando.CommandTimeout = Instellingen().QueryTimeout
         Commando.CommandType = adCmdText

         Set Resultaat = Commando.Execute
      End If

      Do While VerbindingGeopend(Resultaat.ActiveConnection)
         QueryResultaten Resultaat

         If Instellingen().QueryRecordSets Then Set Resultaat = Resultaat.NextRecordset Else Exit Do
      Loop
   End If

EindeProcedure:
   ToonStatus "Uitgevoerde query: " & vbCrLf & Commando.CommandText & vbCrLf

   If Not Resultaat Is Nothing Then
      If VerbindingGeopend(Resultaat.ActiveConnection) Then Resultaat.Close
   End If

   Set Commando = Nothing
   Set Resultaat = Nothing
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False, ExtraInformatie:="Query: " & QueryPad) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure voert een sessie met de opgegeven parameters uit.
Private Sub VoerSessieUit(SessieParameters As String)
On Error GoTo Fout
Static RecensteVerbindingsInformatie As String

   With OpdrachtRegelParameters(SessieParameters)
      If .Verwerkt Then
         If .InstellingenPad = vbNullString Then
            Instellingen BestandsSysteem().BuildPath(App.Path, StandaardInstellingen().Bestand)
         Else
            Instellingen .InstellingenPad
         End If
   
         With Instellingen()
            If Not .VerbindingsInformatie = RecensteVerbindingsInformatie Then
               Verbinding , VerbindingSluiten:=True
               If InStr(UCase$(.VerbindingsInformatie), GEBRUIKER_VARIABEL) > 0 Or InStr(UCase(.VerbindingsInformatie), WACHTWOORD_VARIABEL) > 0 Then
                  InloggenVenster.Show vbModal
               Else
                  Verbinding .VerbindingsInformatie
               End If
            End If
      
            RecensteVerbindingsInformatie = .VerbindingsInformatie
         End With

         If VerbindingGeopend(Verbinding()) Then
            If BatchModusActief() Then
               VoerBatchUit
            Else
               InterfaceVenster.Show
               Do While DoEvents() > 0: ControleerOpAPIFout WaitMessage(): Loop
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



'Deze procedure opent een dialoogvenster waarmee de gebruiker naar het pad voor het te exporteren queryresultaat kan bladeren.
Public Function VraagExportPad(HuidigExportPad As String) As String
On Error GoTo Fout
Dim ExportPadDialoog As OPENFILENAME
Dim NieuwExportPad As String

   NieuwExportPad = HuidigExportPad

   With ExportPadDialoog
      .hInstance = CLng(0)
      .hwndOwner = CLng(0)
      .lCustData = CLng(0)
      .lpfnHook = CLng(0)
      .lpstrCustomFilter = vbNullString
      .lpstrDefExt = vbNullString
      .lpstrFile = String$(MAX_STRING, vbNullChar) & vbNullChar
      .lpstrFileTitle = String$(MAX_STRING, vbNullChar) & vbNullChar
      .lpTemplateName = vbNullString
      .lStructSize = Len(ExportPadDialoog)
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
      .lpstrTitle = "Exporteer het queryresultaat naar:" & vbNullChar
      .lpstrFilter = "Tekstbestand (*.txt)" & vbNullChar & "*.txt" & vbNullChar
      .lpstrFilter = .lpstrFilter & "Microsoft Excel bestand (*.xls)" & vbNullChar & "*.xls" & vbNullChar
      .lpstrFilter = .lpstrFilter & "Microsoft Excel 2007 bestand (*.xlsx)" & vbNullChar & "*.xlsx" & vbNullChar
      .lpstrFilter = .lpstrFilter & vbNullChar
      .lpstrInitialDir = App.Path & vbNullChar

      If CBool(ControleerOpAPIFout(GetSaveFileNameA(ExportPadDialoog))) Then NieuwExportPad = Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
   End With
EindeProcedure:
   VraagExportPad = NieuwExportPad
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure opent een dialoogvenster waarmee de gebruiker naar een querybestand kan bladeren.
Public Function VraagQueryPad() As String
On Error GoTo Fout
Dim QueryPad As String
Dim QueryPadDialoog As OPENFILENAME

   QueryPad = vbNullString

   With QueryPadDialoog
      .hInstance = CLng(0)
      .hwndOwner = CLng(0)
      .lCustData = CLng(0)
      .lpfnHook = CLng(0)
      .lpstrCustomFilter = vbNullString
      .lpstrDefExt = vbNullString
      .lpstrFile = String$(MAX_STRING, vbNullChar) & vbNullChar
      .lpstrFileTitle = String$(MAX_STRING, vbNullChar) & vbNullChar
      .lpTemplateName = vbNullString
      .lStructSize = Len(QueryPadDialoog)
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
      .lpstrTitle = "Selecteer een query:" & vbNullChar
      .lpstrFilter = "Tekstbestanden (*.txt)" & vbNullChar & "*.txt" & vbNullChar
      .lpstrFilter = .lpstrFilter & vbNullChar

      .lpstrInitialDir = App.Path & vbNullChar
      If CBool(ControleerOpAPIFout(GetOpenFileNameA(QueryPadDialoog))) Then QueryPad = Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
   End With

EindeProcedure:
   VraagQueryPad = QueryPad
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure vraagt de gebruiker om de gegevens voor een verbinding met een database op te geven.
Private Function VraagVerbindingsInformatie() As String
On Error GoTo Fout
Dim VerbindingsInformatie As String

   Do While Trim$(VerbindingsInformatie) = vbNullString
      VerbindingsInformatie = InputBox$("Informatie voor een verbinding met een database:")
      If StrPtr(VerbindingsInformatie) = 0 Then
         Exit Do
      ElseIf Trim$(VerbindingsInformatie) = vbNullString Then
         MsgBox "Deze informatie is vereist.", vbExclamation
      End If
   Loop

EindeProcedure:
   VraagVerbindingsInformatie = Trim$(VerbindingsInformatie)
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function


'Deze procedure vult de opgegeven tekst aan met het opgegeven aantal spaties.
Private Function VulAan(Tekst As String, Breedte As Long, Optional LinksAanvullen As Boolean = False) As String
On Error GoTo Fout
Dim AangevuldeTekst As String

   AangevuldeTekst = Tekst
   If Len(AangevuldeTekst) > Breedte Then AangevuldeTekst = Left$(AangevuldeTekst, Breedte)
   If LinksAanvullen Then
      AangevuldeTekst = Space$(Breedte - Len(AangevuldeTekst)) & AangevuldeTekst
   Else
      AangevuldeTekst = AangevuldeTekst & Space$(Breedte - Len(AangevuldeTekst))
   End If

EindeProcedure:
   VulAan = AangevuldeTekst
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure vult de parameter invoer in de querycode in.
Private Function VulParametersIn(QueryCode As String) As String
On Error GoTo Fout
Dim EersteParameter As Long
Dim Index As Long
Dim LaatsteParameter As Long
Dim Positie As Long
Dim QueryMetParameters As String
Dim QueryZonderParameters As String

   QueryParameters , , , EersteParameter, LaatsteParameter
   If EersteParameter = GEEN_PARAMETER And LaatsteParameter = GEEN_PARAMETER Then
      QueryMetParameters = QueryCode
   Else
      Index = EersteParameter
      QueryMetParameters = vbNullString
      QueryZonderParameters = QueryCode
      Do Until Index > LaatsteParameter
         SessieParameters , QueryParameters(, Index).Invoer

         Positie = QueryParameters(, Index).Positie
         QueryMetParameters = QueryMetParameters & Left$(QueryZonderParameters, Positie - 1)
         QueryMetParameters = QueryMetParameters & VervangSymbolen(QueryParameters(, Index).Invoer)
         QueryZonderParameters = Mid$(QueryZonderParameters, Positie + QueryParameters(, Index).Lengte)
         Index = Index + 1
      Loop
      QueryMetParameters = QueryMetParameters & QueryZonderParameters
   End If
EindeProcedure:
   VulParametersIn = QueryMetParameters
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

