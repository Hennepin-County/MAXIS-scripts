'LOADING GLOBAL VARIABLES--------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Access Aging & Disability Services\Scripts\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
'name_of_script = "LTCScrng.vbs"
'start_time = time

'LOADING FUNCTIONS LIBRARY FROM REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
FuncLib_URL = script_repository & "\MMIS FUNCTIONS LIBRARY.vbs"
critical_error_msgbox = MsgBox ("The Functions Library code was not able to be reached by " & name_of_script & vbNewLine & vbNewLine &_
"FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
"The script has stopped. Send issues to " & contact_admin , _
vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
StopScript
ELSE
FuncLib_URL = script_repository & "\MMIS FUNCTIONS LIBRARY.vbs"
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script
END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
call changelog_update("09/09/2025", "Added most recent updates from MN.IT script release. Updates include ALT screen checks and updated field clearning functionality.", "Ilse Ferris, Hennepin County" )
call changelog_update("03/14/2025", "Initial scripts set up in Hennepin County.", "Casey Love, Hennepin County" )
'("10/16/2019", "All infrastructure changed to run locally and stored in BlueZone Scripts ccm. MNIT @ DHS)
'("01/13/2025", "Added code to populate ALT6 and clear out fields before populating screens with new data.
'END CHANGELOG BLOCK =======================================================================================================

CALL file_selection_system_dialog(xmlPath, ".xml")

'xmlPath = "C:\Users\PWTMT03\Desktop\MNChoices\LTC Screening Document.xml"
'xmlPath = "C:\Users\PWTMT03\Desktop\MNChoices\LTC Screening Document14-2.xml"
'xmlPath = "C:\Users\PWTMT03\Desktop\MNChoices\DD Screening Document14-3.xml"

Set ObjFso = CreateObject("Scripting.FileSystemObject")
Set ObjFile = ObjFso.OpenTextFile(xmlPath)

Dim scrngArray()

currFieldValue = "UNDEFINED"
currFieldName = "UNDEFINED"


Function BuildValueArray(tableStr, FieldValues)
z = 0
sectionvalue = Split(tableStr, "<")
For Each sectionvalueitem In sectionvalue
If InStr(CStr(sectionvalueitem), "Group") And InStr(CStr(sectionvalueitem), "FieldName") Then
ReDim Preserve FieldValues(1,z)
fielddata = Split(sectionvalueitem, """")
FieldValues(0,z) = fielddata(3)
If UBound(fielddata) > 4 Then
If Left(LTrim(RTrim(fielddata(4))), 5) = "Value" And fielddata(5) <> "N/A" Then
FieldValues(1,z) = fielddata(5)
Else
FieldValues(1,z) = ""
End If
Else
FieldValues(1,z) = ""
End If
'MsgBox FieldValues(0,z) & " is: " & FieldValues(1,z)
z = z + 1
End If
Next
End Function

Function searchForMenuItemAndSelect(menuItem)
scrX = 1
menuLine = ""
Do
EMReadScreen menuLine, 80, scrX, 1
If inStr(menuLine, menuItem) <> 0 Then
EMWriteScreen "X", scrX, inStr(menuLine, menuItem) - 3
scrX = 100
EMSendKey "<enter>"
End If
scrX = scrX + 1
Loop until scrX >= 24

If scrX = 24 Then
MsgBox "ERROR: '" & menuItem & "' was not found!  Aborting.."
stopscript
End If
End Function

Function findValueInArray(fieldName, fieldValues)
z = 0
Do While z <= ubound(fieldValues,2)
If fieldValues(0,z) = fieldName Then
currFieldName = fieldName
currFieldValue = fieldValues(1,z)
End If
z = z + 1
loop
If currFieldName <> fieldName Then
MsgBox "ERROR: '" & fieldName & "' was not found in array!  Aborting.."
stopscript
End If
End Function

' fieldName is the name of the field on the screen you are looking for
' fieldValue is the value you want to write to the screen if the field is found
Function writeValueOnScreen(fieldName, fieldValue)
scrX = 1
scrLine = ""
If right(fieldName, 1) <> ":" Then  fieldName = fieldName & ":"

IF fieldName = "ACT DT:" Then  fieldName = "ACT DT"

Do
EMReadScreen scrLine, 80, scrX, 1
If inStr(scrLine, fieldName) <> 0 Then
EMWriteScreen fieldValue, scrX, inStr(scrLine, fieldName) + len(fieldName) + 1
scrX = 100
End If
scrX = scrX + 1
Loop until scrX >= 24

If scrX = 24 Then
MsgBox "ERROR: '" & fieldName & "' was not found on current screen!  Aborting.."
stopscript
End If
End Function

Function processXML()
Dim Prog_Type
Do While ObjFile.AtEndOfStream <> True
StrData = ObjFile.ReadLine
StrData = trim(StrData)
CALL BuildValueArray(StrData, scrngArray)
Loop
End Function

Function findXorY(scrKeyword, XorY)
	scrY = 1
	scrLine = ""
	Do
		EMReadScreen scrLine, 80, scrY, 1
		If inStr(scrLine, scrKeyword) <> 0 Then
			If XorY = "X" Then findXorY = inStr(scrLine, scrKeyword)
			If XorY = "Y" Then findXorY = scrY
			If XorY = "both" Then findXorY = inStr(scrLine, scrKeyword) & "," & scrY
			scrY = 100
		End If
	scrY = scrY + 1
	Loop until scrY >= 24
End Function

Function findDocNbrOnScreen()
	scrY = 1
	scrLine = ""
	Do
		EMReadScreen scrLine, 80, scrY, 1
		If inStr(scrLine, "DOCUMENT NBR:") <> 0 Then
		findDocNbrOnScreen = trim(scrLine)
		scrY = 100
	End If
	scrY = scrY + 1
	Loop until scrY >= 24
	If isEmpty(findDocNbrOnScreen) Then findDocNbrOnScreen = "N/A"
End Function

Function checkForErrors()
	errWording = "An error occurred.  Screening Document entry will be cancelled when you press OK.  " &_
				 "Correct error(s) in MnCHOICES, create a new .xml file and rerun the script."

	Transmit
	scrLine = ""
	EMReadScreen scrLine, 80, 24, 1
	If trim(scrLine) <> "" Then
		logMonth = Month(Date)
		logDay = Day(Date)
		If len(logMonth) <= 9 Then logMonth = "0" & logMonth
		If len(logDay) <= 9 Then logDay = "0" & logDay
		logFilename = xmlPath & "." & Year(Date) & logMonth & logDay & "." & CLng(Timer) & ".txt"
		Set ObjFso2 = CreateObject("Scripting.FileSystemObject")
		Set ObjErrorLogfile = ObjFso2.OpenTextFile(logFilename, 2, True)
		CALL findValueInArray("RECIPIENT ID", scrngArray)
		errLine = trim(scrLine)
		EMReadScreen scrLine, 80, 1, 1
		EMReadScreen currScrLine, 26, 1, inStr(scrLine, "MMIS")
		EMReadScreen scrLine, 80, 24, 1
		scrY = 1
		scrLine = ""
		scrError = ""
		Do
			EMReadScreen scrLine, 80, scrY, 1
			If inStr(scrLine, "?") <> 0 Then scrError = scrError & " " & trim(scrLine)
			scrY = scrY + 1
		Loop until scrY >= 24
		docNbr = findDocNbrOnScreen()
		EMReadScreen scrLine, 80, 24, 1
		ObjErrorLogfile.WriteLine "DATE:         " & Date & vbCrlf & "TIME:         " & Time & vbCrlf & "XML:          " & xmlPath &_
								  vbCrlf & "SCREEN:       " & trim(currScrLine) & vbCrlf & trim(docNbr) & vbCrlf &_
								  "RECIPIENT ID: " & currFieldValue & vbCrlf & "ERR MSG:      " & trim(errLine) & vbCrlf &_
								  "SCREEN ERR:   " & trim(scrError)
		ObjErrorLogfile.Close
		MsgBox errWording & vbCrlf & vbCrlf & "ERROR: '" & trim(scrLine) & "'" &_
			   vbCrlf & vbCrlf & "Error Report Created:" & vbCRlf & logFilename, 0, "LTC Script Error"
		PF6
		StopScript
	End IF
	EMReadScreen scrLine, 80, 24, 1
End Function

Function checkForExceptions()
	excWording = "An error occurred.  Click OK to switch to manual entry mode to save or cancel the Screening Document.  " &_
				 "Correct error(s) in MnCHOICES, create a new .xml file and rerun the script."
	PF9
	Dim allExceptions()
	lineItemRow = findXorY("LI EXC ST USER ID", "Y")
	EMReadScreen exceptionLine, 80, lineItemRow + 1, 1
	Do While trim(exceptionLine) <> ""
		z = 0
		currExcLocation = 5
		EMReadScreen excScrollEnd, 80, 24, 1
		Do While inStr(excScrollEnd, "CANNOT SCROLL FORWARD - NO MORE DATA TO DISPLAY") = 0
			Do While currExcLocation <= 62
				EMSetCursor lineItemRow + 1, currExcLocation
				EMReadScreen excFound, 1, lineItemRow + 1, currExcLocation
				If IsNumeric(excFound) Then
					PF1
					EMReadScreen excStr, 80, lineItemRow + 2, 1
					Redim Preserve allExceptions(z)
					allExceptions(z) = trim(excStr)
					z = z + 1
				End IF
				currExcLocation = currExcLocation + 19
			Loop
			currExcLocation = 5
			EMSetCursor lineItemRow, 78
			PF8
			EMReadScreen excScrollEnd, 80, 24, 1
		Loop
		z = 1
		logMonth = Month(Date)
		logDay = Day(Date)
		If len(logMonth) <= 9 Then logMonth = "0" & logMonth
		If len(logDay) <= 9 Then logDay = "0" & logDay
		logFilename = xmlPath & "." & Year(Date) & logMonth & logDay & "." & CLng(Timer) & ".txt"
		Set ObjFso2 = CreateObject("Scripting.FileSystemObject")
		Set ObjErrorLogfile = ObjFso2.OpenTextFile(logFilename, 2, True)
		Do While z <= ubound(allExceptions)
			If currEdits = "" Then currEdits = allExceptions(0)
			currEdits = currEdits & vbCRlf & allExceptions(z)
			z = z + 1
		Loop
		CALL findValueInArray("RECIPIENT ID", scrngArray)
		docNbr = findDocNbrOnScreen()
		ObjErrorLogfile.WriteLine Date & " " & Time
		ObjErrorLogfile.WriteLine xmlPath
		ObjErrorLogfile.WriteLine ""
		ObjErrorLogfile.WriteLine "RECIPIENT ID: " & currFieldValue
		ObjErrorLogfile.WriteLine docNbr
		ObjErrorLogfile.WriteLine ""
		ObjErrorLogfile.WriteLine currEdits
		ObjErrorLogfile.Close
		MsgBox excWording & vbCrlf & vbCrlf & currEdits & vbCrlf & vbCrlf &_
			   "Exception Report created:" & Vbcrlf & logFilename, 0, "Exceptions"
		StopScript
	Loop
End Function

CALL processXML()
ObjFile.close
CALL findValueInArray("DOCUMENT TYPE", scrngArray)

If currFieldValue <> "L" Then
MsgBox """" & xmlPath & """ is not an LTC Document." & vbCrlf & vbCrlf &_
"Please select a valid LTC XML document and try the script again.", 0, "LTC Script Error"
StopScript
End If

EMconnect ""

'MAIN Check
EMReadScreen Main_check, 4, 1, 50
If Main_check <> "MAIN" Then script_end_procedure ("The script must start on MAIN Screen")
'going to screening application
CALL searchForMenuItemAndSelect("SCREENINGS")

'ASCR
CALL writeValueOnScreen("ACTION CODE", "A")
CALL findValueInArray("DOCUMENT TYPE", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("RECIPIENT ID", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue)
Call checkForErrors()

'ALT1
EMwriteScreen     "                 ", 5, 18

EMwriteScreen     "            ", 5, 36

CALL findValueInArray("REF NBR", scrngArray)
CALL writeValueOnScreen(currFieldName, "          ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("DOB", scrngArray)
CALL writeValueOnScreen(currFieldName, "        ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("SEX", scrngArray)
CALL writeValueOnScreen(currFieldName, " ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("REF DATE", scrngArray)
CALL writeValueOnScreen(currFieldName, "      ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("NEXT NF VISIT", scrngArray)
CALL writeValueOnScreen(currFieldName, "      ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("ACTIVITY TYPE", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("ACT DT", scrngArray)
CALL writeValueOnScreen(currFieldName, "      ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("COS", scrngArray)
CALL writeValueOnScreen(currFieldName, "")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("COR", scrngArray)
CALL writeValueOnScreen(currFieldName, "   ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("CFR", scrngArray)
CALL writeValueOnScreen(currFieldName, "   ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("LTCC CTY", scrngArray)
CALL writeValueOnScreen(currFieldName, "   ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("LEGAL REP STAT", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("PRIMARY DIAG", scrngArray)
CALL writeValueOnScreen(currFieldName, "        ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("SECONDARY DIAG", scrngArray)
CALL writeValueOnScreen(currFieldName, "        ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("DD DIAGNOSIS HISTORY", scrngArray)
CALL writeValueOnScreen(currFieldName, " ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("DD DIAGNOSIS", scrngArray)
CALL writeValueOnScreen(currFieldName, "        ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("MI DIAGNOSIS HISTORY", scrngArray)
CALL writeValueOnScreen(currFieldName, " ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("MI DIAGNOSIS", scrngArray)
CALL writeValueOnScreen(currFieldName, "        ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("BI DIAGNOSIS HISTORY", scrngArray)
CALL writeValueOnScreen(currFieldName, " ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("BI DIAGNOSIS", scrngArray)
CALL writeValueOnScreen(currFieldName, "        ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("CM/HP/CA NAME", scrngArray)
CALL writeValueOnScreen(currFieldName, "                                   ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("CM/HP/CA NBR", scrngArray)
CALL writeValueOnScreen(currFieldName, "          ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
Transmit

'ALT2

EMReadScreen ALT2_check, 4, 1, 52
If ALT2_check <> "ALT2" Then script_end_procedure ("The script must CONTINUE on ALT2 Screen")

CALL findValueInArray("PRESENT AT SCRNG-1", scrngArray)
EMwriteScreen     "  ", 5, 20
EMwriteScreen     currFieldValue, 5, 20
CALL findValueInArray("PRESENT AT SCRNG-2", scrngArray)
EMwriteScreen     "  ", 5, 35
EMwriteScreen     currFieldValue, 5, 35
CALL findValueInArray("PRESENT AT SCRNG-3", scrngArray)
EMwriteScreen     "  ", 5, 50
EMwriteScreen     currFieldValue, 5, 50
CALL findValueInArray("PRESENT AT SCRNG-4", scrngArray)
EMwriteScreen     "  ", 5, 65
EMwriteScreen     currFieldValue, 5, 65
CALL findValueInArray("PRESENT AT SCRNG-5", scrngArray)
EMwriteScreen     "  ", 6, 20
EMwriteScreen     currFieldValue, 6, 20
CALL findValueInArray("PRESENT AT SCRNG-6", scrngArray)
EMwriteScreen     "  ", 6, 35
EMwriteScreen     currFieldValue, 6, 35
CALL findValueInArray("PRESENT AT SCRNG-7", scrngArray)
EMwriteScreen     "  ", 6, 50
EMwriteScreen     currFieldValue, 6, 50
CALL findValueInArray("PRESENT AT SCRNG-8", scrngArray)
EMwriteScreen     "  ", 6, 65
EMwriteScreen     currFieldValue, 6, 65
CALL findValueInArray("PRESENT AT SCRNG-9", scrngArray)
EMwriteScreen     "  ", 7, 20
EMwriteScreen     currFieldValue, 7, 20
CALL findValueInArray("PRESENT AT SCRNG-10", scrngArray)
EMwriteScreen     "  ", 7, 35
EMwriteScreen     currFieldValue, 7, 35
CALL findValueInArray("PRESENT AT SCRNG-11", scrngArray)
EMwriteScreen     "  ", 7, 50
EMwriteScreen     currFieldValue, 7, 50
CALL findValueInArray("PRESENT AT SCRNG-12", scrngArray)
EMwriteScreen     "  ", 7, 65
EMwriteScreen     currFieldValue, 7, 65
CALL findValueInArray("INFORMAL CAREGIVER", scrngArray)
CALL writeValueOnScreen(currFieldName, " ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("MARITAL STATUS", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("REASONS FOR REF-1", scrngArray)
EMwriteScreen     "  ", 9, 50
EMwriteScreen     currFieldValue, 9, 50
CALL findValueInArray("REASONS FOR REF-2", scrngArray)
EMwriteScreen     "  ", 9, 64
EMwriteScreen     currFieldValue, 9, 64
CALL findValueInArray("CURRENT LA", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("PLANNED LA", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("TEAM", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("HOSP TRNF", scrngArray)
CALL writeValueOnScreen(currFieldName, " ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("OBRA LVL 1 SCR", scrngArray)
CALL writeValueOnScreen(currFieldName, " ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("DENTAL CONCERNS", scrngArray)
EMwriteScreen     " ", 13, 41
'CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("HAVE DENTIST", scrngArray)
EMwriteScreen     " ", 13, 62
'CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("CURR HOUSING", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("PLANNED HSNG", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("OBRA LVL 2 REF - MI DX", scrngArray)
CALL writeValueOnScreen(currFieldName, " ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("DD DX", scrngArray)
CALL writeValueOnScreen(currFieldName, " ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("BI/CAC REF", scrngArray)
CALL writeValueOnScreen(currFieldName, " ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
'sgBox "pause"
Call checkForErrors()


'ALT3

EMReadScreen ALT3_check, 4, 1, 52
If ALT3_check <> "ALT3" Then script_end_procedure ("The script must CONTINUE on ALT3 Screen")

CALL findValueInArray("DRESSING", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("GROOMING", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("BATHING", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("EATING", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("BED MOB", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("TRANSFER", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("WALKING", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("BEHAVIOR", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("TOILET", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("SPC TRMT", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("CL MONITOR", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("NEURO DX", scrngArray)
CALL writeValueOnScreen(currFieldName, " ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("CASE MIX", scrngArray)
CALL writeValueOnScreen(currFieldName, " ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("ORIENT", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("SLF PRES", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("DIS CERT", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("SLF EVAL", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("HEARING", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("COMM", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("VISION", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("MENT ST EV", scrngArray)
EMwriteScreen     "  ", 12, 15
'CALL writeValueOnScreen(currFieldName, "  ")
'CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("TEL ANS", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("TEL CALL", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("SHOPPING", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("PREP MLS", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("LT HOUSE", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("HY HOUSE", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("LAUNDRY", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("MGMT MEDS", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("INSULIN", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("MONEY MT", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("TRANSP", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("FALLS", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("HOSP", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("ER VISITS", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("NF STAYS", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("VENT DEP", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("FAMILY PLN", scrngArray)
EMwriteScreen     " ", 18, 45
'CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("SEX ACTIVE", scrngArray)
EMwriteScreen     " ", 18, 60
'CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("MINI COG", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
'sgBox "pause"
Call checkForErrors()

'ALT4

EMReadScreen ALT4_check, 4, 1, 52
If ALT4_check <> "ALT4" Then script_end_procedure ("The script must CONTINUE on ALT4 Screen")

CALL findValueInArray("ASSESSMENT RESULTS/EXIT RSNS-1", scrngArray)
EMwriteScreen     "  ", 5, 32
EMwriteScreen     currFieldValue, 5, 32
CALL findValueInArray("ASSESSMENT RESULTS/EXIT RSNS-2", scrngArray)
EMwriteScreen     "  ", 5, 46
EMwriteScreen     currFieldValue, 5, 46
CALL findValueInArray("EFFECTIVE DT", scrngArray)
CALL writeValueOnScreen(currFieldName, "      ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("INFORMED CHOICE", scrngArray)
CALL writeValueOnScreen(currFieldName, " ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("CLIENT CH", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("GUARDIAN CH", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("FAM CH", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("LTCC/IDT RECMND", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("LVL OF CARE", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("NF TRACK", scrngArray)
EMwriteScreen     "  ", 8, 73
CALL writeValueOnScreen(currFieldName, "      ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("CASE MIX AMT", scrngArray)
CALL writeValueOnScreen(currFieldName, " ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("CASE MIX APP (Y/N)", scrngArray)
EMwriteScreen     "  ", 9, 47
'CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("REASON FOR NF CONT STAY/CDCS ENDING-1", scrngArray)
EMwriteScreen     "  ", 10, 39
EMwriteScreen     currFieldValue, 10, 39
CALL findValueInArray("REASON FOR NF CONT STAY/CDCS ENDING-2", scrngArray)
EMwriteScreen     "  ", 10, 49
EMwriteScreen     currFieldValue, 10, 49
'CALL findValueInArray("REL TO COMM", scrngArray)
EMwriteScreen     " ", 10, 80
'CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("ADL COND", scrngArray)
EMwriteScreen     " ", 12, 14
'CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("IADL COND", scrngArray)
EMwriteScreen     " ", 12, 29
'CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("COMP COND", scrngArray)
EMwriteScreen     " ", 12, 44
'CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("COGNITION", scrngArray)
EMwriteScreen     " ", 12, 59
'CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("BEHAVIOR", scrngArray)
CALL writeValueOnScreen(currFieldName, "")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("HYG/SAFETY", scrngArray)
CALL writeValueOnScreen(currFieldName, "")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("NEG/ABUSE", scrngArray)
CALL writeValueOnScreen(currFieldName, "")
CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("FRAILTY", scrngArray)
EMwriteScreen     " ", 13, 44
'CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("INST STAYS", scrngArray)
EMwriteScreen     " ", 13, 59
'CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("HEARING IMP", scrngArray)
EMwriteScreen     "  ", 13, 74
'CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("REST/REHAB", scrngArray)
EMwriteScreen     " ", 14, 14
'CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("UNSTABLE", scrngArray)
EMwriteScreen     " ", 14, 29
'CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("SPEC TREATY", scrngArray)
EMwriteScreen     " ", 14, 44
'CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("CMPLX CARE", scrngArray)
EMwriteScreen     " ", 14, 59
'CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("VISUAL IMP", scrngArray)
EMwriteScreen     " ", 14, 74
'CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("TOL ASSIST", scrngArray)
CALL writeValueOnScreen(currFieldName, "")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("PCA COMPLEX", scrngArray)
CALL writeValueOnScreen(currFieldName, "")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("REQUIRES AC/WVR SVC", scrngArray)
CALL writeValueOnScreen(currFieldName, " ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("SAFE/COST EFFECTIVE", scrngArray)
CALL writeValueOnScreen(currFieldName, "")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("NO OTHER PAYOR IS RESP", scrngArray)
CALL writeValueOnScreen(currFieldName, "")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("PROGRAM TYPE", scrngArray)
CALL writeValueOnScreen(currFieldName, "  ")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("MHM IND", scrngArray)
CALL writeValueOnScreen(currFieldName, "")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("CDCS", scrngArray)
CALL writeValueOnScreen(currFieldName, "")
CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL findValueInArray("CDCS AMOUNT", scrngArray)
CALL writeValueOnScreen(currFieldName, "")
CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("SERVICES", scrngArray)
'EMwriteScreen     "  ", 16, 23
'CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("PROV NPI", scrngArray)
EMwriteScreen     "          ", 18, 29
'CALL writeValueOnScreen(currFieldName, currFieldValue)
'CALL findValueInArray("PERSON", scrngArray)
EMwriteScreen     " ", 18, 48
'CALL writeValueOnScreen(currFieldName, currFieldValue)
'sgBox "pause"
Call checkForErrors()


'ALT5

EMReadScreen ALT5_check, 4, 1, 52
If ALT5_check <> "ALT5" Then script_end_procedure ("The script must CONTINUE on ALT5 Screen")

CALL findValueInArray("CODE-1", scrngArray)
EMwriteScreen     "  ", 6, 3
EMwriteScreen     currFieldValue, 6, 3
CALL findValueInArray("IND-1", scrngArray)
EMwriteScreen     " ", 6, 8
EMwriteScreen     currFieldValue, 6, 8
CALL findValueInArray("CODE-2", scrngArray)
EMwriteScreen     "  ", 6, 30
EMwriteScreen     currFieldValue, 6, 30
CALL findValueInArray("IND-2", scrngArray)
EMwriteScreen     " ", 6, 35
EMwriteScreen     currFieldValue, 6, 35
CALL findValueInArray("CODE-3", scrngArray)
EMwriteScreen     "  ", 6, 57
EMwriteScreen     currFieldValue, 6, 57
CALL findValueInArray("IND-3", scrngArray)
EMwriteScreen     " ", 6, 62
EMwriteScreen     currFieldValue, 6, 62
CALL findValueInArray("CODE-4", scrngArray)
EMwriteScreen     "  ", 7, 3
EMwriteScreen     currFieldValue, 7, 3
CALL findValueInArray("IND-4", scrngArray)
EMwriteScreen     " ", 7, 8
EMwriteScreen     currFieldValue, 7, 8
CALL findValueInArray("CODE-5", scrngArray)
EMwriteScreen     "  ", 7, 30
EMwriteScreen     currFieldValue, 7, 30
CALL findValueInArray("IND-5", scrngArray)
EMwriteScreen     " ", 7, 35
EMwriteScreen     currFieldValue, 7, 35
CALL findValueInArray("CODE-6", scrngArray)
EMwriteScreen     "  ", 7, 57
EMwriteScreen     currFieldValue, 7, 57
CALL findValueInArray("IND-6", scrngArray)
EMwriteScreen     " ", 7, 62
EMwriteScreen     currFieldValue, 7, 62
CALL findValueInArray("CODE-7", scrngArray)
EMwriteScreen     "  ", 8, 3
EMwriteScreen     currFieldValue, 8, 3
CALL findValueInArray("IND-7", scrngArray)
EMwriteScreen     " ", 8, 8
EMwriteScreen     currFieldValue, 8, 8
CALL findValueInArray("CODE-8", scrngArray)
EMwriteScreen     "  ", 8, 30
EMwriteScreen     currFieldValue, 8, 30
CALL findValueInArray("IND-8", scrngArray)
EMwriteScreen     " ", 8, 35
EMwriteScreen     currFieldValue, 8, 35
CALL findValueInArray("CODE-9", scrngArray)
EMwriteScreen     "  ", 8, 57
EMwriteScreen     currFieldValue, 8, 57
CALL findValueInArray("IND-9", scrngArray)
EMwriteScreen     " ", 8, 62
EMwriteScreen     currFieldValue, 8, 62
CALL findValueInArray("CODE-10", scrngArray)
EMwriteScreen     "  ", 9, 3
EMwriteScreen     currFieldValue, 9, 3
CALL findValueInArray("IND-10", scrngArray)
EMwriteScreen     " ", 9, 8
EMwriteScreen     currFieldValue, 9, 8
CALL findValueInArray("CODE-11", scrngArray)
EMwriteScreen     "  ", 9, 30
EMwriteScreen     currFieldValue, 9, 30
CALL findValueInArray("IND-11", scrngArray)
EMwriteScreen     " ", 9, 35
EMwriteScreen     currFieldValue, 9, 35
CALL findValueInArray("CODE-12", scrngArray)
EMwriteScreen     "  ", 9, 57
EMwriteScreen     currFieldValue, 9, 57
CALL findValueInArray("IND-12", scrngArray)
EMwriteScreen     " ", 9, 62
EMwriteScreen     currFieldValue, 9, 62
CALL findValueInArray("CODE-13", scrngArray)
EMwriteScreen     "  ", 10, 3
EMwriteScreen     currFieldValue, 10, 3
CALL findValueInArray("IND-13", scrngArray)
EMwriteScreen     " ", 10, 8
EMwriteScreen     currFieldValue, 10, 8
CALL findValueInArray("CODE-14", scrngArray)
EMwriteScreen     "  ", 10, 30
EMwriteScreen     currFieldValue, 10, 30
CALL findValueInArray("IND-14", scrngArray)
EMwriteScreen     " ", 10, 35
EMwriteScreen     currFieldValue, 10, 35
CALL findValueInArray("CODE-15", scrngArray)
EMwriteScreen     "  ", 10, 57
EMwriteScreen     currFieldValue, 10, 57
CALL findValueInArray("IND-15", scrngArray)
EMwriteScreen     " ", 10, 62
EMwriteScreen     currFieldValue, 10, 62
CALL findValueInArray("CODE-16", scrngArray)
EMwriteScreen     "  ", 11, 3
EMwriteScreen     currFieldValue, 11, 3
CALL findValueInArray("IND-16", scrngArray)
EMwriteScreen     " ", 11, 8
EMwriteScreen     currFieldValue, 11, 8
CALL findValueInArray("CODE-17", scrngArray)
EMwriteScreen     "  ", 11, 30
EMwriteScreen     currFieldValue, 11, 30
CALL findValueInArray("IND-17", scrngArray)
EMwriteScreen     " ", 11, 35
EMwriteScreen     currFieldValue, 11, 35
CALL findValueInArray("CODE-18", scrngArray)
EMwriteScreen     "  ", 11, 57
EMwriteScreen     currFieldValue, 11, 57
CALL findValueInArray("IND-18", scrngArray)
EMwriteScreen     " ", 11, 62
EMwriteScreen     currFieldValue, 11, 62
'sgBox "pause"
Call checkForErrors()

'Alt6

EMReadScreen Alt6, 4, 1, 52
If Alt6 = "ALT6"  THEN
   CALL Populate_Alt6
End If

 'ACMG
' MsgBox "Going to ACMG"
'CALL findValueInArray("CASE MANAGER COMMENTS", scrngArray)
EMwriteScreen     "                                                                      ", 6, 3
'CALL writeValueOnScreen(currFieldName, currFieldValue)
CALL checkForErrors()
CALL checkForExceptions()

'ALT6

Function Populate_Alt6

CALL findValueInArray("STREET ADDRESS-1", scrngArray)
EMwriteScreen     "                  ", 5, 18
EMwriteScreen     currFieldValue, 5, 18
CALL findValueInArray("STREET ADDRESS-2", scrngArray)
EMwriteScreen     "                  ", 6, 18
EMwriteScreen     currFieldValue, 6, 18
CALL findValueInArray("CITY", scrngArray)
EMwriteScreen     "                  ", 7, 18
EMwriteScreen     currFieldValue, 7, 18
CALL findValueInArray("STATE", scrngArray)
EMwriteScreen     "  ", 7, 45
EMwriteScreen     currFieldValue, 7, 45
CALL findValueInArray("ZIP CODE", scrngArray)
EMwriteScreen     "      ", 7, 59
EMwriteScreen     currFieldValue, 7, 59
CALL findValueInArray("CFR", scrngArray)
EMwriteScreen     "   ", 7, 72
EMwriteScreen     currFieldValue, 7, 72
CALL findValueInArray("AC/ECS GROSS INCOME", scrngArray)
EMwriteScreen     "     ", 9, 26
EMwriteScreen     currFieldValue, 9, 26
CALL findValueInArray("AC/ECS GROSS ASSETS", scrngArray)
EMwriteScreen     "     ", 9, 59
EMwriteScreen     currFieldValue, 9, 59
CALL findValueInArray("AC/ECS ADJUSTED INCOME", scrngArray)
EMwriteScreen     "     ", 10, 26
EMwriteScreen     currFieldValue, 10, 26
CALL findValueInArray("AC/ECS ADJUSTED ASSETS", scrngArray)
EMwriteScreen     "     ", 10, 59
EMwriteScreen     currFieldValue, 10, 59
CALL findValueInArray("MEDICARE/MBI ID NUMBER", scrngArray)
EMwriteScreen     "            ", 12, 26
EMwriteScreen     currFieldValue, 12, 26
CALL findValueInArray("MEDICARE/MBI PART A BEGIN DT", scrngArray)
EMwriteScreen     "        ", 13, 32
EMwriteScreen     currFieldValue, 13, 32
CALL findValueInArray("MEDICARE/MBI PART A END DT", scrngArray)
EMwriteScreen     "        ", 13, 50
EMwriteScreen     currFieldValue, 13, 50
CALL findValueInArray("MEDICARE/MBI PART B BEGIN DT", scrngArray)
EMwriteScreen     "        ", 14, 32
EMwriteScreen     currFieldValue, 14, 32
CALL findValueInArray("MEDICARE/MBI PART B END DT", scrngArray)
EMwriteScreen     "        ", 14, 50
EMwriteScreen     currFieldValue, 14, 50
CALL findValueInArray("AC FEE WAIVER REASON", scrngArray)
EMwriteScreen     "  ", 16, 28
EMwriteScreen     currFieldValue, 16, 28
CALL findValueInArray("MED ELIGIBLE", scrngArray)
EMwriteScreen     " ", 16, 56
EMwriteScreen     currFieldValue, 16, 56
CALL findValueInArray("AC FEE ASSESSED", scrngArray)
EMwriteScreen     " ", 16, 77
EMwriteScreen     currFieldValue, 16, 77
CALL findValueInArray("CITIZENSHIP", scrngArray)
EMwriteScreen     " ", 17, 28
EMwriteScreen     currFieldValue, 17, 28

Call checkForErrors()

End Function

StopScript
'PF3
'StopScript