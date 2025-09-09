'LOADING GLOBAL VARIABLES--------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Access Aging & Disability Services\Scripts\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
'name_of_script = ""
'start_time = timer

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
call changelog_update("03/14/2025", "Initial scripts set up in Hennepin County.", "Casey Love, Hennepin County" )
'("10/16/2019", "All infrastructure changed to run locally and stored in BlueZone Scripts ccm. MNIT @ DHS)
'(12/09/2024", Added code to blankout previous current an planned services data. Also added code to fill in the second page "ADD3"
' with current and Planned services data.
'END CHANGELOG BLOCK =======================================================================================================

CALL file_selection_system_dialog(xmlPath, ".xml")

'xmlPath = "C:\Users\PWTMT03\Desktop\MNChoices\LTC Screening Document.xml"
'xmlPath = "C:\Users\PWTMT03\Desktop\MNChoices\LTC Screening Document14-2.xml"
'xmlPath = "C:\Users\PWTMT03\Desktop\MNChoices\DD Screening Document14-3.xml"
'xmlPath = "C:\Users\PWTMT03\Desktop\MNChoices\DD Screening Document ID 13304.xml"

Set ObjFso = CreateObject("Scripting.FileSystemObject")
Set ObjFile = ObjFso.OpenTextFile(xmlPath)

Dim scrngArray()

currFieldValue = "UNDEFINED"
currFieldName = "UNDEFINED"

'tmpACMG = "What is an Array? We know very well that a variable is a container to store a value. Sometimes, developers are in a position to hold more than one value in a single variable at a time. When a series of values is stored in a single variable, then it is known as an array variable. Array Declaration Arrays are declared the same way a variable has been declared except that the declaration of an array variable uses parenthesis. In the following example, the size of the array is mentioned in the brackets. Multi Dimension Arrays Arrays are not just limited to single dimension and can have a maximum of 60 dimensions. Two-dimension arrays are the most commonly used ones. Example in the following example, a multi-dimension array is declared with 3 rows and 4 columns. Array Methods There are various inbuilt functions within VBScript which help the developers to handle arrays effectively. All the methods that are used in conjunction with arrays are listed below. Please click on the method name to know in detail."

'tmpACMG="What_is_an_Array?_We_know_very_well_that_a_variable_is_a_container_to_store_a_value._Sometimes,_developers_are_in_a_position_to_hold_more_than_one_value_in_a_single_variable_at_a_time._When_a_series_of_values_is_stored_in_a_single_variable,_then_it_is_known_as_an_array_variable._Array_Declaration_Arrays_are_declared_the_same_way_a_variable_has_been_declared_except_that_the_declaration_of_an_array_variable_uses_parenthesis._In_the_following_example,_the_size_of_the_array_is_mentioned_in_the_brackets._Multi_Dimension_Arrays_Arrays_are_not_just_limited_to_single_dimension_and_can_have_a_maximum_of_60_dimensions._Two-dimension_arrays_are_the_most_commonly_used_ones._Example_in_the_following_example,_a_multi-dimension_array_is_declared_with_3_rows_and_4_columns._Array_Methods_There_are_various_inbuilt_functions_within_VBScript_which_help_the_developers_to_handle_arrays_effectively._All_the_methods_that_are_used_in_conjunction_with_arrays_are_listed_below._Please_click_on_the_method_name_to_know_in_detail."

'Function createNavDialog(navTitle, navErr, xDim, yDim)
'	BeginDialog navDialog, 0, 0, xDim, yDim, "" & navTitle
'		ButtonGroup ButtonPressed
'		OkButton xDim-100, yDim-20, 50, 15
'		CancelButton xDim-50, yDim-20, 50, 15
'		Text 5, 10, xDim-5, yDim-40, navErr
'	EndDialog
'	Do
'		Dialog navDialog
'		If ButtonPressed = 0 then
'			PF6
'			StopScript
'		End IF
'	LOOP UNTIL ButtonPressed = -1
'	PF3
'	'StopScript
'End Function

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

'Function findKeywordOnScreen(scrKeyword)
'	scrY = 1
'	scrLine = ""
'	Do
'		EMReadScreen scrLine, 80, scry, 1
'		If inStr(scrLine, scrKeyword) <> 0 Then
'			findKeywordOnScreen = scrLine
'			scrY = 100
'		End If
'	scrY = scrY + 1
'	Loop until scrY >= 24
'End Function

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

Function searchForMenuItemAndSelect(menuItem)
	scrY = 1
	menuLine = ""
	Do
		EMReadScreen menuLine, 80, scrY, 1
		If inStr(menuLine, menuItem) <> 0 Then
			EMWriteScreen "X", scrY, inStr(menuLine, menuItem) - 3
			scrY = 100
			EMSendKey "<enter>"
		End If
	scrY = scrY + 1
	Loop until scrY >= 24

	If scrY = 24 Then
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
		'msgbox "currFieldName " & currFieldName & " currfieldValue " & currfieldValue & " Z " & z
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
' offSet is the value of spaces to skip before writing the value (usually 1, unless multi-field)
Function writeValueOnScreen(fieldName, fieldValue, offSet)
	scrX = 1
	scrLine = ""
	If IsNumeric(fieldName) = False Then
		If right(fieldName, 1) <> ":" Then fieldName = fieldName & ":"
	End IF
	If fieldName = "ACT DT:" Then fieldName = "ACT DT"
	Do
		EMReadScreen scrLine, 80, scrX, 1
		If inStr(scrLine, fieldName) <> 0 Then
			EMWriteScreen fieldValue, scrX, inStr(scrLine, fieldName) + len(fieldName) + offSet
			ScrX = 100
		End If
	scrX = scrX + 1
	Loop until scrX >= 24

	If scrX = 24 Then
		MsgBox "ERROR: '" & fieldName & "' was not found on current screen!  Aborting.."
		stopscript
	End If
End Function

Function customWriteADD3()
	CALL findValueInArray("INFORMED CHOICE (Y/N)", scrngArray)
	CALL writeValueOnScreen("INFORMED CHOICE(Y/N)", currFieldValue, 1)
	scrX = 1
	scrLine = ""
	ADD3StartRow = 0
	ADD3EndRow = 0
	ADD3TotalRows = 0
	servicesArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P")
	Do
		EMReadScreen scrLine, 5, scrX, 1
		'Msgbox "scrLine " & scrLine
		If IsNumeric(trim(scrLine)) Then
			If trim(scrLine) = 1 Then
				ADD3StartRow = scrX
				ADD3EndRow = scrX - 1
				'Msgbox "End row " & ADD3EndRow
			End If
			ADD3EndRow = ADD3EndRow + 1
			Msgbox "End row " & ADD3EndRow
		End If
	scrX = scrX + 1
	Loop until scrX >= 24
	ADD3TotalRows = (ADD3EndRow - ADD3StartRow) + 1
	'MsgBox "ADD3StartRow is: " & ADD3StartRow & "/ ADD3EndRow is: " & ADD3EndRow & "/ ADD3TotalRows is: " & ADD3TotalRows
	z = 1
	Do While z <= ADD3TotalRows
		CALL findValueInArray("CURRENT SERVICES-" & servicesArray(z - 1), scrngArray)
		'Msgbox "scrn array " & scrnArray
		If z <= 9 Then
		    Msgbox "curr value " & currFieldValue
			CALL writeValueOnScreen(" 0" & z & " ", currFieldValue, 1)
			'Msgbox "curr value " & currFieldValue
		ELSE
			CALL writeValueOnScreen(" " & z & " ", currFieldValue, 1)
			Msgbox "curr value " & currFieldValue
		End IF
		CALL findValueInArray("PLANNED SERVICES-" & servicesArray(z - 1), scrngArray)
		If z <= 9 Then
			CALL writeValueOnScreen(" 0" & z & "  ", currFieldValue, 39)
		ELSE
			CALL writeValueOnScreen(" " & z & "  ", currFieldValue, 39)
		End IF
		z = z + 1
	Loop
End Function

Function customWrite2ndADD3()

	scrX = 1
	scrLine = ""
	ADD3StartRow = 0
	ADD3EndRow = 0
	ADD3TotalRows = 0
	servicesArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P")
	Do
		EMReadScreen scrLine, 5, scrX, 1
		If IsNumeric(trim(scrLine)) Then
			If trim(scrLine) = 1 Then
				ADD3StartRow = scrX
				ADD3EndRow = scrX - 1
			End If
			ADD3EndRow = ADD3EndRow + 1
		End If
	scrX = scrX + 1
	Loop until scrX >= 24
	ADD3TotalRows = (ADD3EndRow - ADD3StartRow) + 1
	'MsgBox "ADD3StartRow is: " & ADD3StartRow & "/ ADD3EndRow is: " & ADD3EndRow & "/ ADD3TotalRows is: " & ADD3TotalRows
	z = 1
	Do While z <= ADD3TotalRows
		CALL findValueInArray("CURRENT SERVICES-" & servicesArray(z - 1), scrngArray)
		If z <= 9 Then
			CALL writeValueOnScreen(" 0" & z & " ", currFieldValue, 1)
		ELSE
			CALL writeValueOnScreen(" " & z & " ", currFieldValue, 1)
		End IF
		CALL findValueInArray("PLANNED SERVICES-" & servicesArray(z - 1), scrngArray)
		If z <= 9 Then
			CALL writeValueOnScreen(" 0" & z & "  ", currFieldValue, 39)
		ELSE
			CALL writeValueOnScreen(" " & z & "  ", currFieldValue, 39)
		End IF
		z = z + 1
	Loop
End Function


Function customWriteMNChoicesComments(commentFieldName)
	CALL findValueInArray(commentFieldName, scrngArray)
	scrX = 6
	commentLine = ""
	commentArray = split(currFieldValue, " ")
	'commentArray = split(tmpACMG, " ")
	'MsgBox "Words in CASE MANAGER COMMENTS: " & ubound(commentArray)
	For Each commentArrayItem In commentArray
		If len(commentArrayItem) + len(commentLine) >= 69 Then
			If ScrX <= 18 Then EMWriteScreen commentLine, scrX, 3
			commentLine = ""
			scrX = scrX + 1
		End If
		commentLine = commentLine & " " & commentArrayItem
	Next
	If ScrX <= 18 Then EMWriteScreen commentLine, scrX, 3
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
			   vbCrlf & vbCrlf & "Error Report Created:" & vbCRlf & logFilename, 0, "DD Script Error"
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

Function processXML()
	Do While ObjFile.AtEndOfStream <> True
		StrData = ObjFile.ReadLine
		StrData = trim(StrData)
		CALL BuildValueArray(StrData, scrngArray)
	Loop
End Function

CALL processXML()
ObjFile.close
CALL findValueInArray("DOCUMENT TYPE", scrngArray)

If currFieldValue <> "D" Then
	MsgBox """" & xmlPath & """ is not a DD Document." & vbCrlf & vbCrlf &_
	"Please select a valid DD XML document and try the script again.", 0, "DD Script Error"
	StopScript
End If

EMReadScreen scrCheck, 80, 1, 1
If InStr(CStr(scrCheck), "MMIS MAIN MENU - MAIN") Then
	CALL searchForMenuItemAndSelect("SCREENINGS")
End If

EMReadScreen scrCheck, 80, 1, 1
If InStr(CStr(scrCheck), "MMIS MAIN MENU - MAIN") = 0 Then
	If InStr(CStr(scrCheck), "MMIS SCRNG KEY PANEL-ASCR") = 0 Then
		MsgBox "This script must start on MAIN or SCREENINGS.", 0, "DD Script Error"
		StopScript
	End If
End If

'ASCR
CALL writeValueOnScreen("ACTION CODE", " ", 1) 'check for timeout
Transmit 'check for timeout
EMReadScreen scrCheck, 80, 1, 1
If InStr(CStr(scrCheck), "MMIS SCRNG KEY PANEL-ASCR") = 0 Then
	MsgBox "This script must start on MAIN or SCREENINGS.", 0, "DD Script Error"
	StopScript
End If
CALL writeValueOnScreen("ACTION CODE", "A", 1)
CALL findValueInArray("DOCUMENT TYPE", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
'CALL writeValueOnScreen(currFieldName, "X", 1)
CALL findValueInArray("RECIPIENT ID", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
'CALL writeValueOnScreen(currFieldName, "00000008", 1)
CALL checkForErrors()

'ADD1
CALL findValueInArray("CO REF NBR", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("DOB", scrngArray)
CALL writeValueOnScreen("DOB(MMDDYYYY)", currFieldValue, 1)
'CALL writeValueOnScreen("DOB(MMDDYYYY)", "01281984", 1)
CALL findValueInArray("REF DATE", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("GRDN STAT", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
'CALL findValueInArray("CO OF SVC", scrngArray)
'CALL writeValueOnScreen("CO OF SVC/RES", currFieldValue, 1)
'CALL findValueInArray("RES", scrngArray)
'CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
'CALL findValueInArray("CFR", scrngArray)
'CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("DIAG-1", scrngArray)
CALL writeValueOnScreen("DIAG 1-4", "        ", 1)
CALL writeValueOnScreen("DIAG 1-4", currFieldValue, 1)
CALL findValueInArray("DIAG-2", scrngArray)
CALL writeValueOnScreen("DIAG 1-4", "        ", 10)
CALL writeValueOnScreen("DIAG 1-4", currFieldValue, 10)
CALL findValueInArray("DIAG-3", scrngArray)
CALL writeValueOnScreen("DIAG 1-4", "        ", 19)
CALL writeValueOnScreen("DIAG 1-4", currFieldValue, 19)
CALL findValueInArray("DIAG-4", scrngArray)
CALL writeValueOnScreen("DIAG 1-4", "        ", 28)
CALL writeValueOnScreen("DIAG 1-4", currFieldValue, 28)
CALL findValueInArray("CM NAME", scrngArray)
CALL writeValueOnScreen("CM NAME/NBR", currFieldValue, 1)
CALL findValueInArray("CM NBR", scrngArray)
CALL writeValueOnScreen("CM NAME/NBR", currFieldValue, 37)
CALL findValueInArray("PRES AT SCRNG: RECIP", scrngArray)
CALL writeValueOnScreen("PRES AT SCRNG(Y/N)", currFieldValue, 1)
CALL findValueInArray("PRES AT SCRNG: LGL REP", scrngArray)
CALL writeValueOnScreen("PRES AT SCRNG(Y/N)", currFieldValue, 8)
CALL findValueInArray("PRES AT SCRNG: CASE MGR", scrngArray)
CALL writeValueOnScreen("PRES AT SCRNG(Y/N)", currFieldValue, 17)
CALL findValueInArray("PRES AT SCRNG: QDDP", scrngArray)
CALL writeValueOnScreen("PRES AT SCRNG(Y/N)", currFieldValue, 24)
CALL findValueInArray("PRES AT SCRNG: OTHER", scrngArray)
CALL writeValueOnScreen("PRES AT SCRNG(Y/N)", currFieldValue, 29)
CALL findValueInArray("ACTION DT", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("ACTION TYPE", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("TEAM CONVENED (Y/N)", scrngArray)
CALL writeValueOnScreen("TEAM CONVENED(Y/N)", currFieldValue, 1)
CALL findValueInArray("MEDICAL", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("VISION", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("HEARING", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("SEIZURES", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("MOBILITY", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("FINE MOTOR SKILLS", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("EXPRESSIVE", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("RECEPTIVE", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL checkForErrors()

'ADD2
CALL findValueInArray("SELF PRESERVATION", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("VOCATIONAL", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("(A.) SELF CARE", scrngArray)
CALL writeValueOnScreen("SELF CARE (A)", currFieldValue, 1)
CALL findValueInArray("(B.) DAILY LIVING", scrngArray)
CALL writeValueOnScreen("DAILY LIVING (B)", currFieldValue, 1)
CALL findValueInArray("(C.) MONEY MANAGEMENT", scrngArray)
CALL writeValueOnScreen("MONEY MANAGEMENT (C)", currFieldValue, 1)
CALL findValueInArray("(D.) COMMUNITY LIVING", scrngArray)
CALL writeValueOnScreen("COMMUNITY LIVING (D)", currFieldValue, 1)
CALL findValueInArray("(E.) LEISURE RECREATION", scrngArray)
CALL writeValueOnScreen("LEISURE RECREATION (E)", currFieldValue, 1)
CALL findValueInArray("SUPPORT SERVICES", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("(A.) EATING NON-NUTRITIVE", scrngArray)
CALL writeValueOnScreen("EATING NON-NUTRITIVE (A)", currFieldValue, 1)
CALL findValueInArray("(B.) INJURIOUS", scrngArray)
CALL writeValueOnScreen("INJURIOUS (B)", currFieldValue, 1)
CALL findValueInArray("(C.) AGGRESS/PHYSICAL", scrngArray)
CALL writeValueOnScreen("AGGRESS/PHYSICAL (C)", currFieldValue, 1)
CALL findValueInArray("(D.) AGGRESS/VERBAL", scrngArray)
CALL writeValueOnScreen("AGGRESS/VERBAL (D)", currFieldValue, 1)
CALL findValueInArray("(E.) SEXUAL BEHAVIOR", scrngArray)
CALL writeValueOnScreen("SEXUAL BEHAVIOR (E)", currFieldValue, 1)
CALL findValueInArray("(F.) PROPERTY DEST", scrngArray)
CALL writeValueOnScreen("PROPERTY DEST (F)", currFieldValue, 1)
CALL findValueInArray("(G.) RUNS AWAY", scrngArray)
CALL writeValueOnScreen("RUNS AWAY (G)", currFieldValue, 1)
CALL findValueInArray("(H.) BREAKS LAWS", scrngArray)
CALL writeValueOnScreen("BREAKS LAWS (H)", currFieldValue, 1)
CALL findValueInArray("(I.) TEMPER OUTBURSTS", scrngArray)
CALL writeValueOnScreen("TEMPER OUTBURSTS (I)", currFieldValue, 1)
CALL findValueInArray("(J.) OTHER", scrngArray)
CALL writeValueOnScreen("OTHER (J)", currFieldValue, 1)
CALL findValueInArray("LEVEL OF CARE", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL checkForErrors()

'ADD3 Current Services / Planned Services Page 1
'Call customWriteADD3()
'stopscript

CALL findValueInArray("INFORMED CHOICE (Y/N)", scrngArray)
CALL writeValueOnScreen("INFORMED CHOICE(Y/N)", " ", 1)
CALL writeValueOnScreen("INFORMED CHOICE(Y/N)", currFieldValue, 1)

CALL findValueInArray("CURRENT SERVICES-A", scrngArray)
EMwriteScreen     "  ", 8, 6
EMwriteScreen     currFieldValue, 8, 6

CALL findValueInArray("PLANNED SERVICES-A", scrngArray)
EMwriteScreen     "  ", 8, 45
EMwriteScreen     currFieldValue, 8, 45

CALL findValueInArray("CURRENT SERVICES-B", scrngArray)
EMwriteScreen     "  ", 9, 6
EMwriteScreen     currFieldValue, 9, 6

CALL findValueInArray("PLANNED SERVICES-B", scrngArray)
EMwriteScreen     "  ", 9, 45
EMwriteScreen     currFieldValue, 9, 45

CALL findValueInArray("CURRENT SERVICES-C", scrngArray)
EMwriteScreen     "  ", 10, 6
EMwriteScreen     currFieldValue, 10, 6

CALL findValueInArray("PLANNED SERVICES-C", scrngArray)
EMwriteScreen     "  ", 10, 45
EMwriteScreen     currFieldValue, 10, 45

CALL findValueInArray("CURRENT SERVICES-D", scrngArray)
EMwriteScreen     "  ", 11, 6
EMwriteScreen     currFieldValue, 11, 6

CALL findValueInArray("PLANNED SERVICES-D", scrngArray)
EMwriteScreen     "  ", 11, 45
EMwriteScreen     currFieldValue, 11, 45

CALL findValueInArray("CURRENT SERVICES-E", scrngArray)
EMwriteScreen     "  ", 12, 6
EMwriteScreen     currFieldValue, 12, 6

CALL findValueInArray("PLANNED SERVICES-E", scrngArray)
EMwriteScreen     "  ", 12, 45
EMwriteScreen     currFieldValue, 12, 45

CALL findValueInArray("CURRENT SERVICES-F", scrngArray)
EMwriteScreen     "  ", 13, 6
EMwriteScreen     currFieldValue, 13, 6

CALL findValueInArray("PLANNED SERVICES-F", scrngArray)
EMwriteScreen     "  ", 13, 45
EMwriteScreen     currFieldValue, 13, 45

CALL findValueInArray("CURRENT SERVICES-G", scrngArray)
EMwriteScreen     "  ", 14, 6
EMwriteScreen     currFieldValue, 14, 6

CALL findValueInArray("PLANNED SERVICES-G", scrngArray)
EMwriteScreen     "  ", 14, 45
EMwriteScreen     currFieldValue, 14, 45

CALL findValueInArray("CURRENT SERVICES-H", scrngArray)
EMwriteScreen     "  ", 15, 6
EMwriteScreen     currFieldValue, 15, 6

CALL findValueInArray("PLANNED SERVICES-H", scrngArray)
EMwriteScreen     "  ", 15, 45
EMwriteScreen     currFieldValue, 15, 45

CALL findValueInArray("CURRENT SERVICES-I", scrngArray)
EMwriteScreen     "  ", 16, 6
EMwriteScreen     currFieldValue, 16, 6

CALL findValueInArray("PLANNED SERVICES-I", scrngArray)
EMwriteScreen     "  ", 16, 45
EMwriteScreen     currFieldValue, 16, 45

CALL findValueInArray("CURRENT SERVICES-J", scrngArray)
EMwriteScreen     "  ", 17, 6
EMwriteScreen     currFieldValue, 17, 6

CALL findValueInArray("PLANNED SERVICES-J", scrngArray)
EMwriteScreen     "  ", 17, 45
EMwriteScreen     currFieldValue, 17, 45

CALL findValueInArray("CURRENT SERVICES-K", scrngArray)
EMwriteScreen     "  ", 18, 6
EMwriteScreen     currFieldValue, 18, 6

CALL findValueInArray("PLANNED SERVICES-K", scrngArray)
EMwriteScreen     "  ", 18, 45
EMwriteScreen     currFieldValue, 18, 45

'ADD3 Current Services / Planned Services Page 2

EMSetCursor 8, 6
PF8

CALL findValueInArray("CURRENT SERVICES-L", scrngArray)
EMwriteScreen     "  ", 8, 6
EMwriteScreen     currFieldValue, 8, 6

CALL findValueInArray("PLANNED SERVICES-L", scrngArray)
EMwriteScreen     "  ", 8, 45
EMwriteScreen     currFieldValue, 8, 45

CALL findValueInArray("CURRENT SERVICES-M", scrngArray)
EMwriteScreen     "  ", 9, 6
EMwriteScreen     currFieldValue, 9, 6

CALL findValueInArray("PLANNED SERVICES-M", scrngArray)
EMwriteScreen     "  ", 9, 45
EMwriteScreen     currFieldValue, 9, 45

CALL findValueInArray("CURRENT SERVICES-N", scrngArray)
EMwriteScreen     "  ", 10, 6
EMwriteScreen     currFieldValue, 10, 6

CALL findValueInArray("PLANNED SERVICES-N", scrngArray)
EMwriteScreen     "  ", 10, 45
EMwriteScreen     currFieldValue, 10, 45

CALL findValueInArray("CURRENT SERVICES-O", scrngArray)
EMwriteScreen     "  ", 11, 6
EMwriteScreen     currFieldValue, 11, 6

CALL findValueInArray("PLANNED SERVICES-O", scrngArray)
EMwriteScreen     "  ", 11, 45
EMwriteScreen     currFieldValue, 11, 45

CALL findValueInArray("CURRENT SERVICES-P", scrngArray)
EMwriteScreen     "  ", 12, 6
EMwriteScreen     currFieldValue, 12, 6

CALL findValueInArray("PLANNED SERVICES-P", scrngArray)
EMwriteScreen     "  ", 12, 45
EMwriteScreen     currFieldValue, 12, 45

CALL checkForErrors()

'ADD4
'CALL findValueInArray("DT&H SERV AUTH LEVEL", scrngArray)
'CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("WAIVER NEED INDEX-1", scrngArray)
CALL writeValueOnScreen("WAIVER NEED INDEX", currFieldValue, 1)
CALL findValueInArray("WAIVER NEED INDEX-2", scrngArray)
CALL writeValueOnScreen("WAIVER NEED INDEX", currFieldValue, 5)
CALL findValueInArray("WAIVER NEED INDEX-3", scrngArray)
CALL writeValueOnScreen("WAIVER NEED INDEX", currFieldValue, 9)
CALL findValueInArray("(A.) SPEC MEDICAL SERV", scrngArray)
CALL writeValueOnScreen("SPEC MEDICAL SERV (A)", currFieldValue, 1)
CALL findValueInArray("(B.) PHYSICAL THPY", scrngArray)
CALL writeValueOnScreen("PHYSICAL THPY (B)", currFieldValue, 1)
CALL findValueInArray("(C.) OCCUPATIONAL THPY", scrngArray)
CALL writeValueOnScreen("OCCUPATIONAL THPY (C)", currFieldValue, 1)
CALL findValueInArray("(D.) COMM/SPEECH THPY", scrngArray)
CALL writeValueOnScreen("COMM/SPEECH THPY (D)", currFieldValue, 1)
CALL findValueInArray("(E.) TRANSPORTATION", scrngArray)
CALL writeValueOnScreen("TRANSPORTATION (E)", currFieldValue, 1)
CALL findValueInArray("(F.) EXCESSIVE BEHAVIOR", scrngArray)
CALL writeValueOnScreen("EXCESSIVE BEHAVIOR (F)", currFieldValue, 1)
CALL findValueInArray("(G.) MENTAL HEALTH", scrngArray)
CALL writeValueOnScreen("MENTAL HEALTH (G)", currFieldValue, 1)
CALL findValueInArray("(H.) EARLY INTERVENTION", scrngArray)
CALL writeValueOnScreen("EARLY INTERVENTION (H)", currFieldValue, 1)
CALL findValueInArray("(I.) OTHER", scrngArray)
CALL writeValueOnScreen("OTHER (I)", currFieldValue, 1)
CALL findValueInArray("(A.) RCP/L REP", scrngArray)
CALL writeValueOnScreen("RCP/L REP(A)", currFieldValue, 1)
CALL findValueInArray("(B.) CASE MGR", scrngArray)
CALL writeValueOnScreen("CASE MGR(B)", currFieldValue, 1)
CALL findValueInArray("(C.) QDDP", scrngArray)
CALL writeValueOnScreen("QDDP(C)", currFieldValue, 1)
CALL findValueInArray("ASSESSMENT RESULTS", scrngArray)
CALL writeValueOnScreen("ASSESSMENT RESULTS/EXIT RSNS", currFieldValue, 1)
CALL findValueInArray("EXIT RSNS", scrngArray)
CALL writeValueOnScreen("ASSESSMENT RESULTS/EXIT RSNS", currFieldValue, 15)
CALL findValueInArray("EFFECTIVE DT", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("MCAID SVC PROG", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("CO USE ONLY-1", scrngArray)
CALL writeValueOnScreen("CO USE ONLY", currFieldValue, 1)
CALL findValueInArray("CO USE ONLY-2", scrngArray)
CALL writeValueOnScreen("CO USE ONLY", currFieldValue, 6)
CALL findValueInArray("CO USE ONLY-3", scrngArray)
CALL writeValueOnScreen("CO USE ONLY", currFieldValue, 11)
CALL findValueInArray("CASE MGR SIG", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("QDDP SIG", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("PERSON/LGL REP SIG", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("CFR SIG", scrngArray)
CALL writeValueOnScreen(currFieldName, currFieldValue, 1)
CALL findValueInArray("PAYMENT AUTHORIZED", scrngArray)
CALL writeValueOnScreen("PMT AUTHORIZED", currFieldValue, 1)
CALL checkForErrors()

'ADHS
CALL checkForErrors()

'ACMG
customWriteMNChoicesComments("CASE MANAGER COMMENTS")
CALL checkForErrors()

'ARCP
CALL checkForErrors()
CALL checkForExceptions()

'Happy Path
MsgBox "No errors.  No exceptions.  Saving document.."
'PF3