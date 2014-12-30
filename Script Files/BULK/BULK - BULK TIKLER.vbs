'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - BULK TIKLER.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

'----------DIALOGS----------
'The bulk-loading-case numbers dialog
BeginDialog many_case_numbers_dialog, 0, 0, 366, 240, "Enter Many Case Numbers Dialog"
  EditBox 5, 20, 55, 15, case_number_01
  EditBox 65, 20, 55, 15, case_number_02
  EditBox 125, 20, 55, 15, case_number_03
  EditBox 185, 20, 55, 15, case_number_04
  EditBox 245, 20, 55, 15, case_number_05
  EditBox 305, 20, 55, 15, case_number_06
  EditBox 5, 40, 55, 15, case_number_07
  EditBox 65, 40, 55, 15, case_number_08
  EditBox 125, 40, 55, 15, case_number_09
  EditBox 185, 40, 55, 15, case_number_10
  EditBox 245, 40, 55, 15, case_number_11
  EditBox 305, 40, 55, 15, case_number_12
  EditBox 5, 60, 55, 15, case_number_13
  EditBox 65, 60, 55, 15, case_number_14
  EditBox 125, 60, 55, 15, case_number_15
  EditBox 185, 60, 55, 15, case_number_16
  EditBox 245, 60, 55, 15, case_number_17
  EditBox 305, 60, 55, 15, case_number_18
  EditBox 5, 80, 55, 15, case_number_19
  EditBox 65, 80, 55, 15, case_number_20
  EditBox 125, 80, 55, 15, case_number_21
  EditBox 185, 80, 55, 15, case_number_22
  EditBox 245, 80, 55, 15, case_number_23
  EditBox 305, 80, 55, 15, case_number_24
  EditBox 5, 100, 55, 15, case_number_25
  EditBox 65, 100, 55, 15, case_number_26
  EditBox 125, 100, 55, 15, case_number_27
  EditBox 185, 100, 55, 15, case_number_28
  EditBox 245, 100, 55, 15, case_number_29
  EditBox 305, 100, 55, 15, case_number_30
  EditBox 5, 120, 55, 15, case_number_31
  EditBox 65, 120, 55, 15, case_number_32
  EditBox 125, 120, 55, 15, case_number_33
  EditBox 185, 120, 55, 15, case_number_34
  EditBox 245, 120, 55, 15, case_number_35
  EditBox 305, 120, 55, 15, case_number_36
  EditBox 5, 140, 55, 15, case_number_37
  EditBox 65, 140, 55, 15, case_number_38
  EditBox 125, 140, 55, 15, case_number_39
  EditBox 185, 140, 55, 15, case_number_40
  EditBox 245, 140, 55, 15, case_number_41
  EditBox 305, 140, 55, 15, case_number_42
  EditBox 5, 160, 55, 15, case_number_43
  EditBox 65, 160, 55, 15, case_number_44
  EditBox 125, 160, 55, 15, case_number_45
  EditBox 185, 160, 55, 15, case_number_46
  EditBox 245, 160, 55, 15, case_number_47
  EditBox 305, 160, 55, 15, case_number_48
  EditBox 5, 180, 55, 15, case_number_49
  EditBox 65, 180, 55, 15, case_number_50
  EditBox 125, 180, 55, 15, case_number_51
  EditBox 185, 180, 55, 15, case_number_52
  EditBox 245, 180, 55, 15, case_number_53
  EditBox 305, 180, 55, 15, case_number_54
  EditBox 5, 200, 55, 15, case_number_55
  EditBox 65, 200, 55, 15, case_number_56
  EditBox 125, 200, 55, 15, case_number_57
  EditBox 185, 200, 55, 15, case_number_58
  EditBox 245, 200, 55, 15, case_number_59
  EditBox 305, 200, 55, 15, case_number_60
  ButtonGroup ButtonPressed
    PushButton 310, 220, 50, 15, "Next...", next_button
  Text 5, 5, 220, 10, "Enter each MAXIS case number, then press ''Next...'' when finished."
EndDialog

'A very special TIKL dialog
BeginDialog tikl_dialog, 0, 0, 246, 65, "TIKL dialog"
  EditBox 95, 5, 65, 15, tikl_date
  EditBox 5, 45, 235, 15, tikl_text
  ButtonGroup ButtonPressed
    OkButton 190, 5, 50, 15
    CancelButton 190, 25, 50, 15
  Text 5, 10, 90, 10, "TIKL date (MM/DD/YYYY):"
  Text 5, 30, 170, 10, "TIKL text (will automatically wrap to multiple lines):"
EndDialog

'Connects to BlueZone
EMConnect ""

'Checks for MAXIS
maxis_check_function


Do	
	DO
		DO
			Dialog many_case_numbers_dialog
			If buttonpressed = cancel then 
				cancel_confirmation = MsgBox ("Are you sure you want to exit? Answer Yes to exit, and No to return.", vbYesNo)
				If cancel_confirmation = vbYes then stopscript
			End if
			If (isnumeric(case_number_01) = FALSE and case_number_01 <> "") or (isnumeric(case_number_02) = FALSE and case_number_02 <> "") or _ 
			  (isnumeric(case_number_03) = FALSE and case_number_03 <> "") or (isnumeric(case_number_04) = FALSE and case_number_04 <> "") or _ 
			  (isnumeric(case_number_05) = FALSE and case_number_05 <> "") or (isnumeric(case_number_06) = FALSE and case_number_06 <> "") or _ 
			  (isnumeric(case_number_07) = FALSE and case_number_07 <> "") or (isnumeric(case_number_08) = FALSE and case_number_08 <> "") or _ 
			  (isnumeric(case_number_09) = FALSE and case_number_09 <> "") or (isnumeric(case_number_10) = FALSE and case_number_10 <> "") or _ 
			  (isnumeric(case_number_11) = FALSE and case_number_11 <> "") or (isnumeric(case_number_12) = FALSE and case_number_12 <> "") or _ 
			  (isnumeric(case_number_13) = FALSE and case_number_13 <> "") or (isnumeric(case_number_14) = FALSE and case_number_14 <> "") or _ 
			  (isnumeric(case_number_15) = FALSE and case_number_15 <> "") or (isnumeric(case_number_16) = FALSE and case_number_16 <> "") or _ 
			  (isnumeric(case_number_17) = FALSE and case_number_17 <> "") or (isnumeric(case_number_18) = FALSE and case_number_18 <> "") or _ 
			  (isnumeric(case_number_19) = FALSE and case_number_19 <> "") or (isnumeric(case_number_20) = FALSE and case_number_20 <> "") or _ 
			  (isnumeric(case_number_21) = FALSE and case_number_21 <> "") or (isnumeric(case_number_22) = FALSE and case_number_22 <> "") or _ 
			  (isnumeric(case_number_23) = FALSE and case_number_23 <> "") or (isnumeric(case_number_24) = FALSE and case_number_24 <> "") or _ 
			  (isnumeric(case_number_25) = FALSE and case_number_25 <> "") or (isnumeric(case_number_26) = FALSE and case_number_26 <> "") or _ 
			  (isnumeric(case_number_27) = FALSE and case_number_27 <> "") or (isnumeric(case_number_28) = FALSE and case_number_28 <> "") or _ 
			  (isnumeric(case_number_29) = FALSE and case_number_29 <> "") or (isnumeric(case_number_30) = FALSE and case_number_30 <> "") or _ 
			  (isnumeric(case_number_31) = FALSE and case_number_31 <> "") or (isnumeric(case_number_32) = FALSE and case_number_32 <> "") or _ 
			  (isnumeric(case_number_33) = FALSE and case_number_33 <> "") or (isnumeric(case_number_34) = FALSE and case_number_34 <> "") or _ 
			  (isnumeric(case_number_35) = FALSE and case_number_35 <> "") or (isnumeric(case_number_36) = FALSE and case_number_36 <> "") or _ 
			  (isnumeric(case_number_37) = FALSE and case_number_37 <> "") or (isnumeric(case_number_38) = FALSE and case_number_38 <> "") or _ 
			  (isnumeric(case_number_39) = FALSE and case_number_39 <> "") or (isnumeric(case_number_40) = FALSE and case_number_40 <> "") or _ 
			  (isnumeric(case_number_41) = FALSE and case_number_41 <> "") or (isnumeric(case_number_42) = FALSE and case_number_42 <> "") or _ 
			  (isnumeric(case_number_43) = FALSE and case_number_43 <> "") or (isnumeric(case_number_44) = FALSE and case_number_44 <> "") or _ 
			  (isnumeric(case_number_45) = FALSE and case_number_45 <> "") or (isnumeric(case_number_46) = FALSE and case_number_46 <> "") or _ 
			  (isnumeric(case_number_47) = FALSE and case_number_47 <> "") or (isnumeric(case_number_48) = FALSE and case_number_48 <> "") or _ 
			  (isnumeric(case_number_49) = FALSE and case_number_49 <> "") or (isnumeric(case_number_50) = FALSE and case_number_50 <> "") or _ 
			  (isnumeric(case_number_51) = FALSE and case_number_51 <> "") or (isnumeric(case_number_52) = FALSE and case_number_52 <> "") or _ 
			  (isnumeric(case_number_53) = FALSE and case_number_53 <> "") or (isnumeric(case_number_54) = FALSE and case_number_54 <> "") or _ 
			  (isnumeric(case_number_55) = FALSE and case_number_55 <> "") or (isnumeric(case_number_56) = FALSE and case_number_56 <> "") or _ 
			  (isnumeric(case_number_57) = FALSE and case_number_57 <> "") or (isnumeric(case_number_58) = FALSE and case_number_58 <> "") or _ 
			  (isnumeric(case_number_59) = FALSE and case_number_59 <> "") or (isnumeric(case_number_60) = FALSE and case_number_60 <> "") then 
				MsgBox "You must enter a numeric case number for each item, or leave it blank."
			End if
		Loop until (isnumeric(case_number_01) = True or case_number_01 = "") and (isnumeric(case_number_02) = True or case_number_02 = "") and _
		  (isnumeric(case_number_03) = True or case_number_03 = "") and (isnumeric(case_number_04) = True or case_number_04 = "") and _
		  (isnumeric(case_number_05) = True or case_number_05 = "") and (isnumeric(case_number_06) = True or case_number_06 = "") and _
		  (isnumeric(case_number_07) = True or case_number_07 = "") and (isnumeric(case_number_08) = True or case_number_08 = "") and _
		  (isnumeric(case_number_09) = True or case_number_09 = "") and (isnumeric(case_number_10) = True or case_number_10 = "") and _
		  (isnumeric(case_number_11) = True or case_number_11 = "") and (isnumeric(case_number_12) = True or case_number_12 = "") and _
		  (isnumeric(case_number_13) = True or case_number_13 = "") and (isnumeric(case_number_14) = True or case_number_14 = "") and _
		  (isnumeric(case_number_15) = True or case_number_15 = "") and (isnumeric(case_number_16) = True or case_number_16 = "") and _
		  (isnumeric(case_number_17) = True or case_number_17 = "") and (isnumeric(case_number_18) = True or case_number_18 = "") and _
		  (isnumeric(case_number_19) = True or case_number_19 = "") and (isnumeric(case_number_20) = True or case_number_20 = "") and _
		  (isnumeric(case_number_21) = True or case_number_21 = "") and (isnumeric(case_number_22) = True or case_number_22 = "") and _
		  (isnumeric(case_number_23) = True or case_number_23 = "") and (isnumeric(case_number_24) = True or case_number_24 = "") and _
		  (isnumeric(case_number_25) = True or case_number_25 = "") and (isnumeric(case_number_26) = True or case_number_26 = "") and _
		  (isnumeric(case_number_27) = True or case_number_27 = "") and (isnumeric(case_number_28) = True or case_number_28 = "") and _
		  (isnumeric(case_number_29) = True or case_number_29 = "") and (isnumeric(case_number_30) = True or case_number_30 = "") and _
		  (isnumeric(case_number_31) = True or case_number_31 = "") and (isnumeric(case_number_32) = True or case_number_32 = "") and _
		  (isnumeric(case_number_33) = True or case_number_33 = "") and (isnumeric(case_number_34) = True or case_number_34 = "") and _
		  (isnumeric(case_number_35) = True or case_number_35 = "") and (isnumeric(case_number_36) = True or case_number_36 = "") and _
		  (isnumeric(case_number_37) = True or case_number_37 = "") and (isnumeric(case_number_38) = True or case_number_38 = "") and _
		  (isnumeric(case_number_39) = True or case_number_39 = "") and (isnumeric(case_number_40) = True or case_number_40 = "") and _
		  (isnumeric(case_number_41) = True or case_number_41 = "") and (isnumeric(case_number_42) = True or case_number_42 = "") and _
		  (isnumeric(case_number_43) = True or case_number_43 = "") and (isnumeric(case_number_44) = True or case_number_44 = "") and _
		  (isnumeric(case_number_45) = True or case_number_45 = "") and (isnumeric(case_number_46) = True or case_number_46 = "") and _
		  (isnumeric(case_number_47) = True or case_number_47 = "") and (isnumeric(case_number_48) = True or case_number_48 = "") and _
		  (isnumeric(case_number_49) = True or case_number_49 = "") and (isnumeric(case_number_50) = True or case_number_50 = "") and _
		  (isnumeric(case_number_51) = True or case_number_51 = "") and (isnumeric(case_number_52) = True or case_number_52 = "") and _
		  (isnumeric(case_number_53) = True or case_number_53 = "") and (isnumeric(case_number_54) = True or case_number_54 = "") and _
		  (isnumeric(case_number_55) = True or case_number_55 = "") and (isnumeric(case_number_56) = True or case_number_56 = "") and _
		  (isnumeric(case_number_57) = True or case_number_57 = "") and (isnumeric(case_number_58) = True or case_number_58 = "") and _
		  (isnumeric(case_number_59) = True or case_number_59 = "") and (isnumeric(case_number_60) = True or case_number_60 = "")
		Do
			Do
				DIALOG tikl_dialog
				If buttonpressed = cancel then 
					cancel_confirmation = MsgBox ("Are you sure you want to exit? Answer Yes to exit, and No to return.", vbYesNo)
					If cancel_confirmation = vbYes then stopscript
				End if
			Loop until buttonpressed = OK		'Included because if the user selects cancel, then backs out of that, I don't want it trying to validate dates.
			IF isdate(tikl_date) = FALSE THEN MsgBox "Please enter a valid date (MM/DD/YYYY)."
		Loop until isdate(tikl_date) = True		'Doing this here, because below it could introduce an error by trying to datediff a non-date element
		IF datediff("D", date, tikl_date) < 0 THEN MSGBOX "You must set a TIKL date NOT in the past."
	Loop until datediff("D", date, tikl_date) >= 0
	IF len(tikl_text) > 253 THEN MSGBox "Your TIKL message is too long. A TIKL can be 253 characters and this TIKL is " & len(tikl_text) & " characters."
LOOP WHILE len(tikl_text) > 253


'Splits the case_number(s) into a case_number_array
case_number_array = array(case_number_01, case_number_02, case_number_03, case_number_04, case_number_05, _
						case_number_06, case_number_07, case_number_08, case_number_09, case_number_10, _
						case_number_11, case_number_12, case_number_13, case_number_14, case_number_15, _
						case_number_16, case_number_17, case_number_18, case_number_19, case_number_20, _
						case_number_21, case_number_22, case_number_23, case_number_24, case_number_25, _
						case_number_26, case_number_27, case_number_28, case_number_29, case_number_30, _
						case_number_31, case_number_32, case_number_33, case_number_34, case_number_35, _
						case_number_36, case_number_37, case_number_38, case_number_39, case_number_40, _
						case_number_41, case_number_42, case_number_43, case_number_44, case_number_45, _
						case_number_46, case_number_47, case_number_48, case_number_49, case_number_50, _
						case_number_51, case_number_52, case_number_53, case_number_54, case_number_55, _
						case_number_56, case_number_57, case_number_58, case_number_59, case_number_60)


For each case_number in case_number_array
	If case_number <> "" then
		back_to_self
		call navigate_to_screen("DAIL", "WRIT")
		call create_MAXIS_friendly_date(TIKL_date, 0, 5, 18)
		call write_TIKL_function(tikl_text)
		PF3
	End if
Next

call script_end_procedure("Success! All case numbers TIKLed.")

