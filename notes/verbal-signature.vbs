'GATHERING STATS===========================================================================================
name_of_script = "NOTES - Verbal Signature.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 120
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================
'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")

call changelog_update("09/17/2026", "Initial version.", "David Courtright, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call check_for_MAXIS(False)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 146, 75, "Case number dialog"
  EditBox 85, 10, 50, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 30, 50, 50, 15
    CancelButton 85, 50, 50, 15
  Text 30, 15, 45, 10, "Case number:"
EndDialog

Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		If IsNumeric(MAXIS_case_number) = False or Len(MAXIS_case_number) > 8 Then err_msg = err_msg & vbCr & "* Enter a valid case number."
		If err_msg <> "" Then MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False



Call Generate_Client_List(HH_Memb_DropDown, "Select One:")


Dialog1 = ""
BeginDialog Dialog1, 0, 0, 271, 245, "Verbal Signature Record"
  Text 10, 35, 100, 10, "Verbal Signature Accepted for:"
  ComboBox 120, 30, 130, 15, "", member_list, signature_memb
  Text 20, 50, 190, 20, "To record a verbal signature the date, time and resident phone number needs to be recorded. "
  Text 20, 75, 105, 10, "Signature was accepted at:"
  Text 25, 95, 20, 10, "Date: "
  EditBox 50, 90, 50, 15, verbal_sig_date
  Text 25, 115, 20, 10, "Time: "
  EditBox 50, 110, 50, 15, verbal_sig_time
  Text 20, 140, 85, 10, "Resident Phone Number:"
  DropListBox 110, 135, 95, 45, "phone_droplist", verbal_sig_phone_number
  Text 5, 195, 255, 20, "Remember to send the resident a copy of the form they verbally signed and provide instructions for making corrections. "
  DropListBox 150, 155, 30, 15, "Yes"+chr(9)+"No", minor_indicator
  Text 20, 160, 95, 10, "Minor children in SNAP unit?"
  Text 20, 175, 125, 10, "Elderly / Disabled members in unit?"
  DropListBox 150, 170, 30, 15, "Yes"+chr(9)+"No", elderly_indicator
  ButtonGroup ButtonPressed
    OkButton 210, 220, 50, 15
  Text 5, 5, 255, 20, "* ** Do not use this script if the verbal signature information was previously recorded using CSR or Interview Script ***"

EndDialog

Do
  err_msg = ""
  dialog Dialog1
  cancel_without_confirmation
  If IsDate(verbal_sig_date) = False Then err_msg = err_msg & vbCr & "* Enter the date you accepted the verbal signature."
  If IsDate(verbal_sig_time) = True Then
    verbal_sig_time = FormatDateTime(verbal_sig_time, 3)
    If InStr(verbal_sig_time, ":") = 0 Then err_msg = err_msg & vbCr & "* The time information does not appear to be a valid time, review and update."
    verbal_sig_time = replace(verbal_sig_time, ":00 ", " ")
  Else
    err_msg = err_msg & vbCr & "* The time information does not appear to be a valid time, review and update."
  End If
  If verbal_sig_phone_number = "" or verbal_sig_phone_number = "Select or Type" Then err_msg = err_msg & vbCr & "* Phone number detail is required."
  If minor_indicator = "" Then err_msg = err_msg & vbCr & "* Please indicate if there are minor children in the SNAP unit."
  If elderly_indicator = "" Then err_msg = err_msg & vbCr & "* Please indicate if there are elderly or disabled members in the SNAP unit."
  If err_msg <> "" Then MsgBox "*****     NOTICE     *****" & vbCr & "Please resolve to continue:" & vbCr & err_msg
Loop until err_msg = ""

Call start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("* * Verbal Signature Accepted:")
CALL write_variable_in_CASE_NOTE("    - MEMB " & signature_memb)
CALL write_variable_in_CASE_NOTE("    Signature accepted on " & verbal_sig_date & " at " & verbal_sig_time & ".")
CALL write_variable_in_CASE_NOTE("    Resident Phone Number: " & verbal_sig_phone_number)
CALL write_variable_in_CASE_NOTE("    Minor children in SNAP unit: " & minor_indicator)
CALL write_variable_in_CASE_NOTE("    Elderly / Disabled members in unit: " & elderly_indicator)

end_msg = "Verbal signature entered in case/note. Verbal signature accepted on " & verbal_sig_date & " at " & verbal_sig_time & " From: " & signature_memb & " (Minors: " & minor_indicator & ", Elderly/Disabled: " & elderly_indicator & ")"