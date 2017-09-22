'GATHERING STATS===========================================================================================
name_of_script = "NOTES - VENDOR.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 120
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'script gathers information from user regarding vendor payment(s)
'writes case notes in maxis

BeginDialog vendor_dialog, 0, 0, 431, 240, "Vendor Dialog"
  EditBox 70, 5, 45, 15, case_number
  DropListBox 62, 55, 65, 15, "Vendor Type"+chr(9)+"Mandatory Utility/Shelter"+chr(9)+"Voluntary Utility/Shelter", vendor_type_1
  EditBox 140, 55, 50, 15, vendor_number_1
  EditBox 205, 55, 60, 15, vendor_name_1
  EditBox 280, 55, 50, 15, phone_number_1
  EditBox 345, 55, 60, 15, vendor_amount_1
  DropListBox 60, 85, 65, 15, "Vendor Type"+chr(9)+"Mandatory Utility/Shelter"+chr(9)+"Voluntary Utility/Shelter", vendor_type_2
  EditBox 140, 85, 50, 15, vendor_number_2
  EditBox 205, 85, 60, 15, vendor_name_2
  EditBox 280, 85, 50, 15, phone_number_2
  EditBox 345, 85, 60, 15, vendor_amount_2
  DropListBox 60, 115, 65, 15, "Vendor Type"+chr(9)+"Mandatory Utility/Shelter"+chr(9)+"Voluntary Utility/Shelter", vendor_type_3
  EditBox 205, 115, 60, 15, vendor_name_3
  EditBox 140, 115, 50, 15, vendor_number_3
  EditBox 280, 115, 50, 15, phone_number_3
  EditBox 345, 115, 60, 15, vendor_amount_3
  DropListBox 60, 145, 65, 15, "Vendor Type"+chr(9)+"Mandatory Utility/Shelter"+chr(9)+"Voluntary Utility/Shelter", vendor_type_4
  EditBox 205, 145, 60, 15, vendor_name_4
  EditBox 140, 145, 50, 15, vendor_number_4
  EditBox 280, 145, 50, 15, phone_number_4
  EditBox 345, 145, 60, 15, vendor_amount_4
  EditBox 140, 170, 265, 15, other_information
  EditBox 140, 205, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 295, 205, 50, 15
    CancelButton 355, 205, 50, 15
  Text 150, 35, 45, 10, "Vendor #"
  Text 15, 10, 50, 10, "Case Number"
  Text 75, 35, 45, 15, "Vendor Type"
  Text 290, 35, 45, 15, "Phone #"
  Text 215, 35, 45, 15, "Vendor Name"
  Text 60, 175, 60, 15, "Other Information"
  Text 80, 210, 45, 10, "Signature"
  Text 350, 35, 55, 10, "Vendor Amount"
  Text 15, 55, 40, 15, "Vendor #1"
  Text 15, 85, 40, 15, "Vendor #2"
  Text 15, 115, 40, 15, "Vendor #3"
  Text 15, 145, 40, 15, "Vendor #4"
EndDialog

'The script ----------------------------------------------------------------------------------------------------
EMConnect ""

Call MAXIS_case_number_finder(case_number)

DO
	DO
		err_msg = ""
		dialog vendor_dialog
		cancel_confirmation
		IF len(case_number) > 8 or IsNumeric(case_number) = False THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		IF (vendor_type_1 = "Vendor Type" AND vendor_type_2 = "Vendor Type" AND vendor_type_3 = "Vendor Type" AND vendor_type_4 = "Vendor Type") THEN err_msg = err_msg & vbCr & "*At least one vendor type is needed."
		IF (vendor_type_1 <> "Vendor Type" AND (vendor_number_1 = "" OR vendor_name_1 = "" OR phone_number_1 = "" OR vendor_amount_1 = "")) THEN err_msg = err_msg & vbCr & "*All vendor information must be completed for Vendor Number One."
		IF (vendor_type_2 <> "Vendor Type" AND (vendor_number_2 = "" OR vendor_name_2 = "" OR phone_number_2 = "" OR vendor_amount_2 = "")) THEN err_msg = err_msg & vbCr & "*All vendor information must be completed for Vendor Number Two."
		IF (vendor_type_3 <> "Vendor Type" AND (vendor_number_3 = "" OR vendor_name_3 = "" OR phone_number_3 = "" OR vendor_amount_3 = "")) THEN err_msg = err_msg & vbCr & "*All vendor information must be completed for Vendor Number Three."
		IF (vendor_type_4 <> "Vendor Type" AND (vendor_number_4 = "" OR vendor_name_4 = "" OR phone_number_4 = "" OR vendor_amount_4 = "")) THEN err_msg = err_msg & vbCr & "*All vendor information must be completed for Vendor Number Four."
		IF worker_signature = "" then err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
 Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False
 
back_to_SELF
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43

'The case note---------------------------------------------------------------------------------------
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("---Vendor information---")
IF vendor_type_1 <> "Vendor Type" then 
 Call write_bullet_and_variable_in_CASE_NOTE("Vendor type", vendor_type_1)
 Call write_bullet_and_variable_in_CASE_NOTE("Vendor #", vendor_number_1)
 Call write_bullet_and_variable_in_CASE_NOTE("Vendor name", vendor_name_1)
 Call write_bullet_and_variable_in_CASE_NOTE("Phone number", phone_number_1)
 Call write_bullet_and_variable_in_CASE_NOTE("Vendor amount", vendor_amount_1)
 Call write_variable_in_CASE_NOTE ("-")
END IF 
IF vendor_type_2 <> "Vendor Type" then 
 Call write_bullet_and_variable_in_CASE_NOTE("Vendor type", vendor_type_2)
 Call write_bullet_and_variable_in_CASE_NOTE("Vendor #", vendor_number_2)
 Call write_bullet_and_variable_in_CASE_NOTE("Vendor name", vendor_name_2)
 Call write_bullet_and_variable_in_CASE_NOTE("Phone number", phone_number_2)
 Call write_bullet_and_variable_in_CASE_NOTE("Vendor amount", vendor_amount_2)
 Call write_variable_in_CASE_NOTE ("-")
END IF 
IF vendor_type_3 <> "Vendor Type" then 
 Call write_bullet_and_variable_in_CASE_NOTE("Vendor type", vendor_type_3)
 Call write_bullet_and_variable_in_CASE_NOTE("Vendor #", vendor_number_3)
 Call write_bullet_and_variable_in_CASE_NOTE("Vendor name", vendor_name_3)
 Call write_bullet_and_variable_in_CASE_NOTE("Phone number", phone_number_3)
 Call write_bullet_and_variable_in_CASE_NOTE("Vendor amount", vendor_amount_3)
 Call write_variable_in_CASE_NOTE ("-")
END IF 
IF vendor_type_4 <> "Vendor Type" then 
 Call write_bullet_and_variable_in_CASE_NOTE("Vendor type", vendor_type_4)
 Call write_bullet_and_variable_in_CASE_NOTE("Vendor #", vendor_number_4)
 Call write_bullet_and_variable_in_CASE_NOTE("Vendor name", vendor_name_4)
 Call write_bullet_and_variable_in_CASE_NOTE("Phone number", phone_number_4)
 Call write_bullet_and_variable_in_CASE_NOTE("Vendor amount", vendor_amount_4)
 Call write_variable_in_CASE_NOTE ("-")
END IF 

 Call write_bullet_and_variable_in_CASE_NOTE("Other Information", other_information)
 Call write_variable_in_CASE_NOTE ("---")
 
 Call write_variable_in_CASE_NOTE (worker_signature)

 Call script_end_procedure("")