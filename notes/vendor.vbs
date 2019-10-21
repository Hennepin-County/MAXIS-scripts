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
call changelog_update("10/21/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'script gathers information from user regarding vendor payment(s)
'writes case notes in maxis

BeginDialog vendor_dialog, 0, 0, 396, 190, "Vendor Dialog"
  EditBox 60, 5, 55, 15, MAXIS_case_number
  DropListBox 60, 45, 65, 15, "Vendor Type"+chr(9)+"Mandatory Utility/Shelter"+chr(9)+"Voluntary Utility/Shelter", vendor_type_1
  EditBox 140, 45, 50, 15, vendor_number_1
  EditBox 205, 45, 60, 15, vendor_name_1
  EditBox 280, 45, 50, 15, phone_number_1
  EditBox 345, 45, 45, 15, vendor_amount_1
  DropListBox 60, 70, 65, 15, "Vendor Type"+chr(9)+"Mandatory Utility/Shelter"+chr(9)+"Voluntary Utility/Shelter", vendor_type_2
  EditBox 140, 70, 50, 15, vendor_number_2
  EditBox 205, 70, 60, 15, vendor_name_2
  EditBox 280, 70, 50, 15, phone_number_2
  EditBox 345, 70, 45, 15, vendor_amount_2
  DropListBox 60, 95, 65, 15, "Vendor Type"+chr(9)+"Mandatory Utility/Shelter"+chr(9)+"Voluntary Utility/Shelter", vendor_type_3
  EditBox 205, 95, 60, 15, vendor_name_3
  EditBox 140, 95, 50, 15, vendor_number_3
  EditBox 280, 95, 50, 15, phone_number_3
  EditBox 345, 95, 45, 15, vendor_amount_3
  DropListBox 60, 120, 65, 15, "Vendor Type"+chr(9)+"Mandatory Utility/Shelter"+chr(9)+"Voluntary Utility/Shelter", vendor_type_4
  EditBox 205, 120, 60, 15, vendor_name_4
  EditBox 140, 120, 50, 15, vendor_number_4
  EditBox 280, 120, 50, 15, phone_number_4
  EditBox 345, 120, 45, 15, vendor_amount_4
  EditBox 75, 145, 315, 15, other_information
  EditBox 75, 165, 125, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 285, 165, 50, 15
    CancelButton 340, 165, 50, 15
  Text 145, 30, 45, 10, "Vendor #"
  Text 5, 10, 50, 10, "Case Number:"
  Text 70, 30, 45, 10, "Vendor Type"
  Text 290, 30, 45, 10, "Phone #"
  Text 215, 30, 45, 10, "Vendor Name"
  Text 10, 150, 60, 10, "Other Information:"
  Text 10, 170, 60, 10, "Worker signature:"
  Text 345, 30, 55, 10, "Vendor Amt $"
  Text 15, 50, 40, 10, "Vendor #1:"
  Text 15, 75, 40, 10, "Vendor #2:"
  Text 15, 100, 40, 10, "Vendor #3:"
  Text 15, 125, 40, 10, "Vendor #4:"
EndDialog

'The script ----------------------------------------------------------------------------------------------------
EMConnect ""

Call MAXIS_case_number_finder(MAXIS_case_number)

DO
	DO
		err_msg = ""
		dialog vendor_dialog
		cancel_confirmation
		IF len(MAXIS_case_number) > 8 or IsNumeric(MAXIS_case_number) = False THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
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
