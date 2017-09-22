'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - 311.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         	'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

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

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog Property_review_dialog, 0, 0, 301, 245, "311"
  EditBox 85, 10, 55, 15, MAXIS_case_number
  EditBox 235, 10, 55, 15, review_date
  DropListBox 195, 30, 95, 15, "Select one..."+chr(9)+"Called 311"+chr(9)+"Called city"+chr(9)+"Checked website", property_reviewed
  EditBox 85, 50, 205, 15, property_address
  EditBox 85, 75, 205, 15, open_work_orders
  EditBox 85, 100, 205, 15, rental_license
  EditBox 85, 125, 205, 15, rep_name
  EditBox 85, 150, 205, 15, violations
  DropListBox 85, 175, 90, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Inspection pending", passed_inspection
  EditBox 225, 175, 65, 15, vendor_number
  EditBox 85, 200, 205, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 185, 225, 50, 15
    CancelButton 240, 225, 50, 15
  Text 10, 105, 75, 10, "Current rental license:"
  Text 15, 180, 70, 10, "Passed inspection?:"
  Text 150, 15, 80, 10, "Date of property review:"
  Text 20, 55, 60, 10, "Property address:"
  Text 35, 205, 45, 10, "Other notes: "
  Text 10, 130, 75, 10, "Representative name:"
  Text 45, 155, 35, 10, "Violations:"
  Text 35, 15, 45, 10, "Case number:"
  Text 35, 35, 145, 10, "What was the source of the property review:"
  Text 190, 180, 35, 10, "Vendor #:"
  Text 20, 80, 60, 10, "Open work orders:"
EndDialog


'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'autofilling the review_date variable with the current date
review_date = date & ""

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog Property_review_dialog
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If Isdate(review_date) = False then err_msg = err_msg & vbNewLine & "* Please enter the date the property was reviewed."
		If property_reviewed = "Select one..." then err_msg = err_msg & vbNewLine & "* Please select the source of the property review. Called 311 or web-site."
		If property_address = "" then err_msg = err_msg & vbNewLine & "* Enter the property address."
		If open_work_orders = "" then err_msg = err_msg & vbNewLine & "* Enter work order status/information."
		If rental_license = "" then err_msg = err_msg & vbNewLine & "* Enter information about the current rental license."
		If rep_name = "" then err_msg = err_msg & vbNewLine & "* Enter the representative's name."
		If violations = "" then err_msg = err_msg & vbNewLine & "* Enter the household's net income."
		If passed_inspection = "Select one..." then err_msg = err_msg & vbNewLine & "* Has the property passed the inspection?"
		If vendor_number = "" then err_msg = err_msg & vbNewLine & "* Enter the property's vendor #."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					
		
'adding the case number 
back_to_self
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

'The case note'
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("### " & property_reviewed & " on " & review_date & " ###"   )
Call write_bullet_and_variable_in_CASE_NOTE("Property address", property_address)
Call write_bullet_and_variable_in_CASE_NOTE("Open work orders", open_work_orders)
Call write_bullet_and_variable_in_CASE_NOTE("Current rental license", rental_license)
Call write_bullet_and_variable_in_CASE_NOTE("Representative name", rep_name)
Call write_bullet_and_variable_in_CASE_NOTE("Violations", Violations)
Call write_bullet_and_variable_in_CASE_NOTE("Passed inspection", passed_inspection)
Call write_bullet_and_variable_in_CASE_NOTE("Vendor #", vendor_number)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")