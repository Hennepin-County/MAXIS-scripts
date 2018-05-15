'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - EGA APPROVERS.vbs"
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
BeginDialog ega_approval_dialog, 0, 0, 296, 220, "EGA Approval"
  EditBox 60, 10, 55, 15, MAXIS_case_number
  EditBox 200, 10, 90, 15, clients_last_name
  CheckBox 200, 30, 85, 10, "EGA Is For ADS Case", ADS_checkbox
  EditBox 45, 55, 145, 15, ega_paid_to
  EditBox 225, 55, 65, 15, fax_number
  EditBox 45, 80, 60, 15, vendor_number
  EditBox 225, 80, 65, 15, utility_number
  EditBox 65, 120, 40, 15, ega_payment
  DropListBox 230, 120, 60, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", ega_wksheet_completed_droplist
  DropListBox 65, 145, 60, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Does not apply", pins_verified_droplist
  DropListBox 230, 145, 60, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", guarantee_needed_droplist
  EditBox 65, 170, 225, 15, other_notes
  EditBox 65, 195, 115, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 185, 195, 50, 15
    CancelButton 240, 195, 50, 15
  Text 10, 15, 45, 10, "Case number:"
  Text 15, 60, 25, 10, "Pay To:"
  Text 10, 85, 35, 10, "Vendor #:"
  Text 10, 95, 165, 10, "(If no Vendor #, use Federal ID or SSN of Vendor)"
  Text 135, 15, 65, 10, "Client's last name:"
  Text 165, 85, 60, 10, "Utility Account #:"
  Text 5, 200, 60, 10, "Worker signature: "
  Text 150, 150, 70, 10, "Guarantee Needed?:"
  Text 20, 175, 40, 10, "Other notes: "
  Text 10, 150, 50, 10, "PINS Verified?:"
  Text 10, 125, 55, 10, "EGA Payment $:"
  Text 115, 125, 110, 10, "Was EGA worksheet completed?:"
  Text 40, 50, 10, 0, "-85"
  Text 200, 60, 20, 10, "Fax #:"
  GroupBox 5, 40, 290, 70, "Vendor information"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
DO	
	DO		
		err_msg = ""
		Dialog ega_approval_dialog
        cancel_confirmation
		IF len(MAXIS_case_number) > 8 or IsNumeric(MAXIS_case_number) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF clients_last_name = "" then err_msg = err_msg & vbNewLine & "* Client's last name."
		IF ega_payment = "" then err_msg = err_msg & vbNewLine & "* Enter EGA Payment Amount."
		If ega_paid_to = "" then err_msg = err_msg & vbNewLine & "* Enter Who EGA Is Being Paid To."	
		If IsNumeric(vendor_number) = False then err_msg = err_msg & vbNewLine & "* Enter EGA Vendor Number."
		If IsNumeric(utility_number) = False then err_msg = err_msg & vbNewLine & "* Enter Utility Number."
		IF guarantee_needed_droplist = "Select one..." then err_msg = err_msg & vbNewLine & "* Please select if Guarantee of payment was made."
		IF pins_verified_droplist = "Select one..." then err_msg = err_msg & vbNewLine & "* Please select if PINS was verified, or select Does Not Apply."
		IF ega_wksheet_completed_droplist = "Select one..." then err_msg = err_msg & vbNewLine & "* Please select whether EGA Worksheet was completed."	 
		If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
		IF err_msg <> "" then MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
 Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

back_to_SELF
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

'Function that sends an email to Kathryn Fitgerald once the dialog is complete, and MAXIS has been udpated. 
' Here are the parameters for the function: Call create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)

'The case note---------------------------------------------------------------------------------------
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("### MX" & MAXIS_case_number & "/" & clients_last_name & "/ EGA (urgent) ###")
If ADS_checkbox = 1 then call write_variable_in_CASE_NOTE("* EGA is for ADS Case")
Call write_bullet_and_variable_in_CASE_NOTE("EGA Amount Paid", ega_payment)
Call write_bullet_and_variable_in_CASE_NOTE("Pay to", ega_paid_to)
Call write_bullet_and_variable_in_CASE_NOTE("Vendor# (or FedID/SSN, if no Vendor#)", vendor_number)
Call write_bullet_and_variable_in_CASE_NOTE("Utility Account#", utility_number)
Call write_bullet_and_variable_in_CASE_NOTE("Guarantee Needed?", guarantee_needed)
Call write_bullet_and_variable_in_CASE_NOTE("FAX#", fax_number)
Call write_bullet_and_variable_in_CASE_NOTE("PINS Verified?", pins_verified)
Call write_bullet_and_variable_in_CASE_NOTE("Was EGA worksheet completed?", ega_wksheet_completed)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")