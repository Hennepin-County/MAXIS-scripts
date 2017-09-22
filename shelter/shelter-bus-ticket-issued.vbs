'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - BUS TICKET ISSUED.vbs"
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
BeginDialog bus_ticket_dialog, 0, 0, 291, 230, "Bus Ticket Issuances"
  EditBox 60, 5, 60, 15, MAXIS_case_number
  EditBox 225, 5, 50, 15, ticket_amount
  EditBox 120, 30, 155, 15, what_city
  DropListBox 195, 55, 80, 15, "Select one..."+chr(9)+"Central/NE"+chr(9)+"North"+chr(9)+"Northwest"+chr(9)+"South"+chr(9)+"South Suburban"+chr(9)+"West", region_issued
  EditBox 75, 90, 200, 15, staying_with_name
  EditBox 75, 110, 200, 15, staying_with_address
  EditBox 75, 130, 100, 15, staying_with_phone
  EditBox 180, 160, 60, 15, bag_lunches
  DropListBox 180, 180, 60, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", HSM_authorized
  EditBox 50, 205, 125, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 180, 205, 50, 15
    CancelButton 235, 205, 50, 15
  Text 10, 55, 180, 10, "Regional Accountiung Office Where Ticket Was  Issued:"
  Text 45, 95, 25, 10, "Name:"
  Text 10, 35, 105, 10, "Bus ticket destination City/State:"
  Text 35, 115, 35, 10, "Address:"
  Text 45, 185, 125, 10, "EA/ACF Issuance authorized by HSM:"
  Text 15, 135, 50, 10, "Phone Number:"
  GroupBox 5, 75, 280, 80, "Client will be staying with:"
  Text 5, 165, 170, 10, "Number of Bag Lunches Issued for pick up at PSP:"
  Text 5, 210, 40, 10, "Other notes:"
  Text 155, 10, 70, 10, "Bus Ticket Amount: $"
  Text 10, 10, 45, 10, "Case number:"
EndDialog


'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog bus_ticket_dialog
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If ticket_amount = "" then err_msg = err_msg & vbNewLine & "* Enter the Ticket Amount."		
		If what_city = "" then err_msg = err_msg & vbNewLine & "* Enter the City of Destination"
		If region_issued = "Select one..." then err_msg = err_msg & vbNewLine & "* Please select the region where Bus Ticket was issued."
		If staying_with_name = "" then err_msg = err_msg & vbNewLine & "* Enter the name of person client will be staying with"
		If staying_with_address = "" then err_msg = err_msg & vbNewLine & "* Enter the address where the client will be staying."
		If staying_with_phone = "" then err_msg = err_msg & vbNewLine & "* Enter the phone number of the person the client will be staying with."	
		If bag_lunches = "" then err_msg = err_msg & vbNewLine & "* Enter the number of bag lunches issued to the client."
		If HSM_authorized = "Select one..." then err_msg = err_msg & vbNewLine & "* Was EA/ACF approved by HSM?"
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & "(enter NA in all fields that do not apply)" & vbNewLine & err_msg & vbNewLine
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
Call write_variable_in_CASE_NOTE("### All County Funds Request Pending ###")
Call write_bullet_and_variable_in_CASE_NOTE(ticket_amount & " issued for BUS TICKET at " & region_issued & " to", what_city)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE("* Client will be staying with:")
Call write_bullet_and_variable_in_CASE_NOTE("Name", staying_with_name)
Call write_bullet_and_variable_in_CASE_NOTE("Address", staying_with_address)
Call write_bullet_and_variable_in_CASE_NOTE("Phone Number", staying_with_phone)
Call write_variable_in_CASE_NOTE("---")
Call write_bullet_and_variable_in_CASE_NOTE("Number of Bag Lunches issued for pick up at PSP", bag_lunches)
Call write_bullet_and_variable_in_CASE_NOTE("Was EA/ACF approved by HSM?", HSM_authorized)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure("")