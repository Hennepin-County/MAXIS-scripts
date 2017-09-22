'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SHERIFF FORECLOSURE.vbs"
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
BeginDialog sheriff_forclosure, 0, 0, 286, 175, "Sheriff forclosure"
  EditBox 55, 10, 55, 15, MAXIS_case_number
  EditBox 210, 10, 70, 15, date_checked
  EditBox 70, 35, 210, 15, property_address
  EditBox 60, 60, 100, 15, owner_name
  EditBox 225, 60, 55, 15, foreclosure_date
  EditBox 70, 85, 100, 15, occupant_name
  EditBox 95, 110, 185, 15, occupants_whereabouts
  EditBox 50, 135, 230, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 175, 155, 50, 15
    CancelButton 230, 155, 50, 15
  Text 5, 65, 55, 10, "Owner(s) name:"
  Text 5, 90, 60, 10, "Occupant(s) name:"
  Text 125, 15, 80, 10, "Date of property review:"
  Text 5, 40, 60, 10, "Property address:"
  Text 5, 140, 40, 10, "Other notes: "
  Text 170, 65, 55, 10, "Forclosure date:"
  Text 5, 115, 85, 10, "Occupant(s) whereabouts:"
  Text 5, 15, 45, 10, "Case number:"
EndDialog


'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'autofilling the review_date variable with the current date
date_checked = date & ""

'Running the initial dialog
'commented out the foreclosure_date test at reqwust of hennepin shelter Team'
DO
	DO
		err_msg = ""
		Dialog sheriff_forclosure
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If IsDate(date_checked) = False then err_msg = err_msg & vbNewLine & "* Enter the property review date."      
		If property_address = "" then err_msg = err_msg & vbNewLine & "* Enter the property address."
		If owner_name = "" then err_msg = err_msg & vbNewLine & "* Enter the property owner's name."
		'If IsDate(foreclosure_date) = False then err_msg = err_msg & vbNewLine & "* Enter the property's forclosure date."      
		If occupant_name = "" then err_msg = err_msg & vbNewLine & "* Enter the occupant's name."
		If occupants_whereabouts = "" then err_msg = err_msg & vbNewLine & "* Enter the occupant's current whereabouts."
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
Call write_variable_in_CASE_NOTE("### Sheriff foreclosure website checked on: " & date_checked & " ###"   )
Call write_bullet_and_variable_in_CASE_NOTE("Property address", property_address)
Call write_bullet_and_variable_in_CASE_NOTE("Owner(s) name", owner_name)
Call write_bullet_and_variable_in_CASE_NOTE("Foreclosure date", foreclosure_date)
Call write_bullet_and_variable_in_CASE_NOTE("Representative name", rep_name)
Call write_bullet_and_variable_in_CASE_NOTE("Occupant(s) name", occupant_name)
Call write_bullet_and_variable_in_CASE_NOTE("Occupant(s) current whereabouts", occupants_whereabouts)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team") 

script_end_procedure("")