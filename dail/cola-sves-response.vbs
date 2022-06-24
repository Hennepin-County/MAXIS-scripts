'Required for statistical purposes===============================================================================
name_of_script = "DAIL - COLA SVES RESPONSE.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 30          'manual run time in seconds
STATS_denomination = "C"       'C is for case
'END OF stats block==============================================================================================

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
call changelog_update("06/21/2022", "Updated handling for non-disclosure agreement and closing documentation.", "MiKayla Handley") '#493
call changelog_update("05/23/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.
'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

EMReadScreen name_for_dail, 57, 5, 5			'Reading the name of the client
'This next block will determine the name of the client the message is for
'If the message is for someone other than M01 - the name is writen next to the name of M01
other_person = InStr(name_for_dail, "--(")	'This determines if it for someone other than M01
'This is for if the message is for M01'
If other_person = 0 Then
	comma_loc = InStr(name_for_dail, ",")  	'Determines the end of the last name
    dash_loc = InStr(name_for_dail, "-")	'Determines the end of the name
    EMReadscreen last_name, comma_loc - 1, 5, 5									'Reading the last name
	EMReadscreen middle_exists, 1, 5, 5 + (dash_loc - 3)						'Determines if clt's middle initial is listed
    If middle_exists = " " Then 												'If not - reads first name
        EMReadscreen first_name, dash_loc - comma_loc - 3, 5, comma_loc + 5
	Else 																		'If so - reads first name
        EMReadScreen first_name, dash_loc - comma_loc - 1, 5, comma_loc + 5
	End If
'This is for if the message is for a different HH Member
Else
	end_other = InStr(name_for_dail, ")--")
	comma_loc = InStr(other_person, name_for_dail, ",")
	EMReadscreen last_name, comma_loc - other_person - 3, 5, other_person + 7
	EMReadscreen middle_exists, 1, 5, end_other + 2
	If middle_exists = " " Then
		EMReadscreen first_name, end_other - comma_loc - 3, 5, comma_loc + 5
	Else
		EMReadScreen first_name, end_other - comma_loc - 1, 5, comma_loc + 5
	End If
End If

client_name = trim(last_name) & " " & trim(first_name)		'putting the name into one string

'Finding the client SSN for SVES navigation
CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
IF is_this_priv = TRUE THEN script_end_procedure("This case is privileged, the script will now end.")
Do
    EmReadscreen memb_last_name, 25, 6, 30
    memb_last_name = replace(memb_last_name, "_", "")
    EmReadscreen memb_first_name, 12, 6, 63
    memb_first_name = replace(memb_first_name, "_", "")
    memb_client_name = trim(memb_last_name) & " " & trim(memb_first_name)
    If memb_client_name = client_name then
        EmReadscreen client_SSN, 11, 7, 42
        client_SSN = replace(client_SSN, " ", "")
        Exit do
    Else
        transmit
    End if
    EMReadScreen MEMB_error, 5, 24, 2
    If MEMB_error = "ENTER" then script_end_procedure("Unable to find client name in the household. The script will now end.")
Loop

'Going to the SVES Response
Call navigate_to_MAXIS_screen("INFC", "SVES")
EmWriteScreen client_SSN, 4, 68
Call write_value_and_transmit("TPQY", 20, 70)
'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
EMReadScreen agreement_check, 9, 2, 24
IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

script_end_procedure_with_error_report("Success, the script has navigated you to TPQY for: " & first_name & " " & last_name)
'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------06/21/2022
'--Tab orders reviewed & confirmed----------------------------------------------06/21/2022
'--Mandatory fields all present & Reviewed--------------------------------------06/21/2022
'--All variables in dialog match mandatory fields-------------------------------06/21/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------06/21/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------06/21/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------06/21/2022
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------06/21/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------06/21/2022------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------06/21/2022
'--Out-of-County handling reviewed----------------------------------------------06/21/2022------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------06/21/2022
'--BULK - review output of statistics and run time/count (if applicable)--------06/21/2022------------------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---06/21/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------06/21/2022------------------N/A
'--Incrementors reviewed (if necessary)-----------------------------------------06/21/2022------------------N/A
'--Denomination reviewed -------------------------------------------------------06/21/2022
'--Script name reviewed---------------------------------------------------------06/21/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------06/21/2022------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------06/21/2022
'--comment Code-----------------------------------------------------------------06/21/2022
'--Update Changelog for release/update------------------------------------------06/21/2022
'--Remove testing message boxes-------------------------------------------------06/21/2022
'--Remove testing code/unnecessary code-----------------------------------------06/21/2022
'--Review/update SharePoint instructions----------------------------------------06/21/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------06/21/2022 'TODO'
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------06/21/2022
'--Complete misc. documentation (if applicable)---------------------------------06/21/2022
'--Update project team/issue contact (if applicable)----------------------------06/21/2022
