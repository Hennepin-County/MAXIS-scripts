'Required for statistical purposes===============================================================================
name_of_script = "DAIL - TPQY RESPONSE.vbs"
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
call changelog_update("05/03/2022", "Updated script functionality to support IEVS message updates. This DAIL scrubber will work on both older message with SSN's and new messages without.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.
'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""   'Connects to BlueZone

'determining if the old message with the SSN functionality will be needed or not.
EMReadScreen MEMB_check, 7, 6, 20
If left(MEMB_check, 4) = "MEMB" then
    member_number = right(MEMB_check, 2)
    SSN_present = False
Else
    SSN_present = True
End if

If SSN_present = True then
    EMSendKey "I"   'Navigates to INFC
    transmit
    Call write_value_and_transmit("SVES", 20, 71)    'Navigates to SVES
    Call write_value_and_transmit("TPQY", 20, 70)    'Navigates to TPQY
Else
    Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
    If is_this_priv = True then script_end_procedure("This is a privileged case. Cannot access. The script will now end.")

    Call write_value_and_transmit(member_number, 20, 76)
    EmReadscreen client_SSN, 11, 7, 42
    client_SSN = replace(client_SSN, " ", "")

    'Going to the SVES Response
    Call navigate_to_MAXIS_screen("INFC", "SVES")
    EmWriteScreen client_SSN, 4, 68
    Call write_value_and_transmit("TPQY", 20, 70)
End if

script_end_procedure("")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/23/2022------------------N/A
'--Tab orders reviewed & confirmed----------------------------------------------05/23/2022------------------N/A
'--Mandatory fields all present & Reviewed--------------------------------------05/23/2022------------------N/A
'--All variables in dialog match mandatory fields-------------------------------05/23/2022------------------N/A
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------05/23/2022------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------05/23/2022------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------05/23/2022------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/23/2022------------------N/A
'--MAXIS_background_check reviewed (if applicable)------------------------------05/23/2022------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------05/23/2022
'--Out-of-County handling reviewed----------------------------------------------05/23/2022------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/23/2022
'--BULK - review output of statistics and run time/count (if applicable)--------05/23/2022------------------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---05/23/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/23/2022
'--Incrementors reviewed (if necessary)-----------------------------------------05/23/2022
'--Denomination reviewed -------------------------------------------------------05/23/2022
'--Script name reviewed---------------------------------------------------------05/23/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------05/23/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------05/23/2022
'--comment Code-----------------------------------------------------------------05/23/2022
'--Update Changelog for release/update------------------------------------------05/23/2022------------------N/A
'--Remove testing message boxes-------------------------------------------------05/23/2022
'--Remove testing code/unnecessary code-----------------------------------------05/23/2022
'--Review/update SharePoint instructions----------------------------------------05/23/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------05/23/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------05/23/2022
'--Complete misc. documentation (if applicable)---------------------------------05/23/2022
'--Update project team/issue contact (if applicable)----------------------------05/23/2022
