'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - CHECK EDRS.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 49                	'manual run time in seconds
STATS_denomination = "M"       		'M is for each MEMBER
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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
call changelog_update("06/24/2022", "Updated handling for non-disclosure agreement and closing documentation.", "MiKayla Handley") '#493
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-----------------------------------------------------------------------------------------------------------------
'Connects to BLUEZONE
EMConnect ""
'Makes sure we're in MAXIS
Call check_for_MAXIS(False)

'Grabs the MAXIS case number
CALL MAXIS_case_number_finder(MAXIS_case_number)

'changing footer dates to current month to avoid invalid months.
MAXIS_footer_month = datepart("M", date)
IF Len(MAXIS_footer_month) <> 2 THEN MAXIS_footer_month = "0" & MAXIS_footer_month
MAXIS_footer_year = right(datepart("YYYY", date), 2)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 201, 65, "ELECTRONIC DISQUALIFIED RECIPIENT SYSTEM "
  EditBox 70, 5, 50, 15, MAXIS_case_number
  EditBox 70, 25, 125, 15, worker_signature
  ButtonGroup ButtonPressed
    PushButton 130, 5, 65, 15, "ERDS TE02.08.127", POLI_TEMP_ERDS_button
    OkButton 100, 45, 45, 15
    CancelButton 150, 45, 45, 15
  Text 5, 30, 60, 10, "Worker Signature:"
  Text 5, 10, 50, 10, "Case Number:"
EndDialog

Do
    Do
        err_msg = ""
        DIALOG Dialog1  					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
		IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND len(MAXIS_case_number) > 8) OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = False) THEN err_msg = err_msg & vbCr & "Please enter a valid case number."
		IF ButtonPressed = POLI_TEMP_ERDS_button THEN CALL view_poli_temp("02", "08", "127", "") 'TE02.08.127 ELECTRONIC DISQUALIFIED RECIPIENT SYSTEM
        If err_msg <> "" Then MsgBox "*** Resolve to Continue: " & vbNewLine & err_msg
    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

Dim Member_Info_Array()
Redim Member_Info_Array(UBound(HH_member_array), 4)

'Navigate to stat/memb and check for ERRR message
CALL navigate_to_MAXIS_screen("STAT", "MEMB")
For i = 0 to Ubound(HH_member_array)
	Member_Info_Array(i, 0) = HH_member_array(i)
	EMwritescreen HH_member_array(i), 20, 76 	'Navigating to selected memb panel
	transmit

	EMReadScreen no_MEMB, 13, 8, 22 'If this member does not exist, this will stop the script from continuing.
	IF no_MEMB = "Arrival Date:" THEN script_end_procedure("This HH member does not exist.")

	'Reading info and removing spaces
	EMReadscreen First_name, 12, 6, 63
	First_name = replace(First_name, "_", "")
	Member_Info_Array(i, 1) = First_name

	'Reading Last name and removing spaces
	EMReadscreen Last_name, 25, 6, 30
	Last_name = replace(Last_name, "_", "")
	Member_Info_Array(i, 2) = Last_name

	'Reading Middle initial and replacing _ with a blank if empty.
	EMReadscreen Middle_initial, 1, 6, 79
	Middle_initial = replace(Middle_initial, "_", "")
	Member_Info_Array(i, 3) = Middle_initial

	'Reads SSN
	Emreadscreen SSN_number, 11, 7, 42
	SSN_number = replace(SSN_number, " ", "")
	Member_Info_Array(i, 4) = SSN_number
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
Next

'Navigate back to self and to EDRS
CALL Back_to_self

'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
EMReadScreen agreement_check, 9, 2, 24
IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

CALL navigate_to_MAXIS_screen("INFC", "EDRS")

For i = 0 to UBound(HH_member_array)
	'Write in SSN number into EDRS
	EMwritescreen Member_Info_Array(i, 4), 2, 7
	transmit
	Emreadscreen SSN_output, 7, 24, 2

	'Check to see what results you get from entering the SSN. If you get NO DISQ then check the person's name
	IF SSN_output = "NO DISQ" THEN
		EMWritescreen Member_Info_Array(i, 2), 2, 24
		EMWritescreen Member_Info_Array(i, 1), 2, 58
		EMWritescreen Member_Info_Array(i, 3), 2, 76
		transmit
		EMreadscreen NAME_output, 7, 24, 2
		IF NAME_output = "NO DISQ" THEN        'If after entering a name you still get NO DISQ then let worker know otherwise let them know you found a name.
			Hits = Hits & "No disqualifications found for Member #: " & Member_Info_Array(i, 0) & " " & Member_Info_Array(i, 1) & " " & Member_Info_Array(i, 2) & vbcr
		ELSE
			Hits = Hits & "Member #: " & Member_Info_Array(i, 0) & " " & Member_Info_Array(i, 1) & " " & Member_Info_Array(i, 2) & " has a potential name match. " & vbCr
		END IF
	ELSE
		Hits = Hits & "Member #: " & Member_Info_Array(i, 0) & " " & Member_Info_Array(i, 1) & " " & Member_Info_Array(i, 2) & " has SSN Match. " & vbCr     'If after searching a SSN number you don't get the NO DISQ message then let worker know you found the SSN
	END IF
Next

Msgbox Hits

STATS_counter = STATS_counter - 1			'Removing one instance of the STATS Counter
script_end_procedure_with_error_report("Success your request has been completed please see TE02.08.127 for further processing.")
'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------06/24/2022
'--Tab orders reviewed & confirmed----------------------------------------------06/24/2022
'--Mandatory fields all present & Reviewed--------------------------------------06/24/2022
'--All variables in dialog match mandatory fields-------------------------------06/24/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------06/24/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------06/24/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------06/24/2022
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------06/24/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------06/24/2022
'--PRIV Case handling reviewed -------------------------------------------------06/24/2022
'--Out-of-County handling reviewed----------------------------------------------06/24/2022------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------06/24/2022
'--BULK - review output of statistics and run time/count (if applicable)--------06/24/2022------------------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---06/24/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------06/24/2022------------------N/A
'--Incrementors reviewed (if necessary)-----------------------------------------06/24/2022
'--Denomination reviewed -------------------------------------------------------06/24/2022
'--Script name reviewed---------------------------------------------------------06/24/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------06/24/2022 'QUESTION this one is not bulk but still removes'

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------06/24/2022
'--Comment Code-----------------------------------------------------------------06/24/2022
'--Update Changelog for release/update------------------------------------------06/24/2022
'--Remove testing message boxes-------------------------------------------------06/24/2022
'--Remove testing code/unnecessary code-----------------------------------------06/24/2022
'--Review/update SharePoint instructions----------------------------------------06/24/2022 'TODO'
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------06/24/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------06/24/2022
'--Complete misc. documentation (if applicable)---------------------------------06/24/2022
'--Update project team/issue contact (if applicable)----------------------------06/24/2022
