'Required for statistical purposes===============================================================================
name_of_script = "DEU - VIEW INFC.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 30                      'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
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
Call changelog_update("10/01/2024", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""    'CONNECTING TO MAXIS
Call check_for_MAXIS(False)
CALL MAXIS_case_number_finder(MAXIS_case_number) 'Finding current case number if present

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 126, 70, "DEU - INFC NAV"
    EditBox 70, 5, 45, 15, MAXIS_case_number
    DropListBox 70, 25, 45, 15, "IEVS"+chr(9)+"PARIS", INFC_process
    ButtonGroup ButtonPressed
      OkButton 20, 45, 45, 15
      CancelButton 70, 45, 45, 15
    Text 20, 10, 45, 10, "Case number:"
    Text 35, 30, 30, 10, "Process:"
EndDialog

DO
	DO
	   err_msg = ""
	    Dialog Dialog1
        Cancel_without_confirmation
	    Call validate_MAXIS_case_number(err_msg, "*")
	    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""						'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
IF is_this_priv = TRUE THEN script_end_procedure("This case is privileged, the script will now end.")

'Determining if there is a single member (and then we can bypass the HH member selection) or to call the HH member array. 
member_count = 0
row = 5
Do 
    EMReadscreen member_present, 2, row, 5
    If trim(member_present) <> "" then 
        member_count = member_count + 1
        row = row + 1
    Else 
        Exit do
    End if 
    If row = 19 then 
        PF8
        row = 5
    End if 
Loop

If member_count > 1 then 
    'ensuring that users have only selected one member.   
    Do   
		'reads the reference number, last name, first name, and then puts it into a single string then into the array
	    EMReadscreen ref_nbr, 3, 4, 33
	    EMReadscreen last_name, 25, 6, 30
	    EMReadscreen first_name, 12, 6, 63
	    EMReadscreen mid_initial, 1, 6, 79
        EMReadScreen client_DOB, 10, 8, 42
	    last_name = trim(replace(last_name, "_", "")) & " "
	    first_name = trim(replace(first_name, "_", "")) & " "
	    mid_initial = replace(mid_initial, "_", "")
	    client_string = ref_nbr & last_name & first_name
	    client_array = client_array & trim(client_string) & "|"
	    transmit
	    Emreadscreen edit_check, 7, 24, 2
    LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

    client_array = TRIM(client_array)
    test_array = split(client_array, "|")
    total_clients = Ubound(test_array)			'setting the upper bound for how many spaces to use from the array
    
    DIM all_client_array()
    ReDim all_clients_array(total_clients, 1)
    
    FOR x = 0 to total_clients				'using a dummy array to build in the autofilled check boxes into the array used for the dialog.
    	Interim_array = split(client_array, "|")
    	all_clients_array(x, 0) = Interim_array(x)
    	all_clients_array(x, 1) = 0    '0 = unchecked
    NEXT
    
    Dialog1 = ""
    BEGINDIALOG Dialog1, 0, 0, 241, (35 + (total_clients * 15)), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
    	Text 10, 5, 105, 10, "Household members to look at:"
    	FOR i = 0 to total_clients											'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
    		IF all_clients_array(i, 0) <> "" THEN checkbox 10, (20 + (i * 15)), 160, 10, all_clients_array(i, 0), all_clients_array(i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
    	NEXT
    	ButtonGroup ButtonPressed
    	OkButton 185, 10, 50, 15
    	CancelButton 185, 30, 50, 15
    ENDDIALOG

    Do
        Do
            err_msg = ""
            Dialog Dialog1       'runs the dialog that has been dynamically created. Streamlined with new functions.
            Cancel_without_confirmation
            'ensuring that users have
            checked_count = 0
            FOR i = 0 to total_clients												'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
                IF all_clients_array(i, 1) = 1 then checked_count = checked_count + 1 'Ignores and blank scanned in persons/strings to avoid a blank checkbox
            NEXT
            If checked_count <> 1 then err_msg = err_msg & vbcr & "* Select only 1 person."
            IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    FOR i = 0 to total_clients
	    IF all_clients_array(i, 1) = 1 THEN member_number = left(all_clients_array (i,0), 2) 						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
    NEXT
Else 
    'If there's only one person then the person selection doesn't show up. 
    member_number = "01" 
End if    

Call write_value_and_transmit(member_number, 20, 76)
EMReadscreen client_SSN, 11, 7, 42          'Reading and cleaning up SSN for INFC
client_SSN = replace(client_SSN, " ", "")

IF INFC_process = "IEVS" then 
    INFC_nav = "IEVP" 
Elseif INFC_process = "PARIS" then 
    INFC_nav = "INTM"
End If     

'navigating to INFC
CALL navigate_to_MAXIS_screen("INFC" , "____")
EmWriteScreen INFC_nav, 20, 71
CALL write_value_and_transmit(client_SSN, 3, 63)

'Checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
EMReadScreen agreement_check, 9, 2, 24
IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")
script_end_procedure("")