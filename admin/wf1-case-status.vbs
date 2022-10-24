'Required for statistical purposes===============================================================================
name_of_script = "BULK - WF1 CASE STATUS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 120                     'manual run time in seconds
STATS_denomination = "C"       'C is for each Case
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
call changelog_update("10/20/2022", "Script updated to support household's with multiple members, fix out of county handling, add person based SNAP status and enhance background functionality.", "Ilse Ferris, Hennepin County")
call changelog_update("08/16/2018", "Removed default to current month for case status. Users can navigate the footer month/year they wish to review, then run the script.", "Ilse Ferris, Hennepin County")
call changelog_update("11/13/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect "" 'Connects to BlueZone
Call Check_for_MAXIS(False)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 110, "WF1 Case Status"
    ButtonGroup ButtonPressed
    PushButton 200, 45, 50, 15, "Browse...", select_a_file_button
    OkButton 145, 90, 50, 15
    CancelButton 200, 90, 50, 15
    EditBox 15, 45, 180, 15, file_selection_path
    GroupBox 10, 5, 250, 80, "Using the WF1M Case Status script"
    Text 20, 20, 235, 20, "This script should be used when E and T provides you with a list of recipeints that require a status update."
    Text 15, 65, 230, 15, "Select the Excel file that contains the WF1 information by selecting the 'Browse' button, and finding the file."
EndDialog

Do
    Do
        err_msg = ""
        dialog Dialog1
        cancel_without_confirmation
        If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select a file to continue."
        If err_msg <> "" Then MsgBox err_msg
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file

'ARRAY business----------------------------------------------------------------------------------------------------
'Sets up the array to store all the information for each client'
Dim CBO_array ()
ReDim CBO_array (error_reason_const, 0)

'Sets constants for the array to make the script easier to read (and easier to code)'
Const last_name_const         = 0
Const first_name_const        = 1
Const MAXIS_case_number_const = 2
Const client_SSN_const        = 3
Const memb_number_const		  = 4
Const snap_status_const       = 5
Const excel_num_const		  = 6
Const ABAWD_status_const	  = 7
Const error_reason_const      = 8

'Now the script adds all the clients on the excel list into an array for the appropriate county
excel_row = 2 'starting row
entry_record = 0

Do                                                            'Loops until there are no more cases in the Excel list
    last_name = UCASE(trim(objExcel.cells(excel_row, 1).Value)) 'uses client last name since either case number or SSN can be provided
    first_name = UCASE(trim(objExcel.cells(excel_row, 2).Value)) 'uses client last name since either case number or SSN can be provided
	MAXIS_case_number = trim(objExcel.cells(excel_row, 3).Value)
    If MAXIS_case_number = "" then exit do

	client_SSN  = trim(objExcel.cells(excel_row, 4).Value)		'Pulls the SSN and reformats if 9 digits.
    If client_SSN <> "" then
	    If len(client_SSN) = 9 then
           ssn_first = left(client_SSN, 3)
           ssn_mid = right(left(client_SSN, 5), 2)
           ssn_end = right(client_SSN, 4)
           client_SSN = ssn_first & "-" & ssn_mid & "-" & ssn_end
        End if
    End if

	'Adding client information to the array
	ReDim Preserve CBO_array(error_reason_const, entry_record)	'This resizes the array based on if the client is in the selected county
    CBO_array(last_name_const,         entry_record) = last_name
    CBO_array(first_name_const,        entry_record) = first_name
    CBO_array(client_SSN_const,        entry_record) = client_SSN		'The client information is added to the array
	CBO_array(MAXIS_case_number_const, entry_record) = MAXIS_case_number
	CBO_array(excel_num_const, 		   entry_record) = excel_row
	entry_record = entry_record + 1			'This increments to the next entry in the array
    STATS_counter = STATS_counter + 1
	excel_row = excel_row + 1

	'blanking out variables
	client_SSN = ""
	MAXIS_case_number = ""
Loop

If entry_record = 0 then script_end_procedure_with_error_report("No cases have been found on this list. The script wil now end.")

'Gathering info from MAXIS, and making the referrals and case notes if cases are found and active----------------------------------------------------------------------------------------------------
For item = 0 to UBound(CBO_array, 2)
	MAXIS_case_number = CBO_array(MAXIS_case_number_const, item)
    member_found = False    'defaulting to not found/false

    Call navigate_to_MAXIS_screen_review_PRIV("CASE", "PERS", is_this_priv)
    If is_this_priv = True then
        CBO_array(error_reason_const, item) = "PRIV case."
    Else
        EmReadscreen county_code, 4, 20, 14
        If county_code <> UCASE(worker_county_code) then
            CBO_array(error_reason_const, item) = "Out-of-county case. County code is: " & county_code
        Else
            row = 10 '1st row/person in CASE/PERS
            Do
                If CBO_array(client_SSN_const, item) = "" then
                    'Read and connect name to member number
                    EmReadscreen pers_last_name, 15, row, 6
                    EmReadscreen pers_first_name, 11, row, 22
                    pers_last_name = trim(pers_last_name)
                    pers_first_name = trim(pers_first_name)
                    If pers_last_name = "" then exit do
                    'if the name is a match then exiting do (will read person info later)
                    If pers_last_name = CBO_array(last_name_const, item) and pers_first_name = CBO_array(first_name_const, item) then
                        member_found = True
                        Exit do
                    Elseif left(pers_last_name, 4) = left(CBO_array(last_name_const, item), 4) and left(pers_first_name, 4) = left(CBO_array(first_name_const, item), 4) then
                        'if partial name match based on 1st 4 of 1st and last name: because names are long and get cut off.
                        worker_confirm = msgbox("Is this the member you are looking for? " & vbcr & vbcr & pers_first_name & " " & pers_last_name, vbQuestion + vbYesNo, "Confirm WF1 Member")
                        If vbYes then
                            member_found = True
                            Exit do
                        End if
                    End if
                Else
                    'using SSN to connect to member number
                    EmReadscreen pers_SSN, 11, row + 1, 6
                    If pers_SSN = CBO_array(client_SSN_const, item) then
                        member_found = True
                        Exit do
                    End if
                End if
                row = row + 3			'information is 3 rows apart. Will read for the next member.

                If row = 19 then
                    PF8
                    row = 10					'changes MAXIS row if more than one page exists
                END if
                EMReadScreen last_PERS_page, 21, 24, 2
            LOOP until last_PERS_page = "THIS IS THE LAST PAGE"

            'Reading the person information for the multiple member_found = true scenarios above
            If member_found = true then
                EmReadscreen memb_number, 2, row, 3
                CBO_array(memb_number_const, item) = memb_number
                EmReadscreen FS_status, 1, row, 54
                If FS_status = "A" then CBO_array(snap_status_const, item) = "Active"
                If FS_status = "D" then CBO_array(snap_status_const, item) = "Denied"
                If FS_status = "I" then CBO_array(snap_status_const, item) = "Inactive"
                If FS_status = "P" then CBO_array(snap_status_const, item) = "Pending"
                If FS_status = "R" then CBO_array(snap_status_const, item) = "Reinstatement"

                Call navigate_to_MAXIS_screen("STAT", "WREG")
                Call write_value_and_transmit(CBO_array(memb_number_const, item), 20, 76)
                EMReadScreen fset_code, 2, 8, 50
                EMReadScreen abawd_code, 2, 13, 50
                WREG_codes = fset_code & "-" & abawd_code
                If WREG_codes = "30-11" then
                    CBO_array(ABAWD_status_const, item) = "Volunatary"
                Elseif WREG_codes = "30-10" then
                    CBO_array(ABAWD_status_const, item) = "Mandatory - ABAWD"
                ElseIf WREG_codes = "30-13" then
                    CBO_array(ABAWD_status_const, item) = "Mandatory - Banked Months"
                Else
                    CBO_array(ABAWD_status_const, item) = "Exempt"
                End if
            Else
                If CBO_array(memb_number_const, item) = "" then CBO_array(error_reason_const, item) = "Unable to find MEMB in CASE/PERS"
            End if
        End if
    End if
    'Excel Output
    objExcel.cells(CBO_array(excel_num_const, item), 5).Value = CBO_array(snap_status_const,  item)
	objExcel.cells(CBO_array(excel_num_const, item), 6).Value = CBO_array(ABAWD_status_const, item)
	objExcel.cells(CBO_array(excel_num_const, item), 7).Value = CBO_array(error_reason_const, item)
Next

'Formatting the column width.
FOR i = 1 to 7
	objExcel.Columns(i).AutoFit()
NEXT

STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)
script_end_procedure_with_error_report("Success! Review the spreadsheet for accuracy.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------10/24/2022
'--Tab orders reviewed & confirmed----------------------------------------------10/24/2022
'--Mandatory fields all present & Reviewed--------------------------------------10/24/2022
'--All variables in dialog match mandatory fields-------------------------------10/24/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------10/24/2022-----------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------10/24/2022-----------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------10/24/2022-----------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used-10/24/2022-----------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------10/24/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------10/24/2022-----------------N/A
'--PRIV Case handling reviewed -------------------------------------------------10/24/2022
'--Out-of-County handling reviewed----------------------------------------------10/24/2022
'--script_end_procedures (w/ or w/o error messaging)----------------------------10/24/2022
'--BULK - review output of statistics and run time/count (if applicable)--------10/24/2022
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---10/24/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------10/24/2022
'--Incrementors reviewed (if necessary)-----------------------------------------10/24/2022
'--Denomination reviewed -------------------------------------------------------10/24/2022
'--Script name reviewed---------------------------------------------------------10/24/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------10/24/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------10/24/2022
'--comment Code-----------------------------------------------------------------10/24/2022
'--Update Changelog for release/update------------------------------------------10/24/2022
'--Remove testing message boxes-------------------------------------------------10/24/2022
'--Remove testing code/unnecessary code-----------------------------------------10/24/2022
'--Review/update SharePoint instructions----------------------------------------10/24/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------10/24/2022-----------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------10/24/2022
'--Complete misc. documentation (if applicable)---------------------------------10/24/2022
'--Update project team/issue contact (if applicable)----------------------------10/24/2022
