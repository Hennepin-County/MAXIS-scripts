'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - TASK-BASED DAIL CAPTURE.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = 30
STATS_denomination = "C"       			'C is for each CASE
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
call changelog_update("07/21/2023", "Updated function that sends an email through Outlook", "Mark Riegel, Hennepin County")
call changelog_update("10/05/2022", "Ensured correct baskets are pulled for each population. Changed population verbiage (ADAD to adults, etc.)", "Ilse Ferris, Hennepin County")
call changelog_update("07/25/2022", "Removed Laurie's email, added Mary McGuinness.", "Ilse Ferris, Hennepin County")
call changelog_update("07/02/2022", "Updated string for checking to see if no DAILs exist based on options selected (all vs. specified DAIL's).", "Ilse Ferris, Hennepin County")
call changelog_update("04/11/2022", "Added additional handling for moving to another case load if no DAILs are present.", "Ilse Ferris, Hennepin County")
call changelog_update("12/18/2021", "Updated new server name.", "Ilse Ferris, Hennepin County")
call changelog_update("04/26/2021", "Removed emailing Todd Bennington per request.", "Ilse Ferris, Hennepin County")
call changelog_update("02/17/2021", "Added defaults for DAIL type selctions based on if before or on/after ten day cut off. DAIL types selected on/after ten day cut off are only TIKL messages.", "Ilse Ferris, Hennepin County")
call changelog_update("02/02/2021", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""
Call check_for_MAXIS(False)
'Defaulting these populations to checked per current process.
families_checkbox = 1
adults_checkbox = 1

'Defaulting autochecks based on the ten day cut off schedule. On ten day and after, only TIKL's are pulled.
If DateDiff("d", date, ten_day_cutoff_date) > 0 then
    'Defaulting these messages to checked as these are the most assigned cases.
    cola_check = 1
    cses_check = 1
    info_check = 1
    pepr_check = 1
    tikl_check = 1
Else
    tikl_check = 1
End if

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 251, 260, "Task-Based DAIL Capture Main Dialog"
  CheckBox 10, 100, 85, 10, "All Population Baskets", all_baskets_checkbox
  CheckBox 100, 100, 30, 10, "Adults", adults_checkbox
  CheckBox 135, 100, 40, 10, "Families", families_checkbox
  CheckBox 180, 100, 30, 10, "LTC+", LTC_checkbox
  CheckBox 10, 115, 145, 10, "OR check here to process for all workers.", all_workers_check
  CheckBox 10, 150, 25, 10, "ALL", all_check
  CheckBox 40, 150, 30, 10, "COLA", cola_check
  CheckBox 75, 150, 30, 10, "CLMS", clms_check
  CheckBox 110, 150, 30, 10, "CSES", cses_check
  CheckBox 145, 150, 30, 10, "ELIG", elig_check
  CheckBox 180, 150, 30, 10, "IEVS", ievs_check
  CheckBox 210, 150, 30, 10, "INFO", info_check
  CheckBox 10, 165, 25, 10, "IV-E", ive_check
  CheckBox 40, 165, 25, 10, "MA", ma_check
  CheckBox 75, 165, 30, 10, "MEC2", mec2_check
  CheckBox 110, 165, 35, 10, "PARI", pari_chck
  CheckBox 145, 165, 30, 10, "PEPR", pepr_check
  CheckBox 180, 165, 30, 10, "TIKL", tikl_check
  CheckBox 210, 165, 30, 10, "WF1", wf1_check
  ButtonGroup ButtonPressed
    OkButton 155, 185, 40, 15
    CancelButton 200, 185, 40, 15
  GroupBox 5, 80, 240, 50, "Step 1. Select the population"
  Text 65, 5, 95, 10, "---Task-Based DAIL Capture---"
  Text 10, 35, 220, 35, "This script will evaluate and capture actionable DAIL messages from the DAIL type and population selected below. Once the DAIL messages are evaluated, actionable DAIL messages are sent to a SQL Database which feeds the Big Scoop Report."
  GroupBox 5, 20, 240, 55, "Using This Script:"
  GroupBox 5, 135, 240, 45, "Step 2. Select the type(s) of DAIL message to add to the report:"
  Text 10, 215, 220, 25, "The SQL Database takes up to 15 minutes to load. This happens after the DAIL has been evaluated. DO NOT stop the script. Wait until a success message box appears."
  GroupBox 5, 200, 240, 45, "Warning!"
EndDialog

Do
    Do
        err_msg = ""
  	    dialog Dialog1
  	    cancel_without_confirmation
        If all_baskets_checkbox = 1 then
            If adults_checkbox = 1 or LTC_checkbox = 1 or families_checkbox = 1 then err_msg = err_msg & vbcr & "* You cannot select a population and all populations."
        End if
        If all_workers_check = 1 then
            If adults_checkbox = 1 or LTC_checkbox = 1 or families_checkbox = 1 or all_baskets_checkbox = 1 then err_msg = err_msg & vbcr & "* You cannot select a population(s) and all workers."
        End if
        If (all_baskets_checkbox = 0 and all_workers_check = 0 and adults_checkbox = 0 and LTC_checkbox = 0 and families_checkbox = 0 and all_baskets_checkbox = 0) then err_msg = err_msg & vbcr & "* You must select at least one population option."
        If (all_check = 0 and cola_check = 0 and clms_check = 0 and cses_check  = 0 and elig_check  = 0 and ievs_check = 0 and info_check = 0 and ive_check = 0 and ma_check = 0 and mec2_check = 0 and pari_chck = 0 and pepr_check = 0 and tikl_check = 0 and wf1_check = 0) then err_msg = err_msg & vbcr & "* You must select at least one DAIL type."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    Loop Until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

back_to_SELF 'navigates back to self in case the worker is working within the DAIL. All messages for a single number may not be captured otherwise.

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
    adults_baskets = "X127ED8,X127EE1,X127EE2,X127EE3,X127EE4,X127EE5,X127EE6,X127EE7,X127EG4,X127EH8,X127EL1,X127EL2,X127EL3,X127EL4,X127EL5,X127EL6,X127EL7,X127EL8,X127EL9,X127EN1,X127EN2,X127EN3,X127EN4,X127EN5,X127EN7,X127EP6,X127EP7,X127EQ1,X127EQ3,X127EQ4,X127EQ5,X127EQ8,X127EQ9,X127EX1,X127EX2,"
    LTC_plus_baskets = "X127EH1,X127EH3,X127EH4,X127EH5,X127EH6,X127EH7,X127EJ4,X127EJ8,X127EK1,X127EK2,X127EK3,X127EK4,X127EK6,X127EK7,X127EK8,X127EK9,X127EM9,X127EN6,X127EP5,X127EP9,X127EZ5,X127F3F,X127FE5,X127FH4,X127FH5,X127FI2,X127FI7,"
    families_baskets = "X127EA0,X127ES1,X127ES2,X127ES3,X127ES4,X127ES5,X127ES6,X127ES7,X127ES8,X127ES9,X127ET1,X127ET2,X127ET3,X127ET4,X127ET5,X127ET6,X127ET7,X127ET8,X127ET9,X127EZ1,X127EZ7,"

    worker_numbers = ""     'Creating and valuing incrementor variables

    If adults_checkbox = 1 then worker_numbers = worker_numbers & adults_baskets
    If families_checkbox = 1 then worker_numbers = worker_numbers & families_baskets
    If LTC_checkbox = 1 then worker_numbers = worker_numbers & LTC_plus_baskets
    If all_baskets_checkbox = 1 then worker_numbers = adults_baskets & families_baskets & LTC_plus_baskets  'conditional logic in do loop doesn't allow for populations and baskets to be selcted. Not incremented variable.

    x1s_from_dialog = split(worker_numbers, ",")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x1_number))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & "," & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next
	'Split worker_array
	worker_array = split(worker_array, ",")
End if

'----------------------------------------------------------------------------------------------------Setting up and valueing the array
Dim DAIL_array()
ReDim DAIL_array(4, 0)
Dail_count = 0              'Incremental for the array
all_dail_array = "*"    'setting up string to find duplicate DAIL messages. At times there is a glitch in the DAIL, and messages are reviewed a second time.
false_count = 0

'constants for array
const worker_const	            = 0
const maxis_case_number_const   = 1
const dail_type_const 	        = 2
const dail_month_const		    = 3
const dail_msg_const		    = 4

deleted_dails = 0	'establishing the value of the count for deleted deleted_dails

'----------------------------------------------------------------------------------------------------DAIL Actions
CALL navigate_to_MAXIS_screen("DAIL", "PICK")
EMReadscreen pick_confirmation, 26, 4, 29

If pick_confirmation = "View/Pick Selection (PICK)" then
    'selecting the type of DAIl message
    If all_check = 1   then EMWriteScreen "x", 7, 39
	If cola_check = 1  then EMWriteScreen "x", 8, 39
	If clms_check = 1  then EMWriteScreen "x", 9, 39
	If cses_check = 1  then EMWriteScreen "x", 10, 39
	If elig_check = 1  then EMWriteScreen "x", 11, 39
	If ievs_check = 1  then EMWriteScreen "x", 12, 39
	If info_check = 1  then EMWriteScreen "x", 13, 39
	If ive_check = 1   then EMWriteScreen "x", 14, 39
    If ma_check = 1    then EMWriteScreen "x", 15, 39
 	If mec2_check = 1  then EMWriteScreen "x", 16, 39
	If pari_chck = 1   then EMWriteScreen "x", 17, 39
	If pepr_check = 1  then EMWriteScreen "x", 18, 39
	If tikl_check = 1  then EMWriteScreen "x", 19, 39
	If wf1_check = 1   then EMWriteScreen "x", 20, 39
	transmit
Else
    script_end_procedure("Unable to navigate to DAIL/PICK. The script will now end.")
End if

'Ending message when there are no more DAIL's differs based on if you select ALL DAIL's or specific DAILs
If all_check = 1 then
    dail_end_msg = "NO MESSAGES WORK"
Else
    'all specified selection(s) will get this ending user message.
    dail_end_msg = "NO MESSAGES TYPE"
End if

'This for...next contains each worker indicated above
For each worker in worker_array
	EMWriteScreen worker, 21, 6
	transmit
	transmit 'transmit past 'not your dail message'
    EMReadScreen number_of_dails, 1, 3, 67		'Reads where the count of DAILs is listed

	DO
		If number_of_dails = " " Then exit do		'if this space is blank the rest of the DAIL reading is skipped
		dail_row = 6			'Because the script brings each new case to the top of the page, dail_row starts at 6.
		DO
			dail_type = ""
			dail_msg = ""

		    'Determining if there is a new case number...
		    EMReadScreen new_case, 8, dail_row, 63
		    new_case = trim(new_case)
		    IF new_case <> "CASE NBR" THEN '...if there is NOT a new case number, the script will read the DAIL type, month, year, and message...
				Call write_value_and_transmit("T", dail_row, 3)
				dail_row = 6
			ELSEIF new_case = "CASE NBR" THEN
			    '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
			    Call write_value_and_transmit("T", dail_row + 1, 3)
				dail_row = 6
			End if

            'Reading the DAIL Information
			EMReadScreen MAXIS_case_number, 8, dail_row - 1, 73
            MAXIS_case_number = trim(MAXIS_case_number)

            EMReadScreen dail_type, 4, dail_row, 6

            EMReadScreen dail_msg, 61, dail_row, 20
			dail_msg = trim(dail_msg)

            EMReadScreen dail_month, 8, dail_row, 11
            dail_month = trim(dail_month)

            stats_counter = stats_counter + 1   'I increment thee
            Call non_actionable_dails(actionable_dail)   'Function to evaluate the DAIL messages
            'This removes the Ex Parte TIKL's from ES Workflow assignments only. Does not delete for the benefit of other counties. 
            If instr(dail_msg, "PHASE 1 - THE CASE HAS BEEN EVALUATED FOR EX PARTE AND") then actionable_dail = False

            IF actionable_dail = True then      'actionable_dail = True will NOT be deleted and will be captured and reported out as actionable.
                If len(dail_month) = 5 then
                    output_year = ("20" & right(dail_month, 2))
                    output_month = left(dail_month, 2)
                    output_day = "01"
                    dail_month = output_year & "-" & output_month & "-" & output_day
                elseif trim(dail_month) <> "" then
                    'Adjusting data for output to SQL
                    output_year     = DatePart("yyyy",dail_month)   'YYYY-MM-DD format
                    output_month    = right("0" & DatePart("m", dail_month), 2)
                    output_day      = DatePart("d", dail_month)
                    dail_month = output_year & "-" & output_month & "-" & output_day
                End if

                dail_string = worker & " " & MAXIS_case_number & " " & dail_type & " " & dail_month & " " & dail_msg
                'If the case number is found in the string of case numbers, it's not added again.
                If instr(all_dail_array, "*" & dail_string & "*") then
                    If dail_type = "HIRE" then
                        add_to_array = True
                    Else
                        add_to_array = False
                    End if
                else
                    add_to_array = True
                End if

                If add_to_array = True then
                    ReDim Preserve DAIL_array(4, DAIL_count)	'This resizes the array based on the number of rows in the Excel File'
            	    DAIL_array(worker_const,	           DAIL_count) = worker
            	    DAIL_array(maxis_case_number_const,    DAIL_count) = right("00000000" & MAXIS_case_number, 8) 'outputs in 8 digits format
            	    DAIL_array(dail_type_const, 	       DAIL_count) = dail_type
            	    DAIL_array(dail_month_const, 		   DAIL_count) = dail_month
            	    DAIL_array(dail_msg_const, 		       DAIL_count) = dail_msg
                    Dail_count = DAIL_count + 1
                    all_dail_array = trim(all_dail_array & dail_string & "*") 'Adding MAXIS case number to case number string
                    dail_string = ""
                elseif add_to_array = False then
                    false_count = false_count + 1
                End if
			End if

            dail_row = dail_row + 1
			'...going to the next page if necessary
			EMReadScreen next_dail_check, 4, dail_row, 4
			If trim(next_dail_check) = "" then
				PF8
                EMReadScreen last_page_check, 16, 24, 2
                'DAIL/PICK will look for 'no message worker X127XXX as the full message.
                If last_page_check = "THIS IS THE LAST" or last_page_check = dail_end_msg then
					all_done = true
					exit do
				Else
					dail_row = 6
				End if
			End if
		LOOP
		IF all_done = true THEN exit do
	LOOP
Next

'----------------------------------------------------------------------------------------------------SQL Database Actions
'Setting constants
Const adOpenStatic = 3
Const adLockOptimistic = 3

''Creating objects for Database
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

'How to connect to the database
'Provider: the type of connection you are establishing, in this case SQL Server.
'Data Source: The server you are connecting to.
'Initial Catalog: The name of the database.
'user id: your username.
'password: um, your password. ;)

objConnection.Open "Provider = SQLOLEDB.1;Data Source= hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"

'Deleting ALL data fom DAIL table prior to loading new DAIL messages.
objRecordSet.Open "DELETE FROM EWS.DAILDecimator",objConnection, adOpenStatic, adLockOptimistic

'Export informaiton to Excel re: case status
For item = 0 to UBound(DAIL_array, 2)
    worker             = DAIL_array(worker_const, item)
    MAXIS_case_number  = DAIL_array(maxis_case_number_const, item)
    dail_type          = DAIL_array(dail_type_const, item)
    dail_month         = DAIL_array(dail_month_const, item)
    dail_msg           = DAIL_array(dail_msg_const, item)

    If instr(dail_msg, "'") then dail_msg = replace(dail_msg, "'", " ") 'SQL will not allow for an apostrophe
    If instr(dail_msg, "*") then dail_msg = replace(dail_msg, "*", " ") 'SQL will not allow for an apostrophe
    dail_msg = trim(dail_msg)
    'Opening Database and adding a record
    objRecordSet.Open "INSERT INTO EWS.DAILDecimator(EmpStateLogOnID, MaxisCaseNumber, DAILType, DAILMessage, DAILMonth)" & _
    "VALUES ('" & worker & "', '" & MAXIS_case_number & "', '" & dail_type & "', '" & dail_msg & "', '" & dail_month & "')", objConnection, adOpenStatic, adLockOptimistic
Next

'Closing the connection
objConnection.Close

'Function create_outlook_email(email_from, email_recip, email_recip_CC, email_recip_bcc, email_subject, email_importance, include_flag, email_flag_text, email_flag_days, email_flag_reminder, email_flag_reminder_days, email_body, include_email_attachment, email_attachment_array, send_email)
Call create_outlook_email("", "Ilse.Ferris@hennepin.us", "Mary.McGuinness@Hennepin.us", "", "Task-Based DAIL Capture Complete. Actionable DAIL Count: " & DAIL_count & ". EOM.", 1, False, "", "", False, "", "", False, "", True)
stats_counter = stats_counter -1
script_end_procedure("Success! Actionable DAIL's have been added to the database.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------04/12/2022
'--Tab orders reviewed & confirmed----------------------------------------------04/12/2022
'--Mandatory fields all present & Reviewed--------------------------------------04/12/2022
'--All variables in dialog match mandatory fields-------------------------------04/12/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------04/12/2022-------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------04/12/2022-------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------04/12/2022-------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used-10/06/2022-------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------04/12/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------04/12/2022-------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------04/12/2022-------------------N/A
'--Out-of-County handling reviewed----------------------------------------------04/12/2022-------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------04/12/2022
'--BULK - review output of statistics and run time/count (if applicable)--------04/12/2022-------------------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---10/06/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------04/12/2022
'--Incrementors reviewed (if necessary)-----------------------------------------04/12/2022
'--Denomination reviewed -------------------------------------------------------04/12/2022
'--Script name reviewed---------------------------------------------------------04/12/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------04/12/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete-----------------------------------------04/12/2022
'--comment Code-----------------------------------------------------------------04/12/2022
'--Update Changelog for release/update------------------------------------------04/12/2022
'--Remove testing message boxes-------------------------------------------------04/12/2022
'--Remove testing code/unnecessary code-----------------------------------------04/12/2022
'--Review/update SharePoint instructions----------------------------------------04/12/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------04/12/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------04/12/2022
'--Complete misc. documentation (if applicable)---------------------------------04/12/2022
'--Update project team/issue contact (if applicable)----------------------------10/06/2022
