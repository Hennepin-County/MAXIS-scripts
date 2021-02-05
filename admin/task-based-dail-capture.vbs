'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - TASK-BASED DAIL CAPTURE.vbs"
start_time = timer
STATS_counter = 0                       'sets the stats counter at zero
STATS_manualtime = 20
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
call changelog_update("02/02/2021", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

'END CHANGELOG BLOCK =======================================================================================================

Function dail_type_selection
	'selecting the type of DAIl message
    If all_check = 0 then
	    EMWriteScreen "x", 4, 12		'transmits to the PICK screen
	    transmit
	    EMWriteScreen "_", 7, 39		'clears the all selection

	    If cola_check = 1 then EMWriteScreen "x", 8, 39
	    If clms_check = 1 then EMWriteScreen "x", 9, 39
	    If cses_check = 1 then EMWriteScreen "x", 10, 39
	    If elig_check = 1 then EMWriteScreen "x", 11, 39
	    If ievs_check = 1 then EMWriteScreen "x", 12, 39
	    If info_check = 1 then EMWriteScreen "x", 13, 39
	    If ive_check = 1 then EMWriteScreen "x", 14, 39
	    If ma_check = 1 then EMWriteScreen "x", 15, 39
 	    If mec2_check = 1 then EMWriteScreen "x", 16, 39
	    If pari_chck = 1 then EMWriteScreen "x", 17, 39
	    If pepr_check = 1 then EMWriteScreen "x", 18, 39
	    If tikl_check = 1 then EMWriteScreen "x", 19, 39
	    If wf1_check = 1 then EMWriteScreen "x", 20, 39
	    transmit
    End if
End Function

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""

'Defaulting these populations to checked per current process. 
'ADAD_checkbox = 1   
'FAD_checkbox = 1
ADS_checkbox = 1    'For testing - this is my x number 

'Defaulting these messages to checked as these are the most assigned cases. 
cola_check = 1
cses_check = 1
info_check = 1
pepr_check = 1
tikl_check = 1

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 251, 260, "Task-Based DAIL Capture Main Dialog"
  CheckBox 55, 100, 85, 10, "All Population Baskets", all_baskets_checkbox
  CheckBox 140, 100, 30, 10, "ADAD", ADAD_checkbox
  CheckBox 175, 100, 30, 10, "ADS", ADS_checkbox
  CheckBox 205, 100, 30, 10, "FAD", FAD_checkbox
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
  Text 10, 100, 40, 10, "Population:"
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
            If ADAD_checkbox = 1 or ADS_checkbox = 1 or FAD_checkbox = 1 then err_msg = err_msg & vbcr & "* You cannot select a population and all populations."
        End if 
        If all_workers_check = 1 then 
            If ADAD_checkbox = 1 or ADS_checkbox = 1 or FAD_checkbox = 1 or all_baskets_checkbox = 1 then err_msg = err_msg & vbcr & "* You cannot select a population(s) and all workers."
        End if 
        If (all_baskets_checkbox = 0 and all_workers_check = 0 and ADAD_checkbox = 0 and ADS_checkbox = 0 and FAD_checkbox = 0 and all_baskets_checkbox = 0) then err_msg = err_msg & vbcr & "* You must select at least one population option."
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
    ADAD_baskets = "X127EE1,X127EE2,X127EE3,X127EE4,X127EE5,X127EE6,X127EE7,X127EL2,X127EL3,X127EL4,X127EL5,X127EL6,X127EL7,X127EL8,X127EN1,X127EN2,X127EN3,X127EN4,X127EN5,X127EQ1,X127EQ4,X127EQ5,X127EQ8,X127EQ9,X127EL9,X127ED8,X127EH8,X127EG4,X127EQ3,X127EQ2,X127EP6,X127EP7,X127EP8,X127EF8,X127EF9,"
    'ADS_baskets = "X127EH1,X127EH2,X127EH3,X127EH6,X127EJ4,X127EJ6,X127EJ7,X127EJ8,X127EK1,X127EK2,X127EK4,X127EK5,X127EK6,X127EK9,X127EM1,X127EM7,X127EM8,X127EM9,X127EN6,X127EP3,X127EP4,X127EP5,X127EP9,X127F3F,X127FE5,X127FG3,X127FH4,X127FH5,X127FI2,X127FI7,X127EJ5,"
    ADS_baskets = "X127CCL"
    FAD_baskets = "X127ES1,X127ES2,X127ES3,X127ES4,X127ES5,X127ES6,X127ES7,X127ES8,X127ES9,X127ET1,X127ET2,X127ET3,X127ET4,X127ET5,X127ET6,X127ET7,X127ET8,X127ET9,X127FE7,X127FE8,X127FE9,X127FA5,X127FA9,X127FA6,X127FA7,X127FA8"
    
    worker_numbers = ""     'Creating and valuing incrementor for array 

    If ADAD_checkbox = 1 then worker_numbers = worker_numbers & ADAD_baskets
    If ADS_checkbox = 1 then worker_numbers = worker_numbers & ADS_baskets
    If FAD_checkbox = 1 then worker_numbers = worker_numbers & FAD_baskets
    If all_baskets_checkbox = 1 then worker_numbers = ADAD_baskets & "," & ADS_baskets & "," & FAD_baskets  'conditional logic in do loop doesn't allow for populations and baskets to be selcted. Not incremented variable.
        
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

DIM DAIL_array()
ReDim DAIL_array(4, 0)
Dail_count = 0              'Incremental for the array
all_dail_array = "*"    'setting up string to find duplicate DAIL messages. At times there is a glitch in the DAIL, and messages are reviewed a second time.

'constants for array
const worker_const	            = 0
const maxis_case_number_const   = 1
const dail_type_const 	        = 2
const dail_month_const		    = 3
const dail_msg_const		    = 4

Call navigate_to_MAXIS_screen("DAIL", "DAIL")

'CALL navigate_to_MAXIS_screen("DAIL", "PICK")
'EMReadscreen pick_confirmation, 26, 4, 29
'
'If pick_confirmation = "View/Pick Selection (PICK)" then 
'    'selecting the type of DAIl message
'    If all_check = 1   then EMWriteScreen "x", 7, 39
'	If cola_check = 1  then EMWriteScreen "x", 8, 39
'	If clms_check = 1  then EMWriteScreen "x", 9, 39
'	If cses_check = 1  then EMWriteScreen "x", 10, 39
'	If elig_check = 1  then EMWriteScreen "x", 11, 39
'	If ievs_check = 1  then EMWriteScreen "x", 12, 39
'	If info_check = 1  then EMWriteScreen "x", 13, 39
'	If ive_check = 1   then EMWriteScreen "x", 14, 39
'    If ma_check = 1    then EMWriteScreen "x", 15, 39
' 	If mec2_check = 1  then EMWriteScreen "x", 16, 39
'	If pari_chck = 1   then EMWriteScreen "x", 17, 39
'	If pepr_check = 1  then EMWriteScreen "x", 18, 39
'	If tikl_check = 1  then EMWriteScreen "x", 19, 39
'	If wf1_check = 1   then EMWriteScreen "x", 20, 39
'	transmit
'Else
'    script_end_procedure("Unable to navigate to DAIL/PICK. The script will now end.")
'End if 

'----------------------------------------------------------------------------------------------------DAIL Actions
'This for...next contains each worker indicated above
For each worker in worker_array
	EMWriteScreen worker, 21, 6
	transmit
	transmit 'transmit past 'not your dail message'
    
    Call dail_type_selection

    EMReadScreen number_of_dails, 1, 3, 67		'Reads where the count of DAILs is listed
	DO
		If number_of_dails = " " Then exit do		'if this space is blank the rest of the DAIL reading is skipped

		dail_row = 5			'1st line with Case Number  
		DO
			dail_type = ""
			dail_msg = ""
            
		    EMReadScreen new_case, 8, dail_row, 63    'Determining if there is a new case number...
		    new_case = trim(new_case)
            IF new_case = "CASE NBR" THEN
                case_number_row = dail_row
                MAXIS_case_number = ""
                dail_row = dail_row + 1 'incrementing to get to the DAIL information 
            End if 
            
            'Reading the DAIL Information
			EMReadScreen MAXIS_case_number, 8, case_number_row, 73
            MAXIS_case_number = trim(MAXIS_case_number)
            
            EMReadScreen dail_type, 4, dail_row, 6

            EMReadScreen dail_msg, 61, dail_row, 20
			dail_msg = trim(dail_msg)
            
            If dail_msg <> "" then             
                stats_counter = stats_counter + 1   'I increment thee, but only the non-blank messages 
                EMReadScreen dail_month, 8, dail_row, 11
                dail_month = trim(dail_month)
                
                Call non_actionable_dails(actionable_dail)   'Function to evaluate the DAIL messages
                
                'If MAXIS_case_number = "1362618" then 
                '    msgbox "case_number_row: " & case_number_row & " Case Number: " & MAXIS_case_number & vbcr & vbcr & "dail_row: " & dail_row & vbcr & "dail_msg: " & dail_msg & vbcr & vbcr & "actionable_dail: " & actionable_dail  
                '    msgbox "stats_counter: " & stats_counter
                'End if 
                
                IF actionable_dail = True then 
                'actionable_dail = True will be captured and reported out as actionable.  
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
                            add_to_array = True     'Adding all HIRE messages to the SQL output array
                        Else 
                            add_to_array = False    'Not adding other duplicate messages
                        End if 
                    else 
                        add_to_array = True         'Defaulting any other condition to adding to the array 
                    End if 
                    
                    If add_to_array = True then    
                        'msgbox DAIL_count & vbcr & dail_string
                        ReDim Preserve DAIL_array(4, DAIL_count)	'This resizes the array based on the number of rows in the Excel File'
                        DAIL_array(worker_const,	           DAIL_count) = worker
                        DAIL_array(maxis_case_number_const,    DAIL_count) = right("00000000" & MAXIS_case_number, 8) 'outputs in 8 digits format for consistancy in the Database 
                        DAIL_array(dail_type_const, 	       DAIL_count) = dail_type
                        DAIL_array(dail_month_const, 		   DAIL_count) = dail_month
                        DAIL_array(dail_msg_const, 		       DAIL_count) = dail_msg
                        Dail_count = DAIL_count + 1
                        all_dail_array = trim(all_dail_array & dail_string & "*") 'Adding MAXIS case number to case number string
                        dail_string = ""
                    End if 
                End if
            Else 
                'msgbox "Blank message on row: " & dail_row
            End if 
            
            'If at the bottom of the screen, then will navigate to the next screen. Otherwise if at the end, will exit the do...loop. 
            dail_row = dail_row + 1
            If dail_row = 19 then 
                'msgbox "Next Page. Stats counter: " & stats_counter
                PF8 
                EMReadScreen last_page_check, 21, 24, 2
				If last_page_check = "THIS IS THE LAST PAGE" then
                    msgbox "This is the last page."
					all_done = true    'setting variable to exit the second do...loop
					exit do
				Else
					dail_row = 5 'starting at the top of the next page. 
				End if
            End if 
        LOOP
    	IF all_done = true THEN exit do    
	LOOP
Next

msgbox "script will stop now. stats_counter: " & stats_counter & vbcr & dail_count
stopscript 

'EMReadScreen message_error, 11, 24, 2		'Cases can also NAT out for whatever reason if the no messages instruction comes up.
'If message_error = "NO MESSAGES" then
'	CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
'	Call write_value_and_transmit(worker, 21, 6)
'	transmit   'transmit past 'not your dail message'
'    Call dail_type_selection
'	exit do
'End if

'----------------------------------------------------------------------------------------------------SQL Database Output 
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

objConnection.Open "Provider = SQLOLEDB.1;Data Source= HSSQLPW017;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"

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

email_info = "Task-Based DAIL Capture Complete. Number of DAIL's added to database: " & DAIL_count & ". EOM."
'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
Call create_outlook_email("Laurie.Hennen@hennepin.us;Todd.Bennington@hennepin.us", "Ilse.Ferris@hennepin.us", email_info, "", "", True)

script_end_procedure("Success! The Task-Based DAIL capture is complete. Number of DAIL's added to database: " & DAIL_count)