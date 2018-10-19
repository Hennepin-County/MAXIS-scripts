'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - DAIL CLEAN UP.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
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
call changelog_update("03/23/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

'END CHANGELOG BLOCK =======================================================================================================

BeginDialog dail_dialog, 0, 0, 266, 90, "Dail Decimator dialog"
  Text 20, 20, 235, 20, "This script will delete DAILs for cases that are inactive, and capture DAIL messages that cannot be deleted, and may need action."
  EditBox 80, 50, 180, 15, worker_number
  CheckBox 15, 75, 135, 10, "Check here to process for all workers.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 155, 70, 50, 15
    CancelButton 210, 70, 50, 15
  Text 15, 55, 60, 10, "Worker number(s):"
  GroupBox 10, 5, 250, 40, "Using the DAIL Decimator script"
EndDialog

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""

'Grabbing user ID to validate user of script. Only some users are allowed to use this script.
Set objNet = CreateObject("WScript.NetWork") 
user_ID_for_validation = ucase(objNet.UserName)

'Validating user ID
If user_ID_for_validation = "ILFE001" OR _		
	user_ID_for_validation = "WF7638" OR _		
	user_ID_for_validation = "WF1875" OR _ 		
	user_ID_for_validation = "WFQ898" OR _ 		
	user_ID_for_validation = "WFP803" OR _		
	user_ID_for_validation = "WFP106" OR _		
	user_ID_for_validation = "WFK093" OR _ 		
	user_ID_for_validation = "WF1373" OR _ 		
	user_ID_for_validation = "WFU161" OR _ 		
	user_ID_for_validation = "WFS395" OR _ 		
	user_ID_for_validation = "WFU851" OR _ 		
	user_ID_for_validation = "WFX901" OR _ 
	user_ID_for_validation = "CALO001" OR _ 		
	user_ID_for_validation = "WFI021" then 		
    'the dialog
    Do
    	Do
      		err_msg = ""
      		dialog dail_dialog
      		If ButtonPressed = 0 then StopScript
      		If trim(worker_number) = "" and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases."	
      		If trim(worker_number) <> "" and all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases, not both options."							
      	  	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine										
      	LOOP until err_msg = ""		
      	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
      Loop until are_we_passworded_out = false					'loops until user passwords back in		
Else 
	script_end_procedure("This script is for Quality Improvement staff only. You do not have access to use this script.")
End if 

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ", ")	'Splits the worker array based on commas

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

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet to "DAIL List"
ObjExcel.ActiveSheet.Name = "Clean Up!" 

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "X NUMBER"
objExcel.Cells(1, 2).Value = "CASE #"
objExcel.Cells(1, 3).Value = "DAIL TYPE"
objExcel.Cells(1, 4).Value = "DAIL MO."
objExcel.Cells(1, 5).Value = "DAIL MESSAGE"
objExcel.Cells(1, 6).Value = "NOTES"

FOR i = 1 to 6		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Sets variable for all of the Excel stuff
excel_row = 2
deleted_dails = 0	'establishing the value of the count for deleted deleted_dails

MAXIS_case_number = ""
CALL navigate_to_MAXIS_screen("DAIL", "DAIL")

'This for...next contains each worker indicated above
For each worker in worker_array	
	DO 
		EMReadScreen dail_check, 4, 2, 48
		If next_dail_check <> "DAIL" then 
			MAXIS_case_number = ""
			CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
		End if 
	Loop until dail_check = "DAIL"
	
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
			
			EMReadScreen dail_type, 4, dail_row, 6
			EMReadScreen dail_msg, 61, dail_row, 20
            EMReadScreen dail_month, 8, dail_row, 11
            EMReadScreen maxis_case_number, 8, dail_row - 1, 73
            maxis_case_number = trim(maxis_case_number)
			dail_msg = trim(dail_msg)
            dail_month = trim(dail_month)
			stats_counter = stats_counter + 1

            Call write_value_and_transmit("H", dail_row, 3)
            EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
            If PRIV_check = "PRIV" then
                msgbox "PRIV case, navigate to the case after this one, and go into CASE/CURR." & vbcr & MAXIS_case_number
                'clear message, go back to the DAIL and find the case number after it
            End if 
            EMReadScreen case_status, 8, 8, 9
            If case_status = "INACTIVE" then 
                add_to_excel = True
                PF3
                '--------------------------------------------------------------------...and put that in Excel.
				Do 
                    objExcel.Cells(excel_row, 1).Value = worker
				    objExcel.Cells(excel_row, 2).Value = trim(maxis_case_number)
				    objExcel.Cells(excel_row, 3).Value = trim(dail_type)
				    objExcel.Cells(excel_row, 4).Value = trim(dail_month)
				    objExcel.Cells(excel_row, 5).Value = trim(dail_msg)
				    
				    Call write_value_and_transmit("D", dail_row, 3)	
				    EMReadScreen other_worker_error, 13, 24, 2
				    If other_worker_error = "** WARNING **" then 
                        transmit
				        deleted_dails = deleted_dails + 1
                        exit do 
                    Elseif trim(other_worker_error) <> "" then 
                        EMWriteScreen "_", dail_row, 3
                        objExcel.Cells(excel_row, 6).Value = "Unable to delete"
                        dail_row = dail_row + 1
                        Do 
                            EMReadScreen next_case, 8, dail_row, 63
                            'msgbox next_case & vbcr & dail_row
                            If next_case <> "CASE NBR" THEN
                                exit do    'this will make the next dail message will be added to excel 
                            elseIf next_case = "CASE NBR" then 
                                EMReadScreen next_case_number, 8, dail_row, 73
                                'msgbox next_case_number & vbcr & dail_row & vbcr & maxis_case_number
                                If trim(next_case_number) = maxis_case_number then 
                                    matching_case_number = true
                                    dail_row = dail_row + 1 
                                    'msgbox "Matching case number is TRUE" & vbcr & dail_row
                                Elseif IsNumeric(next_case_number) = true then 
                                    'msgbox "exit do"
                                    next_case = True 
                                    exit do 'another case number was found 
                                Else 
                                    'msgbox "Matching case number is FALSE"
                                    matching_case_number = false 
                                    dail_row = dail_row + 1 
                                End if
                            End if 
                            If dail_row = 19 then 
                                PF8
                                dail_row = 6
                            End if 
                        Loop until matching_case_number = False
                    END IF
                    If next_case = True then exit do 
                    If matching_case_number = False then exit do
                Loop
                excel_row = excel_row + 1    
            Else 
                add_to_excel = False 
                PF3
                dail_row = dail_row + 1
                Do 
                    EMReadScreen next_case, 8, dail_row, 63
                    If trim(next_case) = "" then    
                        'msgbox next_case & vbcr & dail_row
                        EMReadScreen dail_content, 4, dail_row, 6
                        'msgbox dail_content & vbcr & dail_row
                        If trim(dail_content) = "" then
                            If dail_row = 18 then 
                                PF8
                                EMReadScreen last_page_check, 21, 24, 2
                				If last_page_check = "THIS IS THE LAST PAGE" then 
                					all_done = true
                                    msgbox "all done: " & all_done
                                    exit do 
                                Else 
                                    dail_row = 6
                                End if 
                            End if 
                        else
                            dail_row = dail_row + 1
                        End if 
                    elseif next_case = "CASE NBR" then 
                        EMReadScreen next_case_number, 8, dail_row, 73
                        'msgbox next_case_number & vbcr & dail_row & vbcr & maxis_case_number
                        If trim(next_case_number) = maxis_case_number then 
                            matching_case_number = true
                            dail_row = dail_row + 1 
                            'msgbox "Matching case number is TRUE" & vbcr & dail_row
                        Elseif IsNumeric(next_case_number) = true then 
                            'msgbox "exit do"
                            exit do 'another case number was found 
                        Else 
                            'msgbox "Matching case number is FALSE"
                            matching_case_number = false 
                            dail_row = dail_row + 1 
                        End if
                    Else 
                        dail_row = dail_row + 1
                    End if 
                    If dail_row = 19 then 
                        PF8
                        dail_row = 6
                    End if 
                    'msgbox matching_case_number & vbcr & maxis_case_number & vbcr & next_case_number
                    'msgbox dail_row
                Loop until matching_case_number = False
                'msgbox "out of loop"
            End if 
                
			''msgbox dail_row
			'EMReadScreen message_error, 11, 24, 2		'Cases can also NAT out for whatever reason if the no messages instruction comes up.
			'If message_error = "NO MESSAGES" then
			'	CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
			'	Call write_value_and_transmit(worker, 21, 6)
			'	transmit   'transmit past 'not your dail message'
			'	exit do
			'End if 
	    	
			''...going to the next page if necessary
			'EMReadScreen next_dail_check, 4, dail_row, 4
			'If trim(next_dail_check) = "" then 
			'	PF8
			'	EMReadScreen last_page_check, 21, 24, 2
			'	If last_page_check = "THIS IS THE LAST PAGE" then 
			'		all_done = true
			'		exit do 
			'	Else 
			'		dail_row = 6
			'	End if
            'End if 
			''End if
		LOOP
		IF all_done = true THEN exit do
	LOOP
Next

STATS_counter = STATS_counter - 1
'Enters info about runtime for the benefit of folks using the script
objExcel.Cells(2, 7).Value = "Number of DAILs deleted:"
objExcel.Cells(3, 7).Value = "Average time to find/select/copy/paste one line (in seconds):"
objExcel.Cells(4, 7).Value = "Estimated manual processing time (lines x average):"
objExcel.Cells(5, 7).Value = "Script run time (in seconds):"
objExcel.Cells(6, 7).Value = "Estimated time savings by using script (in minutes):"
objExcel.Cells(7, 7).Value = "Number of " & dail_to_decimate & " messages reviewed"
objExcel.Columns(7).Font.Bold = true
objExcel.Cells(2, 8).Value = deleted_dails
objExcel.Cells(3, 8).Value = STATS_manualtime
objExcel.Cells(4, 8).Value = STATS_counter * STATS_manualtime
objExcel.Cells(5, 8).Value = timer - start_time
objExcel.Cells(6, 8).Value = ((STATS_counter * STATS_manualtime) - (timer - start_time)) / 60
objExcel.Cells(7, 8).Value = STATS_counter

'Formatting the column width.
FOR i = 1 to 8
	objExcel.Columns(i).AutoFit()
NEXT

script_end_procedure("Success! Please review the list created for accuracy.")

'NEW HIRE msg stuff

''current month -1
'CM_minus_1_mo =  right("0" &          	 DatePart("m",           DateAdd("m", -1, date)            ), 2)
'CM_minus_1_yr =  right(                  DatePart("yyyy",        DateAdd("m", -1, date)            ), 2)


'current_month = CM_mo & " " & CM_yr
'next_month = CM_plus_1_mo & " " & CM_plus_1_yr
'previous_month = CM_minus_1_mo & " " & CM_minus_1_yr

'If case_status = "INACTIVE" then 
'    If (dail_type = "HIRE" AND dail_month = current_month OR dail_month = next_month OR dail_month = previous_month) then                     
'        add_to_excel = False 'does not delete current new hire messages 