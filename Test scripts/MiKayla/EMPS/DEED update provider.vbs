'Required for statistical purposes===============================================================================
name_of_script =  "DEED update provider.vbs"
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("12/26/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=======================================================================================================END CHANGELOG BLOCK 

'-------------------------------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""	
'dialog and dialog DO...Loop	
Do
	Do
		BeginDialog DEED_referral_dialog, 0, 0, 266, 95, "DEED provider update"
	  		ButtonGroup ButtonPressed
	    	PushButton 200, 25, 50, 15, "Browse...", select_a_file_button
	    	OkButton 150, 75, 50, 15
	    	CancelButton 205, 75, 50, 15
	  		EditBox 15, 25, 180, 15, file_selection_path
	  		GroupBox 10, 5, 250, 65, "DEED provider update "
	  		Text 15, 50, 240, 15, "Select the Excel file that contains the DEED information by selecting the 'Browse' button, and finding the file."
		EndDialog
		err_msg = ""
		Dialog DEED_referral_dialog
		cancel_confirmation
		If ButtonPressed = select_a_file_button THEN
			If file_selection_path <> "" THEN 'This is handling for if the BROWSE button is pushed more than once'
				objExcel.Quit 'Closing the Excel file that was opened on the first push'
				objExcel = "" 	'Blanks out the previous file path'
			End If
			call file_selection_system_dialog(file_selection_path, ".xlsx") 'allows the user to select the file'
		End If
		If file_selection_path = "" THEN err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
		If err_msg <> "" THEN MsgBox err_msg
	Loop until err_msg = ""
	If objExcel = "" THEN call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
	If err_msg <> "" THEN MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'ARRAY business----------------------------------------------------------------------------------------------------
'Sets up the array to store all the information for each client'
Dim EMPS_array ()
ReDim EMPS_array (9, 0)

'Sets constants for the array to make the script easier to read (and easier to code)'
Const clt_SSN         	= 1			'Each of the case numbers will be stored at this position'
Const memb_number		= 2
Const case_number       = 3
Const ref_status        = 4
Const EMPS_name         = 5	'ask ilse if these have to be the same'
Const error_reason		= 6
Const EMPS_update 	    = 7
Const excel_num			= 8
Const update_status		= 9

'Now the script adds all the clients on the excel list into an array for the appropriate county
excel_row = 2 're-establishing the row to start checking the members for
entry_record = 0

Do                                                            'Loops until there are no more cases in the Excel list
	
	MAXIS_case_number = objExcel.cells(excel_row, 3).Value
	MAXIS_case_number = trim(MAXIS_case_number)
	client_SSN  = objExcel.cells(excel_row, 4).Value		'Pulls the client's known information 
	client_SSN = replace(client_SSN, "-", "")
	name_of_EMPS = objExcel.cells(excel_row, 5).Value
	name_of_EMPS = trim(name_of_EMPS)
	If name_of_EMPS = "" THEN exit do
	'Adding client information to the array
	ReDim Preserve EMPS_array(9, entry_record)	'This resizes the array based on if the client is in the selected county
	EMPS_array (clt_SSN,     	entry_record) = client_SSN		'The client information is added to the array
	EMPS_array (case_number, 	entry_record) = MAXIS_case_number
	EMPS_array (ref_status,  	entry_record) = true 			'defaults to true
	EMPS_array (EMPS_name,    	entry_record) = name_of_EMPS
	EMPS_array (error_reason, 	entry_record) = ""
	EMPS_array (EMPS_update, 	entry_record) = true				'defaulting to true for now
	EMPS_array (memb_number, 	entry_record) = "01"				'defaults to 01 until it gets to PROG
	EMPS_array (excel_num, 		entry_record) = excel_row
	EMPS_array (update_status, 	entry_record) = ""
	entry_record = entry_record + 1			'This increments to the next entry in the array
	excel_row = excel_row + 1
	
	'blanking out variables
	client_SSN = ""
	MAXIS_case_number = ""
	name_of_EMPS = ""
Loop

If entry_record = 0 THEN script_end_procedure("No cases have been found on this list. The script wil now end.")

'Ensures that user is in current month
back_to_self
EMWriteScreen "________", 18, 43
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46

'Gathering info from MAXIS, and making the referrals and case notes if cases are found and active----------------------------------------------------------------------------------------------------
For item = 0 to UBound(EMPS_array, 2)
	MAXIS_case_number = EMPS_array(case_number, item)			
	client_SSN = EMPS_array(clt_SSN, item)
	
	If client_SSN <> "" THEN 
		'EMPS_array(EMPS_update, item) = False
		call navigate_to_MAXIS_screen("pers", "____")
		
		'changing the formating of the SSN from 123456789 to 123 45 6789 for STAT/MEMB
		If len(client_SSN) < 9 THEN
			EMPS_array(EMPS_update, item) = False
			EMPS_array(ref_status, item) = "Error"
			EMPS_array(error_reason, item) = "SSN not valid."		'Explanation for the rejected report'
		Elseif len(client_SSN) = 9 THEN 
			left_SSN = Left(client_SSN, 3)
			mid_SSN = mid(client_SSN, 4, 2)
			right_SSN = Right(client_SSN, 4)
			client_SSN = left_SSN & " " & mid_SSN & " " & right_SSN
		END IF 
		
		IF EMPS_array(ref_status, item) = True THEN 
		    EMWriteScreen left_SSN, 14, 36
		    EMWriteScreen mid_SSN, 14, 40
		    EMWriteScreen right_SSN, 14, 43
		    Transmit
		    msgbox "DSPL"
		    EMReadscreen DSPL_confirmation, 4, 2, 51
		    If DSPL_confirmation <> "DSPL" THEN 
		    	EMPS_array(EMPS_update, item) = False
		    	EMPS_array(ref_status, item) = "Error"
		    	EMPS_array(error_reason, item) = "Unable to find person in SSN search."		'Explanation for the rejected report'
		    Else 	
		    	
		    	EMWriteScreen "FS", 7, 22	'Selects FS as the program	
		    	Transmit		    	    'checking for an active case
		    	MAXIS_row = 10
		    	Do 
		    		EMReadscreen current_case, 7, MAXIS_row, 35
		    		If current_case = "Current" THEN
		    			EMReadscreen MAXIS_case_number, 8, MAXIS_row, 6
		    			MAXIS_case_number = trim(MAXIS_case_number) 
		    			EMPS_array(case_number, item) = MAXIS_case_number
		    			EMPS_array(EMPS_update, item) = true
		    			Exit do
		    		Else 
		    			MAXIS_row = MAXIS_row + 1
		    			If MAXIS_row = 20 THEN 
		    				PF8
		    				MAXIS_row = 10
		    			END IF
		    			EMReadScreen last_page_check, 21, 24, 2 
		    		END IF 
		    	LOOP until last_page_check = "THIS IS THE LAST PAGE" or last_page_check = "THIS IS THE ONLY PAGE"
		    	If EMPS_array(EMPS_update, item) = False THEN
		    		EMPS_array(EMPS_update, item) = False
		    		EMPS_array(ref_status, item) = "SNAP Inactive" 
				END IF 
		    END IF
		END IF 
	Else 
	 	EMPS_array(EMPS_update, item) = True
		needs_PMI = true
	End if
	msgbox "PROG"
	If EMPS_array(EMPS_update, item) = True THEN 	    'Checking the SNAP status 
	    Call navigate_to_MAXIS_screen("STAT", "PROG")
		EMReadscreen county_code, 2, 21, 23
		If county_code <> "27" THEN 
			EMPS_array(EMPS_update, item) = False
			EMPS_array(ref_status, item) = "Error"
			EMPS_array(error_reason, item) = "Not Hennepin County case, county code is: " & county_code	'Explanation for the rejected report'
		Else 
	        EMReadscreen SNAP_active, 4, 10, 74
	        If SNAP_active <> "ACTV" THEN 
	        	EMPS_array(EMPS_update, item) = False
	        	EMPS_array(ref_status, item) = "SNAP Inactive"
	        Else
			MsgBox "MEMB"
	        	Call navigate_to_MAXIS_screen("STAT", "MEMB")
				if needs_PMI = true THEN 
					row = 5
					HH_count = 0
					Do 
						EMReadScreen member_number, 2, row, 3
						HH_count = HH_count + 1
						transmit
						EMReadScreen MEMB_error, 5, 24, 2
					Loop until MEMB_error = "ENTER"
					If HH_count = 1 THEN 
						EMPS_array(memb_number, item) = member_number
						EMPS_array(EMPS_update, item) = True
					Else
						EMPS_array(EMPS_update, item) = False
						EMPS_array(ref_status, item) = "Error"
						EMPS_array(error_reason, item) = "Process manually, more than one person in HH & SSN not provided."	'Explanation for the rejected report'
					End if 
				Else 	
	        	    Do 
	        	    	EMReadscreen member_SSN, 11, 7, 42
		    	    	member_SSN = replace(member_SSN, " ", "")
	        	    	If member_SSN = EMPS_array(clt_SSN, item) THEN
	        	    		EMReadscreen member_number, 2, 4, 33
	        	    		EMPS_array(memb_number, item) = member_number
	        	    		EMPS_array(EMPS_update, item) = True
	        	    		exit do
	        	    	Else 
	        	    		transmit
							EMPS_array(EMPS_update, item) = False
		    	    	END IF
	        	    Loop until member_SSN = EMPS_array(clt_SSN, item) or MEMB_error = "ENTER"
				End if 
	        	msgbox "EMPS"
				IF EMPS_array(EMPS_update, item) = True THEN 
		    		Call navigate_to_MAXIS_screen("STAT", "EMPS")
					EMWriteScreen member_number, 20, 76				'enters member number
				    transmit
					PF9	
				    EMWriteScreen "X", 19, 25	
					transmit
				    'EMReadScreen other_box, 5, 4, 30		
					'IF other_box <> "Other"	 THEN err_msg "Unable to get into Provider information"
					CALL clear_line_of_text(6, 37)
					CALL clear_line_of_text(7, 37)
					CALL clear_line_of_text(8, 37)
					CALL clear_line_of_text(9, 37)
					CALL clear_line_of_text(10, 47)
					CALL clear_line_of_text(12, 37)
					CALL clear_line_of_text(13, 37)
				    If name_of_EMPS = "AVIVO BC" THEN 
				    	EMPS_array(EMPS_update, item) = True
						EMWriteScreen "HSPH.ESP.20268", 6, 37
						EMWriteScreen "AVIVO BROOKLYN CENTER", 7, 37
						EMWriteScreen "5701 SHINGLE CREEK PARKWAY", 8, 37
						EMWriteScreen "BROOKLYN CENTER", 9, 37
						EMWriteScreen "MN", 10, 37
						EMWriteScreen "55430", 10, 47
						EMWriteScreen "6127528900", 12, 39
						
					Elseif name_of_EMPS = "CAPI" THEN 
		   			 	EMPS_array(EMPS_update, item) = True
		   				EMWriteScreen "HSPH.ESP.20297", 6, 37
		   				EMWriteScreen "CAPI BROOKLYN CENTER", 7, 37
		   				EMWriteScreen "5930 BROOKLYN BLVD", 8, 37
		   				EMWriteScreen "BROOKLYN CENTER", 9, 37
		   				EMWriteScreen "MN", 10, 37
		   				EMWriteScreen "55429", 10, 47
		   				EMWriteScreen "6125883592", 12, 39
						
					Elseif name_of_EMPS = "HIRED" THEN 
    				   	EMPS_array(EMPS_update, item) = True
    					EMWriteScreen "HSPH.ESP.17HIR", 6, 37
    					EMWriteScreen "HIRED", 6, 37
    					EMWriteScreen "1701 EAST 79TH ST", 8, 37
    					EMWriteScreen "BLOOMINGTON,", 9, 37
    					EMWriteScreen "MN", 10, 37
    					EMWriteScreen "55425", 10, 47
    					EMWriteScreen "9528539100", 12, 39	
						    	
				    Elseif name_of_EMPS = "HIRED HENNEPIN NORTH" THEN 
				    	EMPS_array(EMPS_update, item) = True
						EMWriteScreen "HSPH.ESP.1HD10", 6, 37
						EMWriteScreen "HIRED HENNEPIN NORTH", 6, 37
						EMWriteScreen "7225 NORTHLAND DRIVE", 8, 37
						EMWriteScreen "BROOKLYN PARK", 9, 37
						EMWriteScreen "MN", 10, 37
						EMWriteScreen "55428", 10, 47
						EMWriteScreen "7632106200", 12, 39
						
				    Elseif name_of_EMPS = "NORTHPOINT" THEN 	
				    	EMPS_array(EMPS_update, item) = True
						EMWriteScreen "HSPH.ESP.NP027", 6, 37
						EMWriteScreen "NORTHPOINT HEALTH & WELLNESS", 6, 37
						EMWriteScreen "1315 PENN AVE NORTH", 8, 37
						EMWriteScreen "MINNEAPOLIS", 9, 37
						EMWriteScreen "MN", 10, 37
						EMWriteScreen "55411", 10, 47
						EMWriteScreen "6127670321", 12, 39
						
					Elseif name_of_EMPS = "AVIVO BLOOMINGTON" THEN 
					    EMPS_array(EMPS_update, item) = True
					    EMWriteScreen "HSPH.ESP.26AVO", 6, 37
					    EMWriteScreen "AVIVO BLOOMINGTON", 6, 37
					    EMWriteScreen "2626 EAST 82ND ST #370", 8, 37
					    EMWriteScreen "BLOOMINGTON", 9, 37
					    EMWriteScreen "MN", 10, 37
					    EMWriteScreen "55425", 10, 47
					    EMWriteScreen "6127528940", 12, 39
						
				 	Elseif name_of_EMPS = "RISE INC SOUTH" THEN 
						EMPS_array(EMPS_update, item) = True
						EMWriteScreen "HSPH.ESP.1RI50", 6, 37
						EMWriteScreen "RISE, INC SOUTH", 6, 37
						EMWriteScreen "3708 NICOLLET AVE SOUTH", 8, 37
						EMWriteScreen "MINNEAPOLIS", 9, 37
						EMWriteScreen "MN", 10, 37
						EMWriteScreen "55409", 10, 47
						EMWriteScreen "          ", 12, 39
					Else 
				    	EMPS_array(EMPS_update, item) = True
				    End if
				End if 
	        END IF
		End if 
		 
	   	start_a_blank_CASE_NOTE				
			If name_of_EMPS = "AVIVO BC" THEN 
				CALL write_variable_in_CASE_NOTE("ESP CASE TRANSFER TO AVIVO BROOKLYN CENTER DEED NO LONGER AN MFIP ESP 1/1/2018 20268 IS NEW ESP MFIP COORDINATION OFFICE")
			Elseif name_of_EMPS = "CAPI" THEN 
				CALL write_variable_in_CASE_NOTE ("ESP CASE TRANSFER TO CAPI BROOKLYN CENTER DEED NO LONGER AN MFIP ESP 1/1/2018 20297 IS NEW ESP MFIP COORDINATION OFFICE")
			Elseif name_of_EMPS = "HIRED" THEN 
				CALL write_variable_in_CASE_NOTE ("ESP CASE TRANSFER TO HIRED EAST BLOOMINGTON DEED NO LONGER AN MFIP ESP 1/1/2018 17HIR IS NEW ESP MFIP COORDINATION OFFICE")
			Elseif name_of_EMPS = "HIRED HENNEPIN NORTH" THEN 
				CALL write_variable_in_CASE_NOTE ("ESP CASE TRANSFER TO HIRED HENNEPIN NORTH DEED NO LONGER AN MFIP ESP 1/1/2018 1HD10 IS NEW ESP MFIP COORDINATION OFFICE")
			Elseif name_of_EMPS = "NORTHPOINT" THEN 
				CALL write_variable_in_CASE_NOTE ("ESP CASE TRANSFER TO NORTHPOINT HEALTH & WELLNESS DEED NO LONGER AN MFIP ESP 1/1/2018 NP027 IS NEW ESP MFIP COORDINATION OFFICE")
			Elseif name_of_EMPS = "AVIVO BLOOMINGTON" THEN 
				CALL write_variable_in_CASE_NOTE ("ESP CASE TRANSFER TO AVIVO BLOOMINGTON DEED NO LONGER AN MFIP ESP 1/1/2018 26AVO IS NEW ESP MFIP COORDINATION OFFICE")
			Elseif name_of_EMPS = "RISE INC SOUTH" THEN 
				CALL write_variable_in_CASE_NOTE("ESP CASE TRANSFER TO RISE, INC SOUTH DEED NO LONGER AN MFIP ESP 1/1/2018 1RI50 IS NEW ESP MFIP COORDINATION OFFICE")
		    END IF
		STATS_counter = STATS_counter + 1						'adds 1 count to the stats_counter
	END IF
Next 

'Updating the Excel spreadsheet based on what's happening in MAXIS----------------------------------------------------------------------------------------------------
For item = 0 to UBound(EMPS_array, 2)
	excel_row = EMPS_array(excel_num, item)
	objExcel.cells(excel_row, 3).Value = EMPS_array(case_number,	item)
	objExcel.cells(excel_row, 6).Value = EMPS_array(ref_status, 	item)
	objExcel.cells(excel_row, 7).Value = EMPS_array(update_status, 	item)
	objExcel.cells(excel_row, 8).Value = EMPS_array(error_reason, 	item)
Next 
	
STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)
script_end_procedure("Success! Review the spreadsheet for accuracy.")