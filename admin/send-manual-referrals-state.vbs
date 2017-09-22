'Required for statistical purposes===============================================================================
name_of_script = "BULK - SEND MANUAL REFERRAL FOR CBO.vbs "
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 120                     'manual run time in seconds
STATS_denomination = "C"       'C is for each Case
'END OF stats block==============================================================================================

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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("12/12/2016", "Added new BULK script that will send manual E & T referrals for cases that have been identified by DHS as partcipants working with CBO's (Community Based Organizations).", "Ilse Ferris, Hennepin County")
call changelog_update("12/12/2016", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
'Connects to BlueZone and establishing county name
EMConnect ""	
get_county_code

'dialog and dialog DO...Loop	
Do
	Do
			'The dialog is defined in the loop as it can change as buttons are pressed 
			BeginDialog CBO_referral_dialog, 0, 0, 266, 110, "CBO referral"
  				ButtonGroup ButtonPressed
    			PushButton 200, 45, 50, 15, "Browse...", select_a_file_button
    			OkButton 145, 90, 50, 15
    			CancelButton 200, 90, 50, 15
  				EditBox 15, 45, 180, 15, file_selection_path
  				GroupBox 10, 5, 250, 80, "Using the SEND MANUAL REFERRAL script"
  				Text 20, 20, 235, 20, "This script should be used when DHS provides your county with a list of recipeints that are working with CBO's and a manual referral is needed."
  				Text 15, 65, 230, 15, "Select the Excel file that contains the CBO information by selecting the 'Browse' button, and finding the file."
			EndDialog
			err_msg = ""
			Dialog CBO_referral_dialog
			cancel_confirmation
			If ButtonPressed = select_a_file_button then
				If file_selection_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
					objExcel.Quit 'Closing the Excel file that was opened on the first push'
					objExcel = "" 	'Blanks out the previous file path'
				End If
				call file_selection_system_dialog(file_selection_path, ".xlsx") 'allows the user to select the file'
			End If
			If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
			If err_msg <> "" Then MsgBox err_msg
		Loop until err_msg = ""
		If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
		If err_msg <> "" Then MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'ARRAY business----------------------------------------------------------------------------------------------------
'Sets up the array to store all the information for each client'
Dim CBO_array ()
ReDim CBO_array (8, 0)

'Sets constants for the array to make the script easier to read (and easier to code)'
Const clt_SSN         	= 1			'Each of the case numbers will be stored at this position'
Const memb_number		= 2
Const case_number       = 3
Const ref_status        = 4
Const CBO_name          = 5
Const excel_num			= 6
Const error_reason		= 7
Const make_referral 	= 8

'Now the script adds all the clients on the excel list into an array for the appropriate county
excel_row = 2 're-establishing the row to start checking the members for
entry_record = 0

Do                                                            'Loops until there are no more cases in the Excel list
	name_of_county = objExcel.cells(excel_row, 5).Value          're-establishing the name of the county for functions to use
	If name_of_county = "" then exit do
	name_of_county = trim(name_of_county)
	If name_of_county = county_name then 
		client_SSN  = objExcel.cells(excel_row, 4).Value		'Pulls the client's known information 
		client_SSN = replace(client_SSN, "-", "")
		MAXIS_case_number = objExcel.cells(excel_row, 3).Value
		MAXIS_case_number = trim(MAXIS_case_number)
		name_of_CBO = objExcel.cells(excel_row, 6).Value
		'Adding client information to the array
		ReDim Preserve CBO_array(8, entry_record)	'This resizes the array based on if the client is in the selected county
		CBO_array (clt_SSN,     	entry_record) = client_SSN		'The client information is added to the array
		CBO_array (case_number, 	entry_record) = MAXIS_case_number
		CBO_array (ref_status,  	entry_record) = true 			'defaults to true
		CBO_array (CBO_name,    	entry_record) = name_of_CBO
		CBO_array (excel_num, 		entry_record) = excel_row
		CBO_array (error_reason, 	entry_record) = ""
		CBO_array (make_referral, 	entry_record) = true				'defaulting to true for now
		CBO_array (memb_number, 	entry_record) = "01"				'defaults to 01 until it gets to PROG
		entry_record = entry_record + 1			'This increments to the next entry in the array
	End if
	excel_row = excel_row + 1
	'blanking out variables
	client_SSN = ""
	MAXIS_case_number = ""
	name_of_CBO = ""
Loop

If entry_record = 0 then script_end_procedure("No cases have been found on this list for your county. The script wil now end.")

'Ensures that user is in current month
back_to_self
EMWriteScreen "________", 18, 43
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46

'Gathering info from MAXIS, and making the referrals and case notes if cases are found and active----------------------------------------------------------------------------------------------------
For item = 0 to UBound(CBO_array, 2)
	MAXIS_case_number = CBO_array(case_number, item)			
	client_SSN = CBO_array(clt_SSN, item)
	
	If CBO_array(case_number, item) = "" then 
		CBO_array(make_referral, item) = False
		call navigate_to_MAXIS_screen("pers", "____")
		
		'changing the formating of the SSN from 123456789 to 123 45 6789 for STAT/MEMB
		If len(client_SSN) < 9 then
			CBO_array(make_referral, item) = False
			CBO_array(ref_status, item) = "Error"
			CBO_array(error_reason, item) = "SSN not valid."		'Explanation for the rejected report'
		Elseif len(client_SSN) = 9 then 
			left_SSN = Left(client_SSN, 3)
			mid_SSN = mid(client_SSN, 4, 2)
			right_SSN = Right(client_SSN, 4)
			client_SSN = left_SSN & " " & mid_SSN & " " & right_SSN
		END IF 
		
		IF CBO_array(ref_status, item) = True then 
		    EMWriteScreen left_SSN, 14, 36
		    EMWriteScreen mid_SSN, 14, 40
		    EMWriteScreen right_SSN, 14, 43
		    Transmit
		    
		    EMReadscreen DSPL_confirmation, 4, 2, 51
		    If DSPL_confirmation <> "DSPL" then 
		    	CBO_array(make_referral, item) = False
		    	CBO_array(ref_status, item) = "Error"
		    	CBO_array(error_reason, item) = "Unable to find person in SSN search."		'Explanation for the rejected report'
		    Else 	
		    	
		    	EMWriteScreen "FS", 7, 22	'Selects FS as the program	
		    	Transmit
		    	'chekcing for an active case
		    	MAXIS_row = 10
		    	Do 
		    		EMReadscreen current_case, 7, MAXIS_row, 35
		    		If current_case = "Current" then
		    			EMReadscreen MAXIS_case_number, 8, MAXIS_row, 6
		    			MAXIS_case_number = trim(MAXIS_case_number) 
		    			CBO_array(case_number, item) = MAXIS_case_number
		    			CBO_array(make_referral, item) = true
		    			Exit do
		    		Else 
		    			MAXIS_row = MAXIS_row + 1
		    			If MAXIS_row = 20 then 
		    				PF8
		    				MAXIS_row = 10
		    			END IF
		    			EMReadScreen last_page_check, 21, 24, 2 
		    		END IF 
		    	LOOP until last_page_check = "THIS IS THE LAST PAGE" or last_page_check = "THIS IS THE ONLY PAGE"
		    	If CBO_array(make_referral, item) = False then
		    		CBO_array(make_referral, item) = False
		    		CBO_array(ref_status, item) = "SNAP Inactive"
		    	END IF 
		    END IF
		END IF 
	END IF 
	
	If CBO_array(make_referral, item) = True then 
	    'Checking the SNAP status 
	    Call navigate_to_MAXIS_screen("STAT", "PROG")
	    EMReadscreen SNAP_active, 4, 10, 74
	    If SNAP_active <> "ACTV" then 
	    	CBO_array(make_referral, item) = False
	    	CBO_array(ref_status, item) = "SNAP Inactive"
	    Else
	    	Call navigate_to_MAXIS_screen("STAT", "MEMB")
	    	Do 
	    		EMReadscreen member_SSN, 11, 7, 42
				member_SSN = replace(member_SSN, " ", "")
	    		If member_SSN = CBO_array(clt_SSN, item) then
	    			EMReadscreen member_number, 2, 4, 33
	    			CBO_array(memb_number, item) = member_number
	    			CBO_array(make_referral, item) = True
	    			exit do
	    		Else 
	    			transmit
				END IF
	    		EMReadScreen MEMB_error, 5, 24, 2
	    	Loop until member_SSN = CBO_array (clt_SSN, item) or MEMB_error = "ENTER"
	    	IF member_SSN <> CBO_array (clt_SSN, item) then 
	    		CBO_array(make_referral, item) = False
				CBO_array(ref_status, item) = "Error"
	    		CBO_array(error_reason, item) = "Unable to find person's member number."	'Explanation for the rejected report'
	    	END IF 
	    END IF
		 	
		'Manual referral creation if banked months are used
		'Call navigate_to_MAXIS_screen("INFC", "WF1M")				'navigates to WF1M to create the manual referral'
		'EMWriteScreen "01", 4, 47									'this is the manual referral code that DHS has approved
		'EMWriteScreen "FS", 8, 46									'this is a program for ABAWD's for SNAP is the only option for banked months
		'EMWriteScreen CBO_array(memb_number, item), 8, 9							'enters member number
		'EMWriteScreen "Working with CBO: " & CBO_array(CBO_name, item), 17, 6		'enters notes for E & T regarding the name of the CBO  
		'EMWriteScreen "x", 8, 53																				'selects the ES provider
		'transmit																												'navigates to the ES provider selection screen
		'EMWriteScreen "x", 5, 9									'selects the 1st option'
		'transmit												'transmits back to the main WF1M
		'PF3														'saves referral
		'EMWriteScreen "Y", 11, 64								'Y to confirm save
		'transmit												'confirms saving the referral
		CBO_array(ref_status, item) = "Referral Made"
		STATS_counter = STATS_counter + 1						'adds 1 count to the stats_counter
	END IF
Next 

'Updating the Excel spreadsheet based on what's happening in MAXIS----------------------------------------------------------------------------------------------------
For item = 0 to UBound(CBO_array, 2)
	excel_row = CBO_array(excel_num, item)
	objExcel.cells(excel_row, 3).Value = CBO_array(case_number, item)
	objExcel.cells(excel_row, 7).Value = CBO_array(ref_status, item)
	objExcel.cells(excel_row, 8).Value = CBO_array(error_reason, item)
Next 
	
STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)
script_end_procedure("Success! The cases that have active SNAP have a referral that's been made!")