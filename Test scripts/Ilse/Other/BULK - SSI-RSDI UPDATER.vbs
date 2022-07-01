'Required for statistical purposes==========================================================================================
name_of_script = "BULK - SSI-RSDI UPDATER.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 90                	'manual run time in seconds
STATS_denomination = "C"       			'C is for each CASE
'END OF stats block=========================================================================================================

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

''CHANGELOG BLOCK ===========================================================================================================
''Starts by defining a changelog array
'changelog = array()
'
''INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
''Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
'call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")
'
''Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
'changelog_display
''END CHANGELOG BLOCK =======================================================================================================

'THIS SCRIPT IS BEING USED IN A WORKFLOW SO DIALOGS ARE NOT NAMED
'DIALOGS MAY NOT BE DEFINED AT THE BEGINNING OF THE SCRIPT BUT WITHIN THE SCRIPT FILE

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""
renewal_date = CM_plus_1_mo & "/" & CM_plus_1_yr
'file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\UNEA Updater\11-17 ssi rsdi renewals report.xlsx"

'dialog and dialog DO...Loop
Do
	Do
		'The dialog is defined in the loop as it can change as buttons are pressed
		Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 266, 135, "SSI-RSDI Updater"
	  		ButtonGroup ButtonPressed
	  		PushButton 200, 45, 50, 15, "Browse...", select_a_file_button
	  		DropListBox 185, 90, 65, 15, "Select one..."+chr(9)+"Send SVES"+chr(9)+"Update Cases", action_type
	  		ButtonGroup ButtonPressed
	  		OkButton 145, 115, 50, 15
	  		CancelButton 200, 115, 50, 15
	  		GroupBox 10, 5, 250, 105, "Using the SSI-RSDI Updater script:"
	  		Text 20, 20, 235, 20, "This script should be used when a list of recipeints on SSI and/or RSDI has a renewal, or another need for updating this information is needed."
	  		Text 15, 65, 230, 15, "Select the Excel file that contains the PMI inforamtion by selecting the 'Browse' button, and finding the file."
	  		EditBox 15, 45, 180, 15, file_selection_path
	  		Text 80, 95, 100, 10, "Select the script's action type:"
		EndDialog

		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		If ButtonPressed = select_a_file_button then
			If file_selection_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
				objExcel.Quit 'Closing the Excel file that was opened on the first push'
				objExcel = "" 	'Blanks out the previous file path'
			End If
			call file_selection_system_dialog(file_selection_path, ".xlsx") 'allows the user to select the file'
		End If
		If action_type = "Select one..." then err_msg = err_msg & vbNewLine & "Select the type of action you'd like the script to take."
		If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data."
		If err_msg <> "" Then MsgBox err_msg
	Loop until err_msg = ""
	If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
	If err_msg <> "" Then MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'ARRAY business----------------------------------------------------------------------------------------------------
'Sets up the array to store all the information for each client'
Dim PMI_array ()
ReDim PMI_array (6, 0)

'Sets constants for the array to make the script easier to read (and easier to code)'
Const clt_PMI       	= 1			'Each of the case numbers will be stored at this position'
Const case_number   	= 2
Const SVES_status   	= 3
Const clt_SSN 			= 4
Const failure_reason	= 5
Const member_num		= 6

'Now the script adds all the client information  on the excel list into an array
excel_row = 2 're-establishing the row to start checking the members for
entry_record = 0

Do                                                            'Loops until there are no more cases in the Excel list
	'PMI
	client_PMI = objExcel.cells(excel_row, 2).Value          're-establishing the name of the county for functions to use
	If client_PMI = "" then exit do
	'trims off all the zeros to ensure uniformity with PMI on the MEMB panel
	Do
		if left(client_PMI, 1) = "0" then client_PMI = right(client_PMI, len(client_PMI) -1)
	Loop until left(client_PMI, 1) <> "0"
	client_PMI = trim(client_PMI)

	'case number & member number
	MAXIS_case_number = objExcel.cells(excel_row, 1).Value
	MAXIS_case_number = trim(MAXIS_case_number)

	member_number = objExcel.cells(excel_row, 3).Value
	member_number = trim(member_number)

	'Adding client information to the array
	ReDim Preserve PMI_array(6, entry_record)	'This resizes the array based on if the client is in the selected county
	PMI_array (clt_PMI,     	entry_record) = client_PMI			'PMI
	PMI_array (case_number, 	entry_record) = MAXIS_case_number	'case number
	PMI_array (SVES_status,  	entry_record) = true 				'defaults to true
	PMI_array (member_num, 		entry_record) = member_number		'member number

	entry_record = entry_record + 1			'This increments to the next entry in the array
	excel_row = excel_row + 1
	STATS_counter = STATS_counter + 1
	'blanking out variables
	client_PMI = ""
	MAXIS_case_number = ""
	member_number = ""
Loop
'msgbox entry_record
If entry_record = 0 then script_end_procedure("No cases have been found on this list for your county. The script will now end.")

'Closes the excel file
'objExcel.Quit
If action_type = "Send SVES" then
    'Gathering info from MAXIS, and making the referrals and case notes if cases are found and active----------------------------------------------------------------------------------------------------
    For item = 0 to UBound(PMI_array, 2)
    	MAXIS_case_number = PMI_array(case_number, item)
    	client_PMI = PMI_array(clt_PMI, item)

    	Call check_for_MAXIS(False)		'Makes sure we're in MAXIS
    	Call navigate_to_MAXIS_screen("stat", "memb") 'Goes to MEMB to get info

    	'Checking for PRIV cases.
    	EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case, script will end.
    	IF priv_check = "PRIVIL" THEN

    		PMI_array(SVES_status, item) = False
    		PMI_array(failure_reason, item) = "Case is a privliged case."
    		'msgbox "PRIV case" & vbcr & MAXIS_case_number
    	ELSE
			EMReadScreen county_code, 2, 21, 23
			If county_code <> "27" then
				PMI_array(SVES_status, item) = False
				PMI_array(failure_reason, item) = "Not a Hennepin County case."
			Else

    	        Do
    	        	EMReadscreen PMI_confirmation, 8, 4, 46
    	        	PMI_confirmation = trim(PMI_confirmation)
    	        	If PMI_confirmation <> client_PMI then
    	        		transmit
    	        		PMI_array(SVES_status, item) = FALSE
    	        	Else
    	        		'gather SSN info and adding it to the array
						EMReadScreen member_number, 2, 4, 33
    	        		EMReadScreen SSN1, 3, 7, 42
    	        		EMReadScreen SSN2, 2, 7, 46
    	        		EMReadScreen SSN3, 4, 7, 49
    	        		client_SSN = SSN1 & SSN2 & SSN3
    	        		PMI_array(clt_SSN, item) = client_SSN
						PMI_array(member_num, item) = member_number
    		    		PMI_array(SVES_status, item) = True
    	        	END IF
    	        	EMReadScreen MEMB_error, 5, 24, 2
    	        Loop until PMI_confirmation = PMI_array (clt_PMI, item) or MEMB_error = "ENTER"
    	        If PMI_array(SVES_status, item) = FALSE then
					PMI_array(failure_reason, item) = "PMI did not match on STAT/MEMB"
					PMI_array(member_num, item) = ""
				End if
			End if
    	End if
    		'blanking out variables
    	client_PMI = ""
    	MAXIS_case_number = ""
    next

    'Sending the SVES/QURY
    For item = 0 to UBound(PMI_array, 2)

    	IF PMI_array(SVES_status, item) = True then

    		MAXIS_case_number = PMI_array(case_number, item)
    		client_PMI = PMI_array(clt_PMI, item)
    		client_SSN = PMI_array(clt_SSN, item)
    		'establishing values from the array to write into INFC/SVES
    	    Call navigate_to_MAXIS_screen("INFC", "SVES")
			'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
			EMReadScreen agreement_check, 9, 2, 24
			IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

      	    EMWriteScreen client_SSN, 			4, 68
    	    EMWriteScreen client_PMI, 			5, 68
      	    EMWriteScreen "qury",  				20, 70
      	    transmit 'Now we will enter the QURY screen to type the case number.
    	    EMWriteScreen MAXIS_case_number, 	11, 38
    	    EMWriteScreen "y", 					14, 38
     	    transmit  'Now it sends the SVES.
    	    EMReadScreen duplicate_SVES, 	    7, 24, 2
    	    If duplicate_SVES = "WARNING" then transmit
    		EMReadScreen confirm_SVES, 			6, 24, 2
    		if confirm_SVES = "RECORD" then
    	    	PMI_array(SVES_status, item) = True
    			PMI_array(failure_reason, item) = "QURY sent."
    	    Else
    	    	PMI_array(SVES_status, item) = False
    	    	PMI_array(failure_reason, item) = "Attempt to send QURY failed."
    	    END IF
    	END IF
    Next

    For item = 0 to UBound(PMI_array, 2)
    	'establishing values from the array to write into case notes
    	IF PMI_array(SVES_status, item) = True then
    		MAXIS_case_number = PMI_array(case_number, item)
    		client_PMI = PMI_array(clt_PMI, item)

    		start_a_blank_CASE_NOTE		'Now it case notes
    		call write_variable_in_case_note("SVES/QURY sent for PMI# " & client_PMI & " for " & renewal_date & " Recertification")
    		call write_variable_in_case_note("* Used SSN for QURY.")
    		call write_variable_in_case_note("---")
    		call write_variable_in_case_note("QURY sent using script by I. Ferris, QI team")
    	END IF
    next

    excel_row = 2
	ObjExcel.columns(3).NumberFormat = "@" 		'formatting as text
    ObjExcel.Cells(1, 17).Value = "NOTES"
    objExcel.Cells(1, 1).Font.Bold = True		'bold font'

    For item = 0 to UBound(PMI_array, 2)
    	'establishing values from the array to write into case notes
		 ObjExcel.Cells(excel_row, 3).Value = PMI_array (member_num, item)	'Adding the reson why SVES/QURY wasn't sent.
    	 ObjExcel.Cells(excel_row, 17).Value = PMI_array (failure_reason, item)	'Adding the reson why SVES/QURY wasn't sent.
    	excel_row = excel_row + 1
    Next
End if

'If action_type = "Update Cases" then
'
'End if

STATS_counter = STATS_counter - 1
script_end_procedure("Sucess! SVES/QURY has been sent on all cases except for those on the newly created Excel spreadsheet. Please review spreadsheet, and process manually if necessary.")
