'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - HG BGTX.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = "420"                'manual run time in seconds
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
call changelog_update("03/17/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'dialog and dialog DO...Loop	
Do
	Do
			'The dialog is defined in the loop as it can change as buttons are pressed 
			BeginDialog HG_issuance_dialog, 0, 0, 266, 125, "HG expansion background initiation dialog"
  				ButtonGroup ButtonPressed
    			PushButton 200, 60, 50, 15, "Browse...", select_a_file_button
    			OkButton 145, 105, 50, 15
    			CancelButton 200, 105, 50, 15
  				EditBox 15, 60, 180, 15, HG_path
  				GroupBox 10, 5, 250, 95, "Using the script"
  				Text 15, 20, 235, 25, "This script should be used when DHS provides your county with a list of recipeints that are eligible for the HG expansion. It will determine if a case is appropriate for the case to be run through background or not."
  				Text 15, 80, 230, 15, "Select the Excel file that contains the HG inforamtion by selecting the 'Browse' button, and finding the file."
			EndDialog

			err_msg = ""
			Dialog HG_issuance_dialog
			cancel_confirmation
			If ButtonPressed = select_a_file_button then
				If HG_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
					objExcel.Quit 'Closing the Excel file that was opened on the first push'
					objExcel = "" 	'Blanks out the previous file path'
				End If
				call file_selection_system_dialog(HG_path, ".xlsx") 'allows the user to select the file'
			End If
			If HG_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
			If err_msg <> "" Then MsgBox err_msg
		Loop until err_msg = ""
		If objExcel = "" Then call excel_open(HG_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
		If err_msg <> "" Then MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

issuance_month = "03/17"
excel_row = 2 	're-establishing the row to start checking - excel row 2 is the 1st row to start searching for the cases.
 
Do                                                            'Loops until there are no more cases in the Excel list
	MAXIS_case_number = objExcel.cells(excel_row, 1).Value          
	HG_issuance_excel = objExcel.cells(excel_row, 4).Value 
        
	MAXIS_case_number = trim(MAXIS_case_number)
    HG_issuance_excel = trim(HG_issuance_excel)
    If MAXIS_case_number = "" then exit do
    
    If HG_issuance_excel = "110" then
        issuance_found = True 
    Else 
        issuance_found = False
        Call navigate_to_MAXIS_screen("STAT", "PROG")
        EMReadScreen MFIP_prog_1_check, 2, 6, 67		'checking for an active MFIP case
        EMReadScreen MFIP_status_1_check, 4, 6, 74
        EMReadScreen MFIP_prog_2_check, 2, 6, 67		'checking for an active MFIP case
        EMReadScreen MFIP_status_2_check, 4, 6, 74

        'Logic to determine if MFIP is active
        If MFIP_prog_1_check = "MF" Then
            If MFIP_status_1_check = "ACTV" Then MFIP_ACTIVE = TRUE
        ElseIf MFIP_prog_2_check = "MF" Then
            If MFIP_status_2_check = "ACTV" Then MFIP_ACTIVE = TRUE
        End If
            
        'Only looks for SNAP if MFIP is not active
        If MFIP_ACTIVE <> TRUE Then
            objExcel.cells(excel_row, 4).value = ""
            objExcel.cells(excel_row, 4).value = "N/A"
            objExcel.cells(excel_row, 5).value = "MFIP INACTIVE 03/17"
        Else
	        Call navigate_to_MAXIS_screen("MONY", "INQD")
            
	     	row = 6				'establishing the row to start searching for issuance'
	    	Do 
	    		EMReadScreen housing_grant, 2, row, 19		'searching for housing grant issuance
	    		IF housing_grant = "HG" then
	    			'reading the housing grant information
	    			EMReadScreen HG_amt_issued, 3, row, 40
	    			EMReadScreen HG_month, 2, row, 73
	    			EMReadScreen HG_year, 2, row, 79
	    			INQD_issuance_month = HG_month & "/" & HG_year		'creates a new varible for HG month and year
	    			If issuance_month = INQD_issuance_month then 		'if the issuance found matches the issuance month then
	    				HG_amt_issued = trim(HG_amt_issued)				'trims the HG amt issued variable
                        objExcel.cells(excel_row, 4).value = ""
                        objExcel.cells(excel_row, 4).value = HG_amt_issued
                        issuance_found = True 
                        exit do
                    End if 
	    		END IF
                row = row + 1
	    	Loop until row = 18
	    End if 	
    End if 
			         
    If issuance_found = False then 
        Call navigate_to_MAXIS_screen("STAT", "MONT")
        EMReadScreen mont_ID, 1, 2, 73
        If mont_ID = "0" then 
            Call write_value_and_transmit("BGTX", 20, 71)
            EMReadScreen BGTX_error_msg, 8, 24, 2
            IF BGTX_error_msg <> "YOU HAVE" then 
                transmit
            Else 
                objExcel.cells(excel_row, 5).value = "Unable to update case."  
            End if  
        else 
            objExcel.cells(excel_row, 5).value = "Case is monthly reporter. Did not BGTX."     
        End if 
    End if 
            
    excel_row = excel_row + 1
	STATS_counter = STATS_counter + 1		 'adds one instance to the stats counter
	MAXIS_case_number = ""
LOOP UNTIL objExcel.Cells(excel_row, 1).value = ""	'looping until the list of cases to check for recert is complete

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Please review the list generated.")