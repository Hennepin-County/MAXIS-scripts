'Required for statistical purposes===============================================================================
name_of_script = "BULK - SEND CBO MANUAL REFERRALS.vbs"
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


'Sam's PowerShell file function that replaces old file selection function==========
Function file_selection_dialog()

'creates a Windows Script Host object
Set Fshell = CreateObject("WScript.Shell")

' creates a long string of powershell commands that will be executed
shellCmd = "powershell -NoProfile -NonInteractive -WindowStyle Hidden -command " & "Add-Type -AssemblyName System.Windows.Forms; " & _
            "$dlg = New-Object System.Windows.Forms.OpenFileDialog; " & _           
           "$dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop'); " & _
           "$dlg.Filter = 'Excel files (*.xlsx)|*.xlsx'; " & _ 
           "$dlg.ShowDialog() | Out-Null; " & _
           "$dlg.FileName; "

' Sets a variable of the file path selected from the PowerShell script run.
file_selection_path = Fshell.Exec(shellCmd).StdOut.ReadLine

end Function

'CASE/NOTE full routine ==========================================================================================
	
function create_referral_case_note()
	Call start_a_blank_CASE_NOTE()
	Call write_variable_in_CASE_NOTE("***SNAP E & T Referral Processed for MEMB " & member_number & " " & wreg_tlr_status)
	Call write_variable_in_CASE_NOTE("===================")
	Call write_variable_in_CASE_NOTE("Client referral sent through WF1M on " & date_of_today)
	Call write_bullet_and_variable_in_CASE_NOTE("Client is working with the following CBO", CBO_array(CBO_name, CBO_arrays))
	Call write_bullet_and_variable_in_CASE_NOTE("Client's listed STAT/WREG codes are", WREG_codes)

	Call write_variable_in_CASE_NOTE("===================")
	Call write_variable_in_CASE_NOTE("This CASE/NOTE was automatically generated through the bulk CBO referral script")
end function


'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("08/31/2026", "Fixed spaces causing array errors; fixed typos; updated code to use current FuncLib functions. Also added case note functionality.", "Sam Begley-May, Hennepin County")
call changelog_update("07/28/2018", "Fixed bug that was preventing output of ABAWD status. Also cleaned up code in the dialog handling.", "Ilse Ferris, Hennepin County")
call changelog_update("07/28/2017", "Added enhancement to support cases with case number instead of SSN.", "Ilse Ferris, Hennepin County")
call changelog_update("05/08/2017", "Added new BULK script that will send manual E & T referrals for cases that have been identified by E & T as partcipants working with CBO's (Community Based Organizations).", "Ilse Ferris, Hennepin County")
call changelog_update("12/12/2016", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
'Connects to BlueZone and establishing county name
EMConnect ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 110, "CBO referral"
    ButtonGroup ButtonPressed
    PushButton 200, 45, 50, 15, "Browse...", select_a_file_button
    OkButton 145, 90, 50, 15
    CancelButton 200, 90, 50, 15
    EditBox 15, 45, 180, 15, file_selection_path
    GroupBox 10, 5, 250, 80, "Using the SEND MANUAL REFERRAL script"
    Text 20, 20, 235, 20, "This script should be used when E & T provides you with a list of recipeints that are working with CBO's and a manual referral is needed. "
    Text 15, 65, 230, 15, "Select the Excel file that contains the CBO information by selecting the 'Browse' button, and finding the file."
EndDialog

'dialog and dialog DO...Loop
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
    	Dialog Dialog1
    	cancel_without_confirmation
    	If ButtonPressed = select_a_file_button then call file_selection_dialog()
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'ARRAY business----------------------------------------------------------------------------------------------------
'Sets up the array to store all the information for each client'
Dim CBO_array()
ReDim CBO_array(9, 0)

'Sets constants for the array to make the script easier to read (and easier to code)'
Const clt_SSN         	= 1			'Each of the case numbers will be stored at this position'
Const memb_number		= 2
Const case_number       = 3
Const ref_status        = 4
Const CBO_name          = 5
Const error_reason		= 6
Const make_referral 	= 7
Const excel_num			= 8
Const ABAWD_status		= 9

'Now the script adds all the clients on the excel list into an array for the appropriate county
excel_row = 2 're-establishing the row to start checking the members for
entry_record = 0

Do                                                            'Loops until there are no more cases in the Excel list

	MAXIS_case_number = objExcel.cells(excel_row, 3).Value
	MAXIS_case_number = trim(MAXIS_case_number)
	client_SSN  = objExcel.cells(excel_row, 4).Value		'Pulls the client's known information
	client_SSN = replace(client_SSN, "-", "")
    client_SSN = replace(client_SSN, " ", "")
	name_of_CBO = objExcel.cells(excel_row, 5).Value
	name_of_CBO = trim(name_of_CBO)
	If name_of_CBO = "" then exit do
	'Adding client information to the array
	ReDim Preserve CBO_array(9, entry_record)	'This resizes the array based on if the client is in the selected county
	CBO_array(clt_SSN,     	entry_record) = client_SSN		'The client information is added to the array
	CBO_array(case_number, 	entry_record) = MAXIS_case_number
	CBO_array(ref_status,  	entry_record) = true 			'defaults to true
	CBO_array(CBO_name,    	entry_record) = name_of_CBO
	CBO_array(error_reason, 	entry_record) = ""
	CBO_array(make_referral, 	entry_record) = true				'defaulting to true for now
	CBO_array(memb_number, 	entry_record) = "01"				'defaults to 01 until it gets to PROG
	CBO_array(excel_num, 		entry_record) = excel_row
	CBO_array(ABAWD_status, 	entry_record) = ""
	entry_record = entry_record + 1			'This increments to the next entry in the array
	excel_row = excel_row + 1

	'blanking out variables
	client_SSN = ""
	MAXIS_case_number = ""
	name_of_CBO = ""
Loop

If entry_record = 0 then script_end_procedure("No cases have been found on this list. The script wil now end.")

'assigning current month and year for MAXIS navigation footer months; establishing current day for referral date
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
CM_day = right("0" &             DatePart("d",           date                             ), 2)
date_of_today = CM_mo & "/" & CM_day & "/" & CM_yr

'creating tlr status variable for case note title later in script
Dim wreg_tlr_status
If CBO_array(ABAWD_status, CBO_arrays) = "Exempt" then
	wreg_tlr_status = "[Voluntary Participation]"
else
	wreg_tlr_status = "[Mandatory Participation]"
End If


'Starting on SELF to avoid an error that can mess up the notes generated into the Excel sheet -------------------------------------------
Call back_to_SELF()

'Gathering info from MAXIS, and making the referrals and case notes if cases are found and active----------------------------------------------------------------------------------------------------
For CBO_arrays = 0 to UBound(CBO_array, 2)
	MAXIS_case_number = CBO_array(case_number, CBO_arrays)
	client_SSN = CBO_array(clt_SSN, CBO_arrays)

	If client_SSN <> "" then
		CBO_array(make_referral, CBO_arrays) = False
		call navigate_to_MAXIS_screen("pers", clear_line_of_text(18, 43))

		'changing the formating of the SSN from 123456789 to 123 45 6789 for STAT/MEMB 
		If len(client_SSN) < 9 then
			CBO_array(make_referral, CBO_arrays) = False
			CBO_array(ref_status, CBO_arrays) = "Error"
			CBO_array(error_reason, CBO_arrays) = "SSN in spreadsheet is not a 9-digit number."		'Explanation for the rejected report'
		Elseif len(client_SSN) = 9 then
			left_SSN = Left(client_SSN, 3)
			mid_SSN = mid(client_SSN, 4, 2)
			right_SSN = Right(client_SSN, 4)
			client_SSN = left_SSN & " " & mid_SSN & " " & right_SSN
		END IF

		IF CBO_array(ref_status, CBO_arrays) = True then
		    EMWriteScreen left_SSN, 14, 36
		    EMWriteScreen mid_SSN, 14, 40
		    EMWriteScreen right_SSN, 14, 43
		    Transmit
		    EMReadscreen DSPL_confirmation, 4, 2, 51
		    If DSPL_confirmation <> "DSPL" then
		    	CBO_array(make_referral, CBO_arrays) = False
		    	CBO_array(ref_status, CBO_arrays) = "Error"
		    	CBO_array(error_reason, CBO_arrays) = "Unable to find person and case - this can be due to multiple PMI records existing for one SSN; or no results were found with the SSN."		'Explanation for the rejected report'
		    Else
		    	EMWriteScreen "FS", 7, 22	'Selects FS as the program
		    	Transmit
		    	'checking for an active case
		    	MAXIS_row = 10
		    	Do
		    		EMReadscreen current_case, 7, MAXIS_row, 35
		    		If current_case = "Current" then
		    			EMReadscreen MAXIS_case_number, 8, MAXIS_row, 6
		    			MAXIS_case_number = trim(MAXIS_case_number)
		    			CBO_array(case_number, CBO_arrays) = MAXIS_case_number
		    			CBO_array(make_referral, CBO_arrays) = true
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
		    	If CBO_array(make_referral, CBO_arrays) = False then
		    		CBO_array(make_referral, CBO_arrays) = False
		    		CBO_array(ref_status, CBO_arrays) = "SNAP Inactive"
				END IF
		    END IF
		END IF
	Else
	 	CBO_array(make_referral, CBO_arrays) = True
		needs_PMI = true
	End if

	If CBO_array(make_referral, CBO_arrays) = True then
	    'Checking the SNAP status
	    Call navigate_to_MAXIS_screen_review_PRIV("STAT", "PROG", is_this_priv)
		If is_this_priv = True then
			CBO_array(make_referral, CBO_arrays) = False
			CBO_array(ref_status, CBO_arrays) = "Error"
			CBO_array(error_reason, CBO_arrays) = "This case has PRIV status and was not updated."	'Explanation for the rejected report'
		Else
			EMReadscreen county_code, 2, 21, 23
			If county_code <> "27" then
				CBO_array(make_referral, CBO_arrays) = False
				CBO_array(ref_status, CBO_arrays) = "Error"
				CBO_array(error_reason, CBO_arrays) = "Not Hennepin County case, county code is: " & county_code	'Explanation for the rejected report'
			Else
				EMReadscreen SNAP_active, 4, 10, 74
				If SNAP_active <> "ACTV" then
					CBO_array(make_referral, CBO_arrays) = False
					CBO_array(ref_status, CBO_arrays) = "SNAP Inactive"
				Else
					Call navigate_to_MAXIS_screen("STAT", "MEMB")
					if needs_PMI = true then
						row = 5
						HH_count = 0
						Do
							EMReadScreen member_number, 2, row, 3
							HH_count = HH_count + 1
							transmit
							EMReadScreen MEMB_error, 5, 24, 2
						Loop until MEMB_error = "ENTER"
						If HH_count = 1 then
							CBO_array(memb_number, CBO_arrays) = member_number
							CBO_array(make_referral, CBO_arrays) = True
						Else 
							CBO_array(make_referral, CBO_arrays) = False 
							CBO_array(ref_status, CBO_arrays) = "Error"
							CBO_array(error_reason, CBO_arrays) = "Process manually, more than one person in HH & SSN not provided."	'Explanation for the rejected report'
						End if
						Else
						Do
							EMReadscreen member_SSN, 11, 7, 42
							member_SSN = replace(member_SSN, " ", "")
							If member_SSN = CBO_array(clt_SSN, CBO_arrays) then
								EMReadscreen member_number, 2, 4, 33
								CBO_array(memb_number, CBO_arrays) = member_number
								CBO_array(make_referral, CBO_arrays) = True
								exit do
							Else
								transmit
								CBO_array(make_referral, CBO_arrays) = False
							END IF
							EMReadScreen MEMB_error, 5, 24, 2
						Loop until member_SSN = CBO_array(clt_SSN, CBO_arrays) or MEMB_error = "ENTER"
					End if

					'STAT WREG PORTION
					Call navigate_to_MAXIS_screen("STAT", "WREG")
					EMWriteScreen member_number, 20, 76				'enters member number
					transmit
					EMReadScreen fset_code, 2, 8, 50
					EMReadScreen abawd_code, 2, 13, 50
					WREG_codes = fset_code & "-" & abawd_code
					If WREG_codes = "30-11" then
						CBO_array(make_referral, CBO_arrays) = True
						CBO_array(ABAWD_status, CBO_arrays) = "Mandatory - 2nd Set"
					Elseif WREG_codes = "30-10" then
						CBO_array(make_referral, CBO_arrays) = True
						CBO_array(ABAWD_status, CBO_arrays) = "Mandatory - ABAWD"
					Elseif WREG_codes = "30-13" then
						CBO_array(make_referral, CBO_arrays) = True
						CBO_array(ABAWD_status, CBO_arrays) = "Mandatory - Banked Months"
					Else
						CBO_array(make_referral, CBO_arrays) = True
						CBO_array(ABAWD_status, CBO_arrays) = "Exempt"
					End if
					If CBO_array(make_referral, CBO_arrays) = True then 	'if a referral is made, write the date for the "SNAP E&T Referral Date" field. Hennepin County now requires this field to be filled in if a referral is made, even for voluntary participants
						PF9
						EMWriteScreen CM_mo, 9, 50
						EMWriteScreen CM_day, 9, 53
						EMWriteScreen CM_yr, 9, 56
						transmit
					else 
						PF3
						CBO_array(ref_status, CBO_arrays) = "Error"
						CBO_array(error_reason, CBO_arrays) = "No referral is listed for this case"
						CBO_array(make_referral, CBO_arrays) = false
					End If
				END IF
			End if
		End if

		If CBO_array(make_referral, CBO_arrays) = True then
		    'Manual referral creation 
			Call navigate_to_MAXIS_screen("INFC", "WF1M")				'navigates to WF1M to create the manual referral'
		    EMWriteScreen "05", 4, 47									'this is the manual referral code re:Ilse/Sam meeting 7/2026
		    EMWriteScreen "FS", 8, 46									'this is a program for ABAWD's for SNAP is the only option for banked months
		    EMWriteScreen CBO_array(memb_number, CBO_arrays), 8, 9							'enters member number
		    EMWriteScreen "Working with CBO: " & CBO_array(CBO_name, CBO_arrays), 17, 6		'enters notes for E & T regarding the name of the CBO
		    EMWriteScreen "X", 8, 53																				'selects the ES provider
		    transmit																												'navigates to the ES provider selection screen
		    EMWriteScreen "X", 5, 9									'selects the 1st option'
		    transmit												'transmits back to the main WF1M
		    PF3														'saves referral
		    EMWriteScreen "Y", 11, 64								'Y to confirm save
		    transmit												'confirms saving the referral
		    CBO_array(ref_status, CBO_arrays) = "Referral Made"
		    STATS_counter = STATS_counter + 1						'adds 1 count to the stats_counter
			Call create_referral_case_note()
		Elseif CBO_array(make_referral, CBO_arrays) = False And is_this_PRIV = false then
			CBO_array(ref_status, CBO_arrays) = "Error"
			CBO_array(error_reason, CBO_arrays) = "An error occurred while trying to make the referral and CASE/NOTE. Review the case manually."
		End If
	END IF
Next



'Updating the Excel spreadsheet based on what's happening in MAXIS----------------------------------------------------------------------------------------------------
For CBO_arrays = 0 to UBound(CBO_array, 2)
	excel_row = CBO_array(excel_num, CBO_arrays)
	objExcel.cells(excel_row, 3).Value = CBO_array(case_number,		CBO_arrays)
	objExcel.cells(excel_row, 6).Value = CBO_array(ref_status, 		CBO_arrays)
	objExcel.cells(excel_row, 7).Value = CBO_array(ABAWD_status, 	CBO_arrays)
	objExcel.cells(excel_row, 8).Value = CBO_array(error_reason, 	CBO_arrays)

	'wrapping text on the Notes column so it's actually readable
	objExcel.range("H:H").WrapText = True
Next




STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)
script_end_procedure("Success! Review the spreadsheet for accuracy. Some cases may not have had a referral made.")


'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------7/31/26
'--Tab orders reviewed & confirmed----------------------------------------------n/a
'--Mandatory fields all present & Reviewed--------------------------------------
'--All variables in dialog match mandatory fields-------------------------------7/31/26
'Review dialog names for content and content fit in dialog----------------------
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------
'--Include script category and name somewhere on first dialog-------------------
'--Create a button to reference instructions------------------------------------??(talk to Dave and see if we need instructions for this one)
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------
'--CASE:NOTE Header doesn't look funky------------------------------------------7/31/26
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------n/a (BULK script will close after case note)
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------7/31/26
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------n/a
'--MAXIS_background_check reviewed (if applicable)------------------------------n/a
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------7/31/26
'--script_end_procedures (w/ or w/o error messaging)----------------------------
'--BULK - review output of statistics and run time/count (if applicable)--------
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------7/31/26
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------
'--BULK - remove 1 incrementor at end of script reviewed------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------7/31/26
'--Remove testing code/unnecessary code-----------------------------------------
'--Review/update SharePoint instructions----------------------------------------
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------
