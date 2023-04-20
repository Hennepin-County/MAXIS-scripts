'STATS GATHERING=============================================================================================================
name_of_script = "TYPE - PROJECT NOOB SCRIPT.vbs"       'REPLACE TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the anctual manualtime based on time study
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomonation appliable to your script.
'END OF stats block==========================================================================================================

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

'THE SCRIPT==================================================================================================================
EMConnect "" 'Connects to BlueZone

'FIRST DIALOG COLLECTING CASE & MONTH/YEAR===========================================================================
	Dialog1 = "" 'blanking out dialog name ' ' NOTES: First Dialog to capture Case & Month/Year
		'Add dialog here: Add the dialog just before calling the dialog below unless you need it in the dialog due to using COMBO Boxes or other looping reasons. Blank out the dialog name with Dialog1 = "" before adding dialog.
		'Add in all of your mandatory field handling from your dialog here.	
		
		
			BeginDialog Dialog1, 0, 0, 191, 105, "NOOB Test Case"
			Text 5, 10, 50, 10, "Case Number:"
			EditBox 75, 5, 45, 15, MAXIS_case_number
			Text 5, 30, 65, 10, "Footer Month/year:"
			EditBox 75, 25, 20, 15, MAXIS_footer_month
			EditBox 100, 25, 20, 15, MAXIS_footer_year
			ButtonGroup ButtonPressed
				OkButton 75, 85, 50, 15
				CancelButton 135, 85, 50, 15
			EndDialog

		'Shows dialog -----------------------------------------------------------------------------------------------------
		'Notes: DO Loop to ensure all fields are completed
			'MsgBox "Message before Dialog"
			DO
				Do
					err_msg = ""
					Dialog Dialog1
					cancel_confirmation
					IF MAXIS_case_number = "" or (IsNumeric(MAXIS_case_number) = False) or (LEN(MAXIS_case_number) > 8) Then err_msg = "Case Number: Must have numeric entry <8 characters" & vbNewLine
					IF MAXIS_footer_month = "" or (IsNumeric(MAXIS_footer_month) = False) or (LEN(MAXIS_FOOTER_month) <> 2) or Then err_msg = err_msg & vbNewLine & "Month: 2 Characters, numeric & less than 12" & vbNewLine
					'IF MAXIS_footer_month > 12 then err_msg = err_msg & vbNewLine & "Month must be less than 12" & vbNewLine
					IF MAXIS_footer_year = "" or (IsNumeric(MAXIS_footer_year) = False) or (LEN(MAXIS_FOOTER_year) <> 2) Then err_msg = err_msg & vbNewLine & "Year: 2 Characters & numeric" & vbNewLine
					If err_msg <> "" Then Msgbox "***Notice***" & vbNewLine & err_msg 
					'Add in all of your mandatory field handling from your dialog here. Does not restrict user to 2 or 8 digits....gap
						'Add to all dialogs where you need to work within BLUEZONE
				Loop Until err_msg = ""
				CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
			LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
			'End dialog section-----------------------------------------------------------------------------------------------
			'msgbox "Message after Dialog "





'VERIFY YOU ARE IN MAXIS-SELF, if Not NAVIATE TO MAXIS-SELF===========================================================================
		Do
			Do
				EMReadScreen MAXIS_check, 5, 1, 39
				EMReadScreen SELF_check, 4, 2, 50
				If MAXIS_Check = "MAXIS"then PF3
				'If MAXIS_Check <> "MAXIS" then Call Dialog2
				If SELF_check <> "SELF" then PF3
				Loop Until MAXIS_check = "MAXIS" and SELF_check = "SELF"
			CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
	'MsgBox "Do loop for correct screen"
	'End dialog section-----------------------------------------------------------------------------------------------

'NAVIGATING TO THE CORRECT MAXIS SCREEN & ENTERING INFO===========================================================================
	'code snippet example---------------------------------------------------------------------------------------------
	'This is here to show you how we might use the advanced automation library to do something in MAXIS.
	'Feel free to build from this or just take the parts that are helpful.
	'now we are going to STAT/SUMM for a specific case
	'MsgBox "Start of Entring Info"
	EMWriteScreen "STAT", 16, 43				'writing the MAXIS function to enter in the correct place in MAXIS
	EMWriteScreen "        ", 18, 43			'TODO - should I be concerned if there is already information on this line?
	EMWriteScreen MAXIS_case_number, 18, 43		'entering  case number in the 'case number' line
	EMWriteScreen MAXIS_footer_month, 20, 43
	EMWriteScreen MAXIS_footer_year, 20, 46
	EMWriteScreen "SUMM", 21, 70				'writing the MAXIS command to enter in the correct place in MAXIS
	transmit									'function to move in MAXIS
	EMWriteScreen "JOBS", 20, 71
	transmit


' ' 'VERIFY FOOTER MONTH IN JOBS SCREEN IF NOT, ENTER JOBS. When you login the first time it stops on the SUMM page for some reason. This helps work around that. 
	Do
			Do
				EMReadScreen MAXIS_check, 5, 1, 39
				EMReadScreen JOBS_check, 4, 2, 45
				If JOBS_check <> "JOBS" then EMWriteScreen "JOBS:", 20, 71 & transmit	
				Loop Until MAXIS_check = "MAXIS" and JOBS_check = "JOBS"
			CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
	'MsgBox "Do loop for correct screen"
' 	'End dialog section-----------------------------------------------------------------------------------------------


'READING INFO BASED ON MAXIS
	EMReadScreen MAXIS_case_number, 8, 20, 37
	EMReadScreen MAXIS_footer_month, 2, 20, 55
	EMReadScreen MAXIS_footer_year, 2, 20, 58
	EMReadScreen MAXIS_user, 7, 21, 71
	EMReadScreen MAXIS_First, 8, 5, 12
	EMReadScreen MAXIS_Last, 6, 5, 6
	EMReadScreen MAXIS_Retro, 8, 17, 38
	EMReadScreen MAXIS_Pros, 8, 17, 67
	EMReadScreen MAXIS_Empl, 34, 7, 38
	'msgbox "After Read Screen"


' SECOND DIALOG WITH USING WHAT's READ
	Dialog1 = "" 'blanking out dialog name		
		BeginDialog Dialog1, 0, 0, 316, 190, "Displaying Read Info"
		Text 185, 20, 50, 10, "Case Number:"
		Text 260, 20, 45, 10, MAXIS_case_number
		Text 185, 35, 65, 10, "Footer Month/Year:"
		Text 260, 35, 15, 10, MAXIS_footer_month
		Text 280, 35, 15, 10, MAXIS_footer_year
		Text 10, 140, 105, 10, "**Enter Case Notes Here**"
		Text 20, 155, 25, 10, "Notes:"
		EditBox 45, 150, 250, 15, Case_Notes
		ButtonGroup ButtonPressed
			OkButton 195, 170, 50, 15
			CancelButton 255, 170, 50, 15
		Text 15, 20, 45, 10, "First Name:"
		Text 15, 35, 45, 10, "Last Name:"
		Text 90, 20, 45, 10, MAXIS_First
		Text 90, 35, 45, 10, MAXIS_Last
		Text 80, 80, 200, 10, MAXIS_Empl
		Text 15, 80, 45, 10, "Employer:"
		Text 15, 95, 45, 10, "Retro Total"
		GroupBox 0, 65, 310, 70, "JOBS INFO"
		Text 80, 95, 200, 10, MAXIS_Retro
		Text 15, 110, 45, 10, "Pros Total"
		Text 80, 110, 200, 10, MAXIS_Pros
		GroupBox 0, 5, 310, 55, "CASE INFO"
		EndDialog


	DO
		Do
			err_msg = ""
			Dialog Dialog1
			cancel_confirmation
			IF Case_Notes = "" Then err_msg = "Notes are required before proceeding"
			If err_msg <> "" Then Msgbox "***Notice***" & vbNewLine & err_msg 
		Loop Until err_msg = ""
		CALL check_for_password(are_we_passworded_out)			
	LOOP UNTIL are_we_passworded_out = false
	'msgbox "After Dialog 2 "


' CREATING A CASE NOTE FROM ACTV

		PF4
		PF9
		Row = 4
	
		'IF LEN(case_note) > 78 then EMSendKey "<Enter>"
		
		Arr1 = Array(MAXIS_case_number, MAXIS_user, MAXIS_footer_month, MAXIS_footer_year,MAXIS_First, MAXIS_Last, MAXIS_Empl, case_notes)
		
		Row = LEN(arr1) +1
		'Row +=1 is the same as row = row + 1

		EMWRiteScreen ("Case:" & Arr1(0)) , row, 3
		Row = row +1
		EMWRiteScreen ("User:" & Arr1(1)) , row, 3
		Row = row +1
		EMWRiteScreen ("Month:" & Arr1(2)) , row, 3
		Row = row +1
		EMWRiteScreen ("Year:" & Arr1(3)) , row, 3
		Row = row +1
		EMWRiteScreen ("First Name:" & Arr1(4)) , row, 3
		Row = row +1
		EMWRiteScreen ("Last Name:" & Arr1(5)) , row, 3
		Row = row +1
		EMWRiteScreen ("Employee" & Arr1(6)) , row, 3
		Row = row +1
		EMWRiteScreen ("Notes:" & Arr1(7)) , row, 3
		'If Row > 17 then PF8


		'IF LEN(Arr1(0)) > 78 then SPLIT (arr1(0), 78)

		' Row= row +1
		' EMWriteScreen(Arr1(0)), row, 5
		' Row= row +2
		' EMWRiteScreen "User:", row, 3
		' Row= row +1
		' EMWriteScreen(Arr1(1)), row, 5
		' Row= row +2
		' EMWRiteScreen "Month:", row, 3
		' Row= row +1
		' EMWriteScreen(Arr1(2)), row, 5
		' Row= row +2
		' EMWRiteScreen "Year:", row, 3
		' Row= row +1
		' EMWriteScreen(Arr1(3)), row, 5
		' Row= row +2
		' EMWRiteScreen "First Name:", row, 3
		' Row= row +1
		' EMWriteScreen(Arr1(4)), row, 5
		' Row= row +2
		' EMWRiteScreen "Last Name", row, 3
		' Row= row +1
		' EMWriteScreen(Arr1(5)), row, 5
		' Row= row +2
		' EMWRiteScreen "Notes", row, 3
		' Row= row +1
		' EMWriteScreen(Arr1(6)), row, 5



		' EMWriteScreen "Case:" & MAXIS_case_number, row, 3
		' Row = Row + 2
		

		' EMWriteScreen "User:" & MAXIS_user, row, 3
		' Row = row + 2
		' Next

		' EMWriteScreen "***Month/Year***", row, 3
		' Row = Row + 1 
		' EMWRiteScreen MAXIS_footer_month & "/" & MAXIS_footer_year, row, 5
		' Row = row + 2
		
		' EMWriteScreen "***First Name***", row, 3
		' Row = row + 1 
		' EMWriteScreen MAXIS_First, row, 5
		' Row = row + 2

		' EMWriteScreen "***Last Name***", row, 3
		' Row = row + 1 
		' EMWriteScreen MAXIS_Last, row, 5
		' Row = row + 2


		' EMWriteScreen "***Employment:***" & MAXIS_Empl, row, 3
		' Row = row + 2

		' EMWriteScreen "***Retrospective Total:***" & MAXIS_REtro, row, 3
		' Row = row + 2


		' EMWriteScreen "***Prospective Total:***" & MAXIS_Pros, row, 3
		' Row= row + 2
		
		' EMWriteScreen "***Case Notes***", row, 3
		' Row = row + 1
		' EMWriteScreen Case_Notes, row, 5

		' If Case_Notes are longer than one line (80 characters): 
			'Write notes until end is reached, if a full word move to the next line and continue writing. 
		

			

			
'leave the case note open and in edit mode unless you have a business reason not to (BULK scripts, multiple case notes, etc.)

'End the script. Put any success messages in between the quotes


script_end_procedure("Yay, you are done with NOOB script!")