'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - UPDATE BM CASE REVIEW LIST.vbs" 'BULK script that creates a list of cases that require an interview, and the contact phone numbers'
start_time = timer
STATS_counter = 1
STATS_manualtime = 20
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine & _
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

Function earned_income_exemption
	exempt_status = False
    prosp_inc = 0
    prosp_hrs = 0
    prospective_hours = 0
	
    CALL navigate_to_MAXIS_screen("STAT", "JOBS")
    EMWritescreen "01", 20, 79				'ensures that we start at 1st job
    transmit
    EMReadScreen num_of_JOBS, 1, 2, 78
    IF num_of_JOBS <> "0" THEN
        DO
            EMReadScreen jobs_end_dt, 8, 9, 49
            EMReadScreen cont_end_dt, 8, 9, 73
            IF jobs_end_dt = "__ __ __" THEN
            	CALL write_value_and_transmit("X", 19, 38)
            	EMReadScreen prosp_monthly, 8, 18, 56
            	prosp_monthly = trim(prosp_monthly)
            	IF prosp_monthly = "" THEN prosp_monthly = 0
            	prosp_inc = prosp_inc + prosp_monthly
            	EMReadScreen prosp_hrs, 8, 16, 50
            	IF prosp_hrs = "        " THEN prosp_hrs = 0
            	prosp_hrs = prosp_hrs * 1						'Added to ensure that prosp_hrs is a numeric
            	EMReadScreen pay_freq, 1, 5, 64
            	IF pay_freq = "1" THEN
            		prosp_hrs = prosp_hrs
            	ELSEIF pay_freq = "2" THEN
            		prosp_hrs = (2 * prosp_hrs)
            	ELSEIF pay_freq = "3" THEN
            		prosp_hrs = (2.15 * prosp_hrs)
            	ELSEIF pay_freq = "4" THEN
            		prosp_hrs = (4.3 * prosp_hrs)
            	END IF
            	prospective_hours = prospective_hours + prosp_hrs
            ELSE
            	jobs_end_dt = replace(jobs_end_dt, " ", "/")
            	IF DateDiff("D", date, jobs_end_dt) > 0 THEN
            		'Going into the PIC for a job with an end date in the future
            		CALL write_value_and_transmit("X", 19, 38)
            		EMReadScreen prosp_monthly, 8, 18, 56
            		prosp_monthly = trim(prosp_monthly)
            		IF prosp_monthly = "" THEN prosp_monthly = 0
            		prosp_inc = prosp_inc + prosp_monthly
            		EMReadScreen prosp_hrs, 8, 16, 50
            		IF prosp_hrs = "        " THEN prosp_hrs = 0
            		prosp_hrs = prosp_hrs * 1						'Added to ensure that prosp_hrs is a numeric
            		EMReadScreen pay_freq, 1, 5, 64
            		IF pay_freq = "1" THEN
            			prosp_hrs = prosp_hrs
            		ELSEIF pay_freq = "2" THEN
            			prosp_hrs = (2 * prosp_hrs)
            		ELSEIF pay_freq = "3" THEN
            			prosp_hrs = (2.15 * prosp_hrs)
            		ELSEIF pay_freq = "4" THEN
            			prosp_hrs = (4.3 * prosp_hrs)
            		END IF
            		'added seperate incremental variable to account for multiple jobs
            		prospective_hours = prospective_hours + prosp_hrs
            	END IF
            END IF
            transmit		'to exit PIC
            EMReadScreen JOBS_panel_current, 1, 2, 73
            'looping until all the jobs panels are calculated
            If cint(JOBS_panel_current) < cint(num_of_JOBS) then transmit
        Loop until cint(JOBS_panel_current) = cint(num_of_JOBS)
    END IF
    
	EMWriteScreen "BUSI", 20, 71
    CALL write_value_and_transmit(person, 20, 76)
    EMReadScreen num_of_BUSI, 1, 2, 78
    IF num_of_BUSI <> "0" THEN
        DO
            EMReadScreen busi_end_dt, 8, 5, 72
            busi_end_dt = replace(busi_end_dt, " ", "/")
            IF IsDate(busi_end_dt) = True THEN
            	IF DateDiff("D", date, busi_end_dt) > 0 THEN
            		EMReadScreen busi_inc, 8, 10, 69
            		busi_inc = trim(busi_inc)
            		EMReadScreen busi_hrs, 3, 13, 74
            		busi_hrs = trim(busi_hrs)
            		IF InStr(busi_hrs, "?") <> 0 THEN busi_hrs = 0
            		prosp_inc = prosp_inc + busi_inc
            		prosp_hrs = prosp_hrs + busi_hrs
            		prospective_hours = prospective_hours + busi_hrs
            	END IF
            ELSE
            	IF busi_end_dt = "__/__/__" THEN
            		EMReadScreen busi_inc, 8, 10, 69
            		busi_inc = trim(busi_inc)
            		EMReadScreen busi_hrs, 3, 13, 74
            		busi_hrs = trim(busi_hrs)
            		IF InStr(busi_hrs, "?") <> 0 THEN busi_hrs = 0
            		prosp_inc = prosp_inc + busi_inc
            		prosp_hrs = prosp_hrs + busi_hrs
            		prospective_hours = prospective_hours + busi_hrs
            	END IF
            END IF
            transmit
            EMReadScreen enter_a_valid, 13, 24, 2
        LOOP UNTIL enter_a_valid = "ENTER A VALID"
    END IF
	
	EMWriteScreen "RBIC", 20, 71
	CALL write_value_and_transmit(person, 20, 76)
	EMReadScreen num_of_RBIC, 1, 2, 78
	
	IF num_of_RBIC <> "0" THEN exempt_status = False 
	IF prosp_inc >= 935.25 OR prospective_hours >= 129 THEN exempt_status = True 
	'IF prospective_hours >= 80 AND prospective_hours < 129 THEN exempt_status = True 
End FUNCTION

Function disabled_exemption
    CALL navigate_to_MAXIS_screen("STAT", "DISA")
    disa_status = ""
    EMReadScreen num_of_DISA, 1, 2, 78
    IF num_of_DISA <> "0" THEN
    	EMReadScreen disa_end_dt, 10, 6, 69
    	disa_end_dt = replace(disa_end_dt, " ", "/")
    	EMReadScreen cert_end_dt, 10, 7, 69
    	cert_end_dt = replace(cert_end_dt, " ", "/")
    	IF IsDate(disa_end_dt) = True THEN
    		'msgbox isdate(disa_end_dt)
    		IF DateDiff("D", date, disa_end_dt) > 0 THEN disa_status = True
    		IF disa_end_dt = "99/99/9999" THEN disa_status = TRUE
    	elseif disa_end_dt = "__/__/____" then 
    		EMReadScreen disa_begin_dt, 10, 6, 47
    		IF disa_begin_dt <> "__ __ ____" THEN 
    			disa_status = True
    			'msgbox disa_end_dt & vbcr & disa_status
    		End if 
    	elseIF IsDate(cert_end_dt) = True AND disa_status = False THEN
    		IF DateDiff("D", date, cert_end_dt) > 0 THEN disa_status = true
    		IF cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999" THEN 
    			EMReadScreen cert_begin_dt, 8, 7, 47
    			IF cert_begin_dt <> "__ __ __" THEN disa_status = True
    		End if
    	END IF
    End if 
End Function

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

'default file path since this is stationary
file_selection_path = "\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team\Eligibility Support\Restricted\QI - Quality Improvement\SNAP\Banked months data\Master case review list.xlsx"

'dialog and dialog DO...Loop	
Do
	Do
			'The dialog is defined in the loop as it can change as buttons are pressed 
			BeginDialog file_select_dialog, 0, 0, 226, 50, "Select the banked months case review file."
  				ButtonGroup ButtonPressed
    			PushButton 175, 10, 40, 15, "Browse...", select_a_file_button
    			OkButton 110, 30, 50, 15
    			CancelButton 165, 30, 50, 15
  				EditBox 5, 10, 165, 15, file_selection_path
			EndDialog
			err_msg = ""
			Dialog file_select_dialog
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
	
BeginDialog update_banked_month_status_dialog, 0, 0, 191, 60, "Dialog"
  DropListBox 80, 10, 105, 15, "Select one..."+chr(9)+"January"+chr(9)+"February"+chr(9)+"March"+chr(9)+"April"+chr(9)+"May"+chr(9)+"June"+chr(9)+"July"+chr(9)+"August"+chr(9)+"September"+chr(9)+"October"+chr(9)+"November"+chr(9)+"December", month_selection
  ButtonGroup ButtonPressed
    OkButton 80, 35, 50, 15
    CancelButton 135, 35, 50, 15
  Text 5, 15, 70, 10, "Update status month:"
EndDialog

'DISPLAYS DIALOG
DO
	DO
		err_msg = ""
		Dialog update_banked_month_status_dialog
		If ButtonPressed = 0 then StopScript
		If month_selection = "Select one..." then err_msg = err_msg & vbNewLine & "* Please select the status month to update."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in			
	
'resets the case number and footer month/year back to the CM (REVS for current month plus two has is going to be a problem otherwise)
back_to_self
EMwritescreen "________", 18, 43
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46
transmit

'starts adding phone numbers at row selected
Select Case month_selection
Case "January"
	MAXIS_footer_month = "01"
	MAXIS_footer_year = "18"
	excel_col = 29
Case "February"
	MAXIS_footer_month = "02"
	MAXIS_footer_year = "18"
	excel_col = 30
Case "March"
	MAXIS_footer_month = "03"
	MAXIS_footer_year = "18"
	excel_col = 31
Case "April"
	MAXIS_footer_month = "04"
	MAXIS_footer_year = "18"
	excel_col = 32
Case "May"
	MAXIS_footer_month = "05"
	MAXIS_footer_year = "18"
	excel_col = 33
Case "June"
	MAXIS_footer_month = "06"
	MAXIS_footer_year = "18"
	excel_col = 34
Case "July"
	MAXIS_footer_month = "07"
	MAXIS_footer_year = "18"
	excel_col = 35
Case "August"
	MAXIS_footer_month = "08"
	MAXIS_footer_year = "18"
	excel_col = 24
Case "September"
	MAXIS_footer_month = "09"
	MAXIS_footer_year = "18"
	excel_col = 25
Case "October"
	MAXIS_footer_month = "10"
	MAXIS_footer_year = "18"
	excel_col = 26
Case "November"
	MAXIS_footer_month = "11"
	MAXIS_footer_year = "18"
	excel_col = 27
Case "December"
	MAXIS_footer_month = "12"
	MAXIS_footer_year = "18"
	excel_col = 28
End Select

back_to_self
excel_row = 2
DO  
    'Grabs the case number
	MAXIS_case_number = objExcel.cells(excel_row, 3).value
    If MAXIS_case_number = "" then exit do

    Call navigate_to_MAXIS_screen("CASE", "CURR")
    EMReadScreen priv_check, 6, 24, 14 			'If it can't get into the case needs to skip
    IF priv_check = "PRIVIL" THEN
        EMWriteScreen "________", 18, 43		'clears the case number
        transmit
        ObjExcel.Cells(excel_row, excel_col).Value = ""
    Else     
        EMReadScreen case_status, 8, 8, 9
        case_status = trim(case_status)
        If case_status = "INACTIVE" then 
            ObjExcel.Cells(excel_row, excel_col).Value = "Inactive"
	    Elseif case_status = "ACTIVE" then 
            MAXIS_row = 9
            Do 
                EMReadScreen prog_name, 4, MAXIS_row, 3
                prog_name = trim(prog_name)
                if prog_name = "" then exit do
                If prog_name = "FS" then 
                    EMReadScreen case_status, 8, MAXIS_row, 9
                    case_status = trim(case_status)
	                if case_status = "ACTIVE" then 
                        exit do
                    ELSE 
                        MAXIS_row = MAXIS_row + 1
                    END IF 
                Else
                    MAXIS_row = MAXIS_row + 1
                END IF 
	        Loop until MAXIS_row = 17
            If prog_name <> "FS" then ObjExcel.Cells(excel_row, excel_col).Value = "Inactive"
        END If 
	End if 
	'inputs EXEMPT on cases that are active on GA/open on other programs
	If case_status = "ACTIVE" then 
		Call navigate_to_MAXIS_screen("STAT", "PROG")
		EMReadScreen SNAP_prog_status, 4, 10, 74 
		IF SNAP_prog_status = "INAC" or trim(SNAP_prog_status) = "" or SNAP_prog_status = "DENY" then 
			ObjExcel.Cells(excel_row, excel_col).Value = "Inactive"
		Elseif SNAP_prog_status = "PEND" then 
			ObjExcel.Cells(excel_row, excel_col).Value = "Pending"
		Else 
			EMReadScreen cash_prog_status, 4, 6, 74
			IF cash_prog_status = "ACTV" then 
				ObjExcel.Cells(excel_row, excel_col).Value = "Exempt"
			Else
				ObjExcel.Cells(excel_row, excel_col).Value = ""
			End if 
		END IF 	
		
		If ObjExcel.Cells(excel_row, excel_col).Value = "" then
			Call navigate_to_MAXIS_screen("STAT", "MEMB")
			row = 5
			HH_count = 0
			Do 
				EMReadScreen member_number, 2, row, 3
				HH_count = HH_count + 1
				transmit
				EMReadScreen MEMB_error, 5, 24, 2
			Loop until MEMB_error = "ENTER"
			If HH_count <> 1 then
				ObjExcel.Cells(excel_row, excel_col).Value = ""
			 Else 	
				Call navigate_to_MAXIS_screen("STAT", "WREG")
				EMReadScreen fset_code, 2, 8, 50
				If fset_code = "09" then  
					'msgbox "coded 09"
					Call earned_income_exemption
					'msgbox "EI exemption: " & exempt_status
					If exempt_status = true then 
						ObjExcel.Cells(excel_row, excel_col).Value = "Exempt"
						ObjExcel.Cells(excel_row, 4).Value = "09/01 exemption"
					End if 
				elseif fset_code = "03" then
					'msgbox "coded 03"
				   	Call disabled_exemption
				    'msgbox "DISA status: " & disa_status
					If disa_status = true then 
						ObjExcel.Cells(excel_row, excel_col).Value = "Exempt"
						ObjExcel.Cells(excel_row, 4).Value = "03/01 exemption"
					End if 
				Else 
					ObjExcel.Cells(excel_row, excel_col).Value = ""
				End if 
			End if 
		End if 	
	END IF

    MAXIS_case_number = ""
    excel_row = excel_row + 1
	STATS_counter = STATS_counter + 1
LOOP UNTIL objExcel.Cells(excel_row, 3).value = ""	'looping until the list of cases to check for recert is complete

STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)
script_end_procedure("Success! The Excel file now has been updated. Please review the blank case statuses that remain.")
