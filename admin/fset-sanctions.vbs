'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - FSET SANCTIONS.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 120                	'manual run time in seconds
STATS_denomination = "C"       		'M is for each MEMBER
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
CALL changelog_update("06/09/2018", "Made several updates to support using a single master sanction list while processing. Also added text to case note if the case has been identified as potentially homeless for unfit for employment expansion exemption.", "Ilse Ferris, Hennepin County")
CALL changelog_update("05/21/2018", "Added additional handling for when a WCOM exists in the add WCOM option.", "Ilse Ferris, Hennepin County")
CALL changelog_update("05/19/2018", "Added searching for LETR dates when don't match the orientation date.", "Ilse Ferris, Hennepin County")
CALL changelog_update("05/10/2018", "Streamlined text in worker comments based on feedback provided by DHS.", "Ilse Ferris, Hennepin County")
call changelog_update("05/07/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------Custom Functions
Function HCRE_panel_bypass() 
	'handling for cases that do not have a completed HCRE panel
	PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
	Do
		EMReadscreen HCRE_panel_check, 4, 2, 50
		If HCRE_panel_check = "HCRE" then
			PF10	'exists edit mode in cases where HCRE isn't complete for a member
			PF3
		END IF
	Loop until HCRE_panel_check <> "HCRE"
End Function

'----------------------------------------------------------------------------------------------------DIALOG
'The dialog is defined in the loop as it can change as buttons are pressed 
BeginDialog info_dialog, 0, 0, 256, 125, "SNAP ABAWD (FSET) Sanction"
  ButtonGroup ButtonPressed
    PushButton 200, 40, 50, 15, "Browse...", select_a_file_button
  DropListBox 170, 85, 80, 15, "Select one..."+chr(9)+"Review sanctions"+chr(9)+"Update WREG only"+chr(9)+"Add WCOM", sanction_option
  ButtonGroup ButtonPressed
    OkButton 150, 105, 50, 15
    CancelButton 200, 105, 50, 15
  EditBox 15, 40, 180, 15, file_selection_path
  EditBox 65, 105, 80, 15, worker_signature
  Text 90, 90, 80, 10, "Select the script option:"
  Text 15, 60, 230, 15, "Select the Excel file that contains your information by selecting the 'Browse' button, and finding the file."
  Text 20, 20, 225, 15, "This script should be used members have been identified by SNAP E and T as ready for sanction."
  Text 5, 110, 55, 10, "Worker sigature:"
  GroupBox 10, 5, 245, 75, "Using this script:"
EndDialog

BeginDialog excel_row_dialog, 0, 0, 126, 50, "Select the excel row to start"
  EditBox 75, 5, 40, 15, excel_row_to_start
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""

'dialog and dialog DO...Loop	
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
        err_msg = ""
    	Dialog info_dialog
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If sanction_option = "Select one..." then err_msg = err_msg & vbNewLine & "* Select a sanction option."
        If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    Do 
        dialog excel_row_dialog
        If ButtonPressed = cancel then stopscript
        If IsNumeric(excel_row_to_start) = false then msgbox "Enter a numeric excel row to start the script."
    Loop until IsNumeric(excel_row_to_start) = True
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

back_to_SELF
excel_row = excel_row_to_start

'Creating variables for the excel columns as this project has not yet finished evolving. 
date_col = 5
status_col = 6
wreg_col = 7
months_col = 8
referral_col = 9
orient_col = 10
notice_col = 11
sanction_col = 12
notes_col = 13

'----------------------------------------------------------------------------------------------------Actually imposing the sanction
If sanction_option = "Review sanctions" then 
    objExcel.Cells(1, 6).Value = "SNAP Status"
    objExcel.Cells(1, 7).Value = "ABAWD/FSET"
    objExcel.Cells(1, 8).Value = "ABAWD Months Used"
    objExcel.Cells(1, 9).Value = "Referral Date"
    objExcel.Cells(1, 10).Value = "Orient Date"
    objExcel.Cells(1, 11).Value = "Notice Sent?"
    objExcel.Cells(1, 12).Value = "Sanction"
    objExcel.Cells(1, 13).Value = "BULK Notes"
    
    FOR i = 1 to 13 	'formatting the cells'
        objExcel.Cells(1, i).Font.Bold = True		'bold font'
        ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
        objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT
    
    Do 
        sanction_notes = ""
        referral_date = ""
        abawd_counted_months = ""
        abawd_counted_months = ""
        sanction_case = ""
        found_member = ""
        
        PMI_number = ObjExcel.Cells(excel_row, 1).Value
        PMI_number = trim(PMI_number)
        MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
        MAXIS_case_number = trim(MAXIS_case_number)
        If MAXIS_case_number = "" then exit do
        
        MAXIS_footer_month = CM_mo	'establishing footer month/year as next month 
        MAXIS_footer_year = CM_yr 
        Call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year
    
        Call navigate_to_MAXIS_screen("STAT", "PROG")
    	EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
    	If PRIV_check = "PRIV" then
            sanction_notes = sanction_notes & "PRIV case."
            found_member = False 
    	Else
            EmReadscreen county_code, 4, 21, 21
            IF county_code <> ucase(worker_county_code) then 
                'msgbox county_code & vbcr & worker_county_code
                sanction_notes = sanction_notes & " Out-of-county case."
                found_member = False 
            End if 
        End if 
        
        If found_member <> False then 
            EmReadscreen SNAP_actv, 4, 10, 74
            ObjExcel.Cells(excel_row, status_col).Value = SNAP_actv
            If SNAP_actv = "ACTV" then 
                found_member = True
            else 
                sanction_notes = sanction_notes & "SNAP not active."
                found_member = FALSE
            End if
        End if
            
        Call HCRE_panel_bypass  'function to ensure we get past HCRE panel 
        
        If found_member = True then 
            '>>>>>>>>>>ADDR
            CALL navigate_to_MAXIS_screen("STAT", "ADDR")
            EMReadScreen homeless_code, 1, 10, 43
            EmReadscreen addr_line_01, 16, 6, 43
            IF homeless_code = "Y" or addr_line_01 = "GENERAL DELIVERY" THEN sanction_notes = sanction_notes & " Possible homeless exemption."
    
            Call navigate_to_MAXIS_screen ("STAT", "MEMB")
            member_number = ""
            Do 
                EMReadscreen memb_PMI, 8, 4, 46
                memb_PMI = trim(memb_PMI)
                If memb_PMI = PMI_number then
                    EMReadscreen member_number, 2, 4, 33
                    found_member = True 
                    exit do
                Else 
                    transmit
                END IF
                EMReadScreen MEMB_error, 5, 24, 2
            Loop until MEMB_error = "ENTER"
            
            If member_number = "" then 
                sanction_notes = sanction_notes & " Unable to find HH member on case."
                found_member = False 
            End if 
        End if 
            
        If found_member = True then 
    	    call navigate_to_MAXIS_screen("STAT", "WREG")
            Call write_value_and_transmit(member_number, 20, 76)
            
    	    EMReadScreen FSET_code, 2, 8, 50
    	    EMReadScreen ABAWD_code, 2, 13, 50
            wreg_codes = FSET_code & "/" & ABAWD_code
    	    ObjExcel.Cells(excel_row, wreg_col).Value = wreg_codes
            
            '----------------------------------------------------------------------------------------------------Reading the amount of counted months 
            EMReadScreen wreg_total, 1, 2, 78
            IF wreg_total <> "0" THEN
            	EmWriteScreen "x", 13, 57		'Pulls up the WREG tracker'
            	transmit
            	EMREADScreen tracking_record_check, 15, 4, 40  		'adds cases to the rejection list if the ABAWD tracking record cannot be accessed.
            	If tracking_record_check <> "Tracking Record" then
    	       		sanction_notes = sanction_notes & " Cannot access the ABAWD tracking record. Review and process manually."
            	ELSE
            		bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))		'col to search starts at 15, increased by 4 for each footer month
            		bene_yr_row = 10
            		abawd_counted_months = 0					'delclares the variables values at 0
            		month_count = 0
            		DO
            			'establishing variables for specific ABAWD counted month dates
            			If bene_mo_col = "19" then counted_date_month = "01"
            			If bene_mo_col = "23" then counted_date_month = "02"
            			If bene_mo_col = "27" then counted_date_month = "03"
            			If bene_mo_col = "31" then counted_date_month = "04"
            			If bene_mo_col = "35" then counted_date_month = "05"
            			If bene_mo_col = "39" then counted_date_month = "06"
            			If bene_mo_col = "43" then counted_date_month = "07"
            			If bene_mo_col = "47" then counted_date_month = "08"
            			If bene_mo_col = "51" then counted_date_month = "09"
            			If bene_mo_col = "55" then counted_date_month = "10"
            			If bene_mo_col = "59" then counted_date_month = "11"
            			If bene_mo_col = "63" then counted_date_month = "12"
            			'counted date year: this is found on rows 7-10. Row 11 is current year plus one, so this will be exclude this list.
            			If bene_yr_row = "10" then counted_date_year = right(DatePart("yyyy", date), 2)
            			If bene_yr_row = "9"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -1, date)), 2)
            			If bene_yr_row = "8"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -2, date)), 2)
            			If bene_yr_row = "7"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -3, date)), 2)
            			abawd_counted_months_string = counted_date_month & "/" & counted_date_year

            			'reading to see if a month is counted month or not
            			EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col

            			'counting and checking for counted ABAWD months
            			IF is_counted_month = "X" or is_counted_month = "M" THEN
            				EMReadScreen counted_date_year, 2, bene_yr_row, 14			'reading counted year date
            				abawd_counted_months_string = counted_date_month & "/" & counted_date_year
            				abawd_info_list = abawd_info_list & ", " & abawd_counted_months_string			'adding variable to list to add to array
            				abawd_counted_months = abawd_counted_months + 1				'adding counted months
            			END IF

            			'declaring & splitting the abawd months array
            			If left(abawd_info_list, 1) = "," then abawd_info_list = right(abawd_info_list, len(abawd_info_list) - 1)
            			counted_months_array = Split(abawd_info_list, ",")

            			bene_mo_col = bene_mo_col - 4		're-establishing serach once the end of the row is reached
            			IF bene_mo_col = 15 THEN
            				bene_yr_row = bene_yr_row - 1
            				bene_mo_col = 63
            			END IF
            			month_count = month_count + 1
            		LOOP until month_count = 36
            	    PF3
            	End if
    	       	ObjExcel.Cells(excel_row, months_col).Value = abawd_counted_months
    	       END If
        End if 
        
        If found_member = True then  
            Call navigate_to_MAXIS_screen("INFC", "WORK")
            EmReadscreen no_referral, 2, 24, 2
            If no_referral = "NO" then 
                sanction_notes = sanction_notes & " No referral in WF1M for this case."
            Else 
                row = 7
                Do 
                    EMReadscreen work_memb, 2, row, 3
                    If work_memb = member_number then 
                        EmReadscreen referral_date, 8, row, 72
                        EmReadscreen appt_date, 8, row, 59
                        If trim(referral_date) <> "" then referral_date = replace(referral_date, " ", "/")
                        If appt_date <> "__ __ __" then 
                            appt_date = replace(appt_date, " ", "/")
                        Else 
                            appt_date = ""
                        End if 
                        ObjExcel.Cells(excel_row, referral_col).Value = referral_date
                        ObjExcel.Cells(excel_row, orient_col).Value = appt_date
                        found_member = True
                        exit do 
                    Else 
                        'row = row + 1
                        sanction_notes = sanction_notes & " More than one referral. Process manually."
                        found_member = False
                        Exit do
                    End if 
                Loop until trim(work_memb) = ""
            End if 
        End if 
        
        If found_member = True then
            Call navigate_to_MAXIS_screen("SPEC", "WCOM")
            row = 7
            DO
            	EMReadscreen notice_type, 16, row, 30
                If trim(notice_type) = "" then 
                    'msgbox "going to PF7"
                    PF7
                    row = 7
                    sanction_case = false 
                elseIf notice_type = "SPEC/LETR Letter" then 
                    EmReadscreen FS_notice, 2, row, 26
                    If FS_notice = "FS" or FS_notice = "  " then 
                        Call write_value_and_transmit ("x", row, 13)
                        'msgbox "entered into notice?"
                        EmReadscreen in_notice, 4, 1, 45
                        If in_notice = "Copy" then 
                            PF8
                            PF8 'twice to get to the date of the orientation 
                            EmReadscreen orient_date_LETR, 10, 2, 8
                            If isDate(orient_date_LETR) = False then 
                                sanction_case = FALSE
                                PF3
                            Else 
                                Call ONLY_create_MAXIS_friendly_date(orient_date_LETR)
                                If orient_date_LETR = appt_date then 
                                    ObjExcel.Cells(excel_row, notice_col).Value = "Yes"
                                    sanction_case = TRUE
                                    'msgbox appt_date & vbcr & orient_date_LETR
                                    exit do 
                                ELSE
                                    sanction_notes = sanction_notes & " Referral date does not match letter date. Letter date is: " & orient_date_LETR & ". "
                                    sanction_case = FALSE
                                    PF3
                                End if 
                            End if 
                            'msgbox sanction_case
                        End if 
                    else 
                        sanction_case = False 
                    End if 
                else 
                    sanction_case = false
                    'msgbox row 
                End if 
                If sanction_case = False then row = row + 1
                EmReadscreen no_notices, 10, 24, 2 
            Loop until no_notices = "NO NOTICES"
                    
            If sanction_case = true then 
                If wreg_codes = "30/06" or wreg_codes = "30/08" or wreg_codes = "30/10" or wreg_codes = "30/11" then 
                    ObjExcel.Cells(excel_row, sanction_col).Value = "Yes"
                Else 
                    ObjExcel.Cells(excel_row, sanction_col).Value = "No"
                End if 
            End if 
        End if     
        'msgbox MAXIS_case_number & vbcr & sanction_case & vbcr & appt_date & vbcr & orient_date_LETR
        ObjExcel.Cells(excel_row, notes_col).Value = sanction_notes            
    	STATS_counter = STATS_counter + 1
        excel_row = excel_row + 1
    Loop until ObjExcel.Cells(excel_row, 2).Value = ""
End if 

'----------------------------------------------------------------------------------------------------UPDATE WREG ONLY option
If sanction_option = "Update WREG only" then 
    excel_row = excel_row_to_start
    
    Do 
        sanction_notes = ""
        
        PMI_number = ObjExcel.Cells(excel_row, 1).Value
        PMI_number = trim(PMI_number)
        MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
        MAXIS_case_number = trim(MAXIS_case_number)
        
        sanction_code = objExcel.cells(excel_row, sanction_col).Value
        
        agency_informed_sanction = ObjExcel.Cells(excel_row, date_col).Value
        agency_informed_sanction = trim(agency_informed_sanction)
        
        sanction_notes = ObjExcel.Cells(excel_row, notes_col).Value
        
        If MAXIS_case_number = "" then exit do
        If trim(sanction_code) = "Yes" or trim(sanction_code) = "YES" then  
            Call MAXIS_background_check
            MAXIS_footer_month = CM_mo	'establishing footer month/year as next month 
            MAXIS_footer_year = CM_yr 
            call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year
            
            Call navigate_to_MAXIS_screen("CASE", "PERS")
            row = 10
            Do
            	EMReadScreen person_PMI, 8, row, 34
                person_PMI = trim(person_PMI)
                'msgbox person_PMI & vbcr & row
            	IF person_PMI = "" then exit do
            	IF  PMI_number = person_PMI then 
                    EMReadScreen FS_status, 1, row, 54
                    'msgbox FS_status
            		If FS_status <> "A" then 
                        sanction_case = False 
                        sanction_notes = sanction_notes & "Member is not active on SNAP. "
                    Else 
                        sanction_case = True 
                        EMReadScreen member_number, 2, row, 3               'gathers member number
                        exit do 
                        'sanction_array(member_num, item) = member_number
                        'EMReadScreen last_name, 15, row, 6                  'last name
                        'last_name = trim(last_name)
                        'EMReadScreen first_name, 11, row, 13                'first name
                        'first_name = trim(first_name)
                        'client_name = first_name & " " & last_name
                    End if 
            	Else
            		row = row + 3			'information is 3 rows apart
            		If row = 19 then
            			PF8
            			row = 10					'changes MAXIS row if more than one page exists
            		END if
            	END if
            	EMReadScreen last_PERS_page, 21, 24, 2
            LOOP until last_PERS_page = "THIS IS THE LAST PAGE"
            'msgbox sanction_case & vbcr & member_number
    
            MAXIS_footer_month = CM_plus_1_mo	'establishing footer month/year as next month to make the updates to the case
            MAXIS_footer_year = CM_plus_1_yr 
            call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year
            
            Call navigate_to_MAXIS_screen("STAT", "WREG")
            EMWriteScreen member_number, 20, 76
            transmit
            'checking to make sure that WREG is updating for the correct member
            EMReadScreen WREG_MEMB_check, 6, 24, 2
            IF WREG_MEMB_check = "REFERE" OR WREG_MEMB_check = "MEMBER" THEN 
                sanction_case = False
                sanction_notes = sanction_notes & "Member # is not valid on WREG. "
            else  
                'Ensuring that cases are mandatory FSET (ABAWD code "30")
                EMReadScreen ABAWD_status, 2, 13, 50
                If ABAWD_status = "10" or ABAWD_status = "08" or ABAWD_status = "06" or ABAWD_status = "11" then 
                    sanction_case = True
                Else 
                    sanction_case = False
                    sanction_notes = sanction_notes & "Member is not coded as a Mandatory FSET on WREG. "
                End if 
            End if 
    
            If sanction_case = True then 
                'msgbox MAXIS_CASE_NUMBER & " is going to be sanctioned."
                PF9
                EMReadscreen PWE_check, 1, 6, 68                    'who is the PWE?
                'updating WREG to reflect sanction 
                EMWriteScreen "02", 8, 50							'Enters sanction FSET code of "02"
                EMWriteScreen MAXIS_footer_month, 10, 50			'sanction begin month
                EMWriteScreen MAXIS_footer_year, 10, 56			    'sanction begin year
                EMWriteScreen "01", 11, 50	                        'sanction # 
                EMWriteScreen "01", 12, 50		                    'reason for sanction. This adds information to the notice. - If sanction is more than reason 01, then this will be processed indivdually.
                EMWriteScreen "_", 8, 80							'blanks out Defer FSET/No funds field 
                PF3
                '8.21.55 check for update date (if confirmation doesn't stick, then set to false)
                'msgbox "did WREG get updated?"
                '----------------------------------------------------------------------------------------------------The Case note
                Call start_a_blank_CASE_NOTE
                Call write_variable_in_CASE_NOTE("--SNAP sanction imposed for MEMB " & member_number & " for " & MAXIS_footer_month & "/" & MAXIS_footer_year & "--")
                If PWE_check = "Y" THEN Call write_variable_in_CASE_NOTE("* Entire household is sanctioned. Member is the PWE.")
                If PWE_check = "N" THEN Call write_variable_in_CASE_NOTE("* Only the HH MEMB is sanctioned. Memeber is NOT the PWE.")
                Call write_bullet_and_variable_in_CASE_NOTE("Date agency was notified of sanction", agency_informed_sanction)
                Call write_variable_in_CASE_NOTE("* Client does not appear to meet Good Cause criteria.")
                If instr(sanction_notes, "Possible homeless exemption") then 
                    Call write_variable_in_CASE_NOTE("---")
                    Call write_variable_in_CASE_NOTE("Client may meet an ABAWD exemption.")
                    Call write_variable_in_CASE_NOTE("Per CM 11.24: A person is unfit for employment if he or she is currently homeless. Homeless specifically defined for this purpose as:")
                    Call write_variable_in_CASE_NOTE("1. Lacking a fixed and regular nighttime residence, including temporary housing situations AND")
                    Call write_variable_in_CASE_NOTE("2. Lacking access to work-related necessities (i.e. shower or laundry facilities, etc.).")
                else 
                    Call write_variable_in_CASE_NOTE("* Client does not appear to meet any ABAWD exemptions.")
                End if 
                Call write_variable_in_CASE_NOTE("---")
                Call write_variable_in_CASE_NOTE("* Number/occurrence of sanction: 1st")
                Call write_variable_in_CASE_NOTE("* Reason for sanction: Failed to attend orientation.") 
                Call write_variable_in_CASE_NOTE("* Added Good Cause/failure to comply information to the notice.")
                Call write_variable_in_CASE_NOTE("---")
                Call write_variable_in_CASE_NOTE(worker_signature)
                PF3
                sanction_notes = sanction_notes & " APP sanction."
            End if 
            ObjExcel.Cells(Excel_row, notes_col).Value = sanction_notes
        End if     
        excel_row = excel_row + 1     
    Loop until ObjExcel.Cells(excel_row, 2).Value = ""  
End if 

'----------------------------------------------------------------------------------------------------ADD WCOM OPTION                
If sanction_option = "Add WCOM" then
    excel_row = excel_row_to_start
    Do 
        sanction_notes = ""
        
        MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
        MAXIS_case_number = trim(MAXIS_case_number)
        sanction_code = objExcel.cells(excel_row, sanction_col).Value
        If MAXIS_case_number = "" then exit do
        
        If trim(sanction_code) = "Yes" or trim(sanction_code) = "YES" then    
            MAXIS_footer_month = CM_plus_1_mo	'establishing footer month/year as next month to make the updates to the case
            MAXIS_footer_year = CM_plus_1_yr 
            call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year
            Call MAXIS_background_check
            
            'This section will check for whether forms go to AREP and SWKR
            call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
            EMReadscreen forms_to_arep, 1, 10, 45
            call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
            EMReadscreen forms_to_swkr, 1, 15, 63
             
            CALL navigate_to_MAXIS_screen("SPEC", "WCOM")
            'Searching for waiting SNAP notice
            wcom_row = 6
            Do
             	wcom_row = wcom_row + 1
             	Emreadscreen program_type, 2, wcom_row, 26
             	Emreadscreen print_status, 7, wcom_row, 71
             	If program_type = "FS" then
             		If print_status = "Waiting" then
             			Call write_value_and_transmit("x", wcom_row, 13)
             			PF9
             			Emreadscreen fs_wcom_exists, 3, 3, 15
             			If fs_wcom_exists <> "   " then 
                            sanction_notes = sanction_notes & "WCOM already exists on the notice."
                            PF3
                            PF3
                            fs_wcom_writen = true  
                        Else
             		        fs_wcom_writen = true
             				'This will write if the notice is for SNAP only
             				CALL write_variable_in_SPEC_MEMO("******************************************************")
             				CALL write_variable_in_SPEC_MEMO("What to do next:")
             				CALL write_variable_in_SPEC_MEMO("* You must meet the SNAP E&T rules by the end of the month. If you want to meet the rules, contact your team at 612-596-1300, or your SNAP E&T provider at 612-596-7411.")
             				CALL write_variable_in_SPEC_MEMO("* You can tell us why you did not meet the rules. If you had a good reason for not meeting the SNAP E&T rules, contact your SNAP E&T provider right away.")
             				CALL write_variable_in_SPEC_MEMO("******************************************************")
             				PF4
             				PF3
             			End if
             		End If
             	End If
             	If fs_wcom_writen = true then Exit Do
             	If wcom_row = 17 then
             		PF8
             		Emreadscreen spec_edit_check, 6, 24, 2
             		wcom_row = 6
             	end if
             	If spec_edit_check = "NOTICE" THEN no_fs_waiting = true
            Loop until spec_edit_check = "NOTICE"
            'Adding status 
            If no_fs_waiting = true then 
                sanction_notes = sanction_notes & "No waiting FS notice was found for the requested month."
            else 
                sanction_notes = sanction_notes & "WCOM added. Sanction imposed and complete."
            End if 
            ObjExcel.Cells(Excel_row, notes_col).Value = sanction_notes    
        End if   
        excel_row = excel_row + 1 
    Loop until ObjExcel.Cells(excel_row, 2).Value = ""    
End if  

STATS_counter = STATS_counter - 1 'since we start with 1
script_end_procedure("Success! Your list is complete. Please review the list for work that still may be required.")