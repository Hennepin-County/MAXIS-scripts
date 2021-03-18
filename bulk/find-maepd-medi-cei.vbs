'Required for statistical purposes===============================================================================
name_of_script = "BULK - FIND MAEPD MEDI CEI.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 86                      'manual run time in seconds
STATS_denomination = "C"       						 'C is for each CASE
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
'END FUNCTIONS LIBRARY BLOCK=================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("03/18/2021", "Updated the script to include if MSP programs are active. A true or false status will be output into column 5.", "Ilse Ferris, Hennepin County")
call changelog_update("01/31/2019", "Updated the script with increased handling for determining MA-EPD is open in current month, and if they are eligible for reimbursement.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2018", "Several updates to script: Users can enter more than one X number, identification of cases without an active HC span will be identified on the output spreadsheet, as will Medicare part B premiums and active MSP programs. Spreadsheet formatting updated for readability. Back end updates made to ensure password handling and transitions between MAXIS and MMIS.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

function navigate_to_MAXIS_test(maxis_mode)
'--- This function is to be used when navigating back to MAXIS from another function in BlueZone (MMIS, PRISM, INFOPAC, etc.)
'~~~~~ maxis_mode: This parameter needs to be either "PRODUCTION" or 'INQUIRY DB'
'===== Keywords: MAXIS, navigate
    attn
    Do
        EMReadScreen MAI_check, 3, 1, 33
        If MAI_check <> "MAI" then EMWaitReady 1, 1
    Loop until MAI_check = "MAI"
    
    IF maxis_mode = "PRODUCTION" then 
        row = 6
        selection_code = 1
    elseif maxis_mode = "INQUIRY DB" then 
        row = 7
        selection_code = 2
    End if 
    
    EMReadScreen region_check, 7, row, 15
    IF region_check = "RUNNING" THEN
        Call write_value_and_transmit(selection_code, 2, 15)
    ELSE
        EMConnect"A"
        attn
        EMReadScreen region_check, 7, 6, 15
        IF region_check = "RUNNING" THEN
            Call write_value_and_transmit(selection_code, 2, 15)
        ELSE
            EMConnect"B"
            attn
            EMReadScreen region_check, 7, 6, 15
            IF region_check = "RUNNING" THEN
                Call write_value_and_transmit(selection_code, 2, 15)
            Else 
                script_end_procedure("This script will now stop. Error occurred switching between MAXIS and MMIS.")
            END IF
        END IF
    END IF
end function 

get_county_code     'Checks for county info from global variables, or asks if it is not already defined.

'The script----------------------------------------------------------------------------------------------------------------------------------
EMConnect ""

CALL check_for_MAXIS(False)
back_to_SELF
CALL find_variable("Environment: ", production_or_inquiry, 10)			'reading if script was started in production of inquiry, this is used later to navigate back from MMIS.
If production_or_inquiry = "INQUIRY DB" then script_end_procedure("This script must be run in production. Please switch to production, and run the script again.")  'Deletes unapproved HC results from ELIG/HMMM, so must be in production.

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 211, 65, "Find MA-EPD Medicare CEI"
  EditBox 120, 10, 85, 15, x_number
  ButtonGroup ButtonPressed
    OkButton 100, 45, 50, 15
    CancelButton 155, 45, 50, 15
  Text 5, 15, 115, 10, "X Numbers, seperated by comma:"
  Text 10, 30, 200, 10, "This script will check REPT/ACTV for the selected X numbers."
EndDialog
Do 
    DO
    	err_msg = ""								'err message handling to loop until the user has entered the proper information
    	Dialog Dialog1
    	Cancel_without_confirmation
    	IF trim(x_number) = "" or len(x_number) < 7 then err_msg = err_msg & vbCr & "* The X numbers must be the full 7 digits."
    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Resolve for the script to continue."
    LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

MAXIS_footer_month = CM_mo              'ensuring that we're looking in current month/year for MA-EPD information 
MAXIS_footer_year = CM_yr 
Call MAXIS_footer_month_confirmation

current_month = MAXIS_footer_month & "/" & MAXIS_footer_year    'Establishing current month to check for MA-EPD elig in actual current month in ELIG/HC

msp_col = 5
prem_col = 6
reim_elig_col = 7

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True
objExcel.Cells(1, 1).Value = "X NUMBER"						'creating columns to store the information
objExcel.Cells(1, 2).Value = "CASE NUMBER"						'creating columns to store the information
objExcel.Cells(1, 3).Value = "CLIENT NAME"
objExcel.Cells(1, 4).Value = "NEXT REVW"
objExcel.Cells(1, 5).Value = "MSP ACTIVE?"
objExcel.Cells(1, 6).Value = "PART B PREM"
objExcel.Cells(1, 7).Value = "REIMBURSEMENT ELIG?"

FOR i = 1 to 7		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

CALL navigate_to_MAXIS_screen("REPT", "ACTV")						'navigating to rept actv for requested user
rept_row = 7                                                        'setting variables for first run through
excel_row = 2

'Splitting array for use by the for...next statement
worker_number_array = split(x_number, ",")

For each worker in worker_number_array
    worker = trim(worker)
	If worker = "" then exit for
	
    Call write_value_and_transmit(worker, 21, 13) 
    
    DO
    	EMReadScreen last_page, 21, 24, 2											'checking to see if this is the last page, if it is the loop can end.
    	DO
    		EMReadScreen MAXIS_case_number, 8, rept_row, 12						'reading the case numbers from rept/actv
    		MAXIS_case_number = trim(MAXIS_case_number)
    		EMReadScreen hc_case, 1, rept_row, 64
    		IF MAXIS_case_number <> "" AND hc_case <> " " THEN					'checking for HC cases
    			objExcel.Cells(excel_row, 1).Value = worker			'adding read variables to the spreadsheet
                objExcel.Cells(excel_row, 2).Value = MAXIS_case_number			'adding read variables to the spreadsheet
    			EMReadScreen client_name, 21, rept_row, 21						'grabbing client name
    			client_name = trim(client_name)
    			EMReadScreen next_revw_dt, 8, rept_row, 42						'grabbing next review date
    			next_revw_dt = replace(next_revw_dt, " ", "/")
    			objExcel.Cells(excel_row, 3).Value = client_name				'adding read variables to the spreadsheet
    			objExcel.Cells(excel_row, 4).Value = next_revw_dt				'adding read variables to the spreadsheet
    			excel_row = excel_row + 1
    		END IF
    		rept_row = rept_row + 1
    	LOOP UNTIL rept_row = 19								'looping until the script reads through the bottom of the page.
    	PF8														'pf8 navigates to next page of ACTV
    	rept_row = 7											'resetting the row to the top of the page.
    	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
    LOOP UNTIL last_page = "THIS IS THE LAST PAGE"
Next 

back_to_self
excel_row = 2												'resetting excel row so script can review each case number found in previous loops.
DO
    back_to_SELF
	MAXIS_case_number = objExcel.Cells(excel_row, 2).Value					'reading case number from excel spreadsheet
    Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending)
    objExcel.Cells(excel_row, msp_col).Value = msp_case
	CALL navigate_to_MAXIS_screen("ELIG", "HC")
	hhmm_row = 8	                                                         'setting starting point to review all HH members in ELIG HC
    notes = ""
    MaEPD_found = "" 														
	DO																		'the script will now navigate to ELIG HC and begin to search for MA caes with DP as the elig type.
		EMReadScreen hc_type, 2, hhmm_row, 28
		IF hc_type = "MA" THEN												'if it finds MA as the HC type it will go into those results
            EMReadScreen hh_memb_num, 2, hhmm_row, 3
			'Identifies unapproved versions of HC 
            EmReadscreen approval_code, 6, 8, 68
            If trim(approval_code) = "UNAPP" then Call write_value_and_transmit("D", hhmm_row, 26)
            
            Call write_value_and_transmit("X", hhmm_row, 26)
			'start looking for current month budget results
            row = 6
            col = 19
            EMSearch current_month, row, col 
            If row = 0 then 
                'For cases without current HC elig 
                objExcel.Cells(excel_row, reim_elig_col).Value = objExcel.Cells(excel_row, reim_elig_col).Value & ("MEMB " & hh_memb_num & " DOES NOT HAVE HC ELIG covering the current month. Review. ")  
                For col = 1 to 7
                    objExcel.Cells(excel_row, col).Interior.ColorIndex = 3	'Fills the row with red
                Next
            else 
                 'Chekcing for MA-EPD results in the current month
			    EMReadScreen elig_type, 2, 12, col - 2
			    IF elig_type <> "DP" THEN										'once in those HC results it will look for DP as the elig type. DP is for MA-EPD
                    'For cases without current MA-EPD
                    objExcel.Cells(excel_row, reim_elig_col).Value = objExcel.Cells(excel_row, reim_elig_col).Value & ("MEMB " & hh_memb_num & " NOT OPEN ON MA-EPD for "  & current_month & ". ")  
                    
                    MaEPD_found = False                                         'ELIG TYPE DID NOT = DP
    				PF3															'if the MA elig results don't have DP we end up here
    				hhmm_row = hhmm_row + 1	                                    'adding to the read row since we have finished evaluating this particular HH member.        
                else 
                    MaEPD_found = True 									  
                End if 
                
                If MaEPD_found = True then 
                    'Heading into the budget 
                    Call write_value_and_transmit("X", 9, col + 2)
				    EMReadScreen pct_fpg, 4, 18, 38								'here it will check the percert of FPG client is at.
				    pct_fpg = trim(pct_fpg)
                    If pct_fpg = "" then 
                        objExcel.Cells(excel_row, reim_elig_col).Value = objExcel.Cells(excel_row, reim_elig_col).Value & ("MEMB " & hh_memb_num & " Doesn't reflect FPG % in BSUM. Review.")  'writing eligibility status in spreadsheet
                        For col = 1 to 7
        	        		objExcel.Cells(excel_row, col).Interior.ColorIndex = 3	'Fills the row with red
        	            Next
                    Else     
                        pct_fpg = pct_fpg * 1
				        IF pct_fpg < 201 THEN										'If the client is 200% or under they may eligible for reimbursement
				        	PF3														'the script will now grab that person's member number and head into memb to get that person's PMI this will be used later to check MMIS
				        	PF3
                            'Grabbing the Medicare Part B premium 
                            'Getting PMI number 
				        	CALL navigate_to_MAXIS_screen("STAT", "MEMB")
				        	EMWriteScreen hh_memb_num, 20, 76
				        	transmit
				        	EMReadScreen cl_pmi, 8, 4, 46
				        	cl_pmi = replace(cl_pmi, " ", "")
				        	DO
				        		IF len(cl_pmi) <> 8 THEN cl_pmi = "0" & cl_pmi
				        	LOOP UNTIL len(cl_pmi) = 8
                            
                            'Grabbing the Medicare Part B premium 
                           Call navigate_to_MAXIS_screen("STAT", "MEDI")
                           Call write_value_and_transmit(hh_memb_num, 20, 76)
                           EmReadscreen medi_premium, 8, 7, 73
                           objExcel.Cells(excel_row, 6).Value = trim(replace(medi_premium, "_", ""))
                            
				        	Call navigate_to_MMIS_region("CTY ELIG STAFF/UPDATE")	'function to navigate into MMIS, select the HC realm, and enters the prior autorization area
				        	
                            DO
				        		EMReadScreen RKEY, 4, 1, 52
				        		IF RKEY <> "RKEY" THEN EMWaitReady 0, 0
				        	LOOP UNTIL RKEY = "RKEY"
				        	EMWriteScreen "I", 2, 19
				        	EMWriteScreen cl_pmi, 4, 19
				        	transmit
				        	EMWriteScreen "RELG", 1, 8
				        	transmit

				        	'Reading RELG to determine if the CL is active on MA-EPD
				        	EMReadScreen prog01_type, 8, 6, 13
				        		EMReadScreen elig01_type, 2, 6, 33
				        		EMReadScreen elig01_end, 8, 7, 36
				        	EMReadScreen prog02_type, 8, 10, 13
				        		EMReadScreen elig02_type, 2, 10, 33
				        		EMReadScreen elig02_end, 8, 11, 36
				        	EMReadScreen prog03_type, 8, 14, 13
				        		EMReadScreen elig03_type, 2, 14, 33
				        		EMReadScreen elig03_end, 8, 15, 36
				        	EMReadScreen prog04_type, 8, 18, 13
				        		EMReadScreen elig04_type, 2, 18, 33
				        		EMReadScreen elig04_end, 8, 19, 36

				        	IF ((prog01_type = "MEDICAID" AND elig01_type = "DP" AND elig01_end = "99/99/99") OR _
				        		(prog02_type = "MEDICAID" AND elig02_type = "DP" AND elig02_end = "99/99/99") OR _
				        		(prog03_type = "MEDICAID" AND elig03_type = "DP" AND elig03_end = "99/99/99") OR _
				        		(prog04_type = "MEDICAID" AND elig04_type = "DP" AND elig04_end = "99/99/99")) THEN

				        		EMWriteScreen "RMCR", 1, 8							'the script will now check RMCR for an active medicare case
				        		transmit

				        		'-----CHECKING FOR ON-GOING MEDICARE PART B-----
				        		EMReadScreen part_b_begin01, 8, 13, 4
				        			part_b_begin01 = trim(part_b_begin01)
				        		EMReadScreen part_b_end01, 8, 13, 15
				        		EMReadScreen part_b_begin02, 8, 14, 4
				        			part_b_begin02 = trim(part_b_begin02)
				        		EMReadScreen part_b_end02, 8, 14, 15

				        		IF (part_b_begin01 <> "" AND part_b_end01 = "99/99/99") THEN				'lastly the script will check RBYB to see what the client's buy in status is
				        			EMWriteScreen "RBYB", 1, 8
				        			transmit

				        			EMReadScreen accrete_date, 8, 5, 66
				        			EMReadScreen delete_date, 8, 6, 65
				        			accrete_date = replace(accrete_date, " ", "")
                                    
                                    EMReadScreen accrete_date_two, 8, 5, 33
                                    EMReadScreen delete_date_two, 8, 6, 32
                                    accrete_date_two = replace(accrete_date_two, " ", "")

				        			IF (accrete_date = "" and accrete_date_two = "") then
                                        reim_elig = True 
                                    elseif (accrete_date <> "" AND delete_date = "99/99/99") THEN				'if the PMI is found to be open on MA-EPD, under 200% open on medicare and they don't have an end date on the delete date (rbyb) the script marks them as eligible for reimbursement.
                                        reim_elig = false 
                                    else 
                                        'Some cases that have opened & closed on MSP OR new MSP cases will have the accrete/delete information here.
                                        IF (accrete_date_two = "") then 
                                            reim_elig = True 
                                        elseif (accrete_date_two <> "" AND isdate(delete_date_two) = true) then 
                                            reim_elig = True 'this handling is for cases that have closed on the buy in previously, but it has restarted
                                        else 
                                            reim_elig = False
                                        End if 
                                    End if 
                                
                                    If reim_elig = True then objExcel.Cells(excel_row, reim_elig_col).Value = objExcel.Cells(excel_row, reim_elig_col).Value & ("MEMB " & hh_memb_num & " ELIG FOR REIMBURSEMENT. " & notes)  'writing eligibility status in spreadsheet
				        			
				        			CALL write_value_and_transmit("RKEY", 1, 8)
				        		END IF
				        	ELSE
				        		CALL write_value_and_transmit("RKEY", 1, 8)
				        	END IF
				        	CALL navigate_to_MAXIS_test(production_or_inquiry)				'the script now navigates back to the environment the user left MAXIS in to continue searching Household members on the current case.
				        	hhmm_row = hhmm_row + 1
				        	CALL navigate_to_MAXIS_screen("ELIG", "HC")
				        ELSE
				        	DO
				        		EMReadScreen at_hhmm, 4, 3, 51						'making sure the script made it back to ELIG/HC
				        		IF at_hhmm <> "HHMM" THEN PF3
				        	LOOP UNTIL at_hhmm = "HHMM"
				        	hhmm_row = hhmm_row + 1									'adding to the read row since we have finished evaluating this particular HH member.
				        END IF
                    End if 
                End if 
			END IF
		ELSE
			hhmm_row = hhmm_row + 1											'If the elig/hc results aren't MA we end up here and add to the read row since we have finished evaluating this particular HH member.
		END IF
		IF hhmm_row = 20 THEN												'here we are determining that we've read all of the HH members on the current HHMM screen.
			PF8																'pf8 will cause elig hc to move to the next set of HH members if that page is full
			EMReadScreen this_is_the_last_page, 21, 24, 2					'if the script has read everyone on a page and PF8'd and reached the last page the script is done evaulating this case
		END IF
	LOOP UNTIL hc_type = "  " OR this_is_the_last_page = "THIS IS THE LAST PAGE"
    
	'Deleting the blank results to clean up the spreadsheet
	IF objExcel.Cells(excel_row, reim_elig_col).Value = "" THEN
		SET objRange = objExcel.Cells(excel_row, 1).EntireRow
		objRange.Delete
		excel_row = excel_row - 1
	END IF
	excel_row = excel_row + 1										'the script adds 1 to the excel row to move onto the next case to evaluate
LOOP UNTIL objExcel.Cells(excel_row, 2).Value = ""

FOR i = 1 to 7							'making the columns stretch to fit the widest cell
	objExcel.Columns(i).AutoFit()
NEXT

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success!!")