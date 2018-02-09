'Required for statistical purposes===============================================================================
name_of_script = "ACTIONS - ADD GRH RATE 2 TO MMIS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 900                      'manual run time in seconds
STATS_denomination = "C"       				'C is for each CASE
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
call changelog_update("02/08/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

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

Function MMIS_panel_check(panel_name)
	Do 
		EMReadScreen panel_check, 4, 1, 51
		If panel_check <> panel_name then Call write_value_and_transmit(panel_name, 1, 8)
	Loop until panel_check = panel_name
End function

function navigate_to_MMIS_region(group_security_selection)
'--- This function is to be used when navigating to MMIS from another function in BlueZone (MAXIS, PRISM, INFOPAC, etc.)
'~~~~~ group_security_selection: region of MMIS to access - programed options are "CTY ELIG STAFF/UPDATE", "GRH UPDATE", "GRH INQUIRY", "MMIS MCRE"
'===== Keywords: MMIS, navigate
	attn
	Do
		EMReadScreen MAI_check, 3, 1, 33
		If MAI_check <> "MAI" then EMWaitReady 1, 1
	Loop until MAI_check = "MAI"

	EMReadScreen mmis_check, 7, 15, 15
	IF mmis_check = "RUNNING" THEN
		EMWriteScreen "10", 2, 15
		transmit
	ELSE
		EMConnect"A"
		attn
		EMReadScreen mmis_check, 7, 15, 15
		IF mmis_check = "RUNNING" THEN
			EMWriteScreen "10", 2, 15
			transmit
		ELSE
			EMConnect"B"
			attn
			EMReadScreen mmis_b_check, 7, 15, 15
			IF mmis_b_check <> "RUNNING" THEN
				script_end_procedure("You do not appear to have MMIS running. This script will now stop. Please make sure you have an active version of MMIS and re-run the script.")
			ELSE
				EMWriteScreen "10", 2, 15
				transmit
			END IF
		END IF
	END IF

	DO
		PF6
		EMReadScreen password_prompt, 38, 2, 23
		IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then 
			Do 
				CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	 		Loop until are_we_passworded_out = false					'loops until user passwords back in
		End if 
		EMReadScreen session_start, 18, 1, 7
	LOOP UNTIL session_start = "SESSION TERMINATED"

	'Getting back in to MMIS and trasmitting past the warning screen (workers should already have accepted the warning when they logged themselves into MMIS the first time, yo.
	EMWriteScreen "MW00", 1, 2
	transmit
	transmit

	group_security_selection = UCASE(group_security_selection)

	EMReadScreen MMIS_menu, 24, 3, 30
	If MMIS_menu <> "GROUP SECURITY SELECTION" Then
		EMReadScreen mmis_group_selection, 4, 1, 65
		EMReadScreen mmis_group_type, 4, 1, 57

		correct_group = FALSE

		Select Case group_security_selection

		Case "CTY ELIG STAFF/UPDATE"
			mmis_group_selection_part = left(mmis_group_selection, 2)

			If mmis_group_selection_part = "C3" Then correct_group = TRUE
			If mmis_group_selection_part = "C4" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the County Eligibility region. The script will now stop.")

		Case "GRH UPDATE"
			If mmis_group_selection  = "GRHU" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the GRH Update region. The script will now stop.")

		Case "GRH INQUIRY"
			If mmis_group_selection  = "GRHI" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the GRH Inquiry region. The script will now stop.")

		Case "MMIS MCRE"
			If mmis_group_selection  = "EK01" Then correct_group = TRUE
			If mmis_group_selection  = "EKIQ" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the MCRE region. The script will now stop.")

		End Select

	Else
		Select Case group_security_selection

		Case "CTY ELIG STAFF/UPDATE"
			row = 1
			col = 1
			EMSearch " C3", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch " C4", row, col
				If row <> 0 Then
					EMWriteScreen "x", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the County Eligibility area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "GRH UPDATE"
			row = 1
			col = 1
			EMSearch "GRHU", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				script_end_procedure("You do not appear to have access to the GRH area of MMIS, this script requires access to this region. The script will now stop.")
			End If

			'Now it finds the pror authorization application feature and selects it.
			row = 1
			col = 1
			EMSearch "PRIOR AUTHORIZATION   ", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "GRH INQUIRY"
			row = 1
			col = 1
			EMSearch "GRHI", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				script_end_procedure("You do not appear to have access to the GRH Inquiry area of MMIS, this script requires access to this region. The script will now stop.")
			End If

			'Now it finds the pror authorization application feature and selects it.
			row = 1
			col = 1
			EMSearch "PRIOR AUTHORIZATION   ", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "MMIS MCRE"
			row = 1
			col = 1
			EMSearch "EK01", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch "EKIQ", row, col
				If row <> 0 Then
					EMWriteScreen "x", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the MCRE area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "x", row, col - 3
			transmit
		End Select
	End If
end function

'----------------------------------------------------------------------------------------------------DIALOG
BeginDialog case_number_dialog, 0, 0, 251, 155, "MMIS Single Case Conversion"
  ButtonGroup ButtonPressed
    OkButton 100, 50, 50, 15
    CancelButton 155, 50, 50, 15
  EditBox 100, 10, 50, 15, MAXIS_case_number
  EditBox 100, 30, 105, 15, worker_signature
  GroupBox 5, 75, 240, 75, "MMIS Single Case Conversion script:"
  Text 10, 90, 230, 25, "This script is to be used when a GRH Rate 2 case needs to be converted into MMIS intitally. Currently this script only supports cases that were in a Rate 2 facility prior to 02/01/18. Additional tools to come. "
  Text 10, 125, 225, 20, "* Before you use the script, you must have approved GRH results that reflect the SSR information in the SSR pop-up on ELIG/GRFB."
  Text 45, 15, 50, 10, "Case Number:"
  Text 35, 35, 60, 10, "Worker signature:"
EndDialog

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

'Main dialog: user will input case number and initial month/year will default to current month - 1 and member 01 as member number
DO
	DO
		err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
		dialog case_number_dialog				'main dialog
		If buttonpressed = 0 THEN stopscript	'script ends if cancel is selected
		IF len(MAXIS_case_number) > 8 or isnumeric(MAXIS_case_number) = false THEN err_msg = err_msg & vbCr & "Enter a valid case number."		'mandatory field
		If trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "Enter your worker signature."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

DIM Update_MMIS_array()
ReDim Update_MMIS_array(14, 0)

'constants for array
const tot_units 	= 0
const vendor_num 	= 1
const NPI_num 		= 2
const in_date 		= 3
const out_date	 	= 4	
const app_county 	= 5	
const addr_01	 	= 6
const addr_02 		= 7
const addr_city  	= 8
const addr_state 	= 9	
const addr_zip	 	= 10
const SS_rate 		= 11
const auth_number  	= 12
const case_status 	= 13

'const case_number	= 0
'const clt_PMI	 	= 7
'const clt_DOB		= 8
'const update_MMIS 	= 16

MAXIS_footer_month = "02"	'establishing footer month/year 
MAXIS_footer_year = "18"
start_date = "02/01/2018"   'start and end service agreement dates
end_date = "01/31/2019"
total_units = datediff("d", start_date, end_date) + 1
'msgbox total_units

ReDim Preserve Update_MMIS_array(14, 0)	'This resizes the array based on the number of rows in the Excel File'
Update_MMIS_array(tot_units, 	0) = total_units				'STATIC for now. TODO: remove static coding for action script 
Update_MMIS_array(vendor_num, 	0) = ""
Update_MMIS_array(NPI_num, 		0) = ""
Update_MMIS_array(in_date, 		0) = ""
Update_MMIS_array(out_date,		0) = ""
Update_MMIS_array(app_county, 	0) = ""
Update_MMIS_array(addr_01, 		0) = ""
Update_MMIS_array(addr_02, 		0) = ""
Update_MMIS_array(addr_city, 	0) = ""
Update_MMIS_array(addr_state, 	0) = ""
Update_MMIS_array(addr_zip, 	0) = ""
Update_MMIS_array(SS_rate, 		0) = ""
Update_MMIS_array(auth_number,  0) = ""
Update_MMIS_array(case_status,  0) = ""

back_to_self
call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year

call navigate_to_MAXIS_screen("STAT", "PROG")
EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
If PRIV_check = "PRIV" then
	'msgbox "PRIV case, cannot access/update."
	Update_MMIS = False 	
	script_end_procedure("PRIV case, cannot access/update.")
	'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
	Do
		back_to_self
		EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
		If SELF_screen_check <> "SELF" then PF3
	LOOP until SELF_screen_check = "SELF"
	EMWriteScreen "________", 18, 43		'clears the MAXIS case number
	transmit
Else
	EMReadScreen grh_status, 4, 9, 74		'Ensuring that the case is active on GRH. If not, case will not be updated in MMIS. 
	If grh_status <> "ACTV" then
        If trim(grh_status) = "" then grh_status = "Inactive"
		'msgbox "GRH case status is " & grh_status
		Update_MMIS = False
		script_end_procedure("GRH case status is " & grh_status & ".") 	  
	Else 
		Update_MMIS = True 	
	End if
End if 

Call HCRE_panel_bypass			'Function to bypass a janky HCRE panel. If the HCRE panel has fields not completed/'reds up' this gets us out of there. 

'----------------------------------------------------------------------------------------------------DISA: ensuring that client is not on a waiver. If they are, they should not be rate 2.
If Update_MMIS = True then  	
	Call navigate_to_MAXIS_screen("STAT", "DISA")
	Call write_value_and_transmit ("01", 20, 76)	'For member 01 - All GRH cases should be for member 01. 
	EMReadScreen waiver_type, 1, 14, 59
	If waiver_type <> "_" then 
		'msgbox "Client is active on a waiver. Should not be Rate 2."
		Update_MMIS = False
        script_end_procedure("Client is active on a waiver. Should not be Rate 2.")
	Else
	 	Update_MMIS = True 	
	End if 
End if 	
	
'----------------------------------------------------------------------------------------------------FACI: Finding the current FACI, and ensuring they are rate 2	
If Update_MMIS = True then  	
	Call navigate_to_MAXIS_screen("STAT", "FACI")
	EMReadScreen member_number, 2, 4, 33
	If member_number <> "01" then 
		EmWriteScreen "01", 20, 76						'For member 01 - All GRH cases should be for member 01. 
		Call write_value_and_transmit ("01", 20, 79)	'1st version of FACI 
	End if 
    
	'LOOKS FOR MULTIPLE STAT/FACI PANELS, GOES TO THE MOST RECENT ONE
    EMReadScreen FACI_total_check, 1, 2, 78
	If FACI_total_check = "0" then 
		'msgbox "Case does not have a FACI panel."
		Update_MMIS = False
		current_faci = False 
		script_end_procedure("Case does not have a FACI panel.")
	Else 
        row = 14
        Do 
        	EMReadScreen date_out, 10, row, 71
	    	'msgbox "date out: " & date_out 
        	If date_out = "__ __ ____" then 
	    	  	EMReadScreen grh_rate, 1, row, 34
	    	  	If grh_rate = "2" then 
	    	  		'msgbox grh_rate
					Update_MMIS = True 
	    	  		current_faci = TRUE
					EMReadScreen date_in, 10, row, 47
					EMReadScreen faci_vendor_number, 8, 5, 43
					EMReadScreen approval_county, 2, 12, 71
					approval_county = "0" & approval_county
					
					Update_MMIS_array(in_date, item) = date_in			'Adding the FACI information to the array 
					Update_MMIS_array(out_date, item) = date_out
					Update_MMIS_array(app_county, item) = approval_county
        	  		exit do
	    	  	ELSE
					current_faci = False 
					row = row + 1
				End if 
			Else 
        		row = row + 1
				'msgbox row
				current_faci = False	
			End if 	
        	If row = 19 then 
        		transmit
        		row = 14
        	End if 
			EMReadScreen last_panel, 5, 24, 2
        Loop until last_panel = "ENTER"	'This means that there are no other faci panels
	End if 
    
    If current_faci = False then 
		'msgbox "Unable to find FACI panel or current FACI is not Rate 2."
		Update_MMIS = False
        script_end_procedure("Unable to find FACI panel or current FACI is not Rate 2.")
	Else 
		'----------------------------------------------------------------------------------------------------SSRT: ensuring that a panel exists, and the FACI dates match.
    	Call navigate_to_MAXIS_screen ("STAT", "SSRT")				
		call write_value_and_transmit ("01", 20, 76)	'For member 01 - All GRH cases should be for member 01. 
		
		EMReadScreen SSRT_total_check, 1, 2, 78
		If SSRT_total_check = "0" then
			'msgbox "SSRT panel needs to be created."
            Update_MMIS = False
			script_end_procedure("SSRT panel needs to be created.")
		Else
			'Matching the FACI panel vendor number to the SSRT panel vendor number 
			Do 
			    EMReadScreen SSRT_vendor_number, 8, 5, 43		'Enters vendor number 
			    If SSRT_vendor_number = faci_vendor_number then 
			    	SSRT_found = true
					exit do	'when the correct  
				else
				 	SSRT_found = False
				End if  
				transmit
			    EMReadScreen last_panel, 5, 24, 2
        	Loop until last_panel = "ENTER"	'This means that there are no other faci panels
			
			If SSRT_found = False then 
				'msgbox "SSRT panel could not be found for the Rate 2 facility."
                Update_MMIS = False
				script_end_procedure("SSRT panel could not be found for the Rate 2 facility.")
			Else 
				EMReadScreen NPI_number, 10, 7, 43
				If trim(NPI_number) = "" then 
					'msgbox "No NPI number on SSRT panel."
                    Update_MMIS = False
					script_end_procedure("No NPI number on SSRT panel.")
				Else 
			        row = 10 
			        Do 
			        	EMReadScreen ssrt_in_date, 10, row, 47
			        	If ssrt_in_date = date_in then
							EMReadScreen ssrt_out_date, 10, row, 71
							If ssrt_out_date = date_out then 
								'msgbox NPI_number
								'msgbox SSRT_vendor_number
								'trimming the leading 0's from the SSRT vendor numbers. This will be needed to measure against ELIG/GRH 
								If left(SSRT_vendor_number, 1) = "0" then 
									Do 
										SSRT_vendor_number = right(SSRT_vendor_number, len(SSRT_vendor_number) - 1) 
									Loop until left(SSRT_vendor_number, 1) <> "0"
								End if  
								'msgbox SSRT_vendor_number
								Update_MMIS_array(vendor_num, item) = SSRT_vendor_number
								Update_MMIS_array(NPI_num, item) = NPI_number
								Update_MMIS = True
								exit do 
							else 
								'msgbox "FACI dates and SSRT dates do not match."
                                Update_MMIS = False
								script_end_procedure("FACI dates and SSRT dates do not match. Please review these panels for accuracy.")
						  	End if 
						Else 
							'msgbox "Unable to find matching FACI and SSRT dates."
							Update_MMIS = False
							script_end_procedure("Unable to find matching FACI and SSRT dates.")
							row = row + 1
						End if 	
					Loop until row = 15	
				End if 	
			End if         		
		End if
    End if
	
	If Update_MMIS = True then
	 	Call navigate_to_MAXIS_screen("STAT", "MEMB")
		EMReadScreen client_PMI, 8, 4, 46
		client_PMI = trim(client_PMI)
		client_PMI = right("00000000" & client_pmi, 8)
		'Update_MMIS_array(clt_PMI, item) = client_pmi
		
		EMReadScreen client_DOB, 10, 8, 42
		client_DOB = replace(client_DOB, " ", "")
		'Update_MMIS_array(clt_DOB, item) = client_DOB
		'msgbox client_PMI & vbcr & client_DOB
		
		Call navigate_to_MAXIS_screen("STAT", "ADDR")
		EMReadScreen mailing_addr_check, 22, 13, 43
		mailing_addr_check = replace(mailing_addr_check, "_", "")
		If trim(mailing_addr_check) = "" then 
			EMReadScreen addr_line_01, 22, 6, 43
			EMReadScreen addr_line_02, 22, 7, 43
			EMReadScreen city_line, 15, 8, 43
			EMReadScreen state_line, 2, 8, 66
			EMReadScreen zip_line, 5, 9, 43
		Else 
			EMReadScreen addr_line_01, 22, 13, 43
			EMReadScreen addr_line_02, 22, 14, 43
			EMReadScreen city_line, 15, 15, 43
			EMReadScreen state_line, 2, 16, 43
			EMReadScreen zip_line, 5, 16, 52
		End if 
		
		addr_line_01 = replace(addr_line_01, "_", "")
		addr_line_02 = replace(addr_line_02, "_", "")
		city_line = replace(city_line, "_", "")
		
		'msgbox addr_line_01 & vbcr & addr_line_02 & vbcr & city_line & vbcr & state_line & vbcr & zip_line
		'Adding the address information to the array 
		Update_MMIS_array(addr_01, 		item) = addr_line_01
		Update_MMIS_array(addr_02, 		item) = addr_line_02
		Update_MMIS_array(addr_city, 	item) = city_line
		Update_MMIS_array(addr_state, 	item) = state_line
		Update_MMIS_array(addr_zip, 	item) = zip_line
		
		'----------------------------------------------------------------------------------------------------VNDS/VND2
		Call Navigate_to_MAXIS_screen("MONY", "VNDS")
		Call write_value_and_transmit(SSRT_vendor_number, 4, 59)
		Call write_value_and_transmit("VND2", 20, 70)
		EMReadScreen service_rate, 8, 16, 68		'Reading the service rate to input into MMIS
        If IsNumeric(service_rate) = False then EMReadScreen service_rate, 8, 15, 72        'Handling for vendors with Rate 3 information 
		service_rate = replace(service_rate, ".", "")	'removing the period for input into MMIS
		Update_MMIS_array(SS_rate, item) = trim(service_rate)
		
		'msgbox service_rate
		'----------------------------------------------------------------------------------------------------ELIG/GRH
		Call Navigate_to_MAXIS_screen("ELIG", "GRH ")
		EMReadScreen no_grh, 10, 24, 2		'NO GRH version means no conversion to MMIS will take place
		If no_grh = "NO VERSION" then
			'msgbox "No GRH eligibility results."						
			Update_MMIS = False
			script_end_procedure("There are no GRH eligibility results. Please review.")
		Else 
		    Call write_value_and_transmit("99", 20, 79)
		    'This brings up the FS versions of eligibility results to search for approved versions
		    status_row = 7
		    Do
		    	EMReadScreen app_status, 8, status_row, 50
		    	If trim(app_status) = "" then
					'msgbox "There are no GRH eligibility results for " & MAXIS_footer_month & "/" & MAXIS_footer_year
					Update_MMIS = False
					script_end_procedure("There are no GRH eligibility results for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ".")
		    		PF3
		    		exit do 	'if end of the list is reached then exits the do loop
		    	End if
		    	If app_status = "UNAPPROV" Then status_row = status_row + 1
				IF app_status = "APPROVED" then
					EMReadScreen vers_number, 1, status_row, 23
					Call write_value_and_transmit(vers_number, 18, 54)
					'msgbox vers_number
					exit do
			 	End if 
		    Loop until app_status = "APPROVED" or trim(app_status) = ""
		End if 	
		
		If app_status <> "APPROVED" then 
			'msgbox "There are no approved GRH eligibility results for " & MAXIS_footer_month & "/" & MAXIS_footer_year
			Update_MMIS = False	
			script_end_procedure("There are no approved GRH eligibility results for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ".")
		Else
			'----------------------------------------------------------------------------------------------------ELIG/GRFB
			Call write_value_and_transmit("GRFB", 20, 71)
			Call write_value_and_transmit("x", 11, 3)
			'Ensuring a rate 2 is found. If none or more than one are found, MMIS will not be updated.
			EMReadScreen rate_two_check, 8, 15, 8
			rate_two_check = Trim(rate_two_check)
			If rate_two_check = "" then 
				'msgbox "GRH eligibility doesn't reflect rate 2 vendor information."
				Update_MMIS = False	
				script_end_procedure("GRH eligibility doesn't reflect Rate 2 vendor information. The SSR pop-up in ELIG/GRFB must reflect the Rate 2 vendor information. Please review.")
			ElseIf rate_two_check = SSRT_vendor_number then 
				EMReadScreen second_faci, 8, 16, 8
				If trim(second_faci) <> "" then 
					'msgbox "More than one vendor exists in ELIG/GRFB. Process manually."
					Update_MMIS = False	
					script_end_procedure("More than one vendor exists in ELIG/GRFB. Process manually.")
				Else 
					Update_MMIS = True 
				End if
			else 
				'msgbox "SSRT vendor number did not match ELIG/GRFB vendor number."
				Update_MMIS = False 
				script_end_procedure("SSRT vendor number did not match ELIG/GRFB vendor number.")	
			End if 	
		End if 
	End if 	
 End if
  
'msgbox "Going into MMIS"
'----------------------------------------------------------------------------------------------------MMIS portion of the script
If Update_MMIS = True then
	'msgbox "True"
	Call navigate_to_MMIS_region("GRH UPDATE")	'function to navigate into MMIS, select the GRH update realm, and enter the prior autorization area
	Call MMIS_panel_check("AKEY")				'ensuring we are on the right MMIS screen

    EmWriteScreen client_PMI, 10, 36
    Call write_value_and_transmit("C", 3, 22)	'Checking to make sure that more than one agreement is not listed by trying to change (C) the information for the PMI selected.
	EMReadScreen active_agreement, 12, 24, 2
	'msgbox active_agreement & vbcr & MAXIS_case_number
    If active_agreement <> "NO DOCUMENTS" then 
		EMReadScreen AGMT_status, 31, 3, 19 
		AGMT_status = trim(AGMT_status)
		If AGMT_status = "START DT:        END DT:" then 
			AGMT_info = "More than one service agreement exists in MMIS."
		Else 
			AGMT_info = AGMT_status & " agreement already exists in MMIS."
		End if 
    	Update_MMIS = False	
    	script_end_procedure(AGMT_info)
		PF3
    Elseif active_agreement = "NO DOCUMENTS" then   
		Call clear_line_of_text(10, 36) 	'clears out the PMI number. Cannot add new agreement with PMI listed on AKEY.
		'msgbox "did the PMI get cleared?"
		EmWriteScreen "A", 3, 22					'Selects the action code (A)
		EmWriteScreen "T", 3, 71					'Selecs the service agreement option (T)
		'msgbox "AKEY, is everything but the 2 entered?"
		Call write_value_and_transmit("2", 7, 77)	'Enters the agreement type and transmits
		
		'----------------------------------------------------------------------------------------------------ASA1 screen
		Call MMIS_panel_check("ASA1")				'ensuring we are on the right MMIS screen
		EmWriteScreen "020118", 4, 64				'Start date is static for the BULK conversion. TODO: change to date_in which will match the FACI dates. 
		EmWriteScreen "013119", 4, 71				'End date is static for the BULK conversion. TODO: change to date_out which will match the FACI dates. 
		EmWriteScreen client_PMI, 8, 64							'Enters the client's PMI 
		EmWriteScreen client_DOB, 9, 19							'Enters the client's DOB 
		EmWriteScreen approval_county, 11, 19					'Enters 3 digit CO of SVC
		EmWriteScreen approval_county, 11, 39					'Enters 3 digit CO of RES
		'msgbox "ASA1, is everything but CO of FIN entered?"
		Call write_value_and_transmit(approval_county, 11, 64)	'Enters 3 digit CO of FIN RESP and transmits
		
		Call MMIS_panel_check("ASA2")				'ensuring we are on the right MMIS screen
		transmit 	'no action required on ASA2
		'----------------------------------------------------------------------------------------------------ASA3 screen
		Call MMIS_panel_check("ASA3")				'ensuring we are on the right MMIS screen	
		EMWriteScreen "H0043", 7, 36
		EMWriteScreen "U5", 7, 44
		EMWriteScreen "020118", 8, 60				'Start date is static for the BULK conversion. TODO: change to date_in which will match the FACI dates. 
		EMWriteScreen "013119", 8, 67				'End date is static for the BULK conversion. TODO: change to date_out which will match the FACI dates. 
		EMWriteScreen service_rate, 9, 20			'Enters service rate from VND2 
		EMWriteScreen total_units, 9, 60 			'Enters the difference between the start date and end date. TODO: update this coding to work with date_in and date_out after BULK conversion.
		'msgbox "ASA3, is everything but the NPI entered?"
		If NPI_number = "1801986773" then NPI_number = "A767410200"
        If NPI_number = "A096405300" then NPI_number = "A904695300"
        If NPI_number = "A346627201" then NPI_number = "A346627200"
        If NPI_number = "A346627203" then NPI_number = "A346627200"
        If NPI_number = "A346627204" then NPI_number = "A346627200"
        If NPI_number = "A690048500" then NPI_number = "A590048500"
        If NPI_number = "A952618400" then NPI_number = "A186688300"
        
        Call write_value_and_transmit(NPI_number, 10, 20)	'Enters the NPI number then transmits 
		Emreadscreen NPI_issue, 26, 24, 1
		If NPI_issue = "CORRECT HIGHLIGHTED FIELDS" then 
			Update_MMIS = False	
			script_end_procedure("Issue with NPI# in MMIS. Please review case/report issue to the Quality Improvement Team.")
			Call clear_line_of_text(10, 20) 	'clears out the NPI number so that the rest of the information can be saved. 
			PF3
		else 
		    '----------------------------------------------------------------------------------------------------PPOP screen handling
		    EMReadScreen PPOP_check, 4, 1, 52
		    If PPOP_check = "PPOP" then 
		    	'needs to serach for the facility that is associated with GRH facilities
		    	Do 
		    	    row = 1
		    	    col = 1
		    	    EMSearch "SPEC: GR", row, col		'Checking for "SPEC: GR"
		    	    If row <> 0 Then
		    	    	EMWriteScreen "x", row -1, 2	'Selects the correct facility found on the previous row. 
		    			'msgbox "Did the right facility get selected?"
		    			exit do 
		    	    Else 
		    	    	PF8		'going to the next screen if not found on the 1st screen 
		    	    End if
		    	Loop
		    	transmit							'To select match 
				transmit 							'to ACF1. No action required on ACF3.
			End if 
		    
		    '----------------------------------------------------------------------------------------------------ACF1 screen 
		    Call MMIS_panel_check("ACF1")		'ensuring we are on the right MMIS screen
		    EmWriteScreen addr_line_01, 5, 8	'enters the clients address 
		    EmWriteScreen addr_line_02, 5, 37
		    EmWriteScreen city_line, 6, 8
		    EmWriteScreen state_line, 6, 34
		    EmWriteScreen zip_line, 6, 42
		    'msgbox "ACF1, is address entered?"
		    Call write_value_and_transmit("ASA1", 1, 8)		'direct navigating to ASA1
		    
		    '----------------------------------------------------------------------------------------------------ASA1 screen 
		    Call MMIS_panel_check("ASA1")		'ensuring we are on the right MMIS screen
 		    PF9 								'triggering stat edits 	
		    EmreadScreen error_codes, 79, 20, 2	'checking for stat edits
		    If trim(error_codes) <> "00 140  4          01 140  4" then 
		    	'msgbox error_codes
		    	Update_MMIS = False	
		    	script_end_procedure("MMIS stat edits exist. Edit codes are: " & error_codes & vbcr & "PF3 to save what's been updated in MMIS, and report error codes to The Quality Improvement Team.")
		    	'figure out the rest of the steps here. 
		    else 
		    	EMWriteScreen "A", 3, 17						'Updating the AMT type/STAT to A for approved 
		    	Call write_value_and_transmit("ASA3", 1, 8)		'direct navigating to ASA3
		    	Call MMIS_panel_check("ASA3")					'ensuring we are on the right MMIS screen
		    	EMWriteScreen "A", 12, 19						'Updating the STAT CD/DATE to A for approved 
		    	Update_MMIS = true
		        PF3 '	to save changes 

		        Call MMIS_panel_check("AKEY")		'ensuring we are on the right MMIS screen
		        EMReadScreen authorization_number, 13, 9, 36
		        authorization_number = trim(authorization_number)
		        EMReadscreen approval_message, 16, 24, 2
		        If IsNumeric(authorization_number) = True then Update_MMIS_array(auth_number, item) = authorization_number
		    End if 
		End if 
	End if 		
End if     

If Update_MMIS = True then
    ''----------------------------------------------------------------------------------------------------MAXIS 
    Do 
        Call navigate_to_MAXIS("")
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
    
    'msgbox "Are we in MAXIS again?"
    Call navigate_to_MAXIS_screen("CASE", "NOTE")
    '----------------------------------------------------------------------------------------------------CASE NOTE
    Call start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE("GRH Rate 2 SSR updateed in MMIS for NPI #" & npi_number)
    Call write_bullet_and_variable_in_CASE_NOTE("MMIS autorization number", authorization_number)
    Call write_variable_in_CASE_NOTE("* A new GRH span has been created length of certification 2/1/19-1/31/19. This does not change the case's recertification date in MAXIS.")
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)
    PF3
End if 

script_end_procedure("Success! Your case has been updated in MMIS and case noted in MAXIS.")