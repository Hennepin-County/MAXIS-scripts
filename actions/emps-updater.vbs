'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - UPDATE EMPS.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 258                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================

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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS==================================================================================================================
Function Generate_Client_List(list_for_dropdown)

	memb_row = 5

	Call navigate_to_MAXIS_screen ("STAT", "MEMB")
	Do
		EMReadScreen ref_numb, 2, memb_row, 3
		If ref_numb = "  " Then Exit Do
		EMWriteScreen ref_numb, 20, 76
		transmit
		EMReadScreen first_name, 12, 6, 63
		EMReadScreen last_name, 25, 6, 30
		client_info = client_info & "~" & ref_numb & " - " & replace(first_name, "_", "") & " " & replace(last_name, "_", "")
		memb_row = memb_row + 1
	Loop until memb_row = 20

	client_info = right(client_info, len(client_info) - 1)
	client_list_array = split(client_info, "~")

	For each person in client_list_array
		list_for_dropdown = list_for_dropdown & chr(9) & person
	Next

End Function

FUNCTION date_array_generator(initial_month, initial_year, date_array)
	'defines an intial date from the initial_month and initial_year parameters
	initial_date = initial_month & "/1/" & initial_year
	'defines a date_list, which starts with just the initial date
	date_list = initial_date
	'This loop creates a list of dates
	Do
		If datediff("m", date, initial_date) = 1 then exit do		'if initial date is the current month plus one then it exits the do as to not loop for eternity'
		working_date = dateadd("m", 1, right(date_list, len(date_list) - InStrRev(date_list,"|")))	'the working_date is the last-added date + 1 month. We use dateadd, then grab the rightmost characters after the "|" delimiter, which we determine the location of using InStrRev
		date_list = date_list & "|" & working_date	'Adds the working_date to the date_list
	Loop until datediff("m", date, working_date) = 1	'Loops until we're at current month plus one

	'Splits this into an array
	date_array = split(date_list, "|")
End function

'THE SCRIPT=================================================================================================================
'Connecting to MAXIS, and grabbing the case number and footer month'
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

UniversalParticipant = FALSE 		'Setting some boolean variables
ExtensionCase = FALSE
FSSCase = FALSE

If MAXIS_case_number <> "" Then 		'If a case number is found the script will get the list of
	Call Generate_Client_List(HH_Memb_DropDown)
End If

'Running the dialog for case number and client
Do
	err_msg = ""
	'Dialog defined here so the dropdown can be changed
	BeginDialog select_person_dialog, 0, 0, 191, 65, "Update FSS Information from the Status Update"
	  EditBox 55, 5, 50, 15, MAXIS_case_number
	  ButtonGroup ButtonPressed
	    PushButton 135, 5, 50, 15, "search", search_button
	  DropListBox 80, 25, 105, 45, "Select One..." & HH_Memb_DropDown, clt_to_update
	  ButtonGroup ButtonPressed
	    OkButton 115, 45, 35, 15
	    CancelButton 155, 45, 30, 15
	  Text 5, 10, 45, 10, "Case Number"
	  Text 5, 30, 70, 10, "Household member"
	EndDialog
	Dialog select_person_dialog
	If ButtonPressed = cancel Then StopScript
	If ButtonPressed = search_button Then
		If MAXIS_case_number = "" Then
			MsgBox "Cannot search without a case number, please try again."
		Else
			HH_Memb_DropDown = ""
			Call Generate_Client_List(HH_Memb_DropDown)
			err_msg = err_msg & "Start Over"
		End If
	End If
	If MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "You must enter a valid case number."
	If clt_to_update = "Select One..." Then err_msg = err_msg & vbNewLine & "Please pick a client whose EMPS panel you need to update."
	If err_msg <> "" AND left(err_msg, 10) <> "Start Over" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
Loop until err_msg = ""

clt_ref_num = left(clt_to_update, 2)	'Settin the reference number

Fin_Orient_Missing = FALSE 		'Setting variables
ES_referral_Missing = FALSE

Call navigate_to_MAXIS_screen ("STAT", "EMPS")		'Go to EMPS
EMWriteScreen clt_ref_num, 20, 76
transmit
EMReadScreen Fin_Orient_Dt, 8, 5, 39				'Reading and formatting the ES Referral Date and Financial Orientation Date
EMReadScreen ES_Referral_Dt, 8, 16, 40
If Fin_Orient_Dt = "__ __ __" then
	Fin_Orient_Dt = ""
	Fin_Orient_Missing = TRUE
Else
	Fin_Orient_Dt = replace(Fin_Orient_Dt, " ", "/")
End If
If ES_Referral_Dt = "__ __ __" Then
	ES_Referral_Dt = ""
	ES_referral_Missing = TRUE
Else
	ES_Referral_Dt = replace(ES_Referral_Dt, " ", "/")
End If

EMReadScreen ES_Status, 2, 15, 40					'Determining the ES status
ES_Status = abs(ES_Status)
If ES_Status = 20 Then
	UniversalParticipant = TRUE
ElseIf ES_Status < 20 Then
	ExtensionCase = TRUE
Else
	FSSCase = TRUE
End If

EMReadScreen care_of_baby, 1, 12, 76				'Determining if child under 1 is already being used or not
If care_of_baby = "N" Then Current_Using_Exemption = FALSE
If care_of_baby = "Y" Then Current_Using_Exemption = TRUE

baby_on_case = FALSE							'Defaults to false
Do
	Call Navigate_to_MAXIS_screen ("STAT", "PNLP")
	EMReadScreen nav_check, 4, 2, 53
Loop until nav_check = "PNLP"
maxis_row = 3
Do
	EMReadScreen panel_name, 4, maxis_row, 5	'Reads the name of each panel listed on PNLP
	If panel_name = "MEMB" Then 				'Looking for MEMB
		EMReadScreen client_age, 2, maxis_row, 71		'Reads the age on the MEMB line
		If client_age = " 0" Then
			baby_on_case = TRUE	'If a age is listed as 0 then a baby is on the case'
			EMReadScreen Baby_ref_numb, 2, 10, 10
		End If
	End If
	If panel_name = "MEMI" Then Exit Do			'Once it gets to a panel named MEMI, there are no additional MEMB panels
	maxis_row = maxis_row + 1					'Go to next row
	If maxis_row = 20 Then 						'If it gets to row 20 it needs to go to the next page
		transmit
		maxis_row = 3
	End If
Loop until panel_name = "REVW"
If baby_on_case = TRUE Then 		'If there is no baby on the case the script will not update to a child under 12 months exemption - this notifies the worker and unchecks the selector
	Call Navigate_to_MAXIS_screen ("STAT", "MEMB")
	EMWriteScreen Baby_ref_numb, 20, 76
	transmit
	EMReadScreen Baby_DOB, 10, 8, 42
	Baby_DOB = replace(Baby_DOB, " ", "/")
	Baby_is_One = DateAdd("yyyy", 1, Baby_DOB)
	Exemption_Unaavailable = DateAdd("m", 1, Baby_is_One)
	Exemption_End_Month = right("00" & DatePart("m", Exemption_Unaavailable), 2)
	Exemption_End_Year = DatePart("yyyy", Exemption_Unaavailable)
End If


Do
	err_msg = ""
	'This is a very dynamic dialog and gets recreated and resized as buttons are pushed
	dialog_length = 120
	IF ES_referral_Missing = TRUE Then dialog_length = dialog_length + 15
	IF Fin_Orient_Missing = TRUE Then dialog_length = dialog_length + 15
	IF EMPS_Workaround = TRUE Then dialog_length = dialog_length + 20
	IF Child_Under_One = TRUE Then dialog_length = dialog_length + 40
	IF Remove_FSS = TRUE Then dialog_length = dialog_length + 20

	y_pos = 25
	BeginDialog fss_code_detail, 0, 0, 370, dialog_length, "Update FSS Information from the Status Update"
	  Text 5, 10, 195, 10, "This script can update EMPS for the following proceedures:"

	  IF ES_referral_Missing = TRUE Then
		  Text 5, y_pos, 95, 10, "ES Referral Date is Missing:"
		  CheckBox 110, y_pos, 140, 10, "Check Here to have the script update to ", update_ES_ref_checkbox
		  EditBox 255, y_pos - 5, 50, 15, new_es_referral_dt
		  y_pos = y_pos + 15
	  End If
	  IF Fin_Orient_Missing = TRUE Then
		  Text 5, y_pos, 95, 10, "Fin Orient Date is Missing:"
		  CheckBox 110, y_pos, 140, 10, "Check Here to have the script update to ", update_fin_orient_checkbox
		  EditBox 255, y_pos - 5, 50, 15, new_fin_oreient_dt
		  y_pos = y_pos + 15
	  End If

	  ButtonGroup ButtonPressed
	    PushButton 10, y_pos, 185, 10, "Code EMPS to get MFIP results instead of DWP", Intake_MFIP_Button
	  Text 205, y_pos, 105, 10, "Workaround process for Intake"
	  y_pos = y_pos + 15
	  IF EMPS_Workaround = TRUE Then
		  y_pos = y_pos + 5
		  Text 10, y_pos, 65, 10, "Date of Application"
		  EditBox 80, y_pos - 5, 50, 15, date_of_app
		  y_pos = y_pos + 15
	  End If

	  ButtonGroup ButtonPressed
	    PushButton 10, y_pos, 185, 10, "Code EMPS for Child Under 12 Months Exemption", Child_Under_One_Button
	  Text 205, y_pos, 75, 10, "Adding or removing"
	  y_pos = y_pos + 15
	  IF Child_Under_One = TRUE Then
		  y_pos = y_pos + 5
		  If Current_Using_Exemption = FALSE Then Text 10, y_pos, 205, 10, "   It appears you need to ADD the exemption.  Reason:"
		  If Current_Using_Exemption = TRUE  Then Text 10, y_pos, 205, 10, "It appears you need to REMOVE the exemption.  Reason:"
		  DropListBox 220, y_pos - 5, 125, 45, "Select One..."+chr(9)+"Child Age"+chr(9)+"Caregiver request"+chr(9)+"MFIP results approve - complete workaround", child_under_one_reason
		  If Current_Using_Exemption = TRUE  Then
		  	Text 10, y_pos + 20, 75, 10, "First month to remove"
		  	EditBox 90, y_pos + 15, 15, 15, end_month
		  	EditBox 105, y_pos + 15, 15, 15, End_Year
		  End If
		  If Current_Using_Exemption = FALSE Then
		  	Text 135, y_pos + 20, 75, 10, "Date of Client request:"
		  	EditBox 220, y_pos + 15, 50, 15, client_request_date
		  End If
		  y_pos = y_pos + 40
	  End If

	  ButtonGroup ButtonPressed
	    PushButton 10, y_pos, 185, 10, "Code EMPS to remove FSS", Remove_FSS_Button
	  Text 205, y_pos, 125, 10, "Return Caregiver to Regular MFIP-ES"
	  y_pos = y_pos + 15
	  IF Remove_FSS = TRUE Then
		  Text 10, y_pos + 5, 165, 10, "First month client should be Universal Participant"
		  EditBox 180, y_pos, 15, 15, UP_month
		  EditBox 195, y_pos, 15, 15, UP_year
		  y_pos = y_pos +15
	  End If
	  Text 5, y_pos + 15, 40, 10, "Other Notes"
	  EditBox 50, y_pos + 10, 310, 15, other_notes
	  Text 5, y_pos + 35, 60, 10, "Worker Signature"
	  EditBox 75, y_pos + 30, 110, 15, worker_signature
	  ButtonGroup ButtonPressed
	    OkButton 255, y_pos + 30, 50, 15
	    CancelButton 310, y_pos + 30, 50, 15
	EndDialog

	Dialog fss_code_detail
	cancel_confirmation
	If ButtonPressed = Intake_MFIP_Button Then
		err_msg = err_msg & "Start Over"
		If baby_on_case = FALSE Then
			MsgBox "There is no Child Under One listed on this case, this is not the correct workaround to generate MFIP results."
			EMPS_Workaround = FALSE
		Else
			EMPS_Workaround = NOT(EMPS_Workaround)	'Switching the boolean
		End If
	End If
	If ButtonPressed = Child_Under_One_Button Then
		err_msg = err_msg & "Start Over"
		Child_Under_One = NOT(Child_Under_One)		'Switching the boolean
	End If
	If ButtonPressed = Remove_FSS_Button Then
		err_msg = err_msg & "Start Over"
		If FSSCase = FALSE Then
			MsgBox "This client is not coded as using FSS, and so cannot be removed."
			Remove_FSS = FALSE
		Else
			Remove_FSS = NOT(Remove_FSS)			'Switching the boolean
		End If
	End If

	If update_ES_ref_checkbox = checked AND IsDate(new_es_referral_dt) = FALSE Then err_msg = err_msg & vbNewLine & "You must enter the ES Referrak Date to enter into EMPS."
	If update_fin_orient_checkbox = checked AND IsDate(new_fin_oreient_dt) = FALSE Then err_msg = err_msg & vbNewLine & "You must enter the Financial Orientation Date to enter in EMPS."
	If EMPS_Workaround = TRUE AND IsDate(date_of_app) = FALSE Then err_msg = err_msg & vbNewLine & "You must enter the date of application for the script to update EMPS to generate MFIP results."
	If Child_Under_One = TRUE Then
		If child_under_one_reason = "Select One..." Then err_msg = err_msg & vbNewLine & "Select the reason to change the Child Under One Exemption."
		If Current_Using_Exemption = TRUE Then
		 	If end_month = "" OR End_Year = "" Then err_msg = err_msg & vbNewLine & "You must enter the first month that the child under 1 exemption should be removed from."
		ElseIf Current_Using_Exemption = FALSE Then
			If IsDate(client_request_date) = FALSE Then err_msg = err_msg & vbNewLine & "Enter the date te client requested the exemption. (If you are coding to generate MFIP results at application, use the top option)."
		End If
	End If
	If Remove_FSS = TRUE Then
		If UP_month = "" OR UP_year = "" Then err_msg = err_msg & vbNewLine & "Enter the month and year the client should become a universal participant and have FSS ended."
	End If
	If worker_signature = "" Then err_msg = err_msg & vbNewLine & "Please sign your case note."
	If err_msg <> "" AND left(err_msg, 10) <> "Start Over" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
Loop Until err_msg = ""

back_to_self

'Updating ES Referral date
If update_ES_ref_checkbox = checked Then
	Do
		Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
		EMReadScreen nav_check, 4, 2, 50
	Loop until nav_check = "EMPS"
	EMWriteScreen clt_ref_num, 20, 76
	transmit
	PF9													'Edit
	ref_month = right("00" & DatePart("m", new_es_referral_dt), 2)
	ref_date  = right("00" & DatePart("d", new_es_referral_dt), 2)
	ref_year  = right(DatePart("yyyy", new_es_referral_dt), 2)
	EMWriteScreen ref_month, 16, 40						'Write in the date
	EMWriteScreen ref_date,  16, 43
	EMWriteScreen ref_year,  16, 46
	transmit
	back_to_self
End If

'Updating the Financial orientation date
If update_fin_orient_checkbox = checked Then
	Do
		Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
		EMReadScreen nav_check, 4, 2, 50
	Loop until nav_check = "EMPS"
	EMWriteScreen clt_ref_num, 20, 76
	transmit
	PF9
	ref_month = right("00" & DatePart("m", new_fin_oreient_dt), 2)
	ref_date  = right("00" & DatePart("d", new_fin_oreient_dt), 2)
	ref_year  = right(DatePart("yyyy", new_fin_oreient_dt), 2)
	EMWriteScreen ref_month, 5, 39
	EMWriteScreen ref_date,  5, 42
	EMWriteScreen ref_year,  5, 45
	EMWriteScreen "y", 5, 65
	transmit
	back_to_self
End If

'MFIP workaround for intake
'The process here is to code a case with a child under 1 to get MFIP results instead of DWP
'The EMPS panel is changed back after approval unless the client actually requests the exemption
If EMPS_Workaround = TRUE then
	MAXIS_footer_month = right("00" & DatePart("m", date_of_app), 2)		'Update in the month of application
	MAXIS_footer_year = right(DatePart("yyyy", date_of_app), 2)
	Do
		Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
		EMReadScreen nav_check, 4, 2, 50
	Loop until nav_check = "EMPS"
	EMWriteScreen clt_ref_num, 20, 76
	transmit
	PF9
	EMWriteScreen "X", 12, 39							'Open the list of exemption months already taken
	transmit
	emps_row = 7										'Setting the first row and col
	emps_col = 22
	Do
		EMReadScreen month_used, 2, emps_row, emps_col	'reading the first field
		If month_used = "__" Then Exit Do				'if the month was listed as blank, there are no more months listed
		EMReadScreen year_used, 4, emps_row, emps_col + 5		'reads the year associated with the month listed
		emps_exemption_month_used = emps_exemption_month_used & "~" & month_used & "/" & year_used	'adds the month and year to a string seperated by ~
		emps_col = emps_col + 11						'moves to the next month listed spot
		If emps_col = 66 Then 							'Once it has gone through all the fields on this row, it goes to the next row and starts over at the beginning of the columns.
			emps_col = 22
			emps_row = emps_row + 1
		End If
	Loop Until emps_row = 10							'There are only 3 rows of data
	If emps_exemption_month_used <> "" Then
		emps_exemption_month_used = right(emps_exemption_month_used, len(emps_exemption_month_used)-1)	'lops off the extra ~ at the beginning
		used_expemption_months_array = split(emps_exemption_month_used, "~")							'creates an array for the counting
		months_used = Join(used_expemption_months_array, ", ")										'creates a string of months used for case noting
		number_of_months_available = 12 - (ubound(used_expemption_months_array) + 1) & ""				'uses the ubound of the array to determine how many months are left to be used
	Else
		months_used = "NONE"
		number_of_months_available = 12
	End If

	confirm_proceed_msg = MsgBox ("It appears this client has " & number_of_months_available & " months available for the exemption." &_
	   vbNewLine & vbNewLine & "The script will update EMPS so that the client is using months starting in: " &  MAXIS_footer_month & "/" & MAXIS_footer_year, vbOKCancel + vbQuestion, "Child under 1 Months")

	If confirm_proceed_msg = vbOK Then
		Call date_array_generator (MAXIS_footer_month, MAXIS_footer_year, workaround_month_array)	'Will update from month of app to CM + 1

		emps_row = 7												'setting the first location
		emps_col = 22
		Do
			EMReadScreen month_used, 2, emps_row, emps_col			'finding the first blank month to code
			If month_used = "__" Then Exit Do
			emps_col = emps_col + 11
			If emps_col = 66 Then
				emps_col = 22
				emps_row = emps_row + 1
			End If
		Loop Until emps_row = 10
		IF emps_row = 10 Then 										'if there are no blank months then error - cannot code an exemption
			MsgBox "It appears the client has used all of their Exempt Months. EMPS will need to be updated manually."
			PF3
			PF10
		Else
			For each exempt_month in workaround_month_array				'writing each of the months to be exempt in the array into the popup
				EMWriteScreen right("00" & DatePart("m", exempt_month), 2), emps_row, emps_col
				EMWriteScreen right(DatePart("yyyy", exempt_month), 4), emps_row, emps_col + 5
				emps_col = emps_col + 11
				If emps_col = 66 Then
					emps_col = 22
					emps_row = emps_row + 1
				End If
			Next
			PF3
		End IF

		EMWriteScreen "Y", 12, 76		'Coding the EMPS panel with Yes
		transmit

		'Going to TIKL
		Call Navigate_to_MAXIS_screen("DAIL", "WRIT")
		tikl_date = date
		EMWriteScreen right("00" & DatePart("m", tikl_date), 2), 5, 18
		EMWriteScreen right("00" & DatePart("d", tikl_date), 2), 5, 21
		EMWriteScreen right(tikl_date, 2), 5, 24

		tikl_msg = "*** CHANGE EMPS BACK After approval of MFIP results."

		Call Write_Variable_in_TIKL(tikl_msg)
		EMReadScreen tikl_confirm, 4, 24, 2
		If tikl_confirm <> "    " Then workaround_tikl_fail = TRUE
		PF3
	Else
		workaround_aborted = TRUE
	End if
	back_to_self
End if

'Coding for actual use of the exemption
If Child_Under_One = TRUE Then
	If Current_Using_Exemption = FALSE Then 		'This means that we need to add the exemption
		Do
			Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
			EMReadScreen nav_check, 4, 2, 50
		Loop until nav_check = "EMPS"
		EMWriteScreen clt_ref_num, 20, 76
		transmit
		EMWriteScreen "X", 12, 39							'Open the list of exemption months already taken
		transmit
		emps_row = 7										'Setting the first row and col
		emps_col = 22
		Do
			EMReadScreen month_used, 2, emps_row, emps_col	'reading the first field
			If month_used = "__" Then Exit Do				'if the month was listed as blank, there are no more months listed
			EMReadScreen year_used, 4, emps_row, emps_col + 5		'reads the year associated with the month listed
			emps_exemption_month_used = emps_exemption_month_used & "~" & month_used & "/" & year_used	'adds the month and year to a string seperated by ~
			emps_col = emps_col + 11						'moves to the next month listed spot
			If emps_col = 66 Then 							'Once it has gone through all the fields on this row, it goes to the next row and starts over at the beginning of the columns.
				emps_col = 22
				emps_row = emps_row + 1
			End If
		Loop Until emps_row = 10							'There are only 3 rows of data
		If emps_exemption_month_used <> "" Then
			emps_exemption_month_used = right(emps_exemption_month_used, len(emps_exemption_month_used)-1)	'lops off the extra ~ at the beginning
			used_expemption_months_array = split(emps_exemption_month_used, "~")							'creates an array for the counting
			months_used = Join(used_expemption_months_array, ", ")										'creates a string of months used for case noting
			number_of_months_available = 12 - (ubound(used_expemption_months_array) + 1) & ""				'uses the ubound of the array to determine how many months are left to be used
		Else
			months_used = "NONE"
			number_of_months_available = 12
		End If

		For add_month = 1 to number_of_months_available		'using the count determined in the EMPS
			this_month = DatePart("m", DateAdd ("m", add_month, client_request_date))	'first month is the month after the exemption is requested, then adding all the others after'
			If len(this_month) = 1 Then this_month = "0" & this_month		'making 2 digit
			this_year = DatePart("yyyy", DateAdd("m", add_month, client_request_date))	'creating a year
			If trim(this_month) = trim(Exemption_End_Month) AND trim(this_year) = trim(Exemption_End_Year) Then Exit For
			new_exemption_months = new_exemption_months & "~" & this_month & "/" & this_year	'list of all of these months
		Next
		If new_exemption_months <> "" Then
			new_exemption_months = right(new_exemption_months, len(new_exemption_months) - 1)		'taking off the extra ~
			new_exemption_months_array = split(new_exemption_months, "~")							'creating an array of the months to code for future exempt months
			months_to_fill = Join(new_exemption_months_array, ", ")									'list for the edit box
			Impose_Exemption = TRUE
		Else
			months_to_fill = "None available."
			MsgBox "It appears the baby on this case will turn one before the Child Under One Exemption can be put into place. If you contine the script with this date as the request date, this exemption will not be coded. Otherwise review the request date."
			Impose_Exemption = FALSE
		End If

		confirm_proceed_msg = MsgBox ("It appears this client has " & number_of_months_available & " months available for the exemption." &_
		   vbNewLine & vbNewLine & "The script will update EMPS so that the client is using months starting in: " &  left(new_exemption_months_array(0), 2) & "/" & right(new_exemption_months_array(0), 2), vbOKCancel + vbQuestion, "Child under 1 Months")

		If confirm_proceed_msg = vbOK Then
			back_to_self

			last_month = left(new_exemption_months_array(ubound(new_exemption_months_array)), 2)
			last_year = right(new_exemption_months_array(ubound(new_exemption_months_array)), 2)
			child_under_one_tikl_date = last_month & "/01/" & last_year
			MAXIS_footer_month = left(new_exemption_months_array(0), 2)		'getting footer month by using the array of months to be exempt
			MAXIS_footer_year = right(new_exemption_months_array(0), 2)

			Do
				Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
				EMReadScreen nav_check, 4, 2, 50
			Loop until nav_check = "EMPS"
			EMWriteScreen clt_ref_num, 20, 76
			transmit

			PF9
			EMWriteScreen "X", 12, 39							'Open the list of exemption months already taken
			transmit

			emps_row = 7												'setting the first location
			emps_col = 22
			Do
				EMReadScreen month_used, 2, emps_row, emps_col			'finding the first blank month to code
				If month_used = "__" Then Exit Do
				emps_col = emps_col + 11
				If emps_col = 66 Then
					emps_col = 22
					emps_row = emps_row + 1
				End If
			Loop Until emps_row = 10
			IF emps_row = 10 Then 										'if there are no blank months then error - cannot code an exemption
				MsgBox "It appears the client has used all of their Exempt Months. EMPS will need to be updated manually."
				PF3
				PF10
			Else
				For each exempt_month in new_exemption_months_array				'writing each of the months to be exempt in the array into the popup
					EMWriteScreen left(exempt_month, 2), emps_row, emps_col
					EMWriteScreen right(exempt_month, 4), emps_row, emps_col + 5
					emps_col = emps_col + 11
					If emps_col = 66 Then
						emps_col = 22
						emps_row = emps_row + 1
					End If
				Next
				PF3
			End IF

			EMWriteScreen "Y", 12, 76			'EMPS to yes
			transmit

			'Going to TIKL
			Call Navigate_to_MAXIS_screen("DAIL", "WRIT")
			EMWriteScreen right("00" & DatePart("m", child_under_one_tikl_date), 2), 5, 18
			EMWriteScreen right("00" & DatePart("d", child_under_one_tikl_date), 2), 5, 21
			EMWriteScreen right(child_under_one_tikl_date, 2), 5, 24

			tikl_msg = "Review Child Under 12 Months Exemption, appears to be ending/ended. A new MFIP approval may be needed."

			Call Write_Variable_in_TIKL(tikl_msg)
			EMReadScreen tikl_confirm, 4, 24, 2
			If tikl_confirm <> "    " Then add_child_under_one_tikl_fail = TRUE
			PF3
		Else
			add_child_under_one_aborted = TRUE
		End if
		back_to_self
	ElseIf Current_Using_Exemption = TRUE Then 		'This is for if the exemption needs to be ended
		MAXIS_footer_month = right("00" & end_month, 2)		'Update in the first month that it will be removed.
		MAXIS_footer_year = right("00" & End_Year, 2)
		Do
			Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
			EMReadScreen nav_check, 4, 2, 50
		Loop until nav_check = "EMPS"
		EMWriteScreen clt_ref_num, 20, 76
		transmit
		PF9
		EMWriteScreen "N", 12, 76
		EMWriteScreen "X", 12, 39							'Open the list of exemption months already taken
		transmit

		emps_row = 7												'setting the first location
		emps_col = 22
		Do
			EMReadScreen month_used, 2, emps_row, emps_col			'finding where the first month to remove is
			If month_used = MAXIS_footer_month Then
				EMReadScreen year_used, 2, emps_row, emps_col + 7
				If year_used = MAXIS_footer_year Then
					start_row = emps_row
					start_col = emps_col
					Exit Do
				End If
			Else
				emps_col = emps_col + 11
				If emps_col = 66 Then
					emps_col = 22
					emps_row = emps_row + 1
				End If
			End If
		Loop Until emps_row = 10
		IF emps_row = 10 Then 										'if there are no blank months then error - cannot code an exemption
			MsgBox "It appears the client has used all of their Exempt Months. EMPS will need to be updated manually."
			PF3
			PF10
		Else
			del_row = start_row							'Once found, will blank out that one and all future
			del_col = start_col
			Do
				EMWriteScreen "  ", del_row, del_col
				EMWriteScreen "    ", del_row, del_col + 5
				del_col = del_col + 11
				If del_col = 66 Then
					del_col = 22
					del_row = del_row + 1
				End If
			Loop until del_row = 10
		End If
		back_to_self
	End If
End If

'THIS bit JUST does the EMPS portion of ending FSS.
If Remove_FSS = TRUE Then
	Do 		'Go to EMPS
		Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
		EMReadScreen nav_check, 4, 2, 50
	Loop until nav_check = "EMPS"
	EMWriteScreen clt_ref_num, 20, 76
	transmit

	EMReadScreen child_under_one_code, 1, 12, 76				'Handling to force users to use the specific coding for the child under 12 months
	If child_under_one_code = "Y" Then
		MsgBox "Use the Option to remove the Child Under 12 Months Exemption for this case."
		remove_fss_aborted = TRUE
	Else 														'Otherwise updating the fields for FSS
		PF9

		EMWriteScreen "N", 8, 76
		EMWriteScreen "N", 9, 76
		EMWriteScreen "N", 10, 76
		EMWriteScreen "NO", 11, 76
		EMWriteScreen "N", 13, 76
		EMWriteScreen "Y", 14, 76

		transmit
	End If
	back_to_self
End If

'Message about if a process did not complete and allows worker to avoid the case note for a non completed process
aborted_msg = ""
If workaround_aborted = TRUE Then aborted_msg = aborted_msg & vbNewLine & "You chose to stop the update of EMPS for the MFIP Intake workaround."
If add_child_under_one_aborted = TRUE Then aborted_msg = aborted_msg & vbNewLine & "You chose to stop the update of EMPS to add a Child Under 12 Months Exemption"
If aborted_msg <> "" Then
	aborted_msg = "The script did not take all the requested actions because:" & vbNewLine & aborted_msg & vbNewLine & vbNewLine & "Do you still want the script to case note?"
	continue_and_case_note_msg = MsgBox(aborted_msg, vbYesNo + vbAlert, "Case Note?")
	If continue_and_case_note_msg = vbNo Then script_end_procedure("ERROR: Script stopped after updates were aborted.")
End If

'Case Note
Call Navigate_to_MAXIS_screen ("CASE", "NOTE")
Call start_a_blank_CASE_NOTE
Call Write_Variable_in_CASE_NOTE ("EMPS Updated")
Call Write_Variable_in_CASE_NOTE ("* Updated EMPS for Member " & clt_ref_num)
If update_ES_ref_checkbox = checked Then Call Write_Variable_in_CASE_NOTE ("* Updated ES Referral Date. Entered: " & new_es_referral_dt)
If update_fin_orient_checkbox = checked Then Call Write_Variable_in_CASE_NOTE("* Updated Financial Orientation Date. Entered: " & new_fin_oreient_dt)
If EMPS_Workaround = TRUE AND workaround_aborted <> TRUE Then
	Call Write_Variable_in_CASE_NOTE("* Updated EMPS for Intake processing. There is a child under 1 in the household and case should not be DWP. This is the DHS provided workaround.")
	If workaround_tikl_fail <> TRUE Then Call Write_Variable_in_CASE_NOTE ("* TIKL set for today to remove coding after approval.")
End If
IF Child_Under_One = TRUE Then
	If Current_Using_Exemption = FALSE AND add_child_under_one_aborted <> TRUE Then
		Call Write_Variable_in_CASE_NOTE ("* Added Child under 12 months exemption coding starting " & new_exemption_months_array(0))
		Call Write_Bullet_and_Variable_in_Case_Note ("Reason for ending", child_under_one_reason)
		If add_child_under_one_tikl_fail <> TRUE Then Call Write_Variable_in_CASE_NOTE ("* TIKL set for " & new_exemption_months_array(ubound(new_exemption_months_array)) & " to review exemption.")
	End if
	If Current_Using_Exemption = TRUE Then
		Call Write_Variable_in_CASE_NOTE ("* Ended the child under 12 month exemption. Exemption removed eff " & end_month & "/" & End_Year)
		Call Write_Bullet_and_Variable_in_Case_Note ("Reason for ending", child_under_one_reason)
	End if
End If
If Remove_FSS = TRUE Then Call Write_Variable_in_CASE_NOTE("* Updated EMPS only to return clt to regular MFIP from an FSS status. Clt will be Universal Participant eff " & UP_month & "/" & UP_year)
Call Write_Bullet_and_Variable_in_Case_Note ("Notes", other_notes)
Call Write_Variable_in_CASE_NOTE ("---")
Call Write_Variable_in_CASE_NOTE (worker_signature)

script_end_procedure("Success! EMPS has been updated and case noted. You may need to approve new MFIP resuls for the changes to take effect.")
