'This Script generates a OUT OF STATE INQUIRY form in use to fax to the out of state agency.
name_of_script = "NOTICES - OUT OF STATE INQUIRY.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 52         'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("06/28/2016", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""
'Hunts for Maxis case number to autofill it
Call MAXIS_case_number_finder(MAXIS_case_number)

'Error proof functions
Call check_for_MAXIS(true)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 111, 105, "Out of State Inquiry"
  EditBox 55, 5, 50, 15, MAXIS_case_number
  DropListBox 55, 25, 50, 15, "Select One:"+chr(9)+"Sent"+chr(9)+"Received"+chr(9)+"Unknown", out_of_state_request
  DropListBox 55, 45, 50, 15, "Select One:"+chr(9)+"Email"+chr(9)+"Fax"+chr(9)+"Mail"+chr(9)+"Phone", how_sent
  DropListBox 55, 65, 50, 15, "Select One:"+chr(9)+"Alabama"+chr(9)+"Alaska"+chr(9)+"Arizona"+chr(9)+"Arkansas"+chr(9)+"California"+chr(9)+"Colorado"+chr(9)+"Connecticut"+chr(9)+"Delaware"+chr(9)+"Florida"+chr(9)+"Georgia"+chr(9)+"Hawaii"+chr(9)+"Idaho"+chr(9)+"Illinois"+chr(9)+"Indiana"+chr(9)+"Iowa"+chr(9)+"Kansas"+chr(9)+"Kentucky"+chr(9)+"Louisiana"+chr(9)+"Maine"+chr(9)+"Maryland"+chr(9)+"Massachusetts"+chr(9)+"Michigan"+chr(9)+"Mississippi"+chr(9)+"Missouri"+chr(9)+"Montana"+chr(9)+"Nebraska"+chr(9)+"Nevada"+chr(9)+"New Hampshire"+chr(9)+"New Jersey"+chr(9)+"New Mexico"+chr(9)+"New York"+chr(9)+"North Carolina"+chr(9)+"North Dakota"+chr(9)+"Ohio"+chr(9)+"Oklahoma"+chr(9)+"Oregon"+chr(9)+"Pennsylvania"+chr(9)+"Rhode Island"+chr(9)+"South Carolina"+chr(9)+"South Dakota"+chr(9)+"Tennessee"+chr(9)+"Texas"+chr(9)+"Utah"+chr(9)+"Vermont"+chr(9)+"Virginia"+chr(9)+"Washington"+chr(9)+"West Virginia"+chr(9)+"Wisconsin"+chr(9)+"Wyoming", agency_state_droplist
  ButtonGroup ButtonPressed
    OkButton 10, 85, 45, 15
    CancelButton 60, 85, 45, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 20, 30, 30, 10, "Request:"
  Text 20, 50, 30, 10, "Via(How):"
  Text 30, 70, 20, 10, "State:"
EndDialog

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If agency_state_droplist = "Select One:" then err_msg = err_msg & vbnewline & "* Select the state."
		If how_sent = "Select One:" then err_msg = err_msg & vbnewline & "* Select how the request was sent."
		If out_of_state_request = "Select One:" then err_msg = err_msg & vbNewLine & "Please select the status of the out of state inquiry."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false

'Error proof functions
Call check_for_MAXIS(False)

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Error proof functions
Call check_for_MAXIS(False)

'changing footer dates to current month to avoid invalid months.
MAXIS_footer_month = datepart("M", date)
IF Len(MAXIS_footer_month) <> 2 THEN MAXIS_footer_month = "0" & MAXIS_footer_month
MAXIS_footer_year = right(datepart("YYYY", date), 2)

Dim Member_Info_Array()
Redim Member_Info_Array(UBound(HH_member_array), 4)

'Navigate to stat/memb and check for ERRR message
CALL navigate_to_MAXIS_screen("STAT", "MEMB")
For i = 0 to Ubound(HH_member_array)
	Member_Info_Array(i, 0) = HH_member_array(i)
	'Navigating to selected memb panel
	EMwritescreen HH_member_array(i), 20, 76
	transmit
	EMReadScreen no_MEMB, 13, 8, 22 'If this member does not exist, this will stop the script from continuing.
	IF no_MEMB = "Arrival Date:" THEN script_end_procedure("This HH member does not exist.")
	'Reading info and removing spaces
	EMReadscreen First_name, 12, 6, 63
	First_name = replace(First_name, "_", "")
	Member_Info_Array(i, 1) = First_name
	'Reading Last name and removing spaces
	EMReadscreen Last_name, 25, 6, 30
	Last_name = replace(Last_name, "_", "")
	Member_Info_Array(i, 2) = Last_name
	'Reading Middle initial and replacing _ with a blank if empty.
	EMReadscreen Middle_initial, 1, 6, 79
	Middle_initial = replace(Middle_initial, "_", "")
	Member_Info_Array(i, 3) = Middle_initial
	'Reads SSN
	Emreadscreen SSN_number, 11, 7, 42
	SSN_number = replace(SSN_number, " ", "")
	Member_Info_Array(i, 4) = SSN_number
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
Next

'Navigate back to self and to EDRS
Back_to_self
EMWriteScreen MAXIS_case_number, 18, 43
Call navigate_to_MAXIS_screen("CASE", "CURR")
EMReadScreen CURR_panel_check, 4, 2, 55
EMReadScreen case_status, 8, 8, 9
case_status = trim(case_status)

IF case_status = "ACTIVE" THEN active_status = TRUE
IF case_status = "APP OPEN" THEN active_status = TRUE
IF case_status = "APP CLOS" THEN active_status = TRUE
IF case_status = "INACTIVE" THEN active_status = FALSE
If case_status = "CAF2 PEN" THEN active_status = TRUE
If case_status = "CAF1 PEN" THEN active_status = TRUE
IF case_status = "REIN" THEN active_status = TRUE

Call MAXIS_footer_month_confirmation
EmReadscreen original_MAXIS_footer_month, 2, 20, 43
'msgbox original_MAXIS_footer_month
CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG
'Checking for PRIV cases.
EMReadScreen priv_check, 4, 24, 14 'If it can't get into the case needs to skip
IF priv_check = "PRIV" THEN script_end_procedure_with_error_report("This case is privileged. Please request access before running the script again. ")
CASH_STATUS = FALSE 'overall variable'
CCA_STATUS = FALSE
DWP_STATUS = FALSE 'Diversionary Work Program'
ER_STATUS = FALSE
FS_STATUS = FALSE
GA_STATUS = FALSE 'General Assistance'
GRH_STATUS = FALSE
HC_STATUS = FALSE
MS_STATUS = FALSE 'Mn Suppl Aid '
MF_STATUS = FALSE 'Mn Family Invest Program '
RC_STATUS = FALSE 'Refugee Cash Assistance'

'Reading the status and program
EMReadScreen cash1_status_check, 4, 6, 74
'MsgBox  cash1_status_check
EMReadScreen cash2_status_check, 4, 7, 74
'MsgBox cash2_status_check
EMReadScreen emer_status_check, 4, 8, 74
'MsgBox emer_status_check
EMReadScreen grh_status_check, 4, 9, 74
'MsgBox grh_status_check
EMReadScreen fs_status_check, 4, 10, 74
EMReadScreen ive_status_check, 4, 11, 74
EMReadScreen hc_status_check, 4, 12, 74
EMReadScreen cca_status_check, 4, 14, 74
'MsgBox cca_status_check
EMReadScreen cash1_prog_check, 2, 6, 67
'MsgBox cash1_prog_check
EMReadScreen cash2_prog_check, 2, 7, 67
'MsgBox cash2_prog_check
EMReadScreen emer_prog_check, 2, 8, 67
EMReadScreen grh_prog_check, 2, 9, 67
EMReadScreen fs_prog_check, 2, 10, 67
EMReadScreen ive_prog_check, 2, 11, 67
EMReadScreen hc_prog_check, 2, 12, 67

IF FS_status_check = "ACTV" or FS_status_check = "PEND" THEN
	FS_STATUS = TRUE
	FS_CHECKBOX = CHECKED
END IF

IF hc_status_check = "ACTV" or hc_status_check = "PEND" THEN
	HC_STATUS = TRUE
	HC_CHECKBOX   = CHECKED
END IF

IF cca_status_check = "ACTV" or cca_status_check = "PEND" THEN
	CCA_STATUS = TRUE
	CCA_CHECKBOX  = CHECKED
END IF
'Logic to determine if MFIP is active
IF cash1_prog_check = "MF" THEN
	IF cash1_status_check = "ACTV" or cash1_status_check = "PEND" THEN
		MF_STATUS = TRUE
		'MsgBox MF_STATUS
		MFIP_CHECKBOX = CHECKED
		'MsgBox MFIP_CHECKBOX
	END IF
	If cash1_status_check = "INAC" or cash1_status_check = "SUSP" or cash1_status_check = "DENY" or cash1_status_check = "" THEN MF_STATUS = FALSE
END IF

IF cash1_prog_check = "MF" THEN
	IF cash2_status_check = "ACTV" or cash2_status_check = "PEND" THEN
		MF_STATUS = TRUE
		MFIP_CHECKBOX = CHECKED
	END IF
	IF cash2_status_check = "INAC" or cash2_status_check = "SUSP" or cash2_status_check = "DENY" or cash2_status_check = "" THEN MF_STATUS = FALSE
END IF

IF cash1_prog_check = "DW" THEN
	IF cash1_status_check = "ACTV" or cash1_status_check = "PEND" or cash2_status_check = "ACTV" or cash2_status_check = "PEND" THEN
		DWP_STATUS = TRUE
		DWP_CHECKBOX  = CHECKED
	END IF
	If cash1_status_check = "INAC" or cash1_status_check = "SUSP" or cash1_status_check = "DENY" or cash1_status_check = "" THEN DWP_STATUS = FALSE
	If cash2_status_check = "INAC" or cash2_status_check = "SUSP" or cash2_status_check = "DENY" or cash2_status_check = "" THEN DWP_STATUS = FALSE
END IF

If cash1_prog_check = "" THEN
	If cash1_status_check = "PEND" or cash2_status_check = "PEND" THEN
		CASH_STATUS = TRUE
		CASH_CHECKBOX = CHECKED
	END IF
	If cash1_status_check = "INAC" or cash1_status_check = "SUSP" or cash1_status_check = "DENY" or cash1_status_check = "" THEN CASH_STATUS = FALSE
END IF

If cash2_prog_check = "" THEN
	If cash2_status_check = "INAC" or cash2_status_check = "SUSP" or cash2_status_check = "DENY" or cash2_status_check = "" THEN CASH_STATUS = FALSE
END IF

IF emer_status_check = "ACTV" or emer_status_check = "PEND"  THEN ER_STATUS = TRUE
IF grh_status_check = "ACTV" or grh_status_check = "PEND"  THEN GRH_STATUS = TRUE
'can you say and or
IF active_status = FALSE THEN
 	IF MF_STATUS = FALSE and FS_STATUS = FALSE and HC_STATUS = FALSE and DWP_STATUS = FALSE and CASH_STATUS = FALSE THEN
		case_note_only = TRUE
		msgbox "It appears no HC, FS, or Cash are open on this case."
	END IF
END IF

'State information for dialog and notice'
IF FS_CHECKBOX = CHECKED THEN contact_email = "fs@dhr.alabama.gov"

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 231, 230, "OUT OF STATE INQUIRY" &  agency_state_droplist
  CheckBox 50, 20, 25, 10, "Cash", Cash_CHECKBOX
  CheckBox 80, 20, 55, 10, "Commodities", Commodities_CHECKBOX
  CheckBox 135, 20, 55, 10, "Food Support", FS_CHECKBOX
  CheckBox 195, 20, 25, 10, "HC", HC_CHECKBOX
  DropListBox 40, 35, 55, 15, "Select One:"+chr(9)+"Active"+chr(9)+"Closed"+chr(9)+"Unknown", out_of_state_status
  EditBox 160, 35, 55, 15, out_of_state_date
  EditBox 50, 175, 170, 15, other_notes
  PushButton 5, 195, 60, 15, "HSR MANUAL", outofstate_button
  Text 10, 70, 205, 30, "Name:"
  Text 10, 105, 205, 10, "Address:"
  Text 10, 120, 205, 10, "Email:"
  Text 10, 135, 100, 10, "Phone:"
  CheckBox 10, 150, 130, 10, "Different information for contact state", update_state_info_checkbox
  Text 115, 135, 90, 10, "Fax:"
  ButtonGroup ButtonPressed
    OkButton 135, 195, 40, 15
    CancelButton 180, 195, 40, 15
  Text 105, 40, 50, 10, "Last Received:"
  GroupBox 5, 60, 215, 105, "Out of State Agency Contact"
  Text 10, 20, 40, 10, "Programs:"
  Text 10, 40, 25, 10, "Status:"
  Text 5, 180, 45, 10, "Other Notes:"
  Text 15, 215, 185, 10, "*** Reminder: ECF must show verification requested ***"
  GroupBox 5, 5, 215, 50, "Client reported they received assistance (Q5 on CAF):"
EndDialog


'Dialog
DO      'Password DO loop
    DO  'Conditional handling DO loop
        DO  'External resource DO loop
            Dialog Dialog1
            cancel_confirmation
            If ButtonPressed = outofstate_button then CreateObject("WScript.Shell").Run("https://dept.hennepin.us/hsphd/manuals/hsrm/Pages/Out_of_State_Inquiry.aspx")
        Loop until ButtonPressed = -1
        err_msg = ""
		If agency_state_droplist = "Select One:" then err_msg = err_msg & vbnewline & "* Select the state."
        If trim(out_of_state_date) = "" then err_msg = err_msg & vbcr & "* Enter the date the client reported benefits were received."
		'If trim(agency_name) = "" then err_msg = err_msg & vbcr & "* Enter the out of state agency name."
        'If trim(agency_address) = "" then err_msg = err_msg & vbcr & "* Enter the out of state agency address, if there is not one provided enter N/A."
		'If trim(agency_email) = "" then err_msg = err_msg & vbcr & "* Enter the out of state agency email, if there is not one provided enter N/A."
		'If trim(agency_phone) = "" then err_msg = err_msg & vbcr & "* Enter the out of state agency phone, if there is not one provided enter N/A."
		'If trim(agency_fax) = "" then err_msg = err_msg & vbcr & "* Enter the out of state agency fax, if there is not one provided enter N/A."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    LOOP until err_msg = ""
    CALL check_for_password(are_we_passworded_out)                                 'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false                                                                          'loops until user passwords back in

'Goes to MEMB to get info
Call navigate_to_MAXIS_screen("STAT", "MEMB")
'Goes to the right HH member
EMWriteScreen MEMB_number, 20, 76 'It does this to make sure that it navigates to the right HH member.
transmit 'This transmits to STAT/MEMB for the client indicated.

'If this member does not exist, this will stop the script from continuing.
EMReadScreen no_MEMB, 13, 8, 22
If no_MEMB = "Arrival Date:" then script_end_procedure("Error! This HH member does not exist.")

'Reads the SSN pieces
EMReadScreen SSN1, 3, 7, 42
EMReadScreen SSN2, 2, 7, 46
EMReadScreen SSN3, 4, 7, 49
client_ssn = SSN1 & "-" & SSN2 & "-" & SSN3

'Reads Client's DOB
EMReadScreen DOB1, 2, 8, 42
EMReadScreen DOB2, 2, 8, 45
EMReadScreen DOB3, 4, 8, 48
client_dob = DOB1 & "/" & DOB2 & "/" & DOB3

'Reads clients name and coverts to a Variant
EMReadScreen last_name, 24, 06, 30
EMReadScreen first_name, 12, 06, 63
last_name = replace(last_name, "_", "")
first_name = replace(first_name, "_","")
client_name = first_name & " " & last_name

'this reads current mailing address
Call navigate_to_MAXIS_screen("STAT", "ADDR")
EMReadScreen mail_address, 1, 13, 64
If mail_address = "_" then
     EMReadScreen client_1staddress, 21, 06, 43
     EMReadScreen client_2ndaddress, 21, 07, 43
     EMReadScreen client_city, 14, 08, 43
     EMReadScreen client_state, 2, 08, 66
     EMReadScreen client_zip, 7, 09, 43
Else
     EMReadScreen client_1staddress, 21, 13, 43
     EMReadScreen client_2ndaddress, 21, 14, 43
     EMReadScreen client_city, 14, 15, 43
     EMReadScreen client_state, 2, 16, 43
     EMReadScreen client_zip, 7, 16, 52
End If
client_address = replace(client_1staddress, "_","") & " " & replace(client_2ndaddress, "_","") & " " & replace(client_city, "_","") & ", " & replace(client_state, "_","") & " " & replace(client_zip, "_","")


'reads county info.'
EMReadScreen worker_county, 4, 21, 21
If worker_county = "X127" then
hennepin_county = true
Else
hennepin_county = false
End If

'reads assigned worker info
EMSetCursor 21, 21
PF1
EMReadScreen worker_name, 21, 19, 10
EMReadScreen worker_phone, 12, 19, 45
transmit

If hennepin_county = true then
'Generates Word Doc Form
Set objWord = CreateObject("Word.Application")
objWord.Caption = "OUT OF STATE INQUIRY"
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
objSelection.ParagraphFormat.Alignment = 0
objSelection.ParagraphFormat.LineSpacing = 12
objSelection.ParagraphFormat.SpaceBefore = 0
objSelection.ParagraphFormat.SpaceAfter = 0
objSelection.Font.Name = "New York Times"
objSelection.Font.Size = "12"
objSelection.TypeText "OUT OF STATE INQUIRY"
objSelection.TypeParagraph
objSelection.TypeText "Hennepin County Human Services & Public Health Department"
objSelection.TypeParagraph
objSelection.TypeText "PO Box 107, Minneapolis, MN 55440-0107"
objSelection.TypeParagraph
objSelection.TypeText "FAX: 612-288-2981"
objSelection.TypeParagraph
objSelection.TypeText "Phone: 612-596-8500"
objSelection.TypeParagraph

objSelection.ParagraphFormat.Alignment = 2
objSelection.ParagraphFormat.LineSpacing = 12
objSelection.ParagraphFormat.SpaceBefore = 0
objSelection.ParagraphFormat.SpaceAfter = 0
objSelection.Font.Name = "New York Times"
objSelection.Font.Size = "11"
objSelection.TypeText "DATE: " & date()

objSelection.TypeParagraph
objSelection.ParagraphFormat.Alignment = 0
objSelection.Font.Size = "10"
'objSelection.Font.Bold = True
objSelection.TypeText "To: " & agency_name
objSelection.TypeParagraph
objSelection.TypeText "Address: " & agency_address
objSelection.TypeParagraph
objSelection.TypeText "Email: " & agency_email
objSelection.TypeParagraph
objSelection.TypeText "Phone: " & agency_phone
objSelection.TypeParagraph
objSelection.TypeText "Fax: " & agency_fax
objSelection.TypeParagraph
objSelection.TypeText " "
objSelection.TypeParagraph
objSelection.TypeText "RE: " & client_name
objSelection.TypeParagraph
objSelection.TypeText "SSN: " & client_ssn & "			DOB: " & client_dob
objSelection.TypeParagraph
objSelection.TypeText "Current Address: " & client_address
objSelection.TypeParagraph
objSelection.TypeParagraph
objSelection.TypeText "Our records indicate that the above individual received or receives assistance from your state.  We need to verify the number of months of Federally-funded TANF cash assistance issued by your state that count towards the 60 month lifetime limit.  In addition, we need to know the number of months of TANF assistance from other states that your agency has verified.  "
objSelection.TypeText "Please indicate if the client is open on SNAP or Medical Assistance in your state OR the date these programs most recently closed.  Thank you."
objSelection.TypeParagraph

objSelection.TypeParagraph
objSelection.TypeText "Is CASH currently closed?   YES	 NO		Date of closure: "
objSelection.TypeParagraph
objSelection.TypeText "Is SNAP currently closed?   YES	 NO		Date of closure: "
objSelection.TypeParagraph
objSelection.TypeText "Total ABAWD months used:"
objSelection.TypeParagraph
objSelection.TypeText "Please list the month(s)/year(s) of ABAWD months used: "
objSelection.TypeParagraph
objSelection.TypeParagraph
objSelection.TypeText "Please complete the following:"
objSelection.TypeParagraph
objSelection.TypeText "Circle the month(s)/year(s) the person received federally funded TANF cash assistance: "
objSelection.TypeParagraph
objSelection.TypeParagraph
objSelection.TypeText Year(date)-20 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-19 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-18 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-17 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-16 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-15 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-14 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-13 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-12 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-11 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-10 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-9 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-8 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-7 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-6 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-5 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-4 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-3 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-2 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date)-1 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeText Year(date) & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
objSelection.TypeParagraph
objSelection.TypeParagraph
objSelection.TypeText "Is Medical Assistance closed?   YES	NO		Date of closure: "
objSelection.TypeParagraph
objSelection.TypeText "Name of Person verifying information: "
objSelection.TypeParagraph
objSelection.TypeText "Contact Information: "
objSelection.TypeParagraph
objSelection.TypeParagraph
objSelection.TypeText "Please email or fax your response to: " & worker_name & " Hennepin County Human Services and Public Health Services."
objSelection.TypeParagraph
objSelection.TypeText "If you have any questions about this request, you may contact me at: " & worker_phone
objSelection.TypeParagraph
objSelection.TypeParagraph
objSelection.TypeParagraph
objSelection.TypeParagraph
objSelection.TypeParagraph
objSelection.TypeText "Form generated by BlueZone Scripts on: " & Date() & " " & time()
End If

IF	agency_state_droplist = 	"Alabama"	THEN nabbr_state =	"AL"



IF	agency_state_droplist = 	"Alaska"	THEN	abbr_state =	"AK"
IF	agency_state_droplist = 	"Arizona"	THEN	abbr_state =	"AZ"
IF	agency_state_droplist = 	"Arkansas"	THEN	abbr_state =	"AR"
IF	agency_state_droplist = 	"California"THEN	abbr_state = "CA"
IF	agency_state_droplist = 	"Colorado"	THEN	abbr_state =	"CO"
IF	agency_state_droplist = 	"Connecticut"THEN	abbr_state = "CT"
IF	agency_state_droplist = 	"Delaware"	THEN	abbr_state =	"DE"
IF	agency_state_droplist = 	"Florida"	THEN	abbr_state =	"FL"
IF	agency_state_droplist = 	"Georgia"	THEN	abbr_state =	"GA"
IF	agency_state_droplist = 	"Hawaii"	THEN	abbr_state =	"HI"
IF	agency_state_droplist = 	"Idaho"	THEN	abbr_state =	"ID"
IF	agency_state_droplist = 	"Illinois"	THEN	abbr_state =	"IL"
IF	agency_state_droplist = 	"Indiana"	THEN	abbr_state =	"IN"
IF	agency_state_droplist = 	"Iowa"	THEN	abbr_state =	"IA"
IF	agency_state_droplist = 	"Kansas"	THEN	abbr_state =	"KS"
IF	agency_state_droplist = 	"Kentucky"	THEN	abbr_state =	"KY"
IF	agency_state_droplist = 	"Louisiana"	THEN	abbr_state =	"LA"
IF	agency_state_droplist = 	"Maine"	THEN	abbr_state =	"ME"
IF	agency_state_droplist = 	"Maryland"	THEN	abbr_state =	"MD"
IF	agency_state_droplist = 	"Massachusetts"	THEN	abbr_state =	"MA"
IF	agency_state_droplist = 	"Michigan"	THEN	abbr_state =	"MI"
IF	agency_state_droplist = 	"Mississippi"	THEN	abbr_state =	"MS"
IF	agency_state_droplist = 	"Missouri"	THEN	abbr_state =	"MO"
IF	agency_state_droplist = 	"Montana"	THEN	abbr_state =	"MT"
IF	agency_state_droplist = 	"Nebraska"	THEN	abbr_state =	"NE"
IF	agency_state_droplist = 	"Nevada"	THEN	abbr_state =	"NV"
IF	agency_state_droplist = 	"New Hampshire"	THEN	abbr_state =	"NH"
IF	agency_state_droplist = 	"New Jersey"	THEN	abbr_state =	"NJ"
IF	agency_state_droplist = 	"New Mexico"	THEN	abbr_state =	"NM"
IF	agency_state_droplist = 	"New York"	THEN	abbr_state =	"NY"
IF	agency_state_droplist = 	"North Carolina"	THEN	abbr_state =	"NC"
IF	agency_state_droplist = 	"North Dakota"	THEN	abbr_state =	"ND"
IF	agency_state_droplist = 	"Ohio"	THEN	abbr_state =	"OH"
IF	agency_state_droplist = 	"Oklahoma"	THEN	abbr_state =	"OK"
IF	agency_state_droplist = 	"Oregon"	THEN	abbr_state =	"OR"
IF	agency_state_droplist = 	"Pennsylvania"	THEN	abbr_state =	"PA"
IF	agency_state_droplist = 	"Rhode Island"	THEN	abbr_state =	"RI"
IF	agency_state_droplist = 	"South Carolina"	THEN	abbr_state =	"SC"
IF	agency_state_droplist = 	"South Dakota"	THEN	abbr_state =	"SD"
IF	agency_state_droplist = 	"Tennessee"	THEN	abbr_state =	"TN"
IF	agency_state_droplist = 	"Texas"	THEN	abbr_state =	"TX"
IF	agency_state_droplist = 	"Utah"	THEN	abbr_state =	"UT"
IF	agency_state_droplist = 	"Vermont"	THEN	abbr_state =	"VT"
IF	agency_state_droplist = 	"Virginia"	THEN	abbr_state =	"VA"
IF	agency_state_droplist = 	"Washington"	THEN	abbr_state =	"WA"
IF	agency_state_droplist = 	"West Virginia"	THEN	abbr_state =	"WV"
IF	agency_state_droplist = 	"Wisconsin"	THEN	abbr_state =	"WI"
IF	agency_state_droplist = 	"Wyoming"	THEN	abbr_state =	"WY"



'If hennepin_county = true then
''Generates Word Doc Form from share drive
'Set oApp = CreateObject("Word.Application")
'sDocName = "S:\fas\Scripts\Script Files\AGENCY CUSTOMIZED\OUT OF STATE FAX.docx"
'Set oDoc = oApp.Documents.Open(sDocName)
'oApp.Visible = true
'oDoc.FormFields("client_name").Result = client_name
'oDoc.FormFields("client_ssn").Result = client_ssn
'oDoc.FormFields("client_address").Result = client_address
'oDoc.FormFields("worker_name").Result = worker_name
'oDoc.FormFields("worker_phone").Result = worker_phone
'oDoc.FormFields("agency_name").Result = agency_name
'oDoc.FormFields("agency_fax").Result = agency_fax
'oDoc.FormFields("client_dob").Result = client_dob
'oDoc.FormFields("worker_info").Result = worker_info
'
'oDoc.SaveAs("Z:\My Documents\BlueZone\Scripts\OUT OF STATE.doc")
'End If

start_a_blank_case_note
Call write_variable_in_CASE_NOTE("***Out of State Inquiry sent via " & how_sent & " to " & abbr_state & " for M" & memb_number & "***")
CALL write_variable_in_CASE_NOTE("* Client reported they received " & out_of_state_programs & " on " & out_of_state_date & " the case is currently: " & out_of_state_status)
CALL write_bullet_and_variable_in_CASE_NOTE("Name", agency_name)
CALL write_bullet_and_variable_in_CASE_NOTE("Address", agency_adress)
CALL write_bullet_and_variable_in_CASE_NOTE("Email", agency_email)
CALL write_bullet_and_variable_in_CASE_NOTE("Phone", agency_phone)
CALL write_bullet_and_variable_in_CASE_NOTE("Fax", agency_fax)
Call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
PF3

'IF agency_email <> "" THEN
'	EmWriteScreen "x", 5, 3
'	Transmit
'	note_row = 4			'Beginning of the case notes
'	Do 						'Read each line
'		EMReadScreen note_line, 76, note_row, 3
'		note_line = trim(note_line)
'		If trim(note_line) = "" Then Exit Do		'Any blank line indicates the end of the case note because there can be no blank lines in a note
'		message_array = message_array & note_line & vbcr		'putting the lines together
'		note_row = note_row + 1
'		If note_row = 18 then 									'End of a single page of the case note
'			EMReadScreen next_page, 7, note_row, 3
'			If next_page = "More: +" Then 						'This indicates there is another page of the case note
'				PF8												'goes to the next line and resets the row to read'\
'				note_row = 4
'			End If
'		End If
'	Loop until next_page = "More:  " OR next_page = "       "	'No more pages
'	'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
'CALL create_outlook_email(agency_email, "","Out of State Inquiry for case #" &  MAXIS_case_number, "Out of State Inquiry" & vbcr & message_array,"", False)
'END IF
script_end_procedure("Success! Your Out of State Inquiry has been generated, please follow up with the next steps to ensure the request is received timely. The verification request must be reflected in ECF.")
