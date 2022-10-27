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
call changelog_update("11/12/2020", "Updated HSR Manual link for OUT OF STATE due to SharePoint Online Migration.", "Ilse Ferris, Hennepin County")
call changelog_update("11/29/2018", "Updated for current National Directory information and requested changes to word document.", "MiKayla Handley, Hennepin County")
call changelog_update("09/20/2018", "Updated for current content.", "Charles Potter, DHS")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'THE SCRIPT==================================================================================================================
EMConnect ""

'Grabs case number
call MAXIS_case_number_finder(MAXIS_case_number)
'Defaults member number to 01
'If MEMB_number = "" then MEMB_number = "01"

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 226, 255, "OUT OF STATE INQUIRY"
  EditBox 55, 5, 40, 15, MAXIS_case_number
  EditBox 130, 5, 20, 15, MEMB_number
  DropListBox 35, 35, 55, 15, "Select One:"+chr(9)+"Alabama"+chr(9)+"Alaska"+chr(9)+"Arizona"+chr(9)+"Arkansas"+chr(9)+"California"+chr(9)+"Colorado"+chr(9)+"Connecticut"+chr(9)+"Delaware"+chr(9)+"Florida"+chr(9)+"Georgia"+chr(9)+"Hawaii"+chr(9)+"Idaho"+chr(9)+"Illinois"+chr(9)+"Indiana"+chr(9)+"Iowa"+chr(9)+"Kansas"+chr(9)+"Kentucky"+chr(9)+"Louisiana"+chr(9)+"Maine"+chr(9)+"Maryland"+chr(9)+"Massachusetts"+chr(9)+"Michigan"+chr(9)+"Mississippi"+chr(9)+"Missouri"+chr(9)+"Montana"+chr(9)+"Nebraska"+chr(9)+"Nevada"+chr(9)+"New Hampshire"+chr(9)+"New Jersey"+chr(9)+"New Mexico"+chr(9)+"New York"+chr(9)+"North Carolina"+chr(9)+"North Dakota"+chr(9)+"Ohio"+chr(9)+"Oklahoma"+chr(9)+"Oregon"+chr(9)+"Pennsylvania"+chr(9)+"Rhode Island"+chr(9)+"South Carolina"+chr(9)+"South Dakota"+chr(9)+"Tennessee"+chr(9)+"Texas"+chr(9)+"Utah"+chr(9)+"Vermont"+chr(9)+"Virginia"+chr(9)+"Washington"+chr(9)+"West Virginia"+chr(9)+"Wisconsin"+chr(9)+"Wyoming", agency_state_droplist
  DropListBox 35, 55, 55, 15, "Select One:"+chr(9)+"Active"+chr(9)+"Closed"+chr(9)+"Unknown", out_of_state_status
  EditBox 160, 35, 55, 15, out_of_state_programs
  EditBox 160, 55, 55, 15, out_of_state_date
  EditBox 50, 95, 165, 15, agency_name
  EditBox 50, 115, 165, 15, agency_address
  EditBox 50, 135, 165, 15, agency_email
  EditBox 50, 155, 50, 15, agency_phone
  EditBox 165, 155, 50, 15, agency_fax
  DropListBox 160, 180, 55, 15, "Select One:"+chr(9)+"Email"+chr(9)+"Fax"+chr(9)+"Mail"+chr(9)+"Phone", how_sent
  EditBox 50, 200, 165, 15, other_notes
  EditBox 70, 220, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 140, 220, 35, 15
    CancelButton 180, 220, 35, 15
    PushButton 160, 5, 60, 15, "HSR MANUAL", outofstate_button
  Text 5, 10, 50, 10, "Case Number:"
  Text 100, 10, 30, 10, "Memb #:"
  GroupBox 5, 25, 215, 50, "Client reported they received assistance (Q5 on CAF):"
  Text 10, 40, 20, 10, "State:"
  Text 10, 60, 25, 10, "Status:"
  Text 105, 40, 50, 10, "What Benefits:"
  Text 105, 60, 50, 10, "Last Received:"
  GroupBox 5, 80, 215, 95, "Out of State Agency Contact"
  Text 15, 100, 25, 10, "Name:"
  Text 15, 120, 30, 10, "Address:"
  Text 15, 140, 25, 10, "Email:"
  Text 15, 160, 25, 10, "Phone:"
  Text 145, 160, 15, 10, "Fax:"
  Text 70, 185, 90, 10, "How was the request sent:"
  Text 5, 205, 45, 10, "Other Notes:"
  Text 5, 225, 60, 10, "Worker Signature:"
  Text 15, 240, 185, 10, "*** Reminder: ECF must show verification requested ***"
EndDialog

'Dialog
DO      'Password DO loop
    DO  'Conditional handling DO loop
        DO  'External resource DO loop
            Dialog Dialog1
            cancel_confirmation
            If ButtonPressed = outofstate_button then CreateObject("WScript.Shell").Run("https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Out_of_State_Inquiry.aspx")
        Loop until ButtonPressed = -1
        err_msg = ""
		If agency_state_droplist = "Select One:" then err_msg = err_msg & vbnewline & "* Select the state."
        If trim(MAXIS_case_number) = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbcr & "* Enter a valid case number."
        If trim(MEMB_number) = "" or IsNumeric(MEMB_number) = False or len(MEMB_number) <> 2 then err_msg = err_msg & vbcr & "* Enter a valid member number."
        If trim(out_of_state_date) = "" then err_msg = err_msg & vbcr & "* Enter the date the client reported benefits were received."
		If trim(agency_name) = "" then err_msg = err_msg & vbcr & "* Enter the out of state agency name."
        If trim(agency_address) = "" then err_msg = err_msg & vbcr & "* Enter the out of state agency address, if there is not one provided enter N/A."
		If trim(agency_email) = "" then err_msg = err_msg & vbcr & "* Enter the out of state agency email, if there is not one provided enter N/A."
		If trim(agency_phone) = "" then err_msg = err_msg & vbcr & "* Enter the out of state agency phone, if there is not one provided enter N/A."
		If trim(agency_fax) = "" then err_msg = err_msg & vbcr & "* Enter the out of state agency fax, if there is not one provided enter N/A."
        If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Enter your worker signature."
		If how_sent = "Select One:" then err_msg = err_msg & vbnewline & "* Select how the request was sent."
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
Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
If mail_line_one = "" then
     client_address = resi_line_one & " " & resi_line_two & " " & resi_city & ", " & resi_state & " " & resi_zip
Else
	client_address =  mail_line_one & " " & mail_line_two & " " & mail_city & ", " & mail_state & " " & mail_zip
End If

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

IF	agency_state_droplist = 	"Alabama"	THEN	abbr_state =	"AL"
IF	agency_state_droplist = 	"Alaska"	THEN	abbr_state =	"AK"
IF	agency_state_droplist = 	"Arizona"	THEN	abbr_state =	"AZ"
IF	agency_state_droplist = 	"Arkansas"	THEN	abbr_state =	"AR"
IF	agency_state_droplist = 	"California"	THEN	abbr_state = "CA"
IF	agency_state_droplist = 	"Colorado"	THEN	abbr_state =	"CO"
IF	agency_state_droplist = 	"Connecticut"	THEN	abbr_state = "CT"
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
