'This Script generates a OUT OF STATE INQUIRY form in use to fax to the out of state agency.
name_of_script = "NOTICES - OUT OF STATE.vbs"
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
call changelog_update("09/20/2018", "Updated for current content.", "Charles Potter, DHS")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS FOR THE SCRIPT======================================================================================================

    '------Paste any dialogs needed in from the dialog editor here. Dialogs typically include MAXIS_case_number and worker_signature fields
	BeginDialog client_dialog, 0, 0, 161, 165, "OUT OF STATE INQUIRY"
	  EditBox 55, 5, 35, 15, MAXIS_case_number
	  EditBox 135, 5, 20, 15, member_number
	  ButtonGroup ButtonPressed
	    PushButton 25, 25, 105, 20, "National Directory Website", outofstate_button
	  EditBox 95, 50, 60, 15, agency_name
	  EditBox 95, 70, 60, 15, agency_fax
	  EditBox 95, 90, 60, 15, worker_fax
	  CheckBox 5, 110, 150, 10, "Case note that out of state inquiry was sent", case_note_checkbox
	  EditBox 70, 125, 85, 15, worker_signature
	  ButtonGroup ButtonPressed
	    OkButton 50, 145, 50, 15
	    CancelButton 105, 145, 50, 15
	  Text 5, 10, 50, 10, "Case Number:"
	  Text 105, 10, 30, 10, "Memb #:"
	  Text 25, 55, 70, 10, "Out of State Agency:"
	  Text 15, 75, 80, 10, "Out of State Agency fax:"
	  Text 35, 95, 55, 10, "Your County fax:"
	  Text 5, 130, 60, 10, "Worker Signature:"
	EndDialog
'END DIALOGS=================================================================================================================

'THE SCRIPT==================================================================================================================

'Connects to BlueZone
EMConnect ""

'Grabs case number
call MAXIS_case_number_finder(MAXIS_case_number)

'Dialog
Do
	If ButtonPressed = outofstate_button then CreateObject("WScript.Shell").Run("https://dept.hennepin.us/hsphd/manuals/hsrm/Documents/National_Directory.PDF")
	Do
		Dialog client_dialog
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then MsgBox "You need to type a valid case number."
		If case_note_checkbox = 1 and worker_signature = "" then MsgBox "You need to add a signature since you are adding a casenote"
    		Call check_for_MAXIS(False)
    		call check_for_password (are_we_passworded_out) 'adding functionality for MAXIS v.6 Password Out issue'
    		call navigate_to_MAXIS_screen("stat","memb")
    		EMReadScreen invalid_MAXIS_case_number, 7, 24, 2
    		If invalid_MAXIS_case_number = "INVALID" then MsgBox "This is an invalid case number"
	Loop until MAXIS_case_number <> "" and IsNumeric(MAXIS_case_number) = True and len(MAXIS_case_number) <= 8 and invalid_MAXIS_case_number <> "INVALID" and case_note_checkbox = 0 or case_note_checkbox = 1 and worker_signature <> ""
Loop until ButtonPressed = -1

'Defaults member number to 01
If member_number = "" then member_number = "01"

'Goes to MEMB to get info
call navigate_to_MAXIS_screen("stat", "memb")

'Goes to the right HH member
EMWriteScreen member_number, 20, 76 'It does this to make sure that it navigates to the right HH member.
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
call navigate_to_MAXIS_screen("stat", "addr")
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
If worker_county = "X162" then
ramsey_county = true
Else
ramsey_county = false
End If

'reads assigned worker info
EMSetCursor 21, 21
PF1
EMReadScreen worker_name, 21, 19, 10
EMReadScreen worker_phone, 12, 19, 45
transmit

If ramsey_county = false then
'Generates Word Doc Form
Set objWord = CreateObject("Word.Application")
objWord.Caption = "OUT OF STATE INQUIRY"
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
'objSelection.ParagraphFormat.Alignment = 0
'objSelection.ParagraphFormat.LineSpacing = 12
'objSelection.ParagraphFormat.SpaceBefore = 0
'objSelection.ParagraphFormat.SpaceAfter = 0
'objSelection.Font.Name = "New York Times"
'objSelection.Font.Size = "14"
'objSelection.TypeText county_address_line_01
'objSelection.TypeParagraph
'objSelection.TypeText county_address_line_02
'objSelection.TypeParagraph

objSelection.ParagraphFormat.Alignment = 2
objSelection.ParagraphFormat.LineSpacing = 12
objSelection.ParagraphFormat.SpaceBefore = 0
objSelection.ParagraphFormat.SpaceAfter = 0

objSelection.Font.Name = "New York Times"
objSelection.Font.Size = "12"
objSelection.TypeText "DATE: " & date()

objSelection.TypeParagraph
objSelection.ParagraphFormat.Alignment = 0
objSelection.Font.Size = "10"
objSelection.Font.Bold = True
objSelection.TypeText "TO: " & agency_name
objSelection.TypeParagraph
objSelection.TypeText "FAX NUMBER: " & agency_fax
objSelection.TypeParagraph
objSelection.TypeText "RE: " & client_name
objSelection.TypeParagraph
objSelection.TypeText "SSN: " & client_ssn & "			DOB: " & client_dob
objSelection.TypeParagraph
objSelection.TypeText "CURRENT ADDRESS: " & client_address
objSelection.TypeParagraph
objSelection.TypeParagraph
objSelection.TypeText "Our records indicate that the above individual received or receives assistance from your state.  We need to verify the number of months of "
objSelection.Font.Underline = True
objSelection.TypeText "Federally-funded TANF cash assistance "
objSelection.Font.Underline = False
objSelection.TypeText "issued by your state that count towards the 60 month lifetime limit.  In addition, we need to know the number of months of TANF assistance from other states that your agency has verified.  Please indicate if the client is open on SNAP or Medical Assistance in your state OR the date these programs most recently closed.  Thank you."
objSelection.TypeParagraph

objSelection.TypeParagraph
objSelection.TypeText "Is CASH currently closed? 	YES	NO		Date of closure:___________________"
objSelection.TypeParagraph
objSelection.TypeText "Is SNAP currently closed? 	YES	NO		Date of closure:___________________"
objSelection.TypeParagraph
objSelection.TypeText "		TOTAL ABAWD MONTHS USED:________"
objSelection.TypeParagraph
objSelection.TypeText "		Please list the month(s)/year(s) of ABAWD months used:____________________"
objSelection.TypeParagraph
objSelection.TypeParagraph
objSelection.TypeText "Is Medical Assistance closed:	YES	NO		Date of closure:___________________"
objSelection.TypeParagraph
objSelection.TypeText "Name of Person verifying information:__________________________________________________"
objSelection.TypeParagraph
objSelection.TypeText "Phone Number:_____________________________"
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
objSelection.TypeParagraph
objSelection.TypeText "Please FAX your response to: " & worker_name & " MY FAX NUMBER IS: " & worker_fax & "."
objSelection.TypeParagraph
objSelection.TypeText "If you have any questions about this request, you may contact me at: " & worker_phone
objSelection.TypeParagraph
objSelection.TypeParagraph
objSelection.TypeParagraph
objSelection.TypeParagraph
objSelection.TypeParagraph
objSelection.TypeText "Form generated by BlueZone Scripts on: " & Date() & " " & time()
End If

If ramsey_county = true then
'Generates Word Doc Form from share drive
Set oApp = CreateObject("Word.Application")
sDocName = "S:\fas\Scripts\Script Files\AGENCY CUSTOMIZED\OUT OF STATE FAX.docx"
Set oDoc = oApp.Documents.Open(sDocName)
oApp.Visible = true
oDoc.FormFields("client_name").Result = client_name
oDoc.FormFields("client_ssn").Result = client_ssn
oDoc.FormFields("client_address").Result = client_address
oDoc.FormFields("worker_name").Result = worker_name
oDoc.FormFields("worker_phone").Result = worker_phone
oDoc.FormFields("agency_name").Result = agency_name
oDoc.FormFields("agency_fax").Result = agency_fax
oDoc.FormFields("client_dob").Result = client_dob
oDoc.FormFields("worker_fax").Result = worker_fax

oDoc.SaveAs("Z:\My Documents\BlueZone\Scripts\OUT OF STATE.doc")
End If

'Generates a Casenote
If case_note_checkbox = 1 then
pf4
pf9
EMSendKey "***OUT OF STATE INQUIRY SENT***"
CALL write_bullet_and_variable_in_CASE_NOTE("SENT OUT OF STATE INQURY FAX TO: ", agency_name)
CALL write_bullet_and_variable_in_CASE_NOTE("Agency Fax Contact", agency_fax)
CALL write_bullet_and_variable_in_CASE_NOTE("For:", client_name)
CALL write_bullet_and_variable_in_CASE_NOTE("Member Number", member_number)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
END IF

If ramsey_county = false then
script_end_procedure("Success! Your OUT OF STATE FAX form is generated!")
Else
script_end_procedure("Success! Your OUT OF STATE FAX form is generated! A Word Document back up is saved here 'Z:\My Documents\BlueZone\Scripts\OUT OF STATE.doc'")
End If
