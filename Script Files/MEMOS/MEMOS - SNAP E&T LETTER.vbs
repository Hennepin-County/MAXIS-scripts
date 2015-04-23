'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<adding for testing purposes
Worker_county_code = "x127"	 

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - SNAP E&T LETTER.vbs"
start_time = timer

'Option Explicit

DIM beta_agency
DIM FuncLib_URL, req, fso

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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


'Array listed above Dialog as below the dialog, the droplist appeared blank
'Creates an array of county FSET offices, which can be dynamically called in scripts which need it (SNAP ET LETTER for instance)

county_FSET_offices = array("Select one", "Blaine Human Services Building",	"Rural MN CEP Detroit Lakes", "RMCEP", "Leach Lake New", "MCT", "Red Lake Oshkiimaajitahdah",	"Minnesota WorkForce Center: Blue Earth", "Minnesota Valley Action Council New Ulm", "Carlton County Human Services", "Chippewa County Workforce Center", "Rural MN CEP Brainerd", "Northern Service Center", "Burnsville Workforce Center", "Workforce Development Inc. (Kasson)", "Alexandria Workforce Center", "Fairmont Workforce Center Fairbault County", "Workforce Development Office", "Workforce Development Inc. (Redwing)", "Grant County Social Services", "Sabathani Community Center", "Century Plaza", "Workforce Development Inc.", "Cambridge MN Workforce Center", "AEOA – GR Workforce Center", "Kittson County Social Services", "Lace qui Parle Co. Family Services", "Rural MN CEP Lake of the Woods", "AEOA", "MVAC",	"Marshall WorkForce Center", "Mahnomen County Human Services",	"Marshall County Social Services", "Fairmont Workforce Center Martin County",	"Central MN Jobs and Training Services Hutchinson", "Central MN Jobs and Training Services Litchfield",	"Workforce Development Inc. (Austin)", "Marshall WorkForce Center", "Olmstead County Family Support & Assistance", "Rural MN CEP Fergus Falls", "Minnesota WorkForce Center: Theif River Falls", "Pine County Health & Human Services", "Pine Technical & Community College E&T Center", "Southwest MN Private Industry Council Inc. Pipestone", "Polk County Social Services: Crookston", "Polk County Social Services: East Grand Forks", "Polk County Social Services: Fosston", "Minnesota Workforce Center: Red Lake",	"Southwest Health & Human Services", "Central MN Jobs and Training Services Olivia", "Southwest MN Private Industry Council Inc. Luverne.", "Roseau County Social Services", "Minnesota WorkForce Center: Duluth", "Minnesota WorkForce Center: Virginia", "Minnesota WorkForce Center:  Hibbing", "Central MN Jobs and Training Services Monticello", "Steele County Employment Services", "Stevens County Human Services", "SW MN Private Industry Council", "Todd County Health & Human Services: Long Prairie", "Todd County Health & Human Services: Staples", "Rural MN CEP Wheaton", "Workforce Development Inc. (Wabasha)", "Rural MN CEP Wadena", " Minnesota Valley Action Council Waseca", "Washington County Community Services: Stillwater", "Washington County Community Services: Forest Lake", "Washington County Community Services: Cottage Grove", "Washington County Community Services: Woodbury", "Wilkin County Family Services", "Central MN Jobs and Training Services Monticello", "Yellow Medicine County Family Services")																																																																																													

'IF worker_county_code = "x101" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x102" THEN county_FSET_offices = array("Select one", "Blaine Human Services Building")
IF worker_county_code = "x103" THEN county_FSET_offices = array("Select one", "Rural MN CEP Detroit Lakes")
IF worker_county_code = "x104" THEN county_FSET_offices = array("Select one", "RMCEP", "MCT", "Leach Lake New", "Red Lake Oshkiimaajitahdah")
'IF worker_county_code = "x105" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x106" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x107" THEN county_FSET_offices = array("Select one", "Minnesota WorkForce Center: Blue Earth")
IF worker_county_code = "x108" THEN county_FSET_offices = array("Select one", "Minnesota Valley Action Council New Ulm")
IF worker_county_code = "x109" THEN county_FSET_offices = array("Select one", "Carlton County Human Services")
'IF worker_county_code = "x110" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x111" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x112" THEN county_FSET_offices = array("Select one", "Chippewa County Workforce Center")
'IF worker_county_code = "x113" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x114" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x115" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x116" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x117" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x118" THEN county_FSET_offices = array("Select one", "Rural MN CEP Brainerd")
IF worker_county_code = "x119" THEN county_FSET_offices = array("Select one", "Northern Service Center", "Burnsville Workforce Center")
IF worker_county_code = "x120" THEN county_FSET_offices = array("Select one", "Workforce Development Inc. (Kasson)")
IF worker_county_code = "x121" THEN county_FSET_offices = array("Select one", "Alexandria Workforce Center")
IF worker_county_code = "x122" THEN county_FSET_offices = array("Select one", "Fairmont Workforce Center Fairbault County")
IF worker_county_code = "x123" THEN county_FSET_offices = array("Select one", "Workforce Development Office")
'IF worker_county_code = "x124" THEN county_FSET_offices = array("Select one", 
IF worker_county_code = "x125" THEN county_FSET_offices = array("Select one", "Workforce Development Inc. (Redwing)")
IF worker_county_code = "x126" THEN county_FSET_offices = array("Select one", "Grant County Social Services")
IF worker_county_code = "x127" THEN county_FSET_offices = array("Select one", "Century Plaza", "Sabathani Community Center")
IF worker_county_code = "x128" THEN county_FSET_offices = array("Select one", "Workforce Development Inc.")
'IF worker_county_code = "x129" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x130" THEN county_FSET_offices = array("Select one", "Cambridge MN Workforce Center")
IF worker_county_code = "x131" THEN county_FSET_offices = array("Select one", "AEOA – GR Workforce Center")
'IF worker_county_code = "x132" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x133" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x134" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x135" THEN county_FSET_offices = array("Select one", "Kittson County Social Services")
'IF worker_county_code = "x136" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x137" THEN county_FSET_offices = array("Select one", "Lace qui Parle Co. Family Services")
IF worker_county_code = "x138" THEN county_FSET_offices = array("Select one", "AEOA")
IF worker_county_code = "x139" THEN county_FSET_offices = array("Select one", "Rural MN CEP Lake of the Woods")
IF worker_county_code = "x140" THEN county_FSET_offices = array("Select one", "MVAC")
IF worker_county_code = "x141" THEN county_FSET_offices = array("Select one", "Marshall WorkForce Center")                                                     
IF worker_county_code = "x142" THEN county_FSET_offices = array("Select one", "Marshall WorkForce Center") 
IF worker_county_code = "x143" THEN county_FSET_offices = array("Select one", "Mahnomen County Human Services")
IF worker_county_code = "x144" THEN county_FSET_offices = array("Select one", "Marshall County Social Services")
IF worker_county_code = "x145" THEN county_FSET_offices = array("Select one", "Fairmont Workforce Center Martin County")
IF worker_county_code = "x146" THEN county_FSET_offices = array("Select one", "Central MN Jobs and Training Services Hutchinson")
IF worker_county_code = "x147" THEN county_FSET_offices = array("Select one", "Central MN Jobs and Training Services Litchfield")
'IF worker_county_code = "x148" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x149" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x150" THEN county_FSET_offices = array("Select one", "Workforce Development Inc. (Austin)")
IF worker_county_code = "x151" THEN county_FSET_offices = array("Select one", "Marshall WorkForce Center")                                                 
'IF worker_county_code = "x152" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x153" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x154" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x155" THEN county_FSET_offices = array("Select one", "Olmstead County Family Support & Assistance")
IF worker_county_code = "x156" THEN county_FSET_offices = array("Select one", "Rural MN CEP Fergus Falls")
IF worker_county_code = "x157" THEN county_FSET_offices = array("Select one", "Minnesota WorkForce Center: Theif River Falls")
IF worker_county_code = "x158" THEN county_FSET_offices = array("Select one", "Pine County Health & Human Services", "Pine Technical & Community College E&T Center")
IF worker_county_code = "x159" THEN county_FSET_offices = array("Select one", "Southwest MN Private Industry Council Inc. Pipestone")
IF worker_county_code = "x160" THEN county_FSET_offices = array("Select one", "Polk County Social Services: Crookston", "Polk County Social Services: East Grand Forks", "Polk County Social Services: Fosston")
'IF worker_county_code = "x161" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x162" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x163" THEN county_FSET_offices = array("Select one", "Minnesota Workforce Center: Red Lake")
IF worker_county_code = "x164" THEN county_FSET_offices = array("Select one", "Southwest Health & Human Services")
IF worker_county_code = "x165" THEN county_FSET_offices = array("Select one", "Central MN Jobs and Training Services Olivia")
'IF worker_county_code = "x166" THEN county_FSET_offices = array("Select one", 
IF worker_county_code = "x167" THEN county_FSET_offices = array("Select one", "Southwest MN Private Industry Council Inc. Luverne") 
IF worker_county_code = "x168" THEN county_FSET_offices = array("Select one", "Roseau County Social Services")
IF worker_county_code = "x169" THEN county_FSET_offices = array("Select one", "MN Workforce Center: Duluth", "Minnesota WorkForce Center: Virginia", "Minnesota Workforce Center: Hibbing")
'IF worker_county_code = "x170" THEN county_FSET_offices = array("Select one", 
IF worker_county_code = "x171" THEN county_FSET_offices = array("Select one", "Central MN Jobs and Training Services Monticello")
'IF worker_county_code = "x172" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x173" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x174" THEN county_FSET_offices = array("Select one", "Steele County Employment Services")
IF worker_county_code = "x175" THEN county_FSET_offices = array("Select one", "Stevens County Human Services")
IF worker_county_code = "x176" THEN county_FSET_offices = array("Select one", "SW MN Private Industry Council")
IF worker_county_code = "x177" THEN county_FSET_offices = array("Select one", "Todd County Health & Human Services: Long Prairie", "Todd County Health & Human Services: Staples")
IF worker_county_code = "x178" THEN county_FSET_offices = array("Select one", "Rural MN CEP Wadena")
IF worker_county_code = "x179" THEN county_FSET_offices = array("Select one", "Workforce Development Inc.")
IF worker_county_code = "x180" THEN county_FSET_offices = array("Select one", "Rural MN CEP/MN workforce Center")
IF worker_county_code = "x181" THEN county_FSET_offices = array("Select one", "Minnesota Valley Action Council Waseca")
IF worker_county_code = "x182" THEN county_FSET_offices = array("Select one", "Washington County Community Services: Stillwater", "Washington County Community Services: Forest Lake", "Washington County Community Services: Cottage Grove", "Washington County Community Services: Woodbury")
'IF worker_county_code = "x183" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x184" THEN county_FSET_offices = array("Select one", "Wilkin County Family Services")
'IF worker_county_code = "x185" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x186" THEN county_FSET_offices = array("Select one", "Central MN Jobs and Training Services Monticello")
IF worker_county_code = "x187" THEN county_FSET_offices = array("Select one", "Yellow Medicine County Family Services")


call convert_array_to_droplist_items (county_FSET_offices, FSET_list)

If worker_county_code = "x127" THEN 
	SNAPET_contact = "the EZ Info Line"
	SNAPET_phone = "612-596-1300"
END IF

'DIALOGS----------------------------------------------------------------------------------------------------
' FSET_list is a variable not a standard drop down list.  When you copy into dialog editor, it will not work
BeginDialog SNAPET_dialog, 0, 0, 321, 195, "SNAP E&T Appointment Letter"
  EditBox 70, 5, 55, 15, case_number
  EditBox 215, 5, 20, 15, member_number
  EditBox 70, 25, 55, 15, appointment_date
  EditBox 215, 25, 20, 15, appointment_time_prefix_editbox
  EditBox 235, 25, 20, 15, appointment_time_post_editbox
  DropListBox 260, 25, 55, 15, "Select one.."+chr(9)+"AM"+chr(9)+"PM", AM_PM
  DropListBox 175, 45, 140, 15, "county_office_list", interview_location
  EditBox 65, 65, 190, 15, SNAPET_name
  EditBox 65, 85, 190, 15, SNAPET_address_01
  EditBox 65, 105, 95, 15, SNAPET_city
  EditBox 165, 105, 40, 15, SNAPET_ST
  EditBox 210, 105, 45, 15, SNAPET_zip
  EditBox 65, 125, 65, 15, SNAPET_contact
  EditBox 185, 125, 70, 15, SNAPET_phone
  EditBox 140, 175, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 210, 175, 50, 15
    CancelButton 265, 175, 50, 15
  Text 5, 30, 60, 10, "Appointment Date:"
  Text 145, 30, 60, 15, "Appointment Time:"
  Text 5, 50, 170, 10, "Location (select from dropdown, or fill in manually)"
  Text 5, 70, 55, 10, "Provider Name:"
  Text 5, 90, 55, 10, "Address line 1:"
  Text 10, 130, 55, 10, "Contact Name:"
  Text 135, 130, 50, 10, "Contact Phone:"
  Text 80, 180, 60, 10, "Worker Signature:"
  Text 5, 145, 315, 25, "Please note: the dropdown above automatically fills in from your agency office/intake locations.  It may not match your SNAP E&T orientation locations.  Please double check the address before pressing OK. "
  Text 10, 10, 50, 10, "Case Number:"
  Text 10, 110, 55, 10, "City/State/Zip:"
  Text 140, 10, 70, 10, "HH Member Number:"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone default screen
EMConnect ""

'Searches for a case number
call MAXIS_case_number_finder(case_number)

'Shows dialog, checks for password prompt
DO
	DO
		DO
			DO
				DO
					DO
						DO
							DO
								DO
									DO
										DO
											Dialog SNAPET_dialog
											cancel_confirmation 'asks if they really want to cancel script	
											IF case_number = "" then MsgBox "You did not enter a case number. Please try again."
										LOOP UNTIL case_number <> ""
										If isdate(appointment_date) = FALSE then MsgBox "You did not enter a valid appointment date. Please try again."
									LOOP UNTIL isdate(appointment_date) = True
									IF member_number = "" then MsgBox "You did not specify a household member number.  Please try again."
								LOOP UNTIL isnumeric(member_number) = true
								IF SNAPET_name = "" then MsgBox "Please specify the agency name."
							LOOP UNTIL SNAPET_name <> ""
							IF SNAPET_address_01 = "" then MsgBox "Please enter the address for the SNAP ET agency."
						LOOP UNTIL SNAPET_address_01 <> ""
						IF appointment_time_prefix_editbox = "" then MsgBox "Please specify an appointment time."
					LOOP UNTIL appointment_time_prefix_editbox <> ""
					IF appointment_time_post_editbox = "" then MsgBox "Please specify an appointment time."
				LOOP UNTIL appointment_time_post_editbox <> ""	
				If AM_PM = "Select One..." THEN MsgBox "Please choose either a.m. or p.m."
			LOOP UNTIL AM_PM <> "Select One..."					
			IF worker_signature = "" then MsgBox "You did not sign your case note. Please try again."
		LOOP UNTIL worker_signature <> ""
		IF SNAPET_contact = "" THEN MsgBox "You must specify the E&T contact name.  Please try again."
	LOOP UNTIL SNAPET_contact <> ""
	IF SNAPET_phone = "" THEN MsgBox "You must enter a contact phone number.  Please try again."
LOOP UNTIL SNAPET_phone <> ""	

transmit
Call maxis_check_function

'CO #27 Hennepin County addresses
county_FSET_offices = array("Select one", "Century Plaza", "Sabathani Community Center")
IF interview_location = "Century Plaza" THEN 
	SNAPET_name = "Century Plaza"
	SNAPET_address_01 = "330 South 12th Street #3650"
	SNAPET_address_02 = "Minneapolis, MN. 55404"
ElseIf interview_location = "Sabathani Community Center" THEN 
	SNAPET_name = "Sabathani Community Center"
	SNAPET_address_01 = "310 East 38th Street #120"
	SNAPET_address_02 = "Minneapolis, MN. 55409"
END IF

'CO #02 Anoka County address
county_FSET_offices = array("Select one", "Blaine Human Services Building")
IF interview_location = "Century Plaza" THEN 
	SNAPET_name = "Blaine Human Services Building"
	SNAPET_address_01 = "1201 89th Avenue NE Suite 235"
	SNAPET_address_02 = "Blaine, MN. 55434"
END IF

'CO #3 BECKER COUNTY address
county_FSET_offices = array("Select one", "Rural MN CEP")
IF interview_location = "Rural MN CEP" THEN 
	SNAPET_name = "Rural MN CEP"
	SNAPET_address_01 = "1803 Roosevelt Ave"
	SNAPET_address_02 = "Detroit Lakes, MN. 56501"
END IF

'CO #04 BELTRAMI COUNTY addresses
 county_FSET_offices = array("Select one", "RMCEP", "MCT", "Leach Lake New", "Red Lake Oshkiimaajitahdah")
IF interview_location = "RMCEP" THEN 
	SNAPET_name = "RMCEP"
	SNAPET_address_01 = "616 America Ave NW Suite 210"
	SNAPET_address_02 = "Bemedji, MN. 56601"
ElseIf interview_location = "MCT" THEN 
	SNAPET_name = "MCT"
	SNAPET_address_01 = "15542 State Hwy 371 NW"
	SNAPET_address_02 = "Cass Lake, MN. 56633"
ElseIf interview_location = "Leach Lake New" THEN 
	SNAPET_name = "Leach Lake New"
	SNAPET_address_01 = "190 Sail Drive NW"
	SNAPET_address_02 = "Cass Lake, MN. 56633"
ElseIf interview_location = "Red Lake Oshkiimaajitahdah" THEN 
	SNAPET_name = "Red Lake Oshkiimaajitahdah"
	SNAPET_address_01 = "MN-1"
	SNAPET_address_02 = "Redby, MN. 56670"
END IF

'CO #5 BENTON COUNTY address
county_FSET_offices = array("Select one", "Rural MN CEP")
IF interview_location = "Rural MN CEP" THEN 
	SNAPET_name = "Rural MN CEP"
	SNAPET_address_01 = "1803 Roosevelt Ave"
	SNAPET_address_02 = "Detroit Lakes, MN. 56501"
END IF


'Pulls the member name.
call navigate_to_MAXIS_screen("STAT", "MEMB")
EMWriteScreen member_number, 20, 76
transmit
EMReadScreen last_name, 24, 6, 30
EMReadScreen first_name, 11, 6, 63
last_name = trim(replace(last_name, "_", ""))
first_name = trim(replace(first_name, "_", ""))

'Navigates into SPEC/LETR
call navigate_to_MAXIS_screen("SPEC", "LETR") 

'Opens up the SNAP E&T Orientation LETR. If it's unable the script will stop.
EMWriteScreen "x", 8, 12
transmit
EMReadScreen LETR_check, 4, 2, 49
If LETR_check = "LETR" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")


'Writes the info into the LETR. 
EMWriteScreen first_name & " " & last_name, 4, 28
call create_MAXIS_friendly_date_three_spaces_between(appointment_date, 0, 6, 28) 
EMWriteScreen appointment_time_prefix_editbox, 7, 28
EMWriteScreen appointment_time_post_editbox, 7, 33
EMWriteScreen AM_PM, 7, 38
EMWriteScreen SNAPET_name, 9, 28
EMWriteScreen SNAPET_address_01, 10, 28
EMWriteScreen SNAPET_address_02, 11, 28
call create_MAXIS_friendly_phone_number(SNAPET_phone, 13, 28) 'takes out non-digits if listed in variable, and formats phone number for the field
EMWriteScreen SNAPET_contact, 16, 28
PF4		'saves and sends memo

'Navigates to a blank case note
call start_a_blank_CASE_NOTE

'Writes the case note
CALL write_new_line_in_case_note("***SNAP E&T Appointment Letter Sent***")
CALL write_bullet_and_variable_in_case_note("Appointment date", appointment_date)
CALL write_bullet_and_variable_in_case_note("Appointment time", appointment_time_prefix_editbox & ":" & appointment_time_post_editbox & " " & AM_PM)
CALL write_bullet_and_variable_in_case_note("Appointment location", SNAPET_name)
CALL write_new_line_in_case_note("---")
CALL write_new_line_in_case_note(worker_signature)

script_end_procedure("")
