'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - SNAP E AND T LETTER.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 280                     'manual run time in seconds
STATS_denomination = "C"       			   'C is for each CASE
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Checks for county info from global variables, or asks if it is not already defined.
get_county_code

'Creating a blank array to start our process. This will allow for validating whether-or-not the office was assigned later on, because it'll always be an array and not a variable.
county_FSET_offices = array("")

'Array listed above Dialog as below the dialog, the droplist appeared blank
'Creates an array of county FSET offices, which can be dynamically called in scripts which need it (SNAP ET LETTER for instance)
'Certain counties are commented out as they did not submit information about their E & T site, but can be easily rendered if they provide them
'IF worker_county_code = "x101" THEN county_FSET_offices = array("Aitkin Workforce Center")
IF worker_county_code = "x102" THEN county_FSET_offices = array("Minnesota WorkForce Center Blaine")
IF worker_county_code = "x103" THEN county_FSET_offices = array("Rural MN CEP Detroit Lakes")
IF worker_county_code = "x104" THEN county_FSET_offices = array("Select one...", "RMCEP", "MCT", "Leach Lake New", "Red Lake Oshkiimaajitahdah")
'IF worker_county_code = "x105" THEN county_FSET_offices = array("Select one...",
'IF worker_county_code = "x106" THEN county_FSET_offices = array("Select one...",
IF worker_county_code = "x107" THEN county_FSET_offices = array("Blue Earth County Employment Services")
IF worker_county_code = "x108" THEN county_FSET_offices = array("Minnesota Valley Action Council New Ulm")
IF worker_county_code = "x109" THEN county_FSET_offices = array("Carlton County Human Services")
'IF worker_county_code = "x110" THEN county_FSET_offices = array("Select one...",
'IF worker_county_code = "x111" THEN county_FSET_offices = array("Select one...",
IF worker_county_code = "x112" THEN county_FSET_offices = array("Montevideo Workforce Center")
'IF worker_county_code = "x113" THEN county_FSET_offices = array("Select one...",
'IF worker_county_code = "x114" THEN county_FSET_offices = array("Select one...",
'IF worker_county_code = "x115" THEN county_FSET_offices = array("Select one...",
'IF worker_county_code = "x116" THEN county_FSET_offices = array("Select one...",
'IF worker_county_code = "x117" THEN county_FSET_offices = array("Select one...",
IF worker_county_code = "x118" THEN county_FSET_offices = array("Rural MN CEP Brainerd")
IF worker_county_code = "x119" THEN county_FSET_offices = array("Select one...", "Northern Service Center", "Burnsville Workforce Center")
IF worker_county_code = "x120" THEN county_FSET_offices = array("Workforce Development Inc. (Kasson)")
IF worker_county_code = "x121" THEN county_FSET_offices = array("Alexandria Workforce Center")
IF worker_county_code = "x122" THEN county_FSET_offices = array("Fairmont Workforce Center Fairbault County")
IF worker_county_code = "x123" THEN county_FSET_offices = array("Workforce Development Office")
'IF worker_county_code = "x124" THEN county_FSET_offices = array("Select one...",
IF worker_county_code = "x125" THEN county_FSET_offices = array("Workforce Development Inc. (Redwing)")
IF worker_county_code = "x126" THEN county_FSET_offices = array("Grant County Social Services")
'IF worker_county_code = "x127" THEN county_FSET_offices = array("Select one...", "Health Services Building", "Sabathani Community Center")
IF worker_county_code = "x128" THEN county_FSET_offices = array("Workforce Development Inc.")
'IF worker_county_code = "x129" THEN county_FSET_offices = array("Select one...",
IF worker_county_code = "x130" THEN county_FSET_offices = array("Cambridge MN Workforce Center")
'IF worker_county_code = "x131" THEN county_FSET_offices = array("AEOA – GR Workforce Center")
'IF worker_county_code = "x132" THEN county_FSET_offices = array("Select one...",
'IF worker_county_code = "x133" THEN county_FSET_offices = array("Select one...",
'IF worker_county_code = "x134" THEN county_FSET_offices = array("Select one...",
IF worker_county_code = "x135" THEN county_FSET_offices = array("Kittson County Social Services")
'IF worker_county_code = "x136" THEN county_FSET_offices = array("Select one...",
IF worker_county_code = "x137" THEN county_FSET_offices = array("Lace qui Parle Co. Family Services")
IF worker_county_code = "x138" THEN county_FSET_offices = array("AEOA")
'IF worker_county_code = "x139" THEN county_FSET_offices = array("Rural MN CEP Lake of the Woods")
IF worker_county_code = "x140" THEN county_FSET_offices = array("MVAC")
IF worker_county_code = "x141" THEN county_FSET_offices = array("Marshall WorkForce Center")
IF worker_county_code = "x142" THEN county_FSET_offices = array("Marshall WorkForce Center")
IF worker_county_code = "x143" THEN county_FSET_offices = array("Mahnomen County Human Services")
'IF worker_county_code = "x144" THEN county_FSET_offices = array("Marshall County Social Services")
'IF worker_county_code = "x145" THEN county_FSET_offices = array("Fairmont Workforce Center Martin County")
IF worker_county_code = "x146" THEN county_FSET_offices = array("Central MN Jobs and Training Services Hutchinson")
IF worker_county_code = "x147" THEN county_FSET_offices = array("Central MN Jobs and Training Services Litchfield")
'IF worker_county_code = "x148" THEN county_FSET_offices = array("Select one...",
'IF worker_county_code = "x149" THEN county_FSET_offices = array("Select one...",
IF worker_county_code = "x150" THEN county_FSET_offices = array("Workforce Development Inc. (Austin)")
IF worker_county_code = "x151" THEN county_FSET_offices = array("Marshall WorkForce Center")
'IF worker_county_code = "x152" THEN county_FSET_offices = array("Select one...",
'IF worker_county_code = "x153" THEN county_FSET_offices = array("Select one...",
'IF worker_county_code = "x154" THEN county_FSET_offices = array("Select one...",
IF worker_county_code = "x155" THEN county_FSET_offices = array("Olmstead County Family Support & Assistance")
IF worker_county_code = "x156" THEN county_FSET_offices = array("Rural MN CEP Fergus Falls")
IF worker_county_code = "x157" THEN county_FSET_offices = array("Minnesota WorkForce Center: Thief River Falls")
IF worker_county_code = "x158" THEN county_FSET_offices = array("Select one...", "Pine County Public Health Building", "Pine Technical & Community College E&T Center")
IF worker_county_code = "x159" THEN county_FSET_offices = array("Southwest MN Private Industry Council Inc. Pipestone")
IF worker_county_code = "x160" THEN county_FSET_offices = array("Select one...", "Polk County Social Services: Crookston", "Polk County Social Services: East Grand Forks", "Polk County Social Services: Fosston")
IF worker_county_code = "x161" THEN county_FSET_offices = array("Minnesota Workforce Center Alexandria")
'IF worker_county_code = "x162" THEN county_FSET_offices = array("Select one...",
IF worker_county_code = "x163" THEN county_FSET_offices = array("Minnesota Workforce Center: Red Lake")
IF worker_county_code = "x164" THEN county_FSET_offices = array("Southwest Health & Human Services")
IF worker_county_code = "x165" THEN county_FSET_offices = array("Central MN Jobs and Training Services Olivia")
'IF worker_county_code = "x166" THEN county_FSET_offices = array("Select one...",
IF worker_county_code = "x167" THEN county_FSET_offices = array("Southwest MN Private Industry Council Inc. Luverne")
IF worker_county_code = "x168" THEN county_FSET_offices = array("Roseau County Social Services")
IF worker_county_code = "x169" THEN county_FSET_offices = array("Select one...", "Minnesota WorkForce Center: Duluth", "Minnesota WorkForce Center: Virginia", "Minnesota WorkForce Center: Hibbing")
'IF worker_county_code = "x170" THEN county_FSET_offices = array("Select one...",
IF worker_county_code = "x171" THEN county_FSET_offices = array("Central MN Jobs and Training Services Monticello")
'IF worker_county_code = "x172" THEN county_FSET_offices = array("Select one...",
'IF worker_county_code = "x173" THEN county_FSET_offices = array("Select one...",
IF worker_county_code = "x174" THEN county_FSET_offices = array("Steele County Employment Services")
IF worker_county_code = "x175" THEN county_FSET_offices = array("Stevens County Human Services")
IF worker_county_code = "x176" THEN county_FSET_offices = array("SW MN Private Industry Council")
IF worker_county_code = "x177" THEN county_FSET_offices = array("Select one...", "Todd County Health & Human Services: Long Prairie", "Todd County Health & Human Services: Staples")
IF worker_county_code = "x178" THEN county_FSET_offices = array("Rural MN CEP Wadena")
IF worker_county_code = "x179" THEN county_FSET_offices = array("Workforce Development Inc.")
'IF worker_county_code = "x180" THEN county_FSET_offices = array("Rural MN CEP/MN workforce Center")
IF worker_county_code = "x181" THEN county_FSET_offices = array("Minnesota Valley Action Council Waseca")
IF worker_county_code = "x182" THEN county_FSET_offices = array("Select one...", "Washington County Community Services: Stillwater", "Washington County Community Services: Forest Lake", "Washington County Community Services: Cottage Grove", "Washington County Community Services: Woodbury")
'IF worker_county_code = "x183" THEN county_FSET_offices = array("Select one...",
IF worker_county_code = "x184" THEN county_FSET_offices = array("Wilkin County Family Services")
'IF worker_county_code = "x185" THEN county_FSET_offices = array("Select one...",
IF worker_county_code = "x186" THEN county_FSET_offices = array("Central MN Jobs and Training Services Monticello")
IF worker_county_code = "x187" THEN county_FSET_offices = array("Yellow Medicine County Family Services")

'If the array isn't blank, then create a new array called FSET_list containing these items as a droplist. This will be used by the dialog.
IF county_FSET_offices(0) <> "" THEN call convert_array_to_droplist_items (county_FSET_offices, FSET_list)

'DIALOGS----------------------------------------------------------------------------------------------------
' *********FSET_list is a variable not a standard drop down list.  When you copy into dialog editor, it will not work***********
' This dialog is for counties that HAVE provided FSET office addresses
BeginDialog SNAPET_automated_adress_dialog, 0, 0, 306, 240, "SNAP E&T Appointment Letter"
  EditBox 70, 5, 55, 15, MAXIS_case_number
  EditBox 215, 5, 20, 15, member_number
  EditBox 70, 25, 55, 15, appointment_date
  EditBox 195, 25, 20, 15, appointment_time_prefix_editbox
  EditBox 215, 25, 20, 15, appointment_time_post_editbox
  DropListBox 240, 25, 60, 15, "Select one..."+chr(9)+"AM"+chr(9)+"PM", AM_PM
  DropListBox 115, 50, 185, 15, FSET_list, interview_location
  EditBox 60, 70, 110, 15, SNAPET_contact
  EditBox 235, 70, 65, 15, SNAPET_phone
  DropListBox 105, 95, 85, 15, "Select one..."+chr(9)+"Banked months"+chr(9)+"Other manual referral"+chr(9)+"Student"+chr(9)+"Working with CBO", manual_referral
  EditBox 105, 115, 195, 15, other_referral_notes
  EditBox 105, 140, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 195, 140, 50, 15
    CancelButton 250, 140, 50, 15
  Text 130, 10, 70, 10, "HH Member Number:"
  Text 130, 30, 60, 10, "Appointment Time:"
  Text 5, 75, 50, 10, "Contact name: "
  Text 5, 30, 60, 10, "Appointment Date:"
  Text 180, 75, 50, 10, "Contact phone:"
  Text 5, 50, 105, 10, "Location (select from dropdown):"
  GroupBox 5, 165, 295, 70, "When is a manual referral needed"
  Text 15, 180, 275, 20, "If an ABAWD is using banked months, or a student meets criteria under CM0011.18, or receiving E and T services through a Community Based Organization (CBO)."
  Text 10, 10, 50, 10, "Case Number:"
  Text 5, 100, 80, 10, "Manual referral needed:"
  Text 15, 205, 275, 25, "Select a recipient type in the 'Manual referral needed' field, and a manual referral will be created with the information entered into the edit boxes above, and a TIKL will be made for 30 days from the date of manual referral."
  Text 40, 145, 60, 10, "Worker Signature:"
  Text 5, 120, 95, 10, "Other manual referral notes:"
  Text 5, 130, 60, 10, " (for the referral)"
EndDialog

'This dialog is for counties that have not provided FSET office address(s)
BeginDialog SNAPET_manual_address_dialog, 0, 0, 301, 275, "SNAP E&T Appointment Letter"
  EditBox 65, 5, 55, 15, MAXIS_case_number
  EditBox 215, 5, 20, 15, member_number
  EditBox 65, 25, 55, 15, appointment_date
  EditBox 195, 25, 20, 15, appointment_time_prefix_editbox
  EditBox 215, 25, 20, 15, appointment_time_post_editbox
  DropListBox 240, 25, 55, 15, "Select one..."+chr(9)+"AM"+chr(9)+"PM", AM_PM
  EditBox 65, 45, 190, 15, SNAPET_name
  EditBox 65, 65, 190, 15, SNAPET_address_01
  EditBox 65, 85, 95, 15, SNAPET_city
  EditBox 165, 85, 40, 15, SNAPET_ST
  EditBox 210, 85, 45, 15, SNAPET_zip
  EditBox 65, 105, 65, 15, SNAPET_contact
  EditBox 185, 105, 70, 15, SNAPET_phone
  DropListBox 110, 125, 80, 15, "Select one..."+chr(9)+"Banked months"+chr(9)+"Other manual referral"+chr(9)+"Student"+chr(9)+"Working with CBO", manual_referral
  EditBox 110, 145, 185, 15, other_referral_notes
  EditBox 75, 170, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 190, 170, 50, 15
    CancelButton 245, 170, 50, 15
  Text 5, 70, 55, 10, "Address line 1:"
  Text 10, 110, 55, 10, "Contact Name:"
  Text 135, 110, 50, 10, "Contact Phone:"
  Text 10, 175, 60, 10, "Worker Signature:"
  Text 10, 10, 50, 10, "Case Number:"
  Text 10, 90, 55, 10, "City/State/Zip:"
  Text 130, 10, 70, 10, "HH Member Number:"
  Text 5, 30, 60, 10, "Appointment Date:"
  GroupBox 5, 195, 290, 75, "When is a manual referral needed"
  Text 15, 210, 275, 20, "If an ABAWD is using banked months, or a student meets criteria under CM0011.18, or receiving E and T services through a Community Based Organization (CBO)."
  Text 15, 235, 275, 25, "Select a recipient type in the 'Manual referral needed' field, and a manual referral will be created with the information entered into the edit boxes above, and a TIKL will be made for 30 days from the date of manual referral."
  Text 130, 30, 60, 15, "Appointment Time:"
  Text 10, 130, 80, 10, "Manual referral needed:"
  Text 5, 50, 55, 10, "Provider Name:"
  Text 10, 150, 95, 10, "Other manual referral notes:"
  Text 10, 160, 55, 10, " (for the referral)"
EndDialog

'This is a Hennepin specific dialog, should not be used for other counties!!!!!!!!
BeginDialog SNAPET_Hennepin_dialog, 0, 0, 466, 205, "SNAP E&T Appointment Letter"
  EditBox 105, 10, 55, 15, MAXIS_case_number
  EditBox 220, 10, 25, 15, member_number
  DropListBox 105, 35, 195, 15, "Select one..."+chr(9)+"Somali-language (Sabathani, next Tuesday @ 2:00 p.m.)"+chr(9)+"Central NE (HSB, next Wednesday @ 2:00 p.m.)"+chr(9)+"North (HSB, next Wednesday @ 10:00 a.m.)"+chr(9)+"Northwest(Brookdale, next Monday @ 2:00 p.m.)"+chr(9)+"South Mpls (Sabathani, next Tuesday @ 10:00 a.m.)"+chr(9)+"South Suburban (Sabathani, next Tuesday @ 10:00 a.m.)"+chr(9)+"West (Sabathani, next Tuesday @ 10:00 a.m.)", interview_location
  DropListBox 105, 60, 110, 15, "Select one..."+chr(9)+"Banked months"+chr(9)+"Other manual referral"+chr(9)+"Student"+chr(9)+"Working with CBO", manual_referral
  EditBox 105, 80, 195, 15, other_referral_notes
  EditBox 105, 105, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 195, 105, 50, 15
    CancelButton 250, 105, 50, 15
  Text 10, 40, 95, 10, "Client's region of residence: "
  GroupBox 5, 130, 450, 65, "When is a manual referral needed"
  Text 20, 65, 80, 10, "Manual referral needed:"
  Text 15, 145, 435, 20, "If an ABAWD is using banked months, or a student meets criteria under CM0011.18, or receiving E and T services through a Community Based Organization (CBO)."
  GroupBox 310, 10, 145, 115, "For non-English speaking ABAWD's:"
  Text 15, 170, 435, 20, "Select a recipient type in the 'Manual referral needed' field, and a manual referral will be created with the information entered into the edit boxes above, and a TIKL will be made for 30 days from the date of manual referral."
  Text 50, 15, 50, 10, "Case Number:"
  Text 5, 85, 100, 15, "Other manual referral reason:"
  Text 170, 15, 45, 10, "HH Memb #:"
  Text 40, 110, 60, 10, "Worker Signature:"
  Text 320, 25, 130, 35, "If your client is requsting a Somali-language orientation, select this option in the 'client's region of residence' field."
  Text 320, 65, 130, 55, "For all other languages, do not use this script. Contact Mark Scherer, and request language-specific SNAP E and T Orientation/intake. Provide client with Mark’s contact information, and instruct them to contact him to schedule orientation within one week."
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone default screen & 'Searches for a case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

'defaults the member_number to 01
member_number = "01"

'Main dialog
DO
	DO
		'establishes  that the error message is equal to blank (necessary for the DO LOOP to work)
		err_msg = ""
		'these counties are exempt from participation per the FNS'
		If  worker_county_code = "x101" OR _
			worker_county_code = "x111" OR _
			worker_county_code = "x115" OR _
			worker_county_code = "x129" OR _
			worker_county_code = "x131" OR _
			worker_county_code = "x133" OR _
			worker_county_code = "x136" OR _
			worker_county_code = "x139" OR _
			worker_county_code = "x144" OR _
			worker_county_code = "x145" OR _
			worker_county_code = "x148" OR _
			worker_county_code = "x149" OR _
			worker_county_code = "x154" OR _
			worker_county_code = "x158" OR _
			worker_county_code = "x180" THEN
			script_end_procedure ("Your agency is exempt from ABAWD work requirements through 09/30/17." & vbNewLine & vbNewLine & " Please refer to TE02.05.69 for reference.")
		ElseIF worker_county_code = "x127" THEN
			Dialog SNAPET_Hennepin_dialog
			'Hennepin specific information===================================================================================================
			If worker_county_code = "x127" THEN
				SNAPET_contact = "the SNAP Employment and Training team"
				SNAPET_phone = "612-596-7411"
			END IF
			'CO #27 HENNEPIN COUNTY addresses, date and times of orientations
			'Somali-language orientation
			IF interview_location = "Somali-language (Sabathani, next Tuesday @ 2:00 p.m.)" then
				SNAPET_name = "Sabathani Community Center"
				SNAPET_address_01 = "310 East 38th Street #120"
				SNAPET_city = "Minneapolis"
				SNAPET_ST = "MN"
				SNAPET_zip = "55409"
				appointment_time_prefix_editbox = "02"
				appointment_time_post_editbox = "00"
				AM_PM = "PM"
				appointment_date = Date + 8 - Weekday(Date, vbTuesday)
			'Central NE
			Elseif interview_location = "Central NE (HSB, next Wednesday @ 2:00 p.m.)" THEN
				SNAPET_name = "Health Services Building"
				SNAPET_address_01 = "525 Portland Ave, 5th floor"
				SNAPET_city = "Minneapolis"
				SNAPET_ST = "MN"
				SNAPET_zip = "55415"
				appointment_time_prefix_editbox = "02"
				appointment_time_post_editbox = "00"
				AM_PM = "PM"
				appointment_date = Date + 8 - Weekday(Date, vbWednesday)
			'North
			ElseIF interview_location = "North (HSB, next Wednesday @ 10:00 a.m.)" THEN
				SNAPET_name = "Health Services Building"
				SNAPET_address_01 = "525 Portland Ave, 5th floor"
				SNAPET_city = "Minneapolis"
				SNAPET_ST = "MN"
				SNAPET_zip = "55415"
				appointment_time_prefix_editbox = "10"
				appointment_time_post_editbox = "00"
				AM_PM = "AM"
			appointment_date = Date + 8 - Weekday(Date, vbWednesday)
			'Northwest
			ElseIf interview_location = "Northwest(Brookdale, next Monday @ 2:00 p.m.)" THEN
				SNAPET_name = "Brookdale Human Services Center"
				SNAPET_address_01 = "6125 Shingle Creek Parkway, Suite 400"
				SNAPET_city = "Brooklyn Center"
				SNAPET_ST = "MN"
				SNAPET_zip = "55430"
				appointment_time_prefix_editbox = "02"
				appointment_time_post_editbox = "00"
				AM_PM = "PM"
				appointment_date = Date + 8 - Weekday(Date, vbMonday)
			'South Minneapolis
			ElseIf interview_location = "South Mpls (Sabathani, next Tuesday @ 10:00 a.m.)" THEN
				SNAPET_name = "Sabathani Community Center"
				SNAPET_address_01 = "310 East 38th Street #120"
				SNAPET_city = "Minneapolis"
				SNAPET_ST = "MN"
				SNAPET_zip = "55409"
				appointment_time_prefix_editbox = "10"
				appointment_time_post_editbox = "00"
				AM_PM = "AM"
				appointment_date = Date + 8 - Weekday(Date, vbTuesday)
			'South Suburban
			ElseIf interview_location = "South Suburban (Sabathani, next Tuesday @ 10:00 a.m.)" THEN
				SNAPET_name = "Sabathani Community Center"
				SNAPET_address_01 = "310 East 38th Street #120"
				SNAPET_city = "Minneapolis"
				SNAPET_ST = "MN"
				SNAPET_zip = "55409"
				appointment_time_prefix_editbox = "10"
				appointment_time_post_editbox = "00"
				AM_PM = "AM"
				appointment_date = Date + 8 - Weekday(Date, vbTuesday)
			'West
			ElseIf interview_location = "West (Sabathani, next Tuesday @ 10:00 a.m.)" THEN
				SNAPET_name = "Sabathani Community Center"
				SNAPET_address_01 = "310 East 38th Street #120"
				SNAPET_city = "Minneapolis"
				SNAPET_ST = "MN"
				SNAPET_zip = "55409"
				appointment_time_prefix_editbox = "10"
				appointment_time_post_editbox = "00"
				AM_PM = "AM"
				appointment_date = Date + 8 - Weekday(Date, vbTuesday)
			END IF
	'Counties listed here (starting with x105 and ending with x185 did not provide E & T office information, hence will need to use the dialog requiring them to enter in their own address and contact information)
		ELSEIF worker_county_code = "x105" OR _
			worker_county_code = "x106" OR _
			worker_county_code = "x110" OR _
			worker_county_code = "x113" OR _
			worker_county_code = "x114" OR _
			worker_county_code = "x116" OR _
			worker_county_code = "x117" OR _
			worker_county_code = "x124" OR _
			worker_county_code = "x132" OR _
			worker_county_code = "x134" OR _
			worker_county_code = "x149" OR _
			worker_county_code = "x152" OR _
			worker_county_code = "x153" OR _
			worker_county_code = "x162" OR _
			worker_county_code = "x170" OR _
			worker_county_code = "x172" OR _
			worker_county_code = "x173" OR _
			worker_county_code = "x183" OR _
			worker_county_code = "x185" OR _
			worker_county_code = "" THEN
			Dialog SNAPET_manual_address_dialog
		ELSE
			Dialog SNAPET_automated_adress_dialog
			'next 5 lines are tricking the script to read <> "" since they are declared as "_"
			SNAPET_name = "_"
			SNAPET_address_01 = "_"
			SNAPET_city = "_"
			SNAPET_ST = "_"
			SNAPET_zip = "_"
		END IF
		'asks if they really want to cancel script
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If isdate(appointment_date) = FALSE then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		'The DateValue condition does not apply to Hennepin County users which is why it is excluded in the line below
		IF worker_county_code <> "x127" AND DateValue(appointment_date) < date then err_msg = err_msg & vbNewLine & "* Orientation date entered has already passed.  Select a new date."
		IF len(member_number) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid member number."
		IF SNAPET_name = "" then err_msg = err_msg & vbNewLine & "* Enter a E and T office location."
		IF SNAPET_address_01 = "" then err_msg = err_msg & vbNewLine & "* Enter a street address."
		IF appointment_time_prefix_editbox = "" then err_msg = err_msg & vbNewLine & "* Enter a valid appointment time."
		IF appointment_time_post_editbox = "" then err_msg = err_msg & vbNewLine & "* Enter a valid appointment time."
		If AM_PM = "Select one..." then err_msg = err_msg & vbNewLine & "* Select either AM or PM for your appointment time."
		IF SNAPET_contact = "" then err_msg = err_msg & vbNewLine & "* Enter a contact name."
		IF SNAPET_phone = "" then err_msg = err_msg & vbNewLine & "* Enter a phone number."
		If interview_location = "Select one..." then err_msg = err_msg & vbNewLine & "* Enter an interview location."
		IF (manual_referral = "Other manual referral" and other_referral_notes = "") then err_msg = err_msg & vbNewLine & "* Enter other manual referral notes."
		If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'The Hennepin County worker must confirm the appointment time and date, and gives them the option to select another date
If worker_county_code = "x127" THEN
	DO
		DO
			orientation_date_confirmation = MsgBox("Press YES to confirm the orientation date. For the next week, press NO." & vbNewLine & vbNewLine & _
			"                                                  " & appointment_date & " at " & appointment_time_prefix_editbox & ":" & appointment_time_post_editbox & _
			AM_PM, vbYesNoCancel, "Please confirm the SNAP E & T orientation referral date")
			If orientation_date_confirmation = vbCancel then script_end_procedure ("The script has ended. An orientation letter has not been sent.")
			If orientation_date_confirmation = vbYes then exit do
			If orientation_date_confirmation = vbNo then appointment_date = dateadd("d", 7, appointment_date)
		LOOP until orientation_date_confirmation = vbYes
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in
END IF

'County FSET address information which will autofill when option is chosen from county_office_list----------------------------------------------------------------------------------------------------
'CO #01 AITKIN COUNTY address
IF interview_location = "Aitkin Workforce Center" THEN
	SNAPET_name = "Aitkin Workforce Center"
	SNAPET_address_01 = "20 3rd Street NE"
	SNAPET_city = "Aitkin"
	SNAPET_ST = "MN"
	SNAPET_zip = "56431"
END IF

'CO #02 Anoka County address
IF interview_location = "Minnesota WorkForce Center Blaine" THEN
	SNAPET_name = "Minnesota WorkForce Center Blaine"
	SNAPET_address_01 = "1201 89th Avenue NE Suite 235"
	SNAPET_city = "Blaine"
	SNAPET_ST = "MN"
	SNAPET_zip = "55434"
END IF

'CO #3 BECKER COUNTY address
IF interview_location = "Rural MN CEP Detroit Lakes" THEN
	SNAPET_name = "Rural MN CEP Detroit Lakes"
	SNAPET_address_01 = "1803 Roosevelt Ave"
	SNAPET_city = "Detroit Lakes"
	SNAPET_ST = "MN"
	SNAPET_zip = "56501"
END IF

'CO #04 BELTRAMI COUNTY addresses
IF interview_location = "RMCEP" THEN
	SNAPET_name = "RMCEP"
	SNAPET_address_01 = "616 America Ave NW Suite 210"
	SNAPET_city = "Bemedji"
	SNAPET_ST = "MN"
	SNAPET_zip = "56601"
ElseIf interview_location = "MCT" THEN
	SNAPET_name = "MCT"
	SNAPET_address_01 = "15542 State Hwy 371 NW"
	SNAPET_city = "Cass Lake"
	SNAPET_ST = "MN"
	SNAPET_zip = "56633"
ElseIf interview_location = "Leach Lake New" THEN
	SNAPET_name = "Leach Lake New"
	SNAPET_address_01 = "190 Sail Drive NW"
	SNAPET_city = "Cass Lake"
	SNAPET_ST = "MN"
	SNAPET_zip = "56633"
ElseIf interview_location = "Red Lake Oshkiimaajitahdah" THEN
	SNAPET_name = "Red Lake Oshkiimaajitahdah"
	SNAPET_address_01 = "MN-1"
	SNAPET_city = "Redby"
	SNAPET_ST = "MN"
	SNAPET_zip = "56670"
END IF

'CO #7 BLUE EARTH COUNTY address
IF interview_location = "Blue Earth County Employment Services" THEN
	SNAPET_name = "Blue Earth County Employment Services"
	SNAPET_address_01 = "421 E Hickory Street, Suite 400"
	SNAPET_city = "Mankato"
	SNAPET_ST = "MN"
	SNAPET_zip = "56001"
END IF

'CO #8 BROWN COUNTY address
IF interview_location = "Minnesota Valley Action Council New Ulm" THEN
	SNAPET_name = "Minnesota Valley Action Council New Ulm"
	SNAPET_address_01 = "1618 Broadway"
	SNAPET_city = "New Ulm"
	SNAPET_ST = "MN"
	SNAPET_zip = "56073"
END IF

'CO #9 CARLTON COUNTY address
IF interview_location = "Carlton County Human Services" THEN
	SNAPET_name = "Carlton County Human Services"
	SNAPET_address_01 = "14 N. 11th Street"
	SNAPET_city = "Cloquet"
	SNAPET_ST = "MN"
	SNAPET_zip = "55720"
END IF

'CO #12 CHIPPEWA COUNTY address
IF interview_location = "Chippewa County Workforce Center" THEN
	SNAPET_name = "Chippewa County Workforce Center"
	SNAPET_address_01 = "202 N 1st Street Suite 100"
	SNAPET_city = "Montevideo"
	SNAPET_ST = "MN"
	SNAPET_zip = "56265"
END IF

'CO #18 CROW WING COUNTY address
IF interview_location =  "Rural MN CEP Brainerd" THEN
	SNAPET_name = "Rural MN CEP Brainerd"
	SNAPET_address_01 = "204 Laurel Street Suite 21"
	SNAPET_city = "Brainerd"
	SNAPET_ST = "MN"
	SNAPET_zip = "56401"
END IF

'CO #19 DAKOTA COUNTY address
IF interview_location = "Northern Service Center" THEN
	SNAPET_name = "Northern Service Center"
	SNAPET_address_01 = "1 Mendota Road W Suite 170"
	SNAPET_city = "West St. Paul"
	SNAPET_ST = "MN"
	SNAPET_zip = "55118"
ELSEIF interview_location = "Burnsville Workforce Center" THEN
	SNAPET_name = "Burnsville Workforce Center"
	SNAPET_address_01 = "2800 W County Road 42"
	SNAPET_city = "Burnsville"
	SNAPET_ST = "MN"
	SNAPET_zip = "55337"
END IF

'CO #20 DODGE COUNTY address
IF interview_location = "Workforce Development Inc. (Kasson)" THEN
	SNAPET_name = "Workforce Development Inc. (Kasson)"
	SNAPET_address_01 = "504 S Mantorville Ave Suite 4"
	SNAPET_city = "Kasson"
	SNAPET_ST = "MN"
	SNAPET_zip = "55944"
END IF

'CO #21 DOUGLAS COUNTY address
IF interview_location = "Alexandria Workforce Center" THEN
	SNAPET_name = "Alexandria Workforce Center"
	SNAPET_address_01 = "303 22nd Avenue W Suite 107"
	SNAPET_city = "Alexandria"
	SNAPET_ST = "MN"
	SNAPET_zip = "56308"
END IF

'CO #22 FAIRBAULT COUNTY address
IF interview_location =  "Fairmont Workforce Center Fairbault County" THEN
	SNAPET_name = "Fairmont Workforce Center Fairbault County"
	SNAPET_address_01 = "301 N. Main Street"
	SNAPET_city = "Blue Earth"
	SNAPET_ST = "MN"
	SNAPET_zip = "56013"
END IF

'CO #23 FILLMORE COUNTY address
IF interview_location = "Workforce Development Office" THEN
	SNAPET_name = "Workforce Development Office"
	SNAPET_address_01 = "100 South Main"
	SNAPET_city = "Preston"
	SNAPET_ST = "MN"
	SNAPET_zip = "55965"
END IF

'CO #25 GOODHUE COUNTY address
IF interview_location = "Workforce Development Inc. (Redwing)" THEN
	SNAPET_name = "Workforce Development Inc. (Redwing)"
	SNAPET_address_01 = "1606 West 3rd Street"
	SNAPET_city = "Red Wing"
	SNAPET_ST = "MN"
	SNAPET_zip = "55066"
END IF

'CO #26 GRANT COUNTY address
IF interview_location = "Grant County Social Services" THEN
	SNAPET_name = "Grant County Social Services"
	SNAPET_address_01 = "28 Central Avenue S"
	SNAPET_city = "Elbow Lake"
	SNAPET_ST = "MN"
	SNAPET_zip = "56531"
END IF

'CO #27 Hennepin County address is listed at the top of the script as there are some Hennepin specific stuff in the script

'CO #28 HOUSTON COUNTY address
IF interview_location  = "Workforce Development Inc." THEN
    SNAPET_name = "Workforce Development Inc."
    SNAPET_address_01 = "110 E Grove Street"
    SNAPET_city = "Caledonia"
    SNAPET_ST = "MN"
    SNAPET_zip = "55921"
END IF

'CO #30 ISANTI COUNTY address
IF interview_location  = "Cambridge MN Workforce Center" THEN
    SNAPET_name = "Cambridge MN Workforce Center"
    SNAPET_address_01 = "140 Buchanan Street Suite 152"
    SNAPET_city = "Cambridge"
    SNAPET_ST = "MN"
    SNAPET_zip = "55008"
END IF

'CO #31 ITASCA COUNTY address
IF interview_location  = "AEOA – GR Workforce Center" THEN
    SNAPET_name = "AEOA – GR Workforce Center"
    SNAPET_address_01 = "1215 SE 2nd Ave"
    SNAPET_city = "Grand Rapids"
    SNAPET_ST = "MN"
    SNAPET_zip = "55744"
END IF

'CO #35 KITTSON COUNTY address
IF interview_location  = "Kittson County Social Services" THEN
    SNAPET_name = "Kittson County Social Services"
    SNAPET_address_01 = "410 5th Street S #100"
    SNAPET_city = "Hallock"
    SNAPET_ST = "MN"
    SNAPET_zip = "56728"
END IF

'CO #37 LAC QUI PARLE COUNTY address
IF interview_location  = "Lace qui Parle Co. Family Services" THEN
    SNAPET_name = "Lace qui Parle Co. Family Services"
    SNAPET_address_01 = "930 1st Ave"
    SNAPET_city = "Madison"
    SNAPET_ST = "MN"
    SNAPET_zip = "56256"
END IF

'CO #38 LAKE COUNTY address
IF interview_location  = "AEOA" THEN
    SNAPET_name = "AEOA "
    SNAPET_address_01 = "2124 10th Street"
    SNAPET_city = "Two Harbors"
    SNAPET_ST = "MN"
    SNAPET_zip = "55616"
END IF

'CO #39 LAKE OF THE WOODS COUNTY address
IF interview_location  = "Rural MN CEP Lake of the Woods" THEN
    SNAPET_name = "Rural MN CEP Lake of the Woods"
    SNAPET_address_01 = "616 America Ave NW Suite 220"
    SNAPET_city = "Bemedji"
    SNAPET_ST = "MN"
    SNAPET_zip = "55601"
END IF

'CO #40 LE SUEUR COUNTY address
IF interview_location  = "MVAC" THEN
    SNAPET_name = "MVAC"
    SNAPET_address_01 = "125 E. Minnesota Street"
    SNAPET_city = "Le Center"
    SNAPET_ST = "MN"
    SNAPET_zip = "56057"
END IF

'CO #41 LINCOLN COUNTY address
IF interview_location  = "Marshall WorkForce Center" THEN
    SNAPET_name = "Marshall WorkForce Center"
    SNAPET_address_01 = "607 W. Main Street"
    SNAPET_city = "Marshall"
    SNAPET_ST = "MN"
    SNAPET_zip = "56258"
END IF

'CO #42 LYON COUNTY address
IF interview_location  = "Marshall WorkForce Center" THEN
    SNAPET_name = "Marshall WorkForce Center"
    SNAPET_address_01 = "607 W. Main Street"
    SNAPET_city = "Marshall"
    SNAPET_ST = "MN"
    SNAPET_zip = "56258"
END IF

'CO #43 COUNTY address
IF interview_location  = "Mahnomen County Human Services" THEN
    SNAPET_name = "Mahnomen County Human Services"
    SNAPET_address_01 = "311 N. Main Street"
    SNAPET_city = " Mahnomen"
    SNAPET_ST = "MN"
    SNAPET_zip = "56557"
END IF

'CO #44 MARSHALL COUNTY address
IF interview_location  = "Marshall County Social Services" THEN
    SNAPET_name = "Marshall County Social Services"
    SNAPET_address_01 = "208 E Colvin Street Suite 14"
    SNAPET_city = "Warren"
    SNAPET_ST = "MN"
    SNAPET_zip = "56762"
END IF

'CO #45 MARTIN COUNTY address
IF interview_location  = "Fairmont Workforce Center Martin County" THEN
    SNAPET_name = "Fairmont Workforce Center Martin County"
    SNAPET_address_01 = "412 S. State Street"
    SNAPET_city = "Fairmont"
    SNAPET_ST = "MN"
    SNAPET_zip = "56013"
END IF

'CO #46 MCLEOD COUNTY address
IF interview_location  = "Central MN Jobs and Training Services Hutchinson" THEN
    SNAPET_name = "Central MN Jobs and Training Services Hutchinson"
    SNAPET_address_01 = " 2 Century Avenue"
    SNAPET_city = "Hutchinson"
    SNAPET_ST = "MN"
    SNAPET_zip = "55350"
END IF

'CO #47 MEEKER COUNTY address
IF interview_location  = "Central MN Jobs and Training Services Litchfield" THEN
    SNAPET_name = "Central MN Jobs and Training Services Litchfield"
    SNAPET_address_01 = "114 N Holcombe Avenue Suite 170"
    SNAPET_city = "Litchfield"
    SNAPET_ST = "MN"
    SNAPET_zip = "55355"
END IF

'CO #50 MOWER COUNTY address
IF interview_location  = "Workforce Development Inc. (Austin)" THEN
    SNAPET_name = "Workforce Development Inc. (Austin)"
    SNAPET_address_01 = "1600 8th Avenue NW"
    SNAPET_city = "Austin"
    SNAPET_ST = "MN"
    SNAPET_zip = "55912"
END IF

'CO #51 MURRAY COUNTY address
IF interview_location  = "Marshall WorkForce Center" THEN
    SNAPET_name = "Marshall WorkForce Center"
    SNAPET_address_01 = "607 W. Main Street"
    SNAPET_city = "Marshall"
    SNAPET_ST = "MN"
    SNAPET_zip = "56258"
END IF

'CO #55 OLMSTEAD COUNTY address
IF interview_location  = "Olmstead County Family Support & Assistance" THEN
    SNAPET_name = "Olmstead County Family Support & Assistance"
    SNAPET_address_01 = "2117 Campus Drive SE Suite 100"
    SNAPET_city = "Rochester"
    SNAPET_ST = "MN"
    SNAPET_zip = "55904"
END IF

'CO #56 OTTER TAIL COUNTY address
IF interview_location  = "Rural MN CEP Fergus Falls" THEN
    SNAPET_name = "Rural MN CEP Fergus Falls"
    SNAPET_address_01 = "125 W Lincoln Avenue"
    SNAPET_city = "Fergus Falls"
    SNAPET_ST = "MN"
    SNAPET_zip = "56537"
END IF

'CO #57 PENNINGTON COUNTY address
IF interview_location  = "Minnesota WorkForce Center: Thief River Falls" THEN
    SNAPET_name = "Minnesota WorkForce Center: Thief River Falls"
    SNAPET_address_01 = "1301 State Hwy 1"
    SNAPET_city = "Thief River Falls"
    SNAPET_ST = "MN"
    SNAPET_zip = "56701"
END IF

'CO #58 PINE COUNTY address
IF interview_location  = "Pine County Public Health Building" THEN
    SNAPET_name = "Pine County Public Health Building"
    SNAPET_address_01 = "1610 Hwy 23 N"
    SNAPET_city = "Sandstone"
    SNAPET_ST = "MN"
    SNAPET_zip = "55072"
ELSEIF interview_location  = "Pine Technical & Community College E&T Center" THEN
    SNAPET_name = "Pine Technical & Community College E&T Center"
    SNAPET_address_01 = "900 4th St SE"
    SNAPET_city = "Pine City"
    SNAPET_ST = "MN"
    SNAPET_zip = "55063"
END IF

'CO #59 PIPESTONE COUNTY address
IF interview_location  = "Southwest MN Private Industry Council Inc. Pipestone" THEN
    SNAPET_name = "Southwest MN Private Industry Council Inc. Pipestone"
    SNAPET_address_01 = "1091 N Hiawatha Avenue"
    SNAPET_city = "Pipestone"
    SNAPET_ST = "MN"
    SNAPET_zip = "56164"
END IF

'CO #60 POLK COUNTY address
IF interview_location  = "Polk County Social Services: Crookston" THEN
    SNAPET_name = "Polk County Social Services: Crookston"
    SNAPET_address_01 = "612 N Broadway Room 302"
    SNAPET_city = "Crookston"
    SNAPET_ST = "MN"
    SNAPET_zip = "56716"
ELSEIF interview_location  = "Polk County Social Services: East Grand Forks" THEN
    SNAPET_name = "Polk County Social Services: East Grand Forks"
    SNAPET_address_01 = "1424 Central Ave NE Suite 104"
    SNAPET_city = "East Grand Forks"
    SNAPET_ST = "MN"
    SNAPET_zip = "56721"
ELSEIF interview_location  = "Polk County Social Services: Fosston" THEN
    SNAPET_name = "Polk County Social Services: Fosston"
    SNAPET_address_01 = "104 Kaiser Ave"
    SNAPET_city = "Fosston"
    SNAPET_ST = "MN"
    SNAPET_zip = "56542"
END IF

'CO #61 POPE COUNTY address
IF interview_location = "Minnesota Workforce Center Alexandria" THEN
	SNAPET_name = "Minnesota Workforce Center Alexandria"
	SNAPET_address_01 = "303 22nd Avenue W Suite 107"
	SNAPET_city = "Alexandria"
	SNAPET_ST = "MN"
	SNAPET_zip = "56308"
END IF

'CO #63 REDLAKE COUNTY address
IF interview_location  = "Minnesota Workforce Center: Red Lake" THEN
    SNAPET_name = "Minnesota Workforce Center: Red Lake"
    SNAPET_address_01 = "1301 Highway 1 East"
    SNAPET_city = "Thief River Falls"
    SNAPET_ST = "MN"
    SNAPET_zip = "56701"
END IF

'CO #64 REDWOOD COUNTY address
IF interview_location  = "Southwest Health & Human Services" THEN
    SNAPET_name = "Southwest Health & Human Services"
    SNAPET_address_01 = "266 E. Bridge Street"
    SNAPET_city = "Redwood Falls"
    SNAPET_ST = "MN"
    SNAPET_zip = "56283"
END IF

'CO #65 RENVILLE COUNTY address
IF interview_location  = "Central MN Jobs and Training Services Olivia" THEN
    SNAPET_name = "Central MN Jobs and Training Services Olivia"
    SNAPET_address_01 = "1005 W. Elm Ave. Ste. 2"
    SNAPET_city = "Olivia"
    SNAPET_ST = "MN"
    SNAPET_zip = "56277"
END IF

'CO #67 ROCK COUNTY address
IF interview_location  = "Southwest MN Private Industry Council Inc. Luverne" THEN
    SNAPET_name = "Southwest MN Private Industry Council Inc. Luverne"
    SNAPET_address_01 = "2 Roundwind Road"
    SNAPET_city = "Luverne"
    SNAPET_ST = "MN"
    SNAPET_zip = "56156"
END IF

'CO #68 ROSEAU COUNTY address
IF interview_location  = "Roseau County Social Services" THEN
    SNAPET_name = "Roseau County Social Services"
    SNAPET_address_01 = "208 6th Street SW"
    SNAPET_city = "Roseau"
    SNAPET_ST = "MN"
    SNAPET_zip = "56751"
END IF

'CO #69 SAINT LOUIS COUNTY address
IF interview_location  = "Minnesota WorkForce Center: Duluth" THEN
    SNAPET_name = "Minnesota WorkForce Center: Duluth"
    SNAPET_address_01 = "402 W. 1st Street Room 119"
    SNAPET_city = "Duluth"
    SNAPET_ST = "MN"
    SNAPET_zip = "55802"
ELSEIF interview_location  = "Minnesota WorkForce Center: Hibbing" THEN
    SNAPET_name = "Minnesota WorkForce Center:  Hibbing"
    SNAPET_address_01 = "3920 13th Avenue E"
    SNAPET_city = "Hibbing"
    SNAPET_ST = "MN"
    SNAPET_zip = "55746"
ELSEIF interview_location  = "Minnesota WorkForce Center: Virginia" THEN
    SNAPET_name = "Minnesota WorkForce Center: Virginia"
    SNAPET_address_01 = "820 9th St"
    SNAPET_city = "Virginia"
    SNAPET_ST = "MN"
    SNAPET_zip = "55792"
END IF

'CO #71 SHERBURNE COUNTY address
IF interview_location  = "Central MN Jobs and Training Services Monticello" THEN
    SNAPET_name = "Central MN Jobs and Training Services Monticello"
    SNAPET_address_01 = "406 7th Street East"
    SNAPET_city = "Monticello"
    SNAPET_ST = "MN"
    SNAPET_zip = "55362"
END IF

'CO #74 STEELE COUNTY address
IF interview_location  = "Steele County Employment Services" THEN
    SNAPET_name = "Steele County Employment Services "
    SNAPET_address_01 = "630 Florence Avenue Suite 20"
    SNAPET_city = "Owatonna"
    SNAPET_ST = "MN"
    SNAPET_zip = "55060"
END IF

'CO #75 STEVENS COUNTY address
IF interview_location  = "Stevens County Human Services" THEN
    SNAPET_name = "Stevens County Human Services"
    SNAPET_address_01 = "400 Colorado Ave Suite 104"
    SNAPET_city = "Morris"
    SNAPET_ST = "MN"
    SNAPET_zip = "56267"
END IF

'CO #76 SWIFT COUNTY address
IF interview_location  = "SW MN Private Industry Council" THEN
    SNAPET_name = "SW MN Private Industry Council"
    SNAPET_address_01 = "129 W Nichols"
    SNAPET_city = "Montevideo"
    SNAPET_ST = "MN"
    SNAPET_zip = "56265"
END IF

'CO #77 TODD COUNTY address
IF interview_location  = "Todd County Health & Human Services: Long Prairie" THEN
    SNAPET_name = "Todd County Health & Human Services: Long Prairie"
    SNAPET_address_01 = "212 2nd Avenue S."
    SNAPET_city = "Long Prairie"
    SNAPET_ST = "MN"
    SNAPET_zip = "56347"
ELSEIF interview_location  = "Todd County Health & Human Services: Staples" THEN
    SNAPET_name = "Todd County Health & Human Services: Staples"
    SNAPET_address_01 = "200 1st St NE Suite #1"
    SNAPET_city = "Staples"
    SNAPET_ST = "MN"
    SNAPET_zip = "56479"
END IF

'CO #78 TRAVERSE COUNTY address
IF interview_location  = "Rural MN CEP Wheaton" THEN
    SNAPET_name = "Rural MN CEP Wheaton"
    SNAPET_address_01 = "202 8th Street N"
    SNAPET_city = "Wheaton"
    SNAPET_ST = "MN"
    SNAPET_zip = "56296"
END IF

'CO #79 WABASHA COUNTY address
IF interview_location  = "Workforce Development Inc. (Wabasha)" THEN
    SNAPET_name = "Workforce Development Inc. (Wabasha)"
    SNAPET_address_01 = "222 Main Street West"
    SNAPET_city = "Wabasha"
    SNAPET_ST = "MN"
    SNAPET_zip = "55981"
END IF

'CO #80 WADENA COUNTY address
IF interview_location  = "Rural MN CEP Wadena" THEN
    SNAPET_name = "Rural MN CEP Wadena"
    SNAPET_address_01 = "124 First Street SE"
    SNAPET_city = "Wadena"
    SNAPET_ST = "MN"
    SNAPET_zip = "56482"
END IF

'CO #81 WASECA COUNTY address
IF interview_location  = "Minnesota Valley Action Council Waseca" THEN
    SNAPET_name = "Minnesota Valley Action Council Waseca"
    SNAPET_address_01 = "108 10th Avenue SE"
    SNAPET_city = "Waseca"
    SNAPET_ST = "MN"
    SNAPET_zip = "56093"
END IF

'CO #82 WASHINGTON COUNTY address
IF interview_location  = "Washington County Community Services: Stillwater" THEN
    SNAPET_name = "Washington County Community Services: Stillwater"
    SNAPET_address_01 = "14949 62nd Street North"
    SNAPET_city = "Stillwater"
    SNAPET_ST = "MN"
    SNAPET_zip = "55082"
ElseIF interview_location  = "Washington County Community Services: Cottage Grove" THEN
    SNAPET_name = "Washington County Community Services: Cottage Grove"
    SNAPET_address_01 = "13000 Ravine Parkway South"
    SNAPET_city = "Cottage Grove"
    SNAPET_ST = "MN"
    SNAPET_zip = "55016"
ELSEIF interview_location  = "Washington County Community Services: Forest Lake" THEN
    SNAPET_name = "Washington County Community Services: Forest Lake"
    SNAPET_address_01 = "19955 Forest Road North"
    SNAPET_city = "Forest Lake"
    SNAPET_ST = "MN"
    SNAPET_zip = "55025"
ELSEIF interview_location  = "Washington County Community Services: Woodbury" THEN
    SNAPET_name = "Washington County Community Services: Woodbury"
    SNAPET_address_01 = "2150 Radio Drive"
    SNAPET_city = "Woodbury"
    SNAPET_ST = "MN"
    SNAPET_zip = "55125"
END IF

'CO #84 WILKIN COUNTY address
IF interview_location  = "Wilkin County Family Services" THEN
    SNAPET_name = "Wilkin County Family Services"
    SNAPET_address_01 = "300 South 5th Street"
    SNAPET_city = "Breckenridge"
    SNAPET_ST = "MN"
    SNAPET_zip = "56520"
END IF

'CO #86 WRIGHT COUNTY address
IF interview_location  = "Central MN Jobs and Training Services Monticello" THEN
    SNAPET_name = "Central MN Jobs and Training Services Monticello"
    SNAPET_address_01 = "406 E. 7th Street"
    SNAPET_city = "Monticello"
    SNAPET_ST = "MN"
    SNAPET_zip = "55362"
END IF

'CO #87 YELLOW MEDICINE COUNTY address
IF interview_location  = "Yellow Medicine County Family Services" THEN
    SNAPET_name = "Yellow Medicine County Family Services"
    SNAPET_address_01 = "930 4th Street, Suite 4"
    SNAPET_city = "Granite Falls"
    SNAPET_ST = "MN"
    SNAPET_zip = "56241"
END IF
'END COUNTY ADDRESSES----------------------------------------------------------------------------------------------------

'Pulls the member name.
call navigate_to_MAXIS_screen("STAT", "MEMB")
EMWriteScreen member_number, 20, 76
transmit
EMReadScreen memb_error_check, 7, 8, 22
If memb_error_check = "Arrival" then	'checking for valid HH member
	PF3
	PF10
	script_end_procedure("The HH member is invalid. Please review your case, and the HH member number before trying the script again.")
END IF
EMReadScreen last_name, 24, 6, 30
EMReadScreen first_name, 11, 6, 63
last_name = trim(replace(last_name, "_", ""))
first_name = trim(replace(first_name, "_", ""))

'Updates the WREG panel with the appointment_date
Call navigate_to_MAXIS_screen("STAT", "WREG")
EMWriteScreen member_number, 20, 76
transmit

'Ensuring that students have a FSET status of "12" and all others are coded with "30"
EMReadScreen FSET_status, 2, 8, 50
If manual_referral = "Student" then
    if FSET_status <> "12" then script_end_procedure ("Member " & member_number & " is not coded as a student. The script will now end.")
Else
    If FSET_status <> "30" then script_end_procedure("Member " & member_number & " is not coded as a Mandatory FSET Participant. The script will now end.")
End if
'Ensuring that the ABAWD_status is "13" for banked months manual referral recipients
EMReadScreen ABAWD_status, 2, 13, 50
If manual_referral = "Banked months" then
 	if ABAWD_status <> "13" then script_end_procedure ("Member " & member_number & " is not coded as a banked months recipient. The script will now end.")
End if

'Ensuring the orientation date is coding in the with the referral date scheduled
EMReadScreen orientation_date, 8, 9, 50
orientation_date = replace(orientation_date, " ", "/")
If appointment_date <> orientation_date then
	PF9
	Call create_MAXIS_friendly_date(appointment_date, 0, 9, 50)
	PF3
END if

'The CASE/NOTE----------------------------------------------------------------------------------------------------
'Navigates to a blank case note
start_a_blank_CASE_NOTE
CALL write_variable_in_case_note("***SNAP E&T Appointment Letter Sent for MEMB " & member_number & " ***")
Call write_bullet_and_variable_in_case_note("Member referred to E&T", member_number & " " & first_name & " " & last_name)
CALL write_bullet_and_variable_in_case_note("Appointment date", appointment_date)
CALL write_bullet_and_variable_in_case_note("Appointment time", appointment_time_prefix_editbox & ":" & appointment_time_post_editbox & " " & AM_PM)
CALL write_bullet_and_variable_in_case_note("Appointment location", SNAPET_name)
Call write_variable_in_case_note("* The WREG panel has been updated to reflect the E & T orientation date.")
If manual_referral <> "Select one..." then Call write_variable_in_case_note("* Manual referral made for: " & manual_referral & " recipient.")
If manual_referral <> "Select one..." then Call write_variable_in_case_note("* TIKL set for 30 days for proof of compliance with E & T.")
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

'The SPEC/LETR----------------------------------------------------------------------------------------------------
call navigate_to_MAXIS_screen("SPEC", "LETR")
'Opens up the SNAP E&T Orientation LETR. If it's unable the script will stop.
EMWriteScreen "x", 8, 12
transmit
EMReadScreen LETR_check, 4, 2, 49
If LETR_check = "LETR" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")

'Writes the info into the LETR.
IF len(appointment_time_prefix_editbox) = 1 THEN appointment_time_prefix_editbox = "0" & appointment_time_prefix_editbox 'This prevents the letter from being cancelled due to single digit hour
EMWriteScreen first_name & " " & last_name, 4, 28
call create_MAXIS_friendly_date_three_spaces_between(appointment_date, 0, 6, 28)
EMWriteScreen appointment_time_prefix_editbox, 7, 28
EMWriteScreen appointment_time_post_editbox, 7, 33
EMWriteScreen AM_PM, 7, 38
EMWriteScreen SNAPET_name, 9, 28
EMWriteScreen SNAPET_address_01, 10, 28
EMWriteScreen SNAPET_city & ", " & SNAPET_ST & " " &  SNAPET_zip, 11, 28
call create_MAXIS_friendly_phone_number(SNAPET_phone, 13, 28) 'takes out non-digits if listed in variable, and formats phone number for the field
EMWriteScreen SNAPET_contact, 16, 28
PF4		'saves and sends memo
PF3
PF3

'Creates a 30 day TILK to check for compliance with E & T'
If manual_referral <> "Select one..." then
	call navigate_to_MAXIS_screen("dail", "writ")
	call create_MAXIS_friendly_date(date, 30, 5, 18)
	Call write_variable_in_TIKL("Manual referral was made for " & other_referral_notes & " recipient 30 days ago. Please review case to see if verification of E and T compliance was sent to recipient, and that they are complying.")
	transmit
	PF3
End if

'Manual referral creation if banked months are used
If manual_referral <> "Select one..." then 					'if banked months or student are noted, then a manual referral to E & T is needed
	Call navigate_to_MAXIS_screen("INFC", "WF1M")			'navigates to WF1M to create the manual referral'
	EMWriteScreen "01", 4, 47													'this is the manual referral code that DHS has approved
	EMWriteScreen "FS", 8, 46													'this is a program for ABAWD's for SNAP is the only option for banked months
	EMWriteScreen member_number, 8, 9									'enters member number
	Call create_MAXIS_friendly_date(appointment_date, 0, 8, 65)			'enters the E & T referral date
	If manual_referral = "Banked months" then
		EMWriteScreen "Banked ABAWD month referral, initial month", 17, 6	'DHS wants these referrals marked, this marks them
	ELSEIF manual_referral = "Student" then
		EMWriteScreen "Student", 17, 6
	ELSEIF manual_referral = "Working with CBO" then
		EMWriteScreen "Working with Community Based Organization", 17, 6
	ELSEIF manual_referral = "Other manual referral" then
		EMWriteScreen other_referral_notes, 17, 6
	END IF
	EMWriteScreen "x", 8, 53																				'selects the ES provider
	transmit																												'navigates to the ES provider selection screen
		If worker_county_code = "x127" then				'HENNEPIN CO specific info'
			EMWriteScreen "x", 5, 9									'selects the 1st option'
			transmit																'transmits back to the main WF1M
			EMWriteScreen appointment_date & ", " & appointment_time_prefix_editbox & ":" & appointment_time_post_editbox & " " & AM_PM & ", " & SNAPET_name, 18, 6		'enters the location, date and time for Hennepin Co ES providers (per request)'
			PF3																			'saves referral
			EMWriteScreen "Y", 11, 64								'Y to confirm save
			transmit																'confirms saving the referral
			script_end_procedure("Your orientation letter, manual referral, and a 30 day TIKL has been made. Navigate to SPEC/WCOM if you want to review the notice sent to the client." & _
			vbNewLine & vbNewLine & "Make sure that you have sent the form ""ABAWD FS RULES"" to the client.")
		Else
			script_end_procedure("Please select your agency's ES provider, and PF3 to save your referral.")		'if agency is not Hennepin, then user is asked to select the ES provider and save'
		END IF
END IF

If worker_county_code = "x127" then			'specific closing message to Hennepin County message
	script_end_procedure("Your orientation letter and case note have been created. Navigate to SPEC/WCOM if you want to review the notice sent to the client." & _
	vbNewLine & vbNewLine & "Make sure that you have made your E & T referral, and that you have sent the form ""ABAWD FS RULES"" to the client.")
ELSE
	script_end_procedure("If you haven't made the E & T referral, please do so now. Your orientation letter and case note have been created. Navigate to SPEC/WCOM if you want to review the notice sent to the client.")
END IF
