'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - SNAP E AND T LETTER.vbs"
start_time = timer

'Option Explicit

'DIM beta_agency
'DIM FuncLib_URL, req, fso

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

'Creating a blank array to start our process. This will allow for validating whether-or-not the office was assigned later on, because it'll always be an array and not a variable.
county_FSET_offices = array("")

'Array listed above Dialog as below the dialog, the droplist appeared blank
'Creates an array of county FSET offices, which can be dynamically called in scripts which need it (SNAP ET LETTER for instance)
'Certain counties are commented out as they did not submit information about their E & T site, but can be easily rendered if they provide them 

'IF worker_county_code = "x101" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x102" THEN county_FSET_offices = array("Minnesota WorkForce Center Blaine")
IF worker_county_code = "x103" THEN county_FSET_offices = array("Rural MN CEP Detroit Lakes")
IF worker_county_code = "x104" THEN county_FSET_offices = array("Select one", "RMCEP", "MCT", "Leach Lake New", "Red Lake Oshkiimaajitahdah")
'IF worker_county_code = "x105" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x106" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x107" THEN county_FSET_offices = array("Blue Earth County Employment Services")
IF worker_county_code = "x108" THEN county_FSET_offices = array("Minnesota Valley Action Council New Ulm")
IF worker_county_code = "x109" THEN county_FSET_offices = array("Carlton County Human Services")
'IF worker_county_code = "x110" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x111" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x112" THEN county_FSET_offices = array("Chippewa County Workforce Center")
'IF worker_county_code = "x113" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x114" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x115" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x116" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x117" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x118" THEN county_FSET_offices = array("Rural MN CEP Brainerd")
IF worker_county_code = "x119" THEN county_FSET_offices = array("Select one", "Northern Service Center", "Burnsville Workforce Center")
IF worker_county_code = "x120" THEN county_FSET_offices = array("Workforce Development Inc. (Kasson)")
IF worker_county_code = "x121" THEN county_FSET_offices = array("Alexandria Workforce Center")
IF worker_county_code = "x122" THEN county_FSET_offices = array("Fairmont Workforce Center Fairbault County")
IF worker_county_code = "x123" THEN county_FSET_offices = array("Workforce Development Office")
'IF worker_county_code = "x124" THEN county_FSET_offices = array("Select one", 
IF worker_county_code = "x125" THEN county_FSET_offices = array("Workforce Development Inc. (Redwing)")
IF worker_county_code = "x126" THEN county_FSET_offices = array("Grant County Social Services")
IF worker_county_code = "x127" THEN county_FSET_offices = array("Select one", "Century Plaza", "Sabathani Community Center")
IF worker_county_code = "x128" THEN county_FSET_offices = array("Workforce Development Inc.")
'IF worker_county_code = "x129" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x130" THEN county_FSET_offices = array("Cambridge MN Workforce Center")
IF worker_county_code = "x131" THEN county_FSET_offices = array("AEOA – GR Workforce Center")
'IF worker_county_code = "x132" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x133" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x134" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x135" THEN county_FSET_offices = array("Kittson County Social Services")
'IF worker_county_code = "x136" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x137" THEN county_FSET_offices = array("Lace qui Parle Co. Family Services")
IF worker_county_code = "x138" THEN county_FSET_offices = array("AEOA")
IF worker_county_code = "x139" THEN county_FSET_offices = array("Rural MN CEP Lake of the Woods")
IF worker_county_code = "x140" THEN county_FSET_offices = array("MVAC")
IF worker_county_code = "x141" THEN county_FSET_offices = array("Marshall WorkForce Center")
IF worker_county_code = "x142" THEN county_FSET_offices = array("Marshall WorkForce Center")
IF worker_county_code = "x143" THEN county_FSET_offices = array("Mahnomen County Human Services")
IF worker_county_code = "x144" THEN county_FSET_offices = array("Marshall County Social Services")
IF worker_county_code = "x145" THEN county_FSET_offices = array("Fairmont Workforce Center Martin County")
IF worker_county_code = "x146" THEN county_FSET_offices = array("Central MN Jobs and Training Services Hutchinson")
IF worker_county_code = "x147" THEN county_FSET_offices = array("Central MN Jobs and Training Services Litchfield")
'IF worker_county_code = "x148" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x149" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x150" THEN county_FSET_offices = array("Workforce Development Inc. (Austin)")
IF worker_county_code = "x151" THEN county_FSET_offices = array("Marshall WorkForce Center")
'IF worker_county_code = "x152" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x153" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x154" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x155" THEN county_FSET_offices = array("Olmstead County Family Support & Assistance")
IF worker_county_code = "x156" THEN county_FSET_offices = array("Rural MN CEP Fergus Falls")
IF worker_county_code = "x157" THEN county_FSET_offices = array("Minnesota WorkForce Center: Theif River Falls")
IF worker_county_code = "x158" THEN county_FSET_offices = array("Select one", "Pine County Health & Human Services", "Pine Technical & Community College E&T Center")
IF worker_county_code = "x159" THEN county_FSET_offices = array("Southwest MN Private Industry Council Inc. Pipestone")
IF worker_county_code = "x160" THEN county_FSET_offices = array("Select one", "Polk County Social Services: Crookston", "Polk County Social Services: East Grand Forks", "Polk County Social Services: Fosston")
IF worker_county_code = "x161" THEN county_FSET_offices = array("Minnesota Workforce Center Alexandria")
'IF worker_county_code = "x162" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x163" THEN county_FSET_offices = array("Minnesota Workforce Center: Red Lake")
IF worker_county_code = "x164" THEN county_FSET_offices = array("Southwest Health & Human Services")
IF worker_county_code = "x165" THEN county_FSET_offices = array("Central MN Jobs and Training Services Olivia")
'IF worker_county_code = "x166" THEN county_FSET_offices = array("Select one", 
IF worker_county_code = "x167" THEN county_FSET_offices = array("Southwest MN Private Industry Council Inc. Luverne") 
IF worker_county_code = "x168" THEN county_FSET_offices = array("Roseau County Social Services")
IF worker_county_code = "x169" THEN county_FSET_offices = array("Select one", "MN Workforce Center: Duluth", "Minnesota WorkForce Center: Virginia", "Minnesota Workforce Center: Hibbing")
'IF worker_county_code = "x170" THEN county_FSET_offices = array("Select one", 
IF worker_county_code = "x171" THEN county_FSET_offices = array("Central MN Jobs and Training Services Monticello")
'IF worker_county_code = "x172" THEN county_FSET_offices = array("Select one",
'IF worker_county_code = "x173" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x174" THEN county_FSET_offices = array("Steele County Employment Services")
IF worker_county_code = "x175" THEN county_FSET_offices = array("Stevens County Human Services")
IF worker_county_code = "x176" THEN county_FSET_offices = array("SW MN Private Industry Council")
IF worker_county_code = "x177" THEN county_FSET_offices = array("Select one", "Todd County Health & Human Services: Long Prairie", "Todd County Health & Human Services: Staples")
IF worker_county_code = "x178" THEN county_FSET_offices = array("Rural MN CEP Wadena")
IF worker_county_code = "x179" THEN county_FSET_offices = array("Workforce Development Inc.")
IF worker_county_code = "x180" THEN county_FSET_offices = array("Rural MN CEP/MN workforce Center")
IF worker_county_code = "x181" THEN county_FSET_offices = array("Minnesota Valley Action Council Waseca")
IF worker_county_code = "x182" THEN county_FSET_offices = array("Select one", "Washington County Community Services: Stillwater", "Washington County Community Services: Forest Lake", "Washington County Community Services: Cottage Grove", "Washington County Community Services: Woodbury")
'IF worker_county_code = "x183" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x184" THEN county_FSET_offices = array("Wilkin County Family Services")
'IF worker_county_code = "x185" THEN county_FSET_offices = array("Select one",
IF worker_county_code = "x186" THEN county_FSET_offices = array("Central MN Jobs and Training Services Monticello")
IF worker_county_code = "x187" THEN county_FSET_offices = array("Yellow Medicine County Family Services")

'If the array isn't blank, then create a new array called FSET_list containing these items as a droplist. This will be used by the dialog.
IF county_FSET_offices(0) <> "" THEN call convert_array_to_droplist_items (county_FSET_offices, FSET_list)

If worker_county_code = "x127" THEN 
	SNAPET_contact = "the EZ Info Line"
	SNAPET_phone = "612-596-1300"
END IF

'DIALOGS----------------------------------------------------------------------------------------------------
' FSET_list is a variable not a standard drop down list.  When you copy into dialog editor, it will not work
' This dialog is for counties that HAVE provided FSET office addresses
BeginDialog SNAPET_automated_adress_dialog, 0, 0, 306, 110, "SNAP E&T Appointment Letter"
  EditBox 70, 5, 55, 15, case_number
  EditBox 205, 5, 20, 15, member_number
  EditBox 70, 25, 55, 15, appointment_date
  EditBox 195, 25, 20, 15, appointment_time_prefix_editbox
  EditBox 215, 25, 20, 15, appointment_time_post_editbox
  DropListBox 235, 25, 60, 15, "Select one..."+chr(9)+"AM"+chr(9)+"PM", AM_PM
  DropListBox 115, 50, 180, 15, FSET_list, interview_location
  EditBox 60, 70, 65, 15, SNAPET_contact
  EditBox 185, 70, 65, 15, SNAPET_phone
  EditBox 120, 90, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 190, 90, 50, 15
    CancelButton 245, 90, 50, 15
  Text 5, 50, 105, 10, "Location (select from dropdown)"
  Text 60, 95, 60, 10, "Worker Signature:"
  Text 10, 10, 50, 10, "Case Number:"
  Text 130, 10, 70, 10, "HH Member Number:"
  Text 130, 30, 60, 10, "Appointment Time:"
  Text 10, 75, 50, 10, "Contact name: "
  Text 5, 30, 60, 10, "Appointment Date:"
  Text 135, 75, 50, 10, "Contact phone:"
EndDialog



'This dialog is for counties that have not provided FSET office address(s)
BeginDialog SNAPET_manual_address_dialog, 0, 0, 301, 150, "SNAP E&T Appointment Letter"
  EditBox 70, 5, 55, 15, case_number
  EditBox 215, 5, 20, 15, member_number
  EditBox 70, 25, 55, 15, appointment_date
  EditBox 195, 25, 20, 15, appointment_time_prefix_editbox
  EditBox 215, 25, 20, 15, appointment_time_post_editbox
  DropListBox 240, 25, 55, 15, "Select one.."+chr(9)+"AM"+chr(9)+"PM", AM_PM
  EditBox 65, 45, 190, 15, SNAPET_name
  EditBox 65, 65, 190, 15, SNAPET_address_01
  EditBox 65, 85, 95, 15, SNAPET_city
  EditBox 165, 85, 40, 15, SNAPET_ST
  EditBox 210, 85, 45, 15, SNAPET_zip
  EditBox 65, 105, 65, 15, SNAPET_contact
  EditBox 185, 105, 70, 15, SNAPET_phone
  EditBox 120, 125, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 190, 125, 50, 15
    CancelButton 245, 125, 50, 15
  Text 5, 30, 60, 10, "Appointment Date:"
  Text 130, 30, 60, 15, "Appointment Time:"
  Text 5, 50, 55, 10, "Provider Name:"
  Text 5, 70, 55, 10, "Address line 1:"
  Text 10, 110, 55, 10, "Contact Name:"
  Text 135, 110, 50, 10, "Contact Phone:"
  Text 60, 130, 60, 10, "Worker Signature:"
  Text 10, 10, 50, 10, "Case Number:"
  Text 10, 90, 55, 10, "City/State/Zip:"
  Text 130, 10, 70, 10, "HH Member Number:"
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
											'Counties listed here (starting iwth x101 and ending with x185 did not provide E & T office information, hence will need to use the dialog requiring them to enter in their own address and contact information)
											If worker_county_code = "x101" OR _  	
											  worker_county_code = "x105" OR _
											  worker_county_code = "x106" OR _
											  worker_county_code = "x110" OR _
											  worker_county_code = "x111" OR _
											  worker_county_code = "x113" OR _
											  worker_county_code = "x114" OR _
											  worker_county_code = "x115" OR _
											  worker_county_code = "x116" OR _
											  worker_county_code = "x117" OR _
											  worker_county_code = "x124" OR _
											  worker_county_code = "x129" OR _
											  worker_county_code = "x132" OR _
											  worker_county_code = "x133" OR _
											  worker_county_code = "x134" OR _
											  worker_county_code = "x136" OR _
											  worker_county_code = "x148" OR _
											  worker_county_code = "x149" OR _
											  worker_county_code = "x152" OR _
											  worker_county_code = "x153" OR _
											  worker_county_code = "x154" OR _
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

'County FSET address information which will autofill when option is chosen from county_office_list----------------------------------------------------------------------------------------------------

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

'CO #27 HENNEPIN COUNTY addresses
IF interview_location = "Century Plaza" THEN 
	SNAPET_name = "Century Plaza"
	SNAPET_address_01 = "330 South 12th Street #3650"
	SNAPET_city = "Minneapolis"
	SNAPET_ST = "MN"
	SNAPET_zip = "55404"
ElseIf interview_location = "Sabathani Community Center" THEN 
	SNAPET_name = "Sabathani Community Center"
	SNAPET_address_01 = "310 East 38th Street #120"
	SNAPET_city = "Minneapolis"
	SNAPET_ST = "MN"
	SNAPET_zip = "55409"
END IF

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

'CO #50 MOWER COUNTY address
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
IF interview_location  = "Minnesota WorkForce Center: Theif River Falls" THEN
    SNAPET_name = "Minnesota WorkForce Center: Theif River Falls"
    SNAPET_address_01 = "1301 State Hwy 1"
    SNAPET_city = "Theif River Falls"
    SNAPET_ST = "MN"
    SNAPET_zip = "56701"
END IF

'CO #58 PINE COUNTY address
IF interview_location  = "Pine County Health & Human Services" THEN
    SNAPET_name = "Pine County Health & Human Services"
    SNAPET_address_01 = "130 Oriole St E Ste 1"
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
    SNAPET_city = "Theif River Falls"
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
EMWriteScreen SNAPET_city & ", " & SNAPET_ST & " " &  SNAPET_zip, 11, 28
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

MsgBox "If you haven't updated WREG with the FSET Orientation Date, please do so now.  Thank you!"

script_end_procedure("")