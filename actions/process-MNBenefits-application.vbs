'STATS GATHERING=============================================================================================================
name_of_script = "ACTIONS - PROCESS MNBENEFITS APPLICATION.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 30            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the anctual manualtime based on time study
'To do - update manual time calculations
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomonation appliable to your script.
'END OF stats block==========================================================================================================

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

'CHANGELOG BLOCK===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
CALL changelog_update("11/18/25", "Initial version.", "Mark Riegel, Hennepin County") 'REPLACE with release date and your name.

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DEFINING CONSTANTS, VARIABLES, ARRAYS, AND BUTTONS===========================================================================

'Buttons Defined
'--Navigation buttons
next_hh_memb_btn        = 201
previous_hh_memb_button = 202
next_hh_memb_btn        = 203
submit_hh_memb_button   = 204 


'--Other buttons
' instructions_btn
' file_selection_button
hh_memb_1_and_2_button    = 301
hh_memb_3_and_4_button    = 302
hh_memb_5_and_6_button    = 303
hh_memb_7_and_8_button    = 304
hh_memb_9_and_10_button   = 305
hh_memb_11_and_12_button  = 306



'Defining variables


'Dimming variables
Dim folderPath, application_ID, fso, folder, fileList, file, xml_file_path, script_testing

'Initialize variables
script_testing = true


'DEFINING FUNCTIONS===========================================================================

'Creating Household Member dialogs as functions to more easily loop through them 
Function household_members_dialog_1_2()
  hh_memb_dialog_count = 1
  BeginDialog Dialog1, 0, 0, 281, 345, "Verify MNBenefits XML Details - Household Members"
    Text 5, 5, 250, 20, "Please review and verify the household member details for each household member pulled from the XML file below. Make any updates as needed."
    GroupBox 10, 30, 175, 140, "Household Member Information"
    Text 15, 50, 40, 10, "First name:"
    EditBox 70, 45, 100, 15, householdMembers(MEMBER_FIRST_NAME, 0)
    Text 15, 65, 40, 10, "Last name:"
    EditBox 70, 60, 100, 15, householdMembers(MEMBER_LAST_NAME, 0)
    Text 15, 80, 30, 10, "Gender:"
    DropListBox 70, 75, 60, 10, "Male"+chr(9)+"Female"+chr(9)+"Other", householdMembers(MEMBER_GENDER, 0)
    Text 15, 95, 50, 10, "Marital status:"
    DropListBox 70, 90, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", householdMembers(MEMBER_MARITAL_STATUS, 0)
    Text 15, 110, 45, 10, "Date of birth:"
    EditBox 70, 105, 100, 15, householdMembers(MEMBER_DOB, 0)
    Text 15, 125, 20, 10, "SSN:"
    EditBox 70, 120, 100, 15, householdMembers(MEMBER_SSN, 0)
    Text 15, 140, 45, 10, "Citizenship:"
    DropListBox 70, 135, 30, 15, "Yes"+chr(9)+"No", householdMembers(MEMBER_CITIZENSHIP, 0)
    Text 15, 155, 45, 10, "Relationship:"
    DropListBox 70, 150, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", householdMembers(MEMBER_RELATIONSHIP, 0)
    If member_count > 1 Then
      GroupBox 10, 175, 175, 140, "Household Member Information"
      Text 15, 195, 40, 10, "First name:"
      EditBox 70, 190, 100, 15, householdMembers(MEMBER_FIRST_NAME, 1)
      Text 15, 210, 40, 10, "Last name:"
      EditBox 70, 205, 100, 15, householdMembers(MEMBER_LAST_NAME, 1)
      Text 15, 225, 30, 10, "Gender:"
      DropListBox 70, 220, 60, 10, "Male"+chr(9)+"Female"+chr(9)+"Other", householdMembers(MEMBER_GENDER, 1)
      Text 15, 240, 50, 10, "Marital status:"
      DropListBox 70, 235, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", householdMembers(MEMBER_MARITAL_STATUS, 1)
      Text 15, 255, 45, 10, "Date of birth:"
      EditBox 70, 250, 100, 15, householdMembers(MEMBER_DOB, 1)
      Text 15, 270, 20, 10, "SSN:"
      EditBox 70, 265, 100, 15, householdMembers(MEMBER_SSN, 1)
      Text 15, 285, 45, 10, "Citizenship:"
      DropListBox 70, 280, 30, 15, "Yes"+chr(9)+"No", householdMembers(MEMBER_CITIZENSHIP, 1)
      Text 15, 300, 45, 10, "Relationship:"
      DropListBox 70, 295, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", householdMembers(MEMBER_RELATIONSHIP, 1)
    End If
    ButtonGroup ButtonPressed
      If member_count = 1 or member_count = 2 Then
        PushButton 230, 325, 45, 15, "Submit", submit_hh_memb_button
      Else
        PushButton 230, 325, 45, 15, "Next", next_hh_memb_btn
      End If
      PushButton 5, 325, 45, 15, "Previous", previous_hh_memb_button
    GroupBox 195, 30, 70, 105, "Navigation"
    ButtonGroup ButtonPressed
    Call determine_hh_memb_buttons()
  EndDialog
End Function
Dim first_name_hh_memb_1, last_name_hh_memb_1, gender_hh_memb_1_list, marital_status_hh_memb_1, dob_hh_memb_1, SSN_hh_memb_1, citizenship_hh_memb_1_list, relationship_hh_memb_1_list, first_name_hh_memb_2, last_name_hh_memb_2, gender_hh_memb_2_list, marital_status_hh_memb_2, dob_hh_memb_2, SSN_hh_memb_2, citizenship_hh_memb_2_list, relationship_hh_memb_2_list

Function household_members_dialog_3_4()
  hh_memb_dialog_count = 2
  BeginDialog Dialog1, 0, 0, 281, 345, "Verify MNBenefits XML Details - Household Members"
    Text 5, 5, 250, 20, "Please review and verify the household member details for each household member pulled from the XML file below. Make any updates as needed."
    GroupBox 10, 30, 175, 140, "Household Member Information"
    Text 15, 50, 40, 10, "First name:"
    EditBox 70, 45, 100, 15, householdMembers(MEMBER_FIRST_NAME, 2)
    Text 15, 65, 40, 10, "Last name:"
    EditBox 70, 60, 100, 15, householdMembers(MEMBER_LAST_NAME, 2)
    Text 15, 80, 30, 10, "Gender:"
    DropListBox 70, 75, 60, 10, "Male"+chr(9)+"Female"+chr(9)+"Other", householdMembers(MEMBER_GENDER, 2)
    Text 15, 95, 50, 10, "Marital status:"
    DropListBox 70, 90, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", householdMembers(MEMBER_MARITAL_STATUS, 2)
    Text 15, 110, 45, 10, "Date of birth:"
    EditBox 70, 105, 100, 15, householdMembers(MEMBER_DOB, 2)
    Text 15, 125, 20, 10, "SSN:"
    EditBox 70, 120, 100, 15, householdMembers(MEMBER_SSN, 2)
    Text 15, 140, 45, 10, "Citizenship:"
    DropListBox 70, 135, 30, 15, "Yes"+chr(9)+"No", householdMembers(MEMBER_CITIZENSHIP, 2)
    Text 15, 155, 45, 10, "Relationship:"
    DropListBox 70, 150, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", householdMembers(MEMBER_RELATIONSHIP, 2)
    If member_count > 3 Then
      GroupBox 10, 175, 175, 140, "Household Member Information"
      Text 15, 195, 40, 10, "First name:"
      EditBox 70, 190, 100, 15, householdMembers(MEMBER_FIRST_NAME, 3)
      Text 15, 210, 40, 10, "Last name:"
      EditBox 70, 205, 100, 15, householdMembers(MEMBER_LAST_NAME, 3)
      Text 15, 225, 30, 10, "Gender:"
      DropListBox 70, 220, 60, 10, "Male"+chr(9)+"Female"+chr(9)+"Other", householdMembers(MEMBER_GENDER, 3)
      Text 15, 240, 50, 10, "Marital status:"
      DropListBox 70, 235, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", householdMembers(MEMBER_MARITAL_STATUS, 3)
      Text 15, 255, 45, 10, "Date of birth:"
      EditBox 70, 250, 100, 15, householdMembers(MEMBER_DOB, 3)
      Text 15, 270, 20, 10, "SSN:"
      EditBox 70, 265, 100, 15, householdMembers(MEMBER_SSN, 3)
      Text 15, 285, 45, 10, "Citizenship:"
      DropListBox 70, 280, 30, 15, "Yes"+chr(9)+"No", householdMembers(MEMBER_CITIZENSHIP, 3)
      Text 15, 300, 45, 10, "Relationship:"
      DropListBox 70, 295, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", householdMembers(MEMBER_RELATIONSHIP, 3)
    End If
    ButtonGroup ButtonPressed
      If member_count = 3 or member_count = 4 Then
        PushButton 230, 325, 45, 15, "Submit", submit_hh_memb_button
      Else
        PushButton 230, 325, 45, 15, "Next", next_hh_memb_btn
      End If
      PushButton 5, 325, 45, 15, "Previous", previous_hh_memb_button
    GroupBox 195, 30, 70, 105, "Navigation"
    ButtonGroup ButtonPressed
      Call determine_hh_memb_buttons()
  EndDialog
End Function
Dim first_name_hh_memb_3, last_name_hh_memb_3, gender_hh_memb_3_list, marital_status_hh_memb_3, dob_hh_memb_3, SSN_hh_memb_3, citizenship_hh_memb_3_list, relationship_hh_memb_3_list, first_name_hh_memb_4, last_name_hh_memb_4, gender_hh_memb_4_list, marital_status_hh_memb_4, dob_hh_memb_4, SSN_hh_memb_4, citizenship_hh_memb_4_list, relationship_hh_memb_4_list

Function household_members_dialog_5_6()
hh_memb_dialog_count = 3
  BeginDialog Dialog1, 0, 0, 281, 345, "Verify MNBenefits XML Details - Household Members"
    Text 5, 5, 250, 20, "Please review and verify the household member details for each household member pulled from the XML file below. Make any updates as needed."
    GroupBox 10, 30, 175, 140, "Household Member Information"
    Text 15, 50, 40, 10, "First name:"
    EditBox 70, 45, 100, 15, householdMembers(MEMBER_FIRST_NAME, 4)
    Text 15, 65, 40, 10, "Last name:"
    EditBox 70, 60, 100, 15, householdMembers(MEMBER_LAST_NAME, 4)
    Text 15, 80, 30, 10, "Gender:"
    DropListBox 70, 75, 60, 10, "Male"+chr(9)+"Female"+chr(9)+"Other", householdMembers(MEMBER_GENDER, 4)
    Text 15, 95, 50, 10, "Marital status:"
    DropListBox 70, 90, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", householdMembers(MEMBER_MARITAL_STATUS, 4)
    Text 15, 110, 45, 10, "Date of birth:"
    EditBox 70, 105, 100, 15, householdMembers(MEMBER_DOB, 4)
    Text 15, 125, 20, 10, "SSN:"
    EditBox 70, 120, 100, 15, householdMembers(MEMBER_SSN, 4)
    Text 15, 140, 45, 10, "Citizenship:"
    DropListBox 70, 135, 30, 15, "Yes"+chr(9)+"No", householdMembers(MEMBER_CITIZENSHIP, 4)
    Text 15, 155, 45, 10, "Relationship:"
    DropListBox 70, 150, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", householdMembers(MEMBER_RELATIONSHIP, 4)
    If member_count > 5 Then
      GroupBox 10, 175, 175, 140, "Household Member Information"
      Text 15, 195, 40, 10, "First name:"
      EditBox 70, 190, 100, 15, householdMembers(MEMBER_FIRST_NAME, 5)
      Text 15, 210, 40, 10, "Last name:"
      EditBox 70, 205, 100, 15, householdMembers(MEMBER_LAST_NAME, 5)
      Text 15, 225, 30, 10, "Gender:"
      DropListBox 70, 220, 60, 10, "Male"+chr(9)+"Female"+chr(9)+"Other", householdMembers(MEMBER_GENDER, 5)
      Text 15, 240, 50, 10, "Marital status:"
      DropListBox 70, 235, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", householdMembers(MEMBER_MARITAL_STATUS, 5)
      Text 15, 255, 45, 10, "Date of birth:"
      EditBox 70, 250, 100, 15, householdMembers(MEMBER_DOB, 5)
      Text 15, 270, 20, 10, "SSN:"
      EditBox 70, 265, 100, 15, householdMembers(MEMBER_SSN, 5)
      Text 15, 285, 45, 10, "Citizenship:"
      DropListBox 70, 280, 30, 15, "Yes"+chr(9)+"No", householdMembers(MEMBER_CITIZENSHIP, 5)
      Text 15, 300, 45, 10, "Relationship:"
      DropListBox 70, 295, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", householdMembers(MEMBER_RELATIONSHIP, 5)
    End If
    ButtonGroup ButtonPressed
      If member_count = 5 or member_count = 6 Then
        PushButton 230, 325, 45, 15, "Submit", submit_hh_memb_button
      Else
        PushButton 230, 325, 45, 15, "Next", next_hh_memb_btn
      End If
      PushButton 5, 325, 45, 15, "Previous", previous_hh_memb_button
    GroupBox 195, 30, 70, 105, "Navigation"
    ButtonGroup ButtonPressed
      Call determine_hh_memb_buttons()
  EndDialog
End Function
Dim first_name_hh_memb_5, last_name_hh_memb_5, gender_hh_memb_5_list, marital_status_hh_memb_5, dob_hh_memb_5, SSN_hh_memb_5, citizenship_hh_memb_5_list, relationship_hh_memb_5_list, first_name_hh_memb_6, last_name_hh_memb_6, gender_hh_memb_6_list, marital_status_hh_memb_6, dob_hh_memb_6, SSN_hh_memb_6, citizenship_hh_memb_6_list, relationship_hh_memb_6_list

Function household_members_dialog_7_8()   
  hh_memb_dialog_count = 4
  BeginDialog Dialog1, 0, 0, 281, 345, "Verify MNBenefits XML Details - Household Members"
    Text 5, 5, 250, 20, "Please review and verify the household member details for each household member pulled from the XML file below. Make any updates as needed."
    GroupBox 10, 30, 175, 140, "Household Member Information"
    Text 15, 50, 40, 10, "First name:"
    EditBox 70, 45, 100, 15, householdMembers(MEMBER_FIRST_NAME, 6)
    Text 15, 65, 40, 10, "Last name:"
    EditBox 70, 60, 100, 15, householdMembers(MEMBER_LAST_NAME, 6)
    Text 15, 80, 30, 10, "Gender:"
    DropListBox 70, 75, 60, 10, "Male"+chr(9)+"Female"+chr(9)+"Other", householdMembers(MEMBER_GENDER, 6)
    Text 15, 95, 50, 10, "Marital status:"
    DropListBox 70, 90, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", householdMembers(MEMBER_MARITAL_STATUS, 6)
    Text 15, 110, 45, 10, "Date of birth:"
    EditBox 70, 105, 100, 15, householdMembers(MEMBER_DOB, 6)
    Text 15, 125, 20, 10, "SSN:"
    EditBox 70, 120, 100, 15, householdMembers(MEMBER_SSN, 6)
    Text 15, 140, 45, 10, "Citizenship:"
    DropListBox 70, 135, 30, 15, "Yes"+chr(9)+"No", householdMembers(MEMBER_CITIZENSHIP, 6)
    Text 15, 155, 45, 10, "Relationship:"
    DropListBox 70, 150, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", householdMembers(MEMBER_RELATIONSHIP, 6)
    If member_count > 7 Then
      GroupBox 10, 175, 175, 140, "Household Member Information"
      Text 15, 195, 40, 10, "First name:"
      EditBox 70, 190, 100, 15, householdMembers(MEMBER_FIRST_NAME, 7)
      Text 15, 210, 40, 10, "Last name:"
      EditBox 70, 205, 100, 15, householdMembers(MEMBER_LAST_NAME, 7)
      Text 15, 225, 30, 10, "Gender:"
      DropListBox 70, 220, 60, 10, "Male"+chr(9)+"Female"+chr(9)+"Other", householdMembers(MEMBER_GENDER, 7)
      Text 15, 240, 50, 10, "Marital status:"
      DropListBox 70, 235, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", householdMembers(MEMBER_MARITAL_STATUS, 7)
      Text 15, 255, 45, 10, "Date of birth:"
      EditBox 70, 250, 100, 15, householdMembers(MEMBER_DOB, 7)
      Text 15, 270, 20, 10, "SSN:"
      EditBox 70, 265, 100, 15, householdMembers(MEMBER_SSN, 7)
      Text 15, 285, 45, 10, "Citizenship:"
      DropListBox 70, 280, 30, 15, "Yes"+chr(9)+"No", householdMembers(MEMBER_CITIZENSHIP, 7)
      Text 15, 300, 45, 10, "Relationship:"
      DropListBox 70, 295, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", householdMembers(MEMBER_RELATIONSHIP, 7)
    End If
    ButtonGroup ButtonPressed
      If member_count = 7 or member_count = 8 Then
        PushButton 230, 325, 45, 15, "Submit", submit_hh_memb_button
      Else
        PushButton 230, 325, 45, 15, "Next", next_hh_memb_btn
      End If
      PushButton 5, 325, 45, 15, "Previous", previous_hh_memb_button
    GroupBox 195, 30, 70, 105, "Navigation"
    ButtonGroup ButtonPressed
      Call determine_hh_memb_buttons()
  EndDialog
End Function
Dim first_name_hh_memb_7, last_name_hh_memb_7, gender_hh_memb_7_list, marital_status_hh_memb_7, dob_hh_memb_7, SSN_hh_memb_7, citizenship_hh_memb_7_list, relationship_hh_memb_7_list, first_name_hh_memb_8, last_name_hh_memb_8, gender_hh_memb_8_list, marital_status_hh_memb_8, dob_hh_memb_8, SSN_hh_memb_8, citizenship_hh_memb_8_list, relationship_hh_memb_8_list

Function household_members_dialog_9_10()
  hh_memb_dialog_count = 5
  BeginDialog Dialog1, 0, 0, 281, 345, "Verify MNBenefits XML Details - Household Members"
    Text 5, 5, 250, 20, "Please review and verify the household member details for each household member pulled from the XML file below. Make any updates as needed."
    GroupBox 10, 30, 175, 140, "Household Member Information"
    Text 15, 50, 40, 10, "First name:"
    EditBox 70, 45, 100, 15, householdMembers(MEMBER_FIRST_NAME, 8)
    Text 15, 65, 40, 10, "Last name:"
    EditBox 70, 60, 100, 15, householdMembers(MEMBER_LAST_NAME, 8)
    Text 15, 80, 30, 10, "Gender:"
    DropListBox 70, 75, 60, 10, "Male"+chr(9)+"Female"+chr(9)+"Other", householdMembers(MEMBER_GENDER, 8)
    Text 15, 95, 50, 10, "Marital status:"
    DropListBox 70, 90, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", householdMembers(MEMBER_MARITAL_STATUS, 8)
    Text 15, 110, 45, 10, "Date of birth:"
    EditBox 70, 105, 100, 15, householdMembers(MEMBER_DOB, 8)
    Text 15, 125, 20, 10, "SSN:"
    EditBox 70, 120, 100, 15, householdMembers(MEMBER_SSN, 8)
    Text 15, 140, 45, 10, "Citizenship:"
    DropListBox 70, 135, 30, 15, "Yes"+chr(9)+"No", householdMembers(MEMBER_CITIZENSHIP, 8)
    Text 15, 155, 45, 10, "Relationship:"
    DropListBox 70, 150, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", householdMembers(MEMBER_RELATIONSHIP, 8)
    If member_count > 9 Then
      GroupBox 10, 175, 175, 140, "Household Member Information"
      Text 15, 195, 40, 10, "First name:"
      EditBox 70, 190, 100, 15, householdMembers(MEMBER_FIRST_NAME, 9)
      Text 15, 210, 40, 10, "Last name:"
      EditBox 70, 205, 100, 15, householdMembers(MEMBER_LAST_NAME, 9)
      Text 15, 225, 30, 10, "Gender:"
      DropListBox 70, 220, 60, 10, "Male"+chr(9)+"Female"+chr(9)+"Other", householdMembers(MEMBER_GENDER, 9)
      Text 15, 240, 50, 10, "Marital status:"
      DropListBox 70, 235, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", householdMembers(MEMBER_MARITAL_STATUS, 9)
      Text 15, 255, 45, 10, "Date of birth:"
      EditBox 70, 250, 100, 15, householdMembers(MEMBER_DOB, 9)
      Text 15, 270, 20, 10, "SSN:"
      EditBox 70, 265, 100, 15, householdMembers(MEMBER_SSN, 9)
      Text 15, 285, 45, 10, "Citizenship:"
      DropListBox 70, 280, 30, 15, "Yes"+chr(9)+"No", householdMembers(MEMBER_CITIZENSHIP, 9)
      Text 15, 300, 45, 10, "Relationship:"
      DropListBox 70, 295, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", householdMembers(MEMBER_RELATIONSHIP, 9)
    End If
    ButtonGroup ButtonPressed
      If member_count = 9 or member_count = 10 Then
        PushButton 230, 325, 45, 15, "Submit", submit_hh_memb_button
      Else
        PushButton 230, 325, 45, 15, "Next", next_hh_memb_btn
      End If
      PushButton 5, 325, 45, 15, "Previous", previous_hh_memb_button
    GroupBox 195, 30, 70, 105, "Navigation"
    ButtonGroup ButtonPressed
      Call determine_hh_memb_buttons()
  EndDialog
End Function
Dim first_name_hh_memb_9, last_name_hh_memb_9, gender_hh_memb_9_list, marital_status_hh_memb_9, dob_hh_memb_9, SSN_hh_memb_9, citizenship_hh_memb_9_list, relationship_hh_memb_9_list, first_name_hh_memb_10, last_name_hh_memb_10, gender_hh_memb_10_list, marital_status_hh_memb_10, dob_hh_memb_10, SSN_hh_memb_10, citizenship_hh_memb_10_list, relationship_hh_memb_10_list

Function household_members_dialog_11_12()
  hh_memb_dialog_count = 6
  BeginDialog Dialog1, 0, 0, 281, 345, "Verify MNBenefits XML Details - Household Members"
    Text 5, 5, 250, 20, "Please review and verify the household member details for each household member pulled from the XML file below. Make any updates as needed."
    GroupBox 10, 30, 175, 140, "Household Member Information"
    Text 15, 50, 40, 10, "First name:"
    EditBox 70, 45, 100, 15, householdMembers(MEMBER_FIRST_NAME, 10)
    Text 15, 65, 40, 10, "Last name:"
    EditBox 70, 60, 100, 15, householdMembers(MEMBER_LAST_NAME, 10)
    Text 15, 80, 30, 10, "Gender:"
    DropListBox 70, 75, 60, 10, "Male"+chr(9)+"Female"+chr(9)+"Other", householdMembers(MEMBER_GENDER, 10)
    Text 15, 95, 50, 10, "Marital status:"
    DropListBox 70, 90, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", householdMembers(MEMBER_MARITAL_STATUS, 10)
    Text 15, 110, 45, 10, "Date of birth:"
    EditBox 70, 105, 100, 15, householdMembers(MEMBER_DOB, 10)
    Text 15, 125, 20, 10, "SSN:"
    EditBox 70, 120, 100, 15, householdMembers(MEMBER_SSN, 10)
    Text 15, 140, 45, 10, "Citizenship:"
    DropListBox 70, 135, 30, 15, "Yes"+chr(9)+"No", householdMembers(MEMBER_CITIZENSHIP, 10)
    Text 15, 155, 45, 10, "Relationship:"
    DropListBox 70, 150, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", householdMembers(MEMBER_RELATIONSHIP, 10)
    If member_count > 11 Then
      GroupBox 10, 175, 175, 140, "Household Member Information"
      Text 15, 195, 40, 10, "First name:"
      EditBox 70, 190, 100, 15, householdMembers(MEMBER_FIRST_NAME, 11)
      Text 15, 210, 40, 10, "Last name:"
      EditBox 70, 205, 100, 15, householdMembers(MEMBER_LAST_NAME, 11)
      Text 15, 225, 30, 10, "Gender:"
      DropListBox 70, 220, 60, 10, "Male"+chr(9)+"Female"+chr(9)+"Other", householdMembers(MEMBER_GENDER, 11)
      Text 15, 240, 50, 10, "Marital status:"
      DropListBox 70, 235, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", householdMembers(MEMBER_MARITAL_STATUS, 11)
      Text 15, 255, 45, 10, "Date of birth:"
      EditBox 70, 250, 100, 15, householdMembers(MEMBER_DOB, 11)
      Text 15, 270, 20, 10, "SSN:"
      EditBox 70, 265, 100, 15, householdMembers(MEMBER_SSN, 11)
      Text 15, 285, 45, 10, "Citizenship:"
      DropListBox 70, 280, 30, 15, "Yes"+chr(9)+"No", householdMembers(MEMBER_CITIZENSHIP, 11)
      Text 15, 300, 45, 10, "Relationship:"
      DropListBox 70, 295, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", householdMembers(MEMBER_RELATIONSHIP, 11)
    End If
    ButtonGroup ButtonPressed
      PushButton 230, 325, 45, 15, "Submit", submit_hh_memb_button
      PushButton 5, 325, 45, 15, "Previous", previous_hh_memb_button
    GroupBox 195, 30, 70, 105, "Navigation"
    ButtonGroup ButtonPressed
      Call determine_hh_memb_buttons()
  EndDialog
End Function
Dim first_name_hh_memb_11, last_name_hh_memb_11, gender_hh_memb_11_list, marital_status_hh_memb_11, dob_hh_memb_11, SSN_hh_memb_11, citizenship_hh_memb_11_list, relationship_hh_memb_11_list, first_name_hh_memb_12, last_name_hh_memb_12, gender_hh_memb_12_list, marital_status_hh_memb_12, dob_hh_memb_12, SSN_hh_memb_12, citizenship_hh_memb_12_list, relationship_hh_memb_12_list

Function determine_hh_memb_buttons()
  hh_memb_1_and_2_button_text = "HH Memb 1 - 2" 
  hh_memb_3_and_4_button_text = "HH Memb 3 - 4"
  If member_count = 3 Then hh_memb_3_and_4_button_text = "HH Memb 3"
  hh_memb_5_and_6_button_text = "HH Memb 5 - 6"
  If member_count = 5 Then hh_memb_5_and_6_button_text = "HH Memb 5"
  hh_memb_7_and_8_button_text = "HH Memb 7 - 8"
  If member_count = 7 Then hh_memb_7_and_8_button_text = "HH Memb 7"
  hh_memb_9_and_10_button_text = "HH Memb 9 - 10"
  If member_count = 9 Then hh_memb_9_and_10_button_text = "HH Memb 9"
  hh_memb_11_and_12_button_text = "HH Memb 11 - 12"
  If member_count = 11 Then hh_memb_11_and_12_button_text = "HH Memb 11"

  If member_count > 2 Then
    GroupBox 195, 30, 70, 105, "Navigation"
    If dialog_count = 1 Then
      Text 205, 45, 55, 10, hh_memb_1_and_2_button_text
    Else
      PushButton 200, 40, 60, 15, hh_memb_1_and_2_button_text, hh_memb_1_and_2_button
    End If
    If dialog_count = 2 Then
      Text 205, 55, 55, 10, hh_memb_3_and_4_button_text
    Else
      PushButton 200, 55, 60, 15, hh_memb_3_and_4_button_text, hh_memb_3_and_4_button
    End If
  End If
  
  If member_count > 4 Then
    If dialog_count = 3 Then
      Text 205, 70, 55, 10, hh_memb_5_and_6_button_text
    Else
      PushButton 200, 70, 60, 15, hh_memb_5_and_6_button_text, hh_memb_5_and_6_button
    End If
  End If


  If member_count > 6 Then
    If dialog_count = 4 Then
      Text 205, 85, 55, 10, hh_memb_7_and_8_button_text
    Else
      PushButton 200, 85, 60, 15, hh_memb_7_and_8_button_text, hh_memb_7_and_8_button
    End If
  End If

  If member_count > 8 Then
    If dialog_count = 5 then
      Text 205, 100, 55, 10, hh_memb_9_and_10_button_text
    Else
      PushButton 200, 100, 60, 15, hh_memb_9_and_10_button_text, hh_memb_9_and_10_button
    End If
  End If

  If member_count > 10 Then
    If dialog_count = 6 Then
      Text 205, 115, 55, 10, hh_memb_11_and_12_button_text
    Else  
      PushButton 200, 115, 60, 15, hh_memb_11_and_12_button_text, hh_memb_11_and_12_button
    End If
  End If
End Function

function dialog_specific_error_handling()	'Error handling for main dialog of forms
  'Error handling will display at the point of each dialog and will not let the user continue unless the applicable errors are resolved. Had to list all buttons including -1 so ensure the error reporting is called and hit when the script is run.
	If dialog_count = 10 Then
    If ButtonPressed = -1 Then err_msg = err_msg & vbNewLine & "* You must press either the 'Save Info and Return' or the 'Return WITHOUT Saving Assessor Info' buttons."
  End If

  If dialog_count = 11 Then
    If ButtonPressed = -1 Then err_msg = err_msg & vbNewLine & "* You must press either the 'Save Info and Return' or the 'Return WITHOUT Saving Assessor Info' buttons."
  End If

	If ButtonPressed = next_hh_memb_btn or ButtonPressed = previous_btn Or ButtonPressed = -1 OR ButtonPressed = section_a_assessor_return_btn OR ButtonPressed = section_e_assessor_return_btn OR ButtonPressed = section_a_add_assessor_btn Or ButtonPressed = section_e_add_assessor_btn Then
    'section_a_contact_info()
    If dialog_count = 1 then 
      If form_status_dropdown = "Select one:" Then err_msg = err_msg & vbNewLine & "* You must select either 'Complete' or 'Incomplete' from the Form Status dropdown list."
      If trim(section_a_date_form_sent) = "" OR IsDate(section_a_date_form_sent) = FALSE Then err_msg = err_msg & vbNewLine & "* You must fill out the Date Sent to Worker field in the format MM/DD/YYYY."
      If trim(section_a_assessor) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Assessor field." 
      If trim(section_a_lead_agency) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Lead Agency field." 
      If trim(section_a_phone_number) <> "" Then
        If len(trim(section_a_phone_number)) <> 12 OR mid(section_a_phone_number, 4, 1) <> "-" OR mid(section_a_phone_number, 8, 1) <> "-" Then err_msg = err_msg & vbCr & "* You must fill out the Phone Number field in the format ###-###-####."
      End If
      If trim(section_a_state) <> "" Then 
        If len(trim(section_a_state)) <> 2 Then err_msg = err_msg & vbNewLine & "* You must fill out the State field in the two character format, ex. MN."
      End If  
      If trim(section_a_zip_code) <> "" Then
        If len(trim(section_a_zip_code)) <> 5 Then err_msg = err_msg & vbNewLine & "* You must fill out the Zip Code field in a five number format." 
      End If
      If hh_memb = "Select One:" Then err_msg = err_msg & vbNewLine & "* You must select the Household Member from the dropdown." 
    End If
  End If
	If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
end function

Function dialog_selection(dialog_selected) 	
  'Selects the correct dialog based
  If dialog_selected = 1 then call household_members_dialog_1_2()
  If dialog_selected = 2 then call household_members_dialog_3_4()
  If dialog_selected = 3 then call household_members_dialog_5_6()
  If dialog_selected = 4 then call household_members_dialog_7_8()
  If dialog_selected = 5 then call household_members_dialog_9_10()
  If dialog_selected = 6 then call household_members_dialog_11_12()
End Function

function button_movement() 	'Dialog movement handling for buttons displayed on the individual form dialogs.
  'To do - add handling for future dialogs
  If err_msg = "" AND (ButtonPressed = next_hh_memb_btn or ButtonPressed = -1) Then dialog_count = dialog_count + 1
	If err_msg = "" AND ButtonPressed = previous_hh_memb_button Then dialog_count = dialog_count - 1

  If err_msg = "" AND ButtonPressed = submit_hh_memb_button then 
    hh_memb_dialog_loop = "Completed"
    msgbox "Completed triggered"
  End If

  If err_msg = "" AND ButtonPressed = hh_memb_1_and_2_button then dialog_count = 1
  If err_msg = "" AND ButtonPressed = hh_memb_3_and_4_button then dialog_count = 2
  If err_msg = "" AND ButtonPressed = hh_memb_5_and_6_button then dialog_count = 3
  If err_msg = "" AND ButtonPressed = hh_memb_7_and_8_button then dialog_count = 4
  If err_msg = "" AND ButtonPressed = hh_memb_9_and_10_button then dialog_count = 5
  If err_msg = "" AND ButtonPressed = hh_memb_11_and_12_button then dialog_count = 6
end function
Dim hh_memb_dialog_loop

function determine_member_dialogs_display()
  member_dialogs_to_display = "*"
  For member = 1 to member_count 
    member_dialogs_to_display = member_dialogs_to_display & member & "*"
  Next
End function
Dim member_dialogs_to_display


'THE SCRIPT=================================================================================================================
EMConnect "" 'Connects to BlueZone

'Initial Dialog - Instructions
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 281, 220, "Process MNBenefits Application"
  Text 10, 5, 245, 20, "Script Purpose: This script performs a PERS search, APPLs the case using the MNBenefits XML file details, and then moves the case to PND2 status."
  ButtonGroup ButtonPressed
    PushButton 10, 30, 65, 15, "Script Instructions", instructions_btn
  GroupBox 5, 50, 270, 100, "Choose one option"
  CheckBox 15, 65, 250, 10, "Enter 10-digit application ID for the XML file. Then press the Search button.", application_ID_checkbox
  EditBox 20, 80, 55, 15, application_ID
  ButtonGroup ButtonPressed
    PushButton 80, 80, 40, 15, "Search", search_button
  CheckBox 15, 100, 215, 10, "Press button below to locate XML file using Windows Explorer", manual_file_select_checkbox
  ButtonGroup ButtonPressed
    PushButton 25, 115, 85, 15, "Open Windows Explorer", file_selection_button
  CheckBox 15, 135, 205, 10, "Enter application details manually to perform PERS search", enter_app_manually_checkbox
  Text 10, 160, 95, 10, "XML File Path (if applicable):"
  EditBox 10, 170, 265, 15, XML_file_path
  ButtonGroup ButtonPressed
    OkButton 185, 200, 45, 15
    CancelButton 230, 200, 45, 15
EndDialog

DO
	DO
		err_msg = ""					'establishing value of variable, this is necessary for the Do...LOOP
		dialog Dialog1				'main dialog
		cancel_without_confirmation
    If ButtonPressed = file_selection_button then 
      call file_selection_system_dialog(XML_file_path, ".xml")
      err_msg = "LOOP"
    End If
    If ButtonPressed = instructions_btn Then
      'To do - update with instructions 
      Call open_URL_in_browser("https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/") 
      err_msg = "LOOP"
    End IF 

    If trim(application_ID) <> "" and len(application_ID) = 10 and IsNumeric(application_ID) then
      If ButtonPressed = search_button Then

        If script_testing = false Then

          folderPath = "T:\Eligibility Support\EA_ADAD\EA_ADAD_Common\CASE ASSIGNMENT\MNB_XML_files\"

          Set fso = CreateObject("Scripting.FileSystemObject")
          Set folder = fso.GetFolder(folderPath)
          XML_file_found = False

          For Each file In folder.Files
            If InStr(1, file.Name, "_" & application_ID & "_", vbTextCompare) > 0 Then
              msgbox "Found: " & file.Path
              XML_file_path = file.Path
              XML_file_found = True
              err_msg = "LOOP"
              Exit For
            End If
          Next
          If XML_file_found = False Then
            err_msg = err_msg & vbCr & "* The script was unable to locate a MNBenefits XML file with the application ID you provided. You must click the 'Select File' button and select the XML file or manually enter the file path in the field."
          End If
        Else
          startTime = Timer
          folderPath = "C:\Users\mari001\OneDrive - Hennepin County\Desktop\XML Files"

          Set fso = CreateObject("Scripting.FileSystemObject")
          Set folder = fso.GetFolder(folderPath)
          XML_file_found = False
          file_count = 0

          For Each file In folder.Files
            If InStr(1, file.Name, "_" & application_ID & "_", vbTextCompare) > 0 Then
              msgbox "Found: " & file.Path
              XML_file_path = file.Path
              XML_file_found = True
              Exit For
            End If
            file_count = file_count + 1
          Next
          If XML_file_found = False Then
            err_msg = err_msg & vbCr & "* The script was unable to locate a MNBenefits XML file with the application ID you provided. You must click the 'Select File' button and select the XML file or manually enter the file path in the field."
          End If
          endTime = Timer
          duration = endTime - startTime
          msgbox "Search took " & duration & " seconds. It evaluated " & file_count & " files."
        End If
      End If
    End If
    If application_ID_checkbox + manual_file_select_checkbox + enter_app_manually_checkbox = 0 then err_msg = err_msg & vbCr & "* You must check the box for one of the options to locate the MNBenefits XML File or enter application details manually."
    If application_ID_checkbox + manual_file_select_checkbox + enter_app_manually_checkbox > 1 then err_msg = err_msg & vbCr & "* You must only check ONE checkbox."
    If application_ID_checkbox = 1 and (trim(application_ID) = "" OR len(application_ID) <> 10 OR IsNumeric(application_ID) = false) then err_msg = err_msg & vbCr & "* You must enter the 10-digit application ID."
    If manual_file_select_checkbox = 1 and trim(file_path) = "" then err_msg = err_msg & vbCr & "* You must click the 'Select File' button and select the XML file or manually enter the file path in the field."
		If err_msg <> "" and err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Create XML object
Dim xmlDoc
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.async = False

'Load the XML file
xmlDoc.load(XML_file_path)

If xmlDoc.parseError.errorCode <> 0 Then
  'Release the XML DOM object when you're done
  Set xmlDoc = Nothing
  script_end_procedure("Error in XML: " & xmlDoc.parseError.reason)
End If

' Get all of the household members' information
Dim member_count
member_count = 0
Dim householdMembers()
Const MEMBER_FIRST_NAME     = 0
Const MEMBER_LAST_NAME      = 1
Const MEMBER_DOB            = 2
Const MEMBER_SSN            = 3
Const MEMBER_RELATIONSHIP   = 4
Const MEMBER_MARITAL_STATUS = 5
Const MEMBER_CITIZENSHIP    = 6
Const MEMBER_GENDER         = 7

ReDim householdMembers(MEMBER_GENDER, member_count)   'Redimmed to the size of the last constant

Dim objHouseholdMemberNode, objHouseholdMemberNodes
Set objHouseholdMemberNode = xmlDoc.selectSingleNode("//ns4:HouseholdInfo")
Set objHouseholdMemberNodes = objHouseholdMemberNode.selectNodes("ns4:HouseholdMember")

Dim objMemberNode, objRoot
Dim objFirstNameNode, objLastNameNode, objSSNNode, objDOBNode, objRelationshipNode, objMaritalStatusNode, objGenderNode

For Each objMemberNode In objHouseholdMemberNodes
  'MsgBox objMemberNode.InnerText
  Set objFirstNameNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:Person/ns4:FirstName")
  Set objLastNameNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:Person/ns4:LastName")
  Set objSSNNode = objMemberNode.selectSingleNode("ns4:CitizenshipInfo/ns4:SSNInfo/ns4:SSN")
  Set objDOBNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:DOB")
  Set objRelationshipNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:Relationship") 
  Set objMaritalStatusNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:MaritalStatus")
  Set objCitizenshipNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:CitizenshipInfo")
  Set objGenderNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:Gender")

  If Not objFirstNameNode Is Nothing Then
    householdMembers(MEMBER_FIRST_NAME, member_count) = objFirstNameNode.Text
  End If
  If Not objLastNameNode Is Nothing Then
    householdMembers(MEMBER_LAST_NAME, member_count) = objLastNameNode.Text
  End If
  If Not objDOBNode Is Nothing Then
    householdMembers(MEMBER_DOB, member_count) = objDOBNode.Text
  End If
  If Not objSSNNode Is Nothing Then
    householdMembers(MEMBER_SSN, member_count) = objSSNNode.Text
  End If
  If Not objSSNNode Is Nothing Then
    householdMembers(MEMBER_SSN, member_count) = objSSNNode.Text
  End If
  If Not objRelationshipNode Is Nothing Then
    householdMembers(MEMBER_RELATIONSHIP, member_count) = objRelationshipNode.Text
  End If
  If Not objMaritalStatusNode Is Nothing Then
    householdMembers(MEMBER_MARITAL_STATUS, member_count) = objMaritalStatusNode.Text
  End If
  If Not objCitizenshipNode Is Nothing Then
    householdMembers(MEMBER_CITIZENSHIP, member_count) = objMaritalStatusNode.Text
  End If
  If Not objGenderNode Is Nothing Then
    householdMembers(MEMBER_GENDER, member_count) = objGenderNode.Text
  End If

  If householdMembers(MEMBER_FIRST_NAME, member_count) = "" And householdMembers(MEMBER_LAST_NAME, member_count) = "" Then
    Exit For
  End If

  Dim memberNumber
  If member_count < 9 Then
    memberNumber = "0" & member_count + 1
  Else
    memberNumber = member_count + 1
  End If

  member_count = member_count + 1
  ReDim Preserve householdMembers(MEMBER_GENDER, member_count)
Next

'Gather application date and application ID 
Dim formatted_app_Date, objApplicationDate, applicationDate, applicationMonth, applicationDay, applicationYear
' Application Date - First try to get ApplicationDate (new business logic date), then fall back to SubmitDate

' Try to get the new ApplicationDate field first
Set objApplicationDate = xmlDoc.selectSingleNode("//CONTENT/ap:Bytes/io4:ApplicationDate")
If objApplicationDate Is Nothing Then
  ' If ApplicationDate doesn't exist, fall back to SubmitDate for backward compatibility
  Set objApplicationDate = xmlDoc.selectSingleNode("//CONTENT/ap:Bytes/io4:SubmitDate")
End If

If Not objApplicationDate Is Nothing Then
  applicationDate = objApplicationDate.Text
  applicationMonth = Mid(applicationDate, 6, 2)
  applicationDay = Mid(applicationDate, 9, 2)
  applicationYear = Mid(applicationDate, 1, 4)
Else ' Use the current date if neither application date is available    
  applicationMonth = Right("0" & Month(currentDate), 2)
  applicationDay = Right("0" & Day(currentDate), 2)
  applicationYear = Year(currentDate)
End If

formatted_app_Date = applicationMonth & "/" & applicationDay & "/" & applicationYear
MAXIS_footer_month = applicationMonth
MAXIS_footer_year = Mid(applicationYear, 3, 2)

Dim objApplicationId
' Application Id
Set objApplicationId = xmlDoc.selectSingleNode("//CONTENT/ap:Bytes/io4:ApplicationID")
If Not objApplicationId Is Nothing Then
  applicationId = objApplicationId.Text
End If

'Validate the provided application ID against the application ID in the XML file
If application_ID_checkbox = 1 Then
  If applicationId <> application_ID Then script_end_procedure_with_error_report("The application ID provided to locate the MNBenefits XML file does not match the application ID in the XML file. Please try running the script again.")
End If

'Gather the household address and mailing address details from the XML
Dim objHouseholdAddress, objHouseholdCity, objHouseholdState, objHouseholdZip, objMailingAddress, objMailingCity, objMailingState, objMailingZip, objCounty, objPhoneNumber

'Get the household address information from the XML
Set objHouseholdAddress = xmlDoc.selectSingleNode("//CONTENT/ap:Bytes/ns4:ContactInfo/ns4:Address/ns4:Line")
If objHouseholdAddress Is Nothing Then
  'To do - add handling for "No permanent address" with no details for city, zip, etc.
  ' If ApplicationDate doesn't exist, fall back to SubmitDate for backward compatibility
  Set objApplicationDate = xmlDoc.selectSingleNode("//CONTENT/ap:Bytes/io4:SubmitDate")
End If
Set objHouseholdCity = xmlDoc.selectSingleNode("//CONTENT/ap:Bytes/ns4:ContactInfo/ns4:Address/ns4:City")
Set objHouseholdState = xmlDoc.selectSingleNode("//CONTENT/ap:Bytes/ns4:ContactInfo/ns4:Address/ns4:State/ns4:StateCode")
Set objHouseholdZip = xmlDoc.selectSingleNode("//CONTENT/ap:Bytes/ns4:ContactInfo/ns4:Address/ns4:Zip5")
'To do - add handling if household address info blank
'To do - add handling if mailing address info blank

Set objMailingAddress = xmlDoc.selectSingleNode("//CONTENT/ap:Bytes/ns4:ContactInfo/ns4:MailingAddress/ns4:Line")
Set objMailingCity = xmlDoc.selectSingleNode("//CONTENT/ap:Bytes/ns4:ContactInfo/ns4:MailingAddress/ns4:City")
Set objMailingState = xmlDoc.selectSingleNode("//CONTENT/ap:Bytes/ns4:ContactInfo/ns4:MailingAddress/ns4:State/ns4:StateCode")
Set objMailingZip = xmlDoc.selectSingleNode("//CONTENT/ap:Bytes/ns4:ContactInfo/ns4:MailingAddress/ns4:Zip5")

Set objPhoneNumber = xmlDoc.selectSingleNode("//CONTENT/ap:Bytes/ns4:ContactInfo/ns4:Phone/ns4:PhoneNumber")
Set objCounty = xmlDoc.selectSingleNode("//CONTENT/ap:Bytes/ns4:ContactInfo/ns4:CountyResidence/ns4:COR")

'Extract text from object
household_address       = objHouseholdAddress.Text
household_city          = objHouseholdCity.Text
household_state         = objHouseholdState.Text
household_zip           = objHouseholdZip.Text
household_phone_number  = objPhoneNumber.Text
household_county        = objCounty.Text
mailing_address         = objMailingAddress.Text
mailing_city            = objMailingCity.Text
mailing_state           = objMailingState.Text
mailing_zip             = objMailingZip.Text

' Release the XML DOM object when you're done
Set xmlDoc = Nothing

dialog_member_count = 0

'XML File Confirmation Dialog
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 285, 245, "Verify MNBenefits XML Details - Household Members"
  Text 5, 5, 275, 20, "Please verify that the correct XML file has been selected. If you need to change the XML file, please press the 'Reselect XML' button below."  
  GroupBox 10, 35, 270, 155, "MNBenefits XML File Details"
  Text 15, 45, 50, 10, "Application ID:"
  Text 100, 45, 50, 10, application_ID
  Text 15, 55, 60, 10, "Application Date:"
  Text 100, 55, 60, 10, formatted_app_Date
  Text 15, 65, 75, 10, "Household Member 1:"
  Text 100, 65, 175, 10, left(householdMembers(MEMBER_LAST_NAME, dialog_member_count) & ", " & householdMembers(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & householdMembers(MEMBER_DOB, dialog_member_count) & "; " & householdMembers(MEMBER_SSN, dialog_member_count) & ")"  
  dialog_member_count = dialog_member_count + 1 
  If member_count > 1 Then
    Text 15, 75, 75, 10, "Household Member 2:"
    Text 100, 75, 175, 10, left(householdMembers(MEMBER_LAST_NAME, dialog_member_count) & ", " & householdMembers(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & householdMembers(MEMBER_DOB, dialog_member_count) & "; " & householdMembers(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 2 Then
    Text 15, 85, 75, 10, "Household Member 3:"
    Text 100, 85, 175, 10, left(householdMembers(MEMBER_LAST_NAME, dialog_member_count) & ", " & householdMembers(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & householdMembers(MEMBER_DOB, dialog_member_count) & "; " & householdMembers(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 3 Then
    Text 15, 95, 75, 10, "Household Member 4:"
    Text 100, 95, 175, 10, left(householdMembers(MEMBER_LAST_NAME, dialog_member_count) & ", " & householdMembers(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & householdMembers(MEMBER_DOB, dialog_member_count) & "; " & householdMembers(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 4 Then
    Text 15, 105, 75, 10, "Household Member 5:"
    Text 100, 105, 175, 10, left(householdMembers(MEMBER_LAST_NAME, dialog_member_count) & ", " & householdMembers(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & householdMembers(MEMBER_DOB, dialog_member_count) & "; " & householdMembers(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 5 Then
    Text 15, 115, 75, 10, "Household Member 6:"
    Text 100, 115, 175, 10, left(householdMembers(MEMBER_LAST_NAME, dialog_member_count) & ", " & householdMembers(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & householdMembers(MEMBER_DOB, dialog_member_count) & "; " & householdMembers(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 6 Then
    Text 15, 125, 75, 10, "Household Member 7:"
    Text 100, 125, 175, 10, left(householdMembers(MEMBER_LAST_NAME, dialog_member_count) & ", " & householdMembers(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & householdMembers(MEMBER_DOB, dialog_member_count) & "; " & householdMembers(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 7 Then
    Text 15, 135, 75, 10, "Household Member 8:"
    Text 100, 135, 175, 10, left(householdMembers(MEMBER_LAST_NAME, dialog_member_count) & ", " & householdMembers(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & householdMembers(MEMBER_DOB, dialog_member_count) & "; " & householdMembers(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 8 Then
    Text 15, 145, 75, 10, "Household Member 9:"
    Text 100, 145, 175, 10, left(householdMembers(MEMBER_LAST_NAME, dialog_member_count) & ", " & householdMembers(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & householdMembers(MEMBER_DOB, dialog_member_count) & "; " & householdMembers(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 9 Then
    Text 15, 155, 80, 10, "Household Member 10:"
    Text 100, 155, 175, 10, left(householdMembers(MEMBER_LAST_NAME, dialog_member_count) & ", " & householdMembers(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & householdMembers(MEMBER_DOB, dialog_member_count) & "; " & householdMembers(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 10 Then
    Text 15, 165, 80, 10, "Household Member 11:"
    Text 100, 165, 175, 10, left(householdMembers(MEMBER_LAST_NAME, dialog_member_count) & ", " & householdMembers(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & householdMembers(MEMBER_DOB, dialog_member_count) & "; " & householdMembers(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 11 Then
    Text 15, 175, 80, 10, "Household Member 12:"
    Text 100, 175, 175, 10, left(householdMembers(MEMBER_LAST_NAME, dialog_member_count) & ", " & householdMembers(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & householdMembers(MEMBER_DOB, dialog_member_count) & "; " & householdMembers(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  ButtonGroup ButtonPressed
    PushButton 235, 225, 45, 15, "Continue", continue_button
    PushButton 10, 225, 50, 15, "Reselect XML", reselect_xml_button
EndDialog

DO
  dialog Dialog1
  cancel_without_confirmation
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'To do - add handling for cases where not address is provided
'To do - add county of residence
'To do - convert county of residence to county code?
'Dialog to confirm Application and Address Information
BeginDialog Dialog1, 0, 0, 256, 265, "Process MNBenefits Application"
  Text 10, 5, 240, 20, "Please verify the application and address details pulled from the XML file below. Make updates as needed."
  Text 15, 30, 150, 10, "Adjust date to correct business day, if needed"
  Text 15, 45, 60, 10, "Application Date: "
  EditBox 80, 40, 40, 15, formatted_app_date
  GroupBox 10, 60, 175, 105, "Household Address"
  Text 15, 75, 35, 10, "Address:"
  EditBox 70, 70, 100, 15, household_address
  Text 15, 90, 25, 10, "City:"
  EditBox 70, 85, 100, 15, household_city
  Text 15, 105, 30, 10, "State:"
  EditBox 70, 100, 20, 15, household_state
  Text 15, 120, 20, 10, "Zip:"
  EditBox 70, 115, 30, 15, household_zip
  Text 15, 135, 55, 10, "Phone number:"
  EditBox 70, 130, 100, 15, household_phone_number
  GroupBox 10, 165, 175, 75, "Mailing Address"
  Text 15, 180, 35, 10, "Address:"
  EditBox 70, 175, 100, 15, mailing_address
  Text 15, 195, 25, 10, "City:"
  EditBox 70, 190, 100, 15, mailing_city
  Text 15, 210, 30, 10, "State:"
  EditBox 70, 205, 20, 15, mailing_state
  Text 15, 225, 20, 10, "Zip:"
  EditBox 70, 220, 30, 15, mailing_zip
  Text 15, 150, 30, 10, "County:"
  EditBox 70, 145, 100, 15, household_county
  ButtonGroup ButtonPressed
    PushButton 200, 245, 50, 15, "Confirm", confirm_address_button
EndDialog

DO
	DO
		err_msg = ""					'establishing value of variable, this is necessary for the Do...LOOP
		dialog Dialog1				'main dialog
		cancel_without_confirmation
    If ButtonPressed = file_selection_button then 
      call file_selection_system_dialog(XML_file_path, ".xml")
      err_msg = "LOOP"
    End If
    If trim(formatted_app_date) = "" OR IsDate(formatted_app_date) = False then err_msg = err_msg & vbCr & "* You must enter the application date in the format MM/DD/YYYY."
    If trim(household_address) = "" Then err_msg = err_msg & vbCr & "* The household address field cannot be blank."
    If trim(household_city) = "" Then err_msg = err_msg & vbCr & "* The city field cannot be blank."
    If trim(household_state) = "" Then err_msg = err_msg & vbCr & "* The state field cannot be blank."
    If trim(household_zip) = "" Then err_msg = err_msg & vbCr & "* The zip code field cannot be blank."
    'To do - confirm if phone number is required
    ' If trim(household_phone_number) = "" Then then err_msg = err_msg & vbCr & "* The household address field cannot be blank."
    If trim(mailing_address) = "" Then err_msg = err_msg & vbCr & "* The mailing address field cannot be blank."
    If trim(mailing_city) = "" Then err_msg = err_msg & vbCr & "* The mailing address city field cannot be blank."
    If trim(mailing_state) = "" Then err_msg = err_msg & vbCr & "* The mailing address state field cannot be blank."
    If trim(mailing_zip) = "" Then err_msg = err_msg & vbCr & "* The mailing address zip code field cannot be blank."

		If err_msg <> "" and err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Start at the first dialog and start dialog loop
dialog_count = 1
hh_memb_dialog_loop = "Active"
Call determine_member_dialogs_display()

Do
  Do
    Do
      Dialog1 = "" 'Blanking out previous dialog detail

      Call dialog_selection(dialog_count)

      'Blank out variables on each new dialog
      err_msg = ""

      dialog Dialog1 					'Calling a dialog without an assigned variable will call the most recently defined dialog
      cancel_confirmation
      ' Call dialog_specific_error_handling()	'function for error handling of main dialog of forms
      Call button_movement()				'function to move throughout the dialogs
      ' Call incomplete_dialog_handling()     'function to alert worker to incomplete dialogs
    Loop until err_msg = ""
  Loop until hh_memb_dialog_loop = "Completed"
  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
