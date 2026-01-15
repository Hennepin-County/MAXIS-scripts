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
update_xml_button       = 205
back_button             = 206
reselect_xml_button     = 207
continue_button         = 208
proceed_button          = 209
return_button           = 210



'--Other buttons
' instructions_btn
' file_selection_button
hh_memb_1_and_2_button    = 301
hh_memb_3_and_4_button    = 302
hh_memb_5_and_6_button    = 303
hh_memb_7_and_8_button    = 304
hh_memb_9_and_10_button   = 305
hh_memb_11_and_12_button  = 306
case_number_nav_0_button  = 307
case_number_nav_1_button  = 308
case_number_nav_2_button  = 309
case_number_nav_3_button  = 310
case_number_nav_4_button  = 311



'Defining variables


'Dimming variables
Dim folderPath, confirmation_number, fso, folder, fileList, file, xml_file_path, script_testing, proceed_new_case

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
    EditBox 70, 45, 100, 15, household_members(MEMBER_FIRST_NAME, 0)
    Text 15, 65, 40, 10, "Last name:"
    EditBox 70, 60, 100, 15, household_members(MEMBER_LAST_NAME, 0)
    Text 15, 80, 30, 10, "Gender:"
    DropListBox 70, 75, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, 0)
    Text 15, 95, 50, 10, "Marital status:"
    DropListBox 70, 90, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", household_members(MEMBER_MARITAL_STATUS, 0)
    Text 15, 110, 45, 10, "Date of birth:"
    EditBox 70, 105, 100, 15, household_members(MEMBER_DOB, 0)
    Text 15, 125, 20, 10, "SSN:"
    EditBox 70, 120, 100, 15, household_members(MEMBER_SSN, 0)
    Text 15, 140, 45, 10, "Citizenship:"
    DropListBox 70, 135, 100, 15, "Select one:"+chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, 0)
    Text 15, 155, 45, 10, "Relationship:"
    DropListBox 70, 150, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, 0)
    If member_count > 1 Then
      GroupBox 10, 175, 175, 140, "Household Member Information"
      Text 15, 195, 40, 10, "First name:"
      EditBox 70, 190, 100, 15, household_members(MEMBER_FIRST_NAME, 1)
      Text 15, 210, 40, 10, "Last name:"
      EditBox 70, 205, 100, 15, household_members(MEMBER_LAST_NAME, 1)
      Text 15, 225, 30, 10, "Gender:"
      DropListBox 70, 220, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, 1)
      Text 15, 240, 50, 10, "Marital status:"
      DropListBox 70, 235, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", household_members(MEMBER_MARITAL_STATUS, 1)
      Text 15, 255, 45, 10, "Date of birth:"
      EditBox 70, 250, 100, 15, household_members(MEMBER_DOB, 1)
      Text 15, 270, 20, 10, "SSN:"
      EditBox 70, 265, 100, 15, household_members(MEMBER_SSN, 1)
      Text 15, 285, 45, 10, "Citizenship:"
      DropListBox 70, 280, 100, 15, "Select one:"+chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, 1)
      Text 15, 300, 45, 10, "Relationship:"
      DropListBox 70, 295, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, 1)
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

Function household_members_dialog_3_4()
  hh_memb_dialog_count = 2
  BeginDialog Dialog1, 0, 0, 281, 345, "Verify MNBenefits XML Details - Household Members"
    Text 5, 5, 250, 20, "Please review and verify the household member details for each household member pulled from the XML file below. Make any updates as needed."
    GroupBox 10, 30, 175, 140, "Household Member Information"
    Text 15, 50, 40, 10, "First name:"
    EditBox 70, 45, 100, 15, household_members(MEMBER_FIRST_NAME, 2)
    Text 15, 65, 40, 10, "Last name:"
    EditBox 70, 60, 100, 15, household_members(MEMBER_LAST_NAME, 2)
    Text 15, 80, 30, 10, "Gender:"
    DropListBox 70, 75, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, 2)
    Text 15, 95, 50, 10, "Marital status:"
    DropListBox 70, 90, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", household_members(MEMBER_MARITAL_STATUS, 2)
    Text 15, 110, 45, 10, "Date of birth:"
    EditBox 70, 105, 100, 15, household_members(MEMBER_DOB, 2)
    Text 15, 125, 20, 10, "SSN:"
    EditBox 70, 120, 100, 15, household_members(MEMBER_SSN, 2)
    Text 15, 140, 45, 10, "Citizenship:"
    DropListBox 70, 135, 100, 15, "Select one:"+chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, 2)
    Text 15, 155, 45, 10, "Relationship:"
    DropListBox 70, 150, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, 2)
    If member_count > 3 Then
      GroupBox 10, 175, 175, 140, "Household Member Information"
      Text 15, 195, 40, 10, "First name:"
      EditBox 70, 190, 100, 15, household_members(MEMBER_FIRST_NAME, 3)
      Text 15, 210, 40, 10, "Last name:"
      EditBox 70, 205, 100, 15, household_members(MEMBER_LAST_NAME, 3)
      Text 15, 225, 30, 10, "Gender:"
      DropListBox 70, 220, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, 3)
      Text 15, 240, 50, 10, "Marital status:"
      DropListBox 70, 235, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", household_members(MEMBER_MARITAL_STATUS, 3)
      Text 15, 255, 45, 10, "Date of birth:"
      EditBox 70, 250, 100, 15, household_members(MEMBER_DOB, 3)
      Text 15, 270, 20, 10, "SSN:"
      EditBox 70, 265, 100, 15, household_members(MEMBER_SSN, 3)
      Text 15, 285, 45, 10, "Citizenship:"
      DropListBox 70, 280, 100, 15, "Select one:"+chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, 3)
      Text 15, 300, 45, 10, "Relationship:"
      DropListBox 70, 295, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, 3)
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

Function household_members_dialog_5_6()
hh_memb_dialog_count = 3
  BeginDialog Dialog1, 0, 0, 281, 345, "Verify MNBenefits XML Details - Household Members"
    Text 5, 5, 250, 20, "Please review and verify the household member details for each household member pulled from the XML file below. Make any updates as needed."
    GroupBox 10, 30, 175, 140, "Household Member Information"
    Text 15, 50, 40, 10, "First name:"
    EditBox 70, 45, 100, 15, household_members(MEMBER_FIRST_NAME, 4)
    Text 15, 65, 40, 10, "Last name:"
    EditBox 70, 60, 100, 15, household_members(MEMBER_LAST_NAME, 4)
    Text 15, 80, 30, 10, "Gender:"
    DropListBox 70, 75, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, 4)
    Text 15, 95, 50, 10, "Marital status:"
    DropListBox 70, 90, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", household_members(MEMBER_MARITAL_STATUS, 4)
    Text 15, 110, 45, 10, "Date of birth:"
    EditBox 70, 105, 100, 15, household_members(MEMBER_DOB, 4)
    Text 15, 125, 20, 10, "SSN:"
    EditBox 70, 120, 100, 15, household_members(MEMBER_SSN, 4)
    Text 15, 140, 45, 10, "Citizenship:"
    DropListBox 70, 135, 100, 15, "Select one:"+chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, 4)
    Text 15, 155, 45, 10, "Relationship:"
    DropListBox 70, 150, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, 4)
    If member_count > 5 Then
      GroupBox 10, 175, 175, 140, "Household Member Information"
      Text 15, 195, 40, 10, "First name:"
      EditBox 70, 190, 100, 15, household_members(MEMBER_FIRST_NAME, 5)
      Text 15, 210, 40, 10, "Last name:"
      EditBox 70, 205, 100, 15, household_members(MEMBER_LAST_NAME, 5)
      Text 15, 225, 30, 10, "Gender:"
      DropListBox 70, 220, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, 5)
      Text 15, 240, 50, 10, "Marital status:"
      DropListBox 70, 235, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", household_members(MEMBER_MARITAL_STATUS, 5)
      Text 15, 255, 45, 10, "Date of birth:"
      EditBox 70, 250, 100, 15, household_members(MEMBER_DOB, 5)
      Text 15, 270, 20, 10, "SSN:"
      EditBox 70, 265, 100, 15, household_members(MEMBER_SSN, 5)
      Text 15, 285, 45, 10, "Citizenship:"
      DropListBox 70, 280, 100, 15, "Select one:"+chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, 5)
      Text 15, 300, 45, 10, "Relationship:"
      DropListBox 70, 295, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, 5)
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

Function household_members_dialog_7_8()   
  hh_memb_dialog_count = 4
  BeginDialog Dialog1, 0, 0, 281, 345, "Verify MNBenefits XML Details - Household Members"
    Text 5, 5, 250, 20, "Please review and verify the household member details for each household member pulled from the XML file below. Make any updates as needed."
    GroupBox 10, 30, 175, 140, "Household Member Information"
    Text 15, 50, 40, 10, "First name:"
    EditBox 70, 45, 100, 15, household_members(MEMBER_FIRST_NAME, 6)
    Text 15, 65, 40, 10, "Last name:"
    EditBox 70, 60, 100, 15, household_members(MEMBER_LAST_NAME, 6)
    Text 15, 80, 30, 10, "Gender:"
    DropListBox 70, 75, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, 6)
    Text 15, 95, 50, 10, "Marital status:"
    DropListBox 70, 90, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", household_members(MEMBER_MARITAL_STATUS, 6)
    Text 15, 110, 45, 10, "Date of birth:"
    EditBox 70, 105, 100, 15, household_members(MEMBER_DOB, 6)
    Text 15, 125, 20, 10, "SSN:"
    EditBox 70, 120, 100, 15, household_members(MEMBER_SSN, 6)
    Text 15, 140, 45, 10, "Citizenship:"
    DropListBox 70, 135, 100, 15, "Select one:"+chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, 6)
    Text 15, 155, 45, 10, "Relationship:"
    DropListBox 70, 150, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, 6)
    If member_count > 7 Then
      GroupBox 10, 175, 175, 140, "Household Member Information"
      Text 15, 195, 40, 10, "First name:"
      EditBox 70, 190, 100, 15, household_members(MEMBER_FIRST_NAME, 7)
      Text 15, 210, 40, 10, "Last name:"
      EditBox 70, 205, 100, 15, household_members(MEMBER_LAST_NAME, 7)
      Text 15, 225, 30, 10, "Gender:"
      DropListBox 70, 220, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, 7)
      Text 15, 240, 50, 10, "Marital status:"
      DropListBox 70, 235, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", household_members(MEMBER_MARITAL_STATUS, 7)
      Text 15, 255, 45, 10, "Date of birth:"
      EditBox 70, 250, 100, 15, household_members(MEMBER_DOB, 7)
      Text 15, 270, 20, 10, "SSN:"
      EditBox 70, 265, 100, 15, household_members(MEMBER_SSN, 7)
      Text 15, 285, 45, 10, "Citizenship:"
      DropListBox 70, 280, 100, 15, "Select one:"+chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, 7)
      Text 15, 300, 45, 10, "Relationship:"
      DropListBox 70, 295, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, 7)
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

Function household_members_dialog_9_10()
  hh_memb_dialog_count = 5
  BeginDialog Dialog1, 0, 0, 281, 345, "Verify MNBenefits XML Details - Household Members"
    Text 5, 5, 250, 20, "Please review and verify the household member details for each household member pulled from the XML file below. Make any updates as needed."
    GroupBox 10, 30, 175, 140, "Household Member Information"
    Text 15, 50, 40, 10, "First name:"
    EditBox 70, 45, 100, 15, household_members(MEMBER_FIRST_NAME, 8)
    Text 15, 65, 40, 10, "Last name:"
    EditBox 70, 60, 100, 15, household_members(MEMBER_LAST_NAME, 8)
    Text 15, 80, 30, 10, "Gender:"
    DropListBox 70, 75, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, 8)
    Text 15, 95, 50, 10, "Marital status:"
    DropListBox 70, 90, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", household_members(MEMBER_MARITAL_STATUS, 8)
    Text 15, 110, 45, 10, "Date of birth:"
    EditBox 70, 105, 100, 15, household_members(MEMBER_DOB, 8)
    Text 15, 125, 20, 10, "SSN:"
    EditBox 70, 120, 100, 15, household_members(MEMBER_SSN, 8)
    Text 15, 140, 45, 10, "Citizenship:"
    DropListBox 70, 135, 100, 15, "Select one:"+chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, 8)
    Text 15, 155, 45, 10, "Relationship:"
    DropListBox 70, 150, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, 8)
    If member_count > 9 Then
      GroupBox 10, 175, 175, 140, "Household Member Information"
      Text 15, 195, 40, 10, "First name:"
      EditBox 70, 190, 100, 15, household_members(MEMBER_FIRST_NAME, 9)
      Text 15, 210, 40, 10, "Last name:"
      EditBox 70, 205, 100, 15, household_members(MEMBER_LAST_NAME, 9)
      Text 15, 225, 30, 10, "Gender:"
      DropListBox 70, 220, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, 9)
      Text 15, 240, 50, 10, "Marital status:"
      DropListBox 70, 235, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", household_members(MEMBER_MARITAL_STATUS, 9)
      Text 15, 255, 45, 10, "Date of birth:"
      EditBox 70, 250, 100, 15, household_members(MEMBER_DOB, 9)
      Text 15, 270, 20, 10, "SSN:"
      EditBox 70, 265, 100, 15, household_members(MEMBER_SSN, 9)
      Text 15, 285, 45, 10, "Citizenship:"
      DropListBox 70, 280, 100, 15, "Select one:"+chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, 9)
      Text 15, 300, 45, 10, "Relationship:"
      DropListBox 70, 295, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, 9)
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

Function household_members_dialog_11_12()
  hh_memb_dialog_count = 6
  BeginDialog Dialog1, 0, 0, 281, 345, "Verify MNBenefits XML Details - Household Members"
    Text 5, 5, 250, 20, "Please review and verify the household member details for each household member pulled from the XML file below. Make any updates as needed."
    GroupBox 10, 30, 175, 140, "Household Member Information"
    Text 15, 50, 40, 10, "First name:"
    EditBox 70, 45, 100, 15, household_members(MEMBER_FIRST_NAME, 10)
    Text 15, 65, 40, 10, "Last name:"
    EditBox 70, 60, 100, 15, household_members(MEMBER_LAST_NAME, 10)
    Text 15, 80, 30, 10, "Gender:"
    DropListBox 70, 75, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, 10)
    Text 15, 95, 50, 10, "Marital status:"
    DropListBox 70, 90, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", household_members(MEMBER_MARITAL_STATUS, 10)
    Text 15, 110, 45, 10, "Date of birth:"
    EditBox 70, 105, 100, 15, household_members(MEMBER_DOB, 10)
    Text 15, 125, 20, 10, "SSN:"
    EditBox 70, 120, 100, 15, household_members(MEMBER_SSN, 10)
    Text 15, 140, 45, 10, "Citizenship:"
    DropListBox 70, 135, 100, 15, "Select one:"+chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, 10)
    Text 15, 155, 45, 10, "Relationship:"
    DropListBox 70, 150, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, 10)
    If member_count > 11 Then
      GroupBox 10, 175, 175, 140, "Household Member Information"
      Text 15, 195, 40, 10, "First name:"
      EditBox 70, 190, 100, 15, household_members(MEMBER_FIRST_NAME, 11)
      Text 15, 210, 40, 10, "Last name:"
      EditBox 70, 205, 100, 15, household_members(MEMBER_LAST_NAME, 11)
      Text 15, 225, 30, 10, "Gender:"
      DropListBox 70, 220, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, 11)
      Text 15, 240, 50, 10, "Marital status:"
      DropListBox 70, 235, 100, 15, "Select one:"+chr(9)+"Never married"+chr(9)+"Married"+chr(9)+"Married living with spouse"+chr(9)+"Married living apart"+chr(9)+"Separated"+chr(9)+"Legally separated"+chr(9)+"Divorced"+chr(9)+"Widowed", household_members(MEMBER_MARITAL_STATUS, 11)
      Text 15, 255, 45, 10, "Date of birth:"
      EditBox 70, 250, 100, 15, household_members(MEMBER_DOB, 11)
      Text 15, 270, 20, 10, "SSN:"
      EditBox 70, 265, 100, 15, household_members(MEMBER_SSN, 11)
      Text 15, 285, 45, 10, "Citizenship:"
      DropListBox 70, 280, 100, 15, "Select one:"+chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, 11)
      Text 15, 300, 45, 10, "Relationship:"
      DropListBox 70, 295, 100, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, 11)
    End If
    ButtonGroup ButtonPressed
      PushButton 230, 325, 45, 15, "Submit", submit_hh_memb_button
      PushButton 5, 325, 45, 15, "Previous", previous_hh_memb_button
    GroupBox 195, 30, 70, 105, "Navigation"
    ButtonGroup ButtonPressed
      Call determine_hh_memb_buttons()
  EndDialog
End Function

Function confirm_xml_update_dialog()
  BeginDialog Dialog1, 0, 0, 281, 70, "Update XML File with Updates"
    Text 10, 5, 265, 35, "The script will now update the XML file with any changes made to the address and/or household member details. Press 'Update XML with changes' to update the XML file. If you want to review the changes to the XML file before changing, press 'Back'. To cancel the script entirely, press 'Cancel script'."
    ButtonGroup ButtonPressed
      PushButton 185, 50, 90, 15, "Update XML with changes", update_xml_button
      PushButton 160, 50, 25, 15, "Back", back_button
      CancelButton 10, 50, 50, 15
  EndDialog
End Function

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
  If hh_memb_dialog_loop = "Active" Then
    If dialog_count = 1 Then
      If trim(household_members(MEMBER_FIRST_NAME, 0)) = "" Then err_msg = err_msg & vbNewLine & "* The first name field cannot be left blank."
      If trim(household_members(MEMBER_LAST_NAME, 0)) = "" Then err_msg = err_msg & vbNewLine & "* The last name field cannot be left blank."
      If trim(household_members(MEMBER_GENDER, 0)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A gender option must be selected from the dropdown list."
      If trim(household_members(MEMBER_MARITAL_STATUS, 0)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A marital status option must be selected from the dropdown list."
      If trim(household_members(MEMBER_DOB, 0)) = "" OR (Not IsDate(trim(household_members(MEMBER_DOB, 0)))) Then err_msg = err_msg & vbNewLine & "* The date of birth field must be filled out in the MM/DD/YYYY format."
      If trim(household_members(MEMBER_SSN, 0)) <> "" AND Len(trim(household_members(MEMBER_SSN, 0))) <> 11 Then err_msg = err_msg & vbNewLine & "* The SSN must be in the format ###-##-####."
      If trim(household_members(MEMBER_RELATIONSHIP, 0)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A relationship option must be selected from the dropdown list."
      If member_count > 1 Then
        If trim(household_members(MEMBER_FIRST_NAME, 1)) = "" Then err_msg = err_msg & vbNewLine & "* The first name field cannot be left blank."
        If trim(household_members(MEMBER_LAST_NAME, 1)) = "" Then err_msg = err_msg & vbNewLine & "* The last name field cannot be left blank."
        If trim(household_members(MEMBER_GENDER, 1)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A gender option must be selected from the dropdown list."
        If trim(household_members(MEMBER_MARITAL_STATUS, 1)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A marital status option must be selected from the dropdown list."
        If trim(household_members(MEMBER_DOB, 1)) = "" OR (Not IsDate(trim(household_members(MEMBER_DOB, 1)))) Then err_msg = err_msg & vbNewLine & "* The date of birth field must be filled out in the MM/DD/YYYY format."
        If trim(household_members(MEMBER_SSN, 1)) <> "" AND Len(trim(household_members(MEMBER_SSN, 1))) <> 11 Then err_msg = err_msg & vbNewLine & "* The SSN must be in the format ###-##-####."
        If trim(household_members(MEMBER_RELATIONSHIP, 1)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A relationship option must be selected from the dropdown list."
      End If
    End If
    
    If dialog_count = 2 Then
      If trim(household_members(MEMBER_FIRST_NAME, 2)) = "" Then err_msg = err_msg & vbNewLine & "* The first name field cannot be left blank."
      If trim(household_members(MEMBER_LAST_NAME, 2)) = "" Then err_msg = err_msg & vbNewLine & "* The last name field cannot be left blank."
      If trim(household_members(MEMBER_GENDER, 2)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A gender option must be selected from the dropdown list."
      If trim(household_members(MEMBER_MARITAL_STATUS, 2)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A marital status option must be selected from the dropdown list."
      If trim(household_members(MEMBER_DOB, 2)) = "" OR (Not IsDate(trim(household_members(MEMBER_DOB, 2)))) Then err_msg = err_msg & vbNewLine & "* The date of birth field must be filled out in the MM/DD/YYYY format."
      If trim(household_members(MEMBER_SSN, 2)) <> "" AND Len(trim(household_members(MEMBER_SSN, 2))) <> 11 Then err_msg = err_msg & vbNewLine & "* The SSN must be in the format ###-##-####."
      If trim(household_members(MEMBER_RELATIONSHIP, 2)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A relationship option must be selected from the dropdown list."
      If member_count > 3 Then
        If trim(household_members(MEMBER_FIRST_NAME, 3)) = "" Then err_msg = err_msg & vbNewLine & "* The first name field cannot be left blank."
        If trim(household_members(MEMBER_LAST_NAME, 3)) = "" Then err_msg = err_msg & vbNewLine & "* The last name field cannot be left blank."
        If trim(household_members(MEMBER_GENDER, 3)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A gender option must be selected from the dropdown list."
        If trim(household_members(MEMBER_MARITAL_STATUS, 3)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A marital status option must be selected from the dropdown list."
        If trim(household_members(MEMBER_DOB, 3)) = "" OR (Not IsDate(trim(household_members(MEMBER_DOB, 3)))) Then err_msg = err_msg & vbNewLine & "* The date of birth field must be filled out in the MM/DD/YYYY format."
        If trim(household_members(MEMBER_SSN, 3)) <> "" AND Len(trim(household_members(MEMBER_SSN, 3))) <> 11 Then err_msg = err_msg & vbNewLine & "* The SSN must be in the format ###-##-####."
        If trim(household_members(MEMBER_RELATIONSHIP, 3)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A relationship option must be selected from the dropdown list."
      End If
    End If
    
    If dialog_count = 3 Then
      If trim(household_members(MEMBER_FIRST_NAME, 4)) = "" Then err_msg = err_msg & vbNewLine & "* The first name field cannot be left blank."
      If trim(household_members(MEMBER_LAST_NAME, 4)) = "" Then err_msg = err_msg & vbNewLine & "* The last name field cannot be left blank."
      If trim(household_members(MEMBER_GENDER, 4)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A gender option must be selected from the dropdown list."
      If trim(household_members(MEMBER_MARITAL_STATUS, 4)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A marital status option must be selected from the dropdown list."
      If trim(household_members(MEMBER_DOB, 4)) = "" OR (Not IsDate(trim(household_members(MEMBER_DOB, 4)))) Then err_msg = err_msg & vbNewLine & "* The date of birth field must be filled out in the MM/DD/YYYY format."
      If trim(household_members(MEMBER_SSN, 4)) <> "" AND Len(trim(household_members(MEMBER_SSN, 4))) <> 11 Then err_msg = err_msg & vbNewLine & "* The SSN must be in the format ###-##-####."
      If trim(household_members(MEMBER_RELATIONSHIP, 4)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A relationship option must be selected from the dropdown list."
      If member_count > 5 Then
        If trim(household_members(MEMBER_FIRST_NAME, 5)) = "" Then err_msg = err_msg & vbNewLine & "* The first name field cannot be left blank."
        If trim(household_members(MEMBER_LAST_NAME, 5)) = "" Then err_msg = err_msg & vbNewLine & "* The last name field cannot be left blank."
        If trim(household_members(MEMBER_GENDER, 5)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A gender option must be selected from the dropdown list."
        If trim(household_members(MEMBER_MARITAL_STATUS, 5)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A marital status option must be selected from the dropdown list."
        If trim(household_members(MEMBER_DOB, 5)) = "" OR (Not IsDate(trim(household_members(MEMBER_DOB, 5)))) Then err_msg = err_msg & vbNewLine & "* The date of birth field must be filled out in the MM/DD/YYYY format."
        If trim(household_members(MEMBER_SSN, 5)) <> "" AND Len(trim(household_members(MEMBER_SSN, 5))) <> 11 Then err_msg = err_msg & vbNewLine & "* The SSN must be in the format ###-##-####."
        If trim(household_members(MEMBER_RELATIONSHIP, 5)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A relationship option must be selected from the dropdown list."
      End If
    End If
    
    If dialog_count = 4 Then
      If trim(household_members(MEMBER_FIRST_NAME, 6)) = "" Then err_msg = err_msg & vbNewLine & "* The first name field cannot be left blank."
      If trim(household_members(MEMBER_LAST_NAME, 6)) = "" Then err_msg = err_msg & vbNewLine & "* The last name field cannot be left blank."
      If trim(household_members(MEMBER_GENDER, 6)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A gender option must be selected from the dropdown list."
      If trim(household_members(MEMBER_MARITAL_STATUS, 6)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A marital status option must be selected from the dropdown list."
      If trim(household_members(MEMBER_DOB, 6)) = "" OR (Not IsDate(trim(household_members(MEMBER_DOB, 6)))) Then err_msg = err_msg & vbNewLine & "* The date of birth field must be filled out in the MM/DD/YYYY format."
      If trim(household_members(MEMBER_SSN, 6)) <> "" AND Len(trim(household_members(MEMBER_SSN, 6))) <> 11 Then err_msg = err_msg & vbNewLine & "* The SSN must be in the format ###-##-####."
      If trim(household_members(MEMBER_RELATIONSHIP, 6)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A relationship option must be selected from the dropdown list."
      If member_count > 7 Then
        If trim(household_members(MEMBER_FIRST_NAME, 7)) = "" Then err_msg = err_msg & vbNewLine & "* The first name field cannot be left blank."
        If trim(household_members(MEMBER_LAST_NAME, 7)) = "" Then err_msg = err_msg & vbNewLine & "* The last name field cannot be left blank."
        If trim(household_members(MEMBER_GENDER, 7)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A gender option must be selected from the dropdown list."
        If trim(household_members(MEMBER_MARITAL_STATUS, 7)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A marital status option must be selected from the dropdown list."
        If trim(household_members(MEMBER_DOB, 7)) = "" OR (Not IsDate(trim(household_members(MEMBER_DOB, 7)))) Then err_msg = err_msg & vbNewLine & "* The date of birth field must be filled out in the MM/DD/YYYY format."
        If trim(household_members(MEMBER_SSN, 7)) <> "" AND Len(trim(household_members(MEMBER_SSN, 7))) <> 11 Then err_msg = err_msg & vbNewLine & "* The SSN must be in the format ###-##-####."
        If trim(household_members(MEMBER_RELATIONSHIP, 7)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A relationship option must be selected from the dropdown list."
      End If
    End If
    
    If dialog_count = 5 Then
      If trim(household_members(MEMBER_FIRST_NAME, 8)) = "" Then err_msg = err_msg & vbNewLine & "* The first name field cannot be left blank."
      If trim(household_members(MEMBER_LAST_NAME, 8)) = "" Then err_msg = err_msg & vbNewLine & "* The last name field cannot be left blank."
      If trim(household_members(MEMBER_GENDER, 8)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A gender option must be selected from the dropdown list."
      If trim(household_members(MEMBER_MARITAL_STATUS, 8)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A marital status option must be selected from the dropdown list."
      If trim(household_members(MEMBER_DOB, 8)) = "" OR (Not IsDate(trim(household_members(MEMBER_DOB, 8)))) Then err_msg = err_msg & vbNewLine & "* The date of birth field must be filled out in the MM/DD/YYYY format."
      If trim(household_members(MEMBER_SSN, 8)) <> "" AND Len(trim(household_members(MEMBER_SSN, 8))) <> 11 Then err_msg = err_msg & vbNewLine & "* The SSN must be in the format ###-##-####."
      If trim(household_members(MEMBER_RELATIONSHIP, 8)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A relationship option must be selected from the dropdown list."
      If member_count > 9 Then
        If trim(household_members(MEMBER_FIRST_NAME, 9)) = "" Then err_msg = err_msg & vbNewLine & "* The first name field cannot be left blank."
        If trim(household_members(MEMBER_LAST_NAME, 9)) = "" Then err_msg = err_msg & vbNewLine & "* The last name field cannot be left blank."
        If trim(household_members(MEMBER_GENDER, 9)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A gender option must be selected from the dropdown list."
        If trim(household_members(MEMBER_MARITAL_STATUS, 9)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A marital status option must be selected from the dropdown list."
        If trim(household_members(MEMBER_DOB, 9)) = "" OR (Not IsDate(trim(household_members(MEMBER_DOB, 9)))) Then err_msg = err_msg & vbNewLine & "* The date of birth field must be filled out in the MM/DD/YYYY format."
        If trim(household_members(MEMBER_SSN, 9)) <> "" AND Len(trim(household_members(MEMBER_SSN, 9))) <> 11 Then err_msg = err_msg & vbNewLine & "* The SSN must be in the format ###-##-####."
        If trim(household_members(MEMBER_RELATIONSHIP, 9)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A relationship option must be selected from the dropdown list."
      End If
    End If
    
    If dialog_count = 6 Then
      If trim(household_members(MEMBER_FIRST_NAME, 10)) = "" Then err_msg = err_msg & vbNewLine & "* The first name field cannot be left blank."
      If trim(household_members(MEMBER_LAST_NAME, 10)) = "" Then err_msg = err_msg & vbNewLine & "* The last name field cannot be left blank."
      If trim(household_members(MEMBER_GENDER, 10)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A gender option must be selected from the dropdown list."
      If trim(household_members(MEMBER_MARITAL_STATUS, 10)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A marital status option must be selected from the dropdown list."
      If trim(household_members(MEMBER_DOB, 10)) = "" OR (Not IsDate(trim(household_members(MEMBER_DOB, 10)))) Then err_msg = err_msg & vbNewLine & "* The date of birth field must be filled out in the MM/DD/YYYY format."
      If trim(household_members(MEMBER_SSN, 10)) <> "" AND Len(trim(household_members(MEMBER_SSN, 10))) <> 11 Then err_msg = err_msg & vbNewLine & "* The SSN must be in the format ###-##-####."
      If trim(household_members(MEMBER_RELATIONSHIP, 10)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A relationship option must be selected from the dropdown list."
      If member_count > 11 Then
        If trim(household_members(MEMBER_FIRST_NAME, 11)) = "" Then err_msg = err_msg & vbNewLine & "* The first name field cannot be left blank."
        If trim(household_members(MEMBER_LAST_NAME, 11)) = "" Then err_msg = err_msg & vbNewLine & "* The last name field cannot be left blank."
        If trim(household_members(MEMBER_GENDER, 11)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A gender option must be selected from the dropdown list."
        If trim(household_members(MEMBER_MARITAL_STATUS, 11)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A marital status option must be selected from the dropdown list."
        If trim(household_members(MEMBER_DOB, 11)) = "" OR (Not IsDate(trim(household_members(MEMBER_DOB, 11)))) Then err_msg = err_msg & vbNewLine & "* The date of birth field must be filled out in the MM/DD/YYYY format."
        If trim(household_members(MEMBER_SSN, 11)) <> "" AND Len(trim(household_members(MEMBER_SSN, 11))) <> 11 Then err_msg = err_msg & vbNewLine & "* The SSN must be in the format ###-##-####."
        If trim(household_members(MEMBER_RELATIONSHIP, 11)) = "Select one:" Then err_msg = err_msg & vbNewLine & "* A relationship option must be selected from the dropdown list."
      End If
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

Function GetMAXISRelationshipCode(relationship, gender)
    Dim returnCode
    Select Case relationship
        Case "applicant", "self"
            returnCode = "01"     
        Case "spouse"
            returnCode = "02"
        Case "child"
            returnCode = "03"
        Case "parent"
            returnCode = "04"
        Case "sibling", "brother or sister", "brother", "sister", "half-brother or half-sister", "half-brother", "half-sister"
            returnCode = "05"
        Case "step sibling", "step brother or sister", "step brother", "step sister"
            returnCode = "06"
        Case "step child", "step-child", "step son", "step daughter"
            returnCode = "08"
        Case "step parent"
            returnCode = "09"
        Case "aunt", "aunt or uncle"
            returnCode = "10"
            If gender = "male" Then
                returnCode = "11" 'uncle
            End If
        Case "uncle"
            returnCode = "11"
        Case "niece", "niece or nephew", "nephew or niece"
            returnCode = "12"
            If gender = "male" Then
                returnCode = "13" 'nephew
            End If
        Case "nephew"
            returnCode = "13"
        Case "cousin"
            returnCode = "14"
        Case "grandparent"
            returnCode = "15"
        Case "grandchild"
            returnCode = "16"
        Case "other relative"
            returnCode = "17"
        Case "legal guardian", "parent or guardian", "guardian"
            returnCode = "18"
        Case "live-in attendent"
            returnCode = "25"
        Case "unknown on caf i", "other"
            returnCode = "27"
        Case "child's parent"
            returnCode = "24" ' Not Related for now, but should find out from group what they use.
        Case "partner"
            returnCode = "24" ' Not Related for now, but should find out from group what they use.
        Case "roommate", "friend"
            returnCode = "24" ' Not Related
        Case Else
            returnCode = "27" ' Unknown/Not Indc On CAF I      
    End Select

	GetMAXISRelationshipCode = returnCode
		
End Function


'THE SCRIPT=================================================================================================================
EMConnect "" 'Connects to BlueZone

'Initial Dialog - Instructions
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 341, 70, "Process MNBenefits Application"
  Text 10, 5, 255, 20, "Script Purpose: This script performs a PERS search, APPLs the case using the MNBenefits XML file details, and then moves the case to PND2 status."
  GroupBox 10, 35, 255, 30, "Enter 10-digit confirmation number for XML file and then press 'Search'."
  Text 15, 50, 75, 10, "Confirmation Number:"
  EditBox 95, 45, 60, 15, confirmation_number
  ButtonGroup ButtonPressed
    PushButton 175, 45, 40, 15, "Search", search_button
    PushButton 270, 5, 65, 15, "Script Instructions", instructions_btn
EndDialog


DO
	DO
		err_msg = ""					'establishing value of variable, this is necessary for the Do...LOOP
		dialog Dialog1				'main dialog
		cancel_without_confirmation
    If ButtonPressed = instructions_btn Then
      'To do - update with instructions 
      Call open_URL_in_browser("https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/") 
      err_msg = "LOOP"
    End IF 

    If trim(confirmation_number) = "" OR len(confirmation_number) <> 10 OR Not IsNumeric(confirmation_number) then err_msg = err_msg & vbCr & "* You must enter the 10-digit confirmation number before pressing 'Search'."


    If trim(confirmation_number) <> "" and len(confirmation_number) = 10 and IsNumeric(confirmation_number) then
      If ButtonPressed = search_button Then
        If script_testing = false Then
          startTime = Timer

          folderPath = "T:\Eligibility Support\EA_ADAD\EA_ADAD_Common\CASE ASSIGNMENT\MNB_XML_files\"

          Set fso = CreateObject("Scripting.FileSystemObject")
          Set folder = fso.GetFolder(folderPath)
          XML_file_found = False
          file_count = 0

          For Each file In folder.Files
            If InStr(1, file.Name, "_" & confirmation_number & "_", vbTextCompare) > 0 Then
              ' msgbox "Found: " & file.Path
              XML_file_path = file.Path
              XML_file_found = True
              ' err_msg = "LOOP"
              Exit For
            End If
            file_count = file_count + 1
          Next
          If XML_file_found = False Then
            err_msg = err_msg & vbCr & "* The script was unable to locate a MNBenefits XML file with the application ID you provided. You must click the 'Select File' button and select the XML file or manually enter the file path in the field."
          End If
          'To do - delete after testing
          endTime = Timer
          duration = endTime - startTime
          ' msgbox "Search took " & duration & " seconds. It evaluated " & file_count & " files."
        Else
          startTime = Timer
          folderPath = "C:\Users\mari001\OneDrive - Hennepin County\Desktop\XML Files"

          Set fso = CreateObject("Scripting.FileSystemObject")
          Set folder = fso.GetFolder(folderPath)
          XML_file_found = False
          file_count = 0

          For Each file In folder.Files
            If InStr(1, file.Name, "_" & confirmation_number & "_", vbTextCompare) > 0 Then
              ' msgbox "Found: " & file.Path
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
          ' msgbox "Search took " & duration & " seconds. It evaluated " & file_count & " files."
        End If
      End If
    End If
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
Dim household_members()
Const MEMBER_FIRST_NAME     = 0
Const MEMBER_LAST_NAME      = 1
Const MEMBER_DOB            = 2
Const MEMBER_SSN            = 3
Const MEMBER_RELATIONSHIP   = 4
Const MEMBER_MARITAL_STATUS = 5
Const MEMBER_CITIZENSHIP    = 6
Const MEMBER_PMI            = 7
Const MEMBER_EXISTS_MAXIS   = 8
Const APPL_CASE_NUMBER      = 9
Const MEMBER_GENDER         = 10

ReDim household_members(MEMBER_GENDER, member_count)   'Redimmed to the size of the last constant

Dim objHouseholdMemberNode, objHouseholdMemberNodes
Set objHouseholdMemberNode = xmlDoc.selectSingleNode("//ns4:HouseholdInfo")
Set objHouseholdMemberNodes = objHouseholdMemberNode.selectNodes("ns4:HouseholdMember")

Dim objMemberNode, objRoot
Dim objFirstNameNode, objLastNameNode, objSSNNode, objDOBNode, objRelationshipNode, objMaritalStatusNode, objGenderNode

For Each objMemberNode In objHouseholdMemberNodes
  Set objFirstNameNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:Person/ns4:FirstName")
  Set objLastNameNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:Person/ns4:LastName")
  Set objSSNNode = objMemberNode.selectSingleNode("ns4:CitizenshipInfo/ns4:SSNInfo/ns4:SSN")
  Set objDOBNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:DOB")
  Set objRelationshipNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:Relationship") 
  Set objMaritalStatusNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:MaritalStatus")
  Set objGenderNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:Gender")

  If Not objFirstNameNode Is Nothing Then household_members(MEMBER_FIRST_NAME, member_count) = objFirstNameNode.Text
  If Not objLastNameNode Is Nothing Then household_members(MEMBER_LAST_NAME, member_count) = objLastNameNode.Text
  If Not objDOBNode Is Nothing Then household_members(MEMBER_DOB, member_count) = objDOBNode.Text
  If Not objSSNNode Is Nothing Then 
    household_members(MEMBER_SSN, member_count) = objSSNNode.Text
    If trim(objSSNNode.Text) = "" Then 
      household_members(MEMBER_CITIZENSHIP, member_count) = "Select one:"
    Else
      household_members(MEMBER_CITIZENSHIP, member_count) = "Yes"
    End If
  End If
  If Not objRelationshipNode Is Nothing Then household_members(MEMBER_RELATIONSHIP, member_count) = objRelationshipNode.Text
  If Not objMaritalStatusNode Is Nothing Then household_members(MEMBER_MARITAL_STATUS, member_count) = objMaritalStatusNode.Text
  If Not objGenderNode Is Nothing Then household_members(MEMBER_GENDER, member_count) = objGenderNode.Text

  If household_members(MEMBER_FIRST_NAME, member_count) = "" And household_members(MEMBER_LAST_NAME, member_count) = "" Then Exit For

  Dim memberNumber
  If member_count < 9 Then
    memberNumber = "0" & member_count + 1
  Else
    memberNumber = member_count + 1
  End If

  member_count = member_count + 1
  ReDim Preserve household_members(MEMBER_GENDER, member_count)
Next

'Remove the empty household member from the array
' member_count = member_count - 1
ReDim Preserve household_members(MEMBER_GENDER, member_count - 1)

'Gather application date and application ID 
Dim formatted_app_date, objApplicationDate, applicationDate, applicationMonth, applicationDay, applicationYear
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

formatted_app_date = applicationMonth & "/" & applicationDay & "/" & applicationYear
MAXIS_footer_month = applicationMonth
MAXIS_footer_year = Mid(applicationYear, 3, 2)

Dim objApplicationId
' Application Id
Set objApplicationId = xmlDoc.selectSingleNode("//CONTENT/ap:Bytes/io4:ApplicationID")
If Not objApplicationId Is Nothing Then
  applicationId = objApplicationId.Text
End If

'Validate the provided application ID against the application ID in the XML file
If confirmation_number_checkbox = 1 Then
  If applicationId <> confirmation_number Then script_end_procedure_with_error_report("The application ID provided to locate the MNBenefits XML file does not match the application ID in the XML file. Please try running the script again.")
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
'To do - determine if best practice to release the XML doc now, or wait until update to avoid need to reopen
' Set xmlDoc = Nothing

dialog_member_count = 0

'XML File Confirmation Dialog
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 285, 245, "Verify MNBenefits XML Details - Household Members"
  Text 5, 5, 275, 20, "Please verify that the correct XML file has been selected. If you need to change the XML file, please press the 'Reselect XML' button below."  
  GroupBox 10, 35, 270, 155, "MNBenefits XML File Details"
  Text 15, 45, 50, 10, "Application ID:"
  Text 100, 45, 50, 10, confirmation_number
  Text 15, 55, 60, 10, "Application Date:"
  Text 100, 55, 60, 10, formatted_app_date
  Text 15, 65, 75, 10, "Household Member 1:"
  Text 100, 65, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"  
  dialog_member_count = dialog_member_count + 1 
  If member_count > 1 Then
    Text 15, 75, 75, 10, "Household Member 2:"
    Text 100, 75, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 2 Then
    Text 15, 85, 75, 10, "Household Member 3:"
    Text 100, 85, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 3 Then
    Text 15, 95, 75, 10, "Household Member 4:"
    Text 100, 95, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 4 Then
    Text 15, 105, 75, 10, "Household Member 5:"
    Text 100, 105, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 5 Then
    Text 15, 115, 75, 10, "Household Member 6:"
    Text 100, 115, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 6 Then
    Text 15, 125, 75, 10, "Household Member 7:"
    Text 100, 125, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 7 Then
    Text 15, 135, 75, 10, "Household Member 8:"
    Text 100, 135, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 8 Then
    Text 15, 145, 75, 10, "Household Member 9:"
    Text 100, 145, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 9 Then
    Text 15, 155, 80, 10, "Household Member 10:"
    Text 100, 155, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 10 Then
    Text 15, 165, 80, 10, "Household Member 11:"
    Text 100, 165, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 11 Then
    Text 15, 175, 80, 10, "Household Member 12:"
    Text 100, 175, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"  
    dialog_member_count = dialog_member_count + 1
  End If
  ButtonGroup ButtonPressed
    PushButton 235, 225, 45, 15, "Continue", continue_button
    PushButton 10, 225, 50, 15, "Reselect XML", reselect_xml_button
    'To do - add functionality for reselect xml button
EndDialog

DO
  dialog Dialog1
  cancel_without_confirmation
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
'To do - put dialog above and XML selection dialog into functions so that they can be called as a loop so user can move back and forth if needed

'To do - add handling for cases where not address is provided
'To do - add county of residence
'To do - convert county of residence to county code?
'Dialog to confirm Application and Address Information

'To do - move this to end once PERS search completed
' BeginDialog Dialog1, 0, 0, 256, 265, "Process MNBenefits Application"
'   Text 10, 5, 240, 20, "Please verify the application and address details pulled from the XML file below. Make updates as needed."
'   Text 15, 30, 150, 10, "Adjust date to correct business day, if needed"
'   Text 15, 45, 60, 10, "Application Date: "
'   EditBox 80, 40, 40, 15, formatted_app_date
'   GroupBox 10, 60, 175, 105, "Household Address"
'   Text 15, 75, 35, 10, "Address:"
'   EditBox 70, 70, 100, 15, household_address
'   Text 15, 90, 25, 10, "City:"
'   EditBox 70, 85, 100, 15, household_city
'   Text 15, 105, 30, 10, "State:"
'   EditBox 70, 100, 20, 15, household_state
'   Text 15, 120, 20, 10, "Zip:"
'   EditBox 70, 115, 30, 15, household_zip
'   Text 15, 135, 55, 10, "Phone number:"
'   EditBox 70, 130, 100, 15, household_phone_number
'   GroupBox 10, 165, 175, 75, "Mailing Address"
'   Text 15, 180, 35, 10, "Address:"
'   EditBox 70, 175, 100, 15, mailing_address
'   Text 15, 195, 25, 10, "City:"
'   EditBox 70, 190, 100, 15, mailing_city
'   Text 15, 210, 30, 10, "State:"
'   EditBox 70, 205, 20, 15, mailing_state
'   Text 15, 225, 20, 10, "Zip:"
'   EditBox 70, 220, 30, 15, mailing_zip
'   Text 15, 150, 30, 10, "County:"
'   EditBox 70, 145, 100, 15, household_county
'   ButtonGroup ButtonPressed
'     PushButton 200, 245, 50, 15, "Confirm", confirm_address_button
' EndDialog

' DO
' 	DO
' 		err_msg = ""					'establishing value of variable, this is necessary for the Do...LOOP
' 		dialog Dialog1				'main dialog
' 		cancel_without_confirmation
'     If ButtonPressed = file_selection_button then 
'       call file_selection_system_dialog(XML_file_path, ".xml")
'       err_msg = "LOOP"
'     End If
'     If trim(formatted_app_date) = "" OR IsDate(formatted_app_date) = False OR Len(trim(formatted_app_date)) <> 10 then err_msg = err_msg & vbCr & "* You must enter the application date in the format MM/DD/YYYY."
'     If trim(household_address) = "" Then err_msg = err_msg & vbCr & "* The household address field cannot be blank."
'     If trim(household_city) = "" Then err_msg = err_msg & vbCr & "* The city field cannot be blank."
'     If trim(household_state) = "" Then err_msg = err_msg & vbCr & "* The state field cannot be blank."
'     If trim(household_zip) = "" Then err_msg = err_msg & vbCr & "* The zip code field cannot be blank."
'     'To do - confirm if phone number is required
'     ' If trim(household_phone_number) = "" Then then err_msg = err_msg & vbCr & "* The household address field cannot be blank."
'     If trim(mailing_address) = "" Then err_msg = err_msg & vbCr & "* The mailing address field cannot be blank."
'     If trim(mailing_city) = "" Then err_msg = err_msg & vbCr & "* The mailing address city field cannot be blank."
'     If trim(mailing_state) = "" Then err_msg = err_msg & vbCr & "* The mailing address state field cannot be blank."
'     If trim(mailing_zip) = "" Then err_msg = err_msg & vbCr & "* The mailing address zip code field cannot be blank."

' 		If err_msg <> "" and err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
' 	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
' 	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
' Loop until are_we_passworded_out = false					'loops until user passwords back in

' 'Start at the first dialog and start dialog loop
' dialog_count = 1
' hh_memb_dialog_loop = "Active"
' Call determine_member_dialogs_display()

' Do
'   Do
'     Do
'       Dialog1 = "" 'Blanking out previous dialog detail

'       Call dialog_selection(dialog_count)

'       'Blank out variables on each new dialog
'       err_msg = ""

'       dialog Dialog1 					'Calling a dialog without an assigned variable will call the most recently defined dialog
'       cancel_confirmation
'       Call dialog_specific_error_handling()	'function for error handling of main dialog of forms
'       Call button_movement()				'function to move throughout the dialogs
'     Loop until err_msg = ""
'     If hh_memb_dialog_loop = "Completed" Then
'       Dialog1 = "" 'Blanking out previous dialog detail
'       Call confirm_xml_update_dialog()
'       dialog Dialog1 					'Calling a dialog without an assigned variable will call the most recently defined dialog
'       cancel_without_confirmation
'       If ButtonPressed = back_button Then 
'         dialog_count = 1
'         hh_memb_dialog_loop = "Active"
'       End If
'     End If
'   Loop until hh_memb_dialog_loop = "Completed"
'   CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
' Loop until are_we_passworded_out = false					'loops until user passwords back in

' member_array_index = 0

' For Each objMemberNode In objHouseholdMemberNodes
'   Set objFirstNameNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:Person/ns4:FirstName")
'   Set objLastNameNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:Person/ns4:LastName")
'   Set objSSNNode = objMemberNode.selectSingleNode("ns4:CitizenshipInfo/ns4:SSNInfo/ns4:SSN")
'   Set objDOBNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:DOB")
'   Set objRelationshipNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:Relationship") 
'   Set objMaritalStatusNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:MaritalStatus")
'   Set objCitizenshipNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:CitizenshipInfo")
'   Set objGenderNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:Gender")

'   If household_members(MEMBER_FIRST_NAME, member_array_index) <> "" Then objFirstNameNode.Text = household_members(MEMBER_FIRST_NAME, member_array_index)

'   If household_members(MEMBER_LAST_NAME, member_array_index) <> "" Then objLastNameNode.Text = household_members(MEMBER_LAST_NAME, member_array_index)

'   If household_members(MEMBER_DOB, member_array_index) <> "" Then objDOBNode.Text = household_members(MEMBER_DOB, member_array_index)

'   If household_members(MEMBER_SSN, member_array_index) <> "" Then objSSNNode.Text = household_members(MEMBER_SSN, member_array_index)

'   If household_members(MEMBER_RELATIONSHIP, member_array_index) <> "" Then objRelationshipNode.Text = household_members(MEMBER_RELATIONSHIP, member_array_index)

'   If household_members(MEMBER_MARITAL_STATUS, member_array_index) <> "" Then objMaritalStatusNode.Text = household_members(MEMBER_MARITAL_STATUS, member_array_index)

'   If household_members(MEMBER_GENDER, member_array_index) <> "" Then objGenderNode.Text = household_members(MEMBER_GENDER, member_array_index)

'   If household_members(MEMBER_FIRST_NAME, member_array_index) = "" Then Exit For

'   member_array_index = member_array_index + 1
' Next

' 'Replace the application date
' 'Format 2025-11-26

' current_XML_app_date = left(objApplicationDate.Text, 10)
' 'Convert the formatted_app_date to XML format and replace
' updated_XML_app_date = right(formatted_app_date, 4) & "-" & left(formatted_app_date, 2) & "-" & mid(formatted_app_date, 4, 2)
' objApplicationDate.Text       = replace(objApplicationDate.Text, current_XML_app_date, updated_XML_app_date)
' objHouseholdAddress.Text      = household_address
' objHouseholdCity.Text         = household_city
' objHouseholdState.Text        = household_state
' objHouseholdZip.Text          = household_zip
' objPhoneNumber.Text           = household_phone_number
' objCounty.Text                = household_county
' objMailingAddress.Text        = mailing_address
' objMailingCity.Text           = mailing_city
' objMailingState.Text          = mailing_state
' objMailingZip.Text            = mailing_zip

' ' Save the updated XML to a file
' ' To do - update with actual file path once done testing
' xmlDoc.Save "C:\Users\mari001\OneDrive - Hennepin County\Desktop\New XML Files\new xml file success.xml"
' xmlDoc.Save replace(XML_file_path, confirmation_number, confirmation_number & "_" & "processed")

' Save the XML document with 'processed' in file name
' xmlDoc.Save replace(XML_file_path, confirmation_number, confirmation_number & "_" & "processed")

' ' To do - uncomment after testing, this is where file is saved and moved

' On Error Resume Next

' ' Attempt to save the XML document
' Dim XML_file_path_processed
' XML_file_path_processed = Replace(XML_file_path, confirmation_number, confirmation_number & "_processed")
' xmlDoc.Save XML_file_path_processed

' ' Check for errors
' If Err.Number <> 0 Then
'   WScript.Echo "Error saving file: " & Err.Description
'   ' Optional: log the error or take corrective action
'   script_end_procedure_with_error_report("Script failed to save the processed XML file. The script will now end.")
' Else
'   msgbox "Success!"
' End If

' On Error GoTo 0 ' Reset error handling

' Set fso = CreateObject("Scripting.FileSystemObject")

' If fso.FileExists(XML_file_path) Then
'   fso.DeleteFile XML_file_path
' Else
'   script_end_procedure_with_error_report("Script failed to delete XML file.")
' End If

' ' Clean up
' Set objMemberNode           = Nothing
' Set objFirstNameNode        = Nothing
' Set objLastNameNode         = Nothing
' Set objDOBNode              = Nothing
' Set objSSNNode              = Nothing
' Set objRelationshipNode     = Nothing
' Set objMaritalStatusNode    = Nothing
' Set objGenderNode           = Nothing
' Set objCitizenshipNode      = Nothing
' Set objHouseholdMemberNodes = Nothing
' Set objHouseholdMemberNode  = Nothing
' Set xmlDoc                  = Nothing

' ' MsgBox "XML file saved and updated successfully from array."

' Complete PERS search for every member listed on the application
'Navigate to PERS
Call navigate_to_MAXIS_screen("PERS", "")

'Validation to confirm PERS search reached
EmReadScreen PERS_panel_check, 4, 2, 47
If PERS_panel_check <> "PERS" Then 
  Call back_to_SELF
End If
EmReadScreen PERS_panel_check, 4, 2, 47
If PERS_panel_check <> "PERS" Then 
  script_end_procedure_with_error_report("Script was unable to navigate to PERS search. Script will now end")
End If

' Script will search for person using all details provided EXCEPT SSN
'   --> Script reads through all results on first page until end reached or match found
'   --> Script matches based on the first and last name and then DOB (if provided) and SSN (if provided)

' msgbox "Ubound(household_members, 2) " & Ubound(household_members, 2) 
' msgbox "first name of last member " & household_members(MEMBER_FIRST_NAME, Ubound(household_members, 2))
' msgbox "last name of last member " & household_members(MEMBER_LAST_NAME, Ubound(household_members, 2))

For member = 0 to Ubound(household_members, 2)
  'Setting variables for search
  ssn_match_found = False
  PERS_search_results_string = ""
  MTCH_row = 8
  SSN_search = True
  PERS_second_search = False
  ' NOTES: DOB can be blank ('          '), SSN can be blank ('   -  -    ')

  Do 
    'Conduct initial search with all details provided EXCEPT SSN
    EmWriteScreen household_members(MEMBER_LAST_NAME, member), 4, 36
    EmWriteScreen household_members(MEMBER_FIRST_NAME, member), 10, 36
    EmWriteScreen Left(household_members(MEMBER_DOB, member), 2), 11, 53
    EmWriteScreen Mid(household_members(MEMBER_DOB, member), 4, 2), 11, 56
    EmWriteScreen Mid(household_members(MEMBER_DOB, member), 7, 4), 11, 59
    EmWriteScreen Left(household_members(MEMBER_GENDER, member), 1), 11, 36
    'To do - ssn search
    If PERS_second_search = True Then
      EmWriteScreen Left(household_members(MEMBER_SSN, member), 3), 14, 36
      EmWriteScreen Mid(household_members(MEMBER_SSN, member), 5, 2), 14, 40
      EmWriteScreen right(household_members(MEMBER_SSN, member), 4), 14, 43
    End If
    transmit

    Do
      'Read the data from the corresponding MTCH row
      match_rating = 0

      EmReadScreen SSN_MTCH_panel, 11, MTCH_row, 7
      SSN_MTCH_panel = trim(SSN_MTCH_panel)
      If SSN_MTCH_panel = household_members(MEMBER_SSN, member) Then match_rating = match_rating + .2

      EmReadScreen last_name_MTCH_panel, 20, MTCH_row, 21
      last_name_MTCH_panel = trim(last_name_MTCH_panel)
      If last_name_MTCH_panel = UCase(household_members(MEMBER_LAST_NAME, member)) Then match_rating = match_rating + .2
      
      EmReadScreen first_name_MTCH_panel, 12, MTCH_row, 42
      first_name_MTCH_panel = trim(first_name_MTCH_panel)
      If first_name_MTCH_panel = UCase(household_members(MEMBER_FIRST_NAME, member)) Then match_rating = match_rating + .1
      
      EmReadScreen gender_MTCH_panel, 1, MTCH_row, 58
      gender_MTCH_panel = trim(gender_MTCH_panel)
      ' If gender_MTCH_panel = household_members(MEMBER_GENDER, member) Then match_rating = match_rating + .1
      'To do - does it make sense to use gender to match?

      EmReadScreen dob_MTCH_panel, 10, MTCH_row, 60
      dob_MTCH_panel = trim(dob_MTCH_panel)
      If dob_MTCH_panel = replace(household_members(MEMBER_DOB, member), "/", "-") Then match_rating = match_rating + .2
        
      EmReadScreen pmi_MTCH_panel, 10, MTCH_row, 71
      pmi_MTCH_panel = trim(pmi_MTCH_panel)

      If match_rating > .2 Then         
        'Validate the PMI number. Script will only display a potential match if the PMI number exists
        CALL write_value_and_transmit("X", MTCH_row, 5)
        EMReadScreen PMI_exists_check, 24, 24, 2
        If Instr(PMI_exists_check, "PMI NBR ASSIGNED") = 0 Then
          If Instr(PERS_search_results_string, first_name_MTCH_panel & " " & last_name_MTCH_panel & " " & "(DOB: " & dob_MTCH_panel & "; SSN: " & SSN_MTCH_panel & "; PMI: " & pmi_MTCH_panel & "; Gender: " & gender_MTCH_panel & ")") = 0 Then 

            ' Read all of the case numbers and add to array
            DSPL_row = 10
            DSPL_case_number_string = "*"
            Do 
              EmReadScreen DSPL_case_number, 12, DSPL_row, 6
              DSPL_case_number = trim(DSPL_case_number)
              If Instr(DSPL_case_number_string, DSPL_case_number) = 0 Then DSPL_case_number_string = DSPL_case_number_string & DSPL_case_number & "*"  
              DSPL_row = DSPL_row + 1
              EmReadScreen blank_case_number_check, 12, DSPL_row, 6
              If trim(blank_case_number_check) = "" then Exit Do
              If DSPL_row = 20 then 
                EMReadScreen more_check, 9, 20, 3
                more_check = trim(more_check)
                If more_check = "" or more_check = "More: -" Then Exit Do
                If more_check = "More: +" OR more_check = "More: +/-" Then 
                  PF8
                  DSPL_row = 10
                End If
              End If
            Loop
            PERS_search_results_string = PERS_search_results_string & first_name_MTCH_panel & " " & last_name_MTCH_panel & " " & "(DOB: " & dob_MTCH_panel & "; SSN: " & SSN_MTCH_panel & "; PMI: " & pmi_MTCH_panel & "; Gender: " & gender_MTCH_panel & ")" & DSPL_case_number_string & "#"
          End If
          PF3   'Back to MTCH panel
        Else
          'Clear the X
          EMWriteScreen "_", MTCH_row, 5
        End If
      End If

      MTCH_row = MTCH_row + 1
      If MTCH_row = 17 then 
        Exit Do
      End If
    Loop
    ' End If

    'If we made it through second search then we need to exit loop
    If PERS_second_search = True Then Exit Do

    'If no SSN match found, then we will search again with SSN
    If PERS_second_search = False Then
      'Conduct a second search using SSN for search criteria
      PERS_second_search = True
      'Setting variables for search
      MTCH_row = 8
      SSN_search = True

      Call back_to_SELF
      Call navigate_to_MAXIS_screen("PERS", "")

      'Validation to confirm PERS search reached
      EmReadScreen PERS_panel_check, 4, 2, 47
      If PERS_panel_check <> "PERS" Then 
        Call back_to_SELF
      End If
      EmReadScreen PERS_panel_check, 4, 2, 47
      If PERS_panel_check <> "PERS" Then 
        script_end_procedure_with_error_report("Script was unable to navigate to PERS search. Script will now end")
      End If
    End If
  LOOP

  'Need to determine if any matches found and if so, format for the dialog
  PERS_match_found = false
  If Instr(PERS_search_results_string, "#") Then PERS_match_found = True
  
  If PERS_match_found Then 
    'Remove trailing * before splitting  
    PERS_search_results_string = Left(PERS_search_results_string, Len(PERS_search_results_string) - 1)
    PERS_search_results_string_array = split(PERS_search_results_string, "#")
  End If
  
  PERS_search_criteria = household_members(MEMBER_FIRST_NAME, member) & " " & household_members(MEMBER_LAST_NAME, member) & " (DOB: " & household_members(MEMBER_DOB, member) & "; SSN: " & household_members(MEMBER_SSN, member) & "; Gender: " & household_members(MEMBER_GENDER, member) & ")"
  
  checkbox_y = 85
  groupbox_height = 35
  If PERS_match_found Then groupbox_height = groupbox_height + (Ubound(PERS_search_results_string_array) * 10)
  dialog_height = 130
  If PERS_match_found Then dialog_height = dialog_height + (Ubound(PERS_search_results_string_array) * 10)

  'Call dialog to display matches (if found) and option to indicate if no match found 
  Dialog1 = "" 'Blanking out previous dialog detail
  BeginDialog Dialog1, 0, 0, 350, dialog_height, "PERS Search Results"
    Text 5, 5, 330, 10, "Please review the potential matches found, if any, and select the best applicable checkbox."
    GroupBox 5, 30, 340, 30, "Household Member Details from XML File"
    Text 15, 45, 325, 10, PERS_search_criteria
    GroupBox 5, 75, 340, groupbox_height, "Review potential PERS matches (select ONE option):"
    If PERS_match_found = False Then
      Text 15, checkbox_y, 325, 10, "No potential matches found. You must complete a manual search. Press 'OK' to continue"
    ElseIf PERS_match_found = True Then
      CheckBox 15, checkbox_y, 325, 10, mid(PERS_search_results_string_array(0), 1, instr(PERS_search_results_string_array(0), "*") - 1), pers_search_results_0
      If UBound(PERS_search_results_string_array) > 0 Then
        ' If PERS_search_results_string_array(1) <> "" Then
          checkbox_y = checkbox_y + 10
          CheckBox 15, checkbox_y, 325, 10, mid(PERS_search_results_string_array(1), 1, instr(PERS_search_results_string_array(1), "*") - 1), pers_search_results_1
        ' End If 
      End If
      If UBound(PERS_search_results_string_array) > 1 Then
        ' If PERS_search_results_string_array(2) <> "" Then
          checkbox_y = checkbox_y + 10
          CheckBox 15, checkbox_y, 325, 10, mid(PERS_search_results_string_array(2), 1, instr(PERS_search_results_string_array(2), "*") - 1), pers_search_results_2
        ' End If 
      End If
      If UBound(PERS_search_results_string_array) > 2 Then
        ' If PERS_search_results_string_array(3) <> "" Then
          checkbox_y = checkbox_y + 10
          CheckBox 15, checkbox_y, 325, 10, mid(PERS_search_results_string_array(3), 1, instr(PERS_search_results_string_array(3), "*") - 1), pers_search_results_3
        ' End If 
      End If
      If UBound(PERS_search_results_string_array) > 3 Then
        ' If PERS_search_results_string_array(4) <> "" Then
          checkbox_y = checkbox_y + 10
          CheckBox 15, checkbox_y, 325, 10, mid(PERS_search_results_string_array(4), 1, instr(PERS_search_results_string_array(4), "*") - 1), pers_search_results_4
        ' End If 
      End If
      If UBound(PERS_search_results_string_array) > 4 Then
        ' If PERS_search_results_string_array(5) <> "" Then
          checkbox_y = checkbox_y + 10
          CheckBox 15, checkbox_y, 325, 10, mid(PERS_search_results_string_array(5), 1, instr(PERS_search_results_string_array(5), "*") - 1), pers_search_results_5
        ' End If 
      End If
      If UBound(PERS_search_results_string_array) > 5 Then
        ' If PERS_search_results_string_array(6) <> "" Then
          checkbox_y = checkbox_y + 10
          CheckBox 15, checkbox_y, 325, 10, mid(PERS_search_results_string_array(6), 1, instr(PERS_search_results_string_array(6), "*") - 1), pers_search_results_6
        ' End If 
      End If
      If UBound(PERS_search_results_string_array) > 6 Then
        ' If PERS_search_results_string_array(7) <> "" Then
          checkbox_y = checkbox_y + 10
          CheckBox 15, checkbox_y, 325, 10, mid(PERS_search_results_string_array(7), 1, instr(PERS_search_results_string_array(7), "*") - 1), pers_search_results_7
        ' End If 
      End If
      If UBound(PERS_search_results_string_array) > 7 Then
        ' If PERS_search_results_string_array(8) <> "" Then
          checkbox_y = checkbox_y + 10
          CheckBox 15, checkbox_y, 325, 10, mid(PERS_search_results_string_array(8), 1, instr(PERS_search_results_string_array(8), "*") - 1), pers_search_results_8
        ' End If 
      End If
      If UBound(PERS_search_results_string_array) > 8 Then
        ' If PERS_search_results_string_array(9) <> "" Then
          checkbox_y = checkbox_y + 10
          CheckBox 15, checkbox_y, 325, 10, mid(PERS_search_results_string_array(9), 1, instr(PERS_search_results_string_array(9), "*") - 1), pers_search_results_9
        ' End If 
      End If
      CheckBox 15, checkbox_y + 10, 325, 10, "None of these matches are correct. I will complete a manual search.", no_match_search_manually
    End If
    ButtonGroup ButtonPressed
      OkButton 245, checkbox_y + 30, 50, 15
      CancelButton 295, checkbox_y + 30, 50, 15
  EndDialog

  DO
    DO
      err_msg = ""					'establishing value of variable, this is necessary for the Do...LOOP
      dialog Dialog1				'main dialog
      cancel_without_confirmation
      'to do - add error handling
      If PERS_match_found = True Then
        If pers_search_results_0 + pers_search_results_1 + pers_search_results_2 + pers_search_results_3 + pers_search_results_4 + pers_search_results_5 + pers_search_results_6 + pers_search_results_7 + pers_search_results_8 + pers_search_results_9 + no_match_search_manually > 1 Then err_msg = err_msg & vbNewLine & "* You can only check one checkbox for the PERS results section."
        If pers_search_results_0 + pers_search_results_1 + pers_search_results_2 + pers_search_results_3 + pers_search_results_4 + pers_search_results_5 + pers_search_results_6 + pers_search_results_7 + pers_search_results_8 + pers_search_results_9 + no_match_search_manually = 0 Then err_msg = err_msg & vbNewLine & "* You must check one of the checkboxes in the PERS results section."
      End If
      If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
  Loop until are_we_passworded_out = false					'loops until user passwords back in

  'Determine which option selected
  If PERS_match_found = False Or no_match_search_manually = 1 Then
    'Call dialog for worker to identify the PMI or indicate if a new person

    'Enter the PMI number. If script is on the applicant then worker will need to select whether case already exists or if a new one is needed
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 321, 180, "Manual PERS Search"
      Text 5, 5, 300, 25, "You indicated that none of the matches identified by the script were correct. Please complete a manual search for the household member identified below and provide the corresponding PMI and MAXIS case number, if applicable."
      GroupBox 5, 40, 310, 30, "Household Member Details from XML File"
      Text 15, 55, 295, 10, PERS_search_criteria
      GroupBox 5, 75, 310, 65, "Check one option below:"
      CheckBox 10, 90, 60, 10, "I found a match", manual_match_found_checkbox
      Text 20, 105, 20, 10, "PMI:"
      EditBox 40, 100, 50, 15, PMI_number
      CheckBox 10, 125, 185, 10, "No match found. This person does not exist in MAXIS.", manual_no_match_found_checkbox
      ButtonGroup ButtonPressed
        OkButton 215, 160, 50, 15
        CancelButton 265, 160, 50, 15
    EndDialog

    DO
      DO
        err_msg = ""					'establishing value of variable, this is necessary for the Do...LOOP
        dialog Dialog1				'main dialog
        cancel_without_confirmation
        'to do - add error handling
        If manual_match_found_checkbox = 1 Then
          ' If trim(MAXIS_case_number) = "" Then err_msg = err_msg & vbNewLine & "* You must enter the case number."
          If trim(PMI_number) = "" Then err_msg = err_msg & vbNewLine & "* You must enter the PMI number."
        End If
        If manual_match_found_checkbox + manual_no_match_found_checkbox = 2 then err_msg = err_msg & vbNewLine & "* You can only check one checkbox."
        If manual_match_found_checkbox + manual_no_match_found_checkbox = 0 then err_msg = err_msg & vbNewLine & "* You must check one of the checkboxes."
        If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
      LOOP UNTIL err_msg = ""									'loops until all errors are resolved
      CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    'If manual match found, need to verify PMI number entered correctly. If this is applicant, worker needs to select case number or indicate that this is a new case
    If manual_match_found_checkbox = 1 Then
      'Conduct a PERS search using the PMI number to find match and display for worker
      pmi_number = trim(pmi_number)

      'Navigate to PERS
      Call navigate_to_MAXIS_screen("PERS", "")

      'Validation to confirm PERS search reached
      EmReadScreen PERS_panel_check, 4, 2, 47
      If PERS_panel_check <> "PERS" Then 
        Call back_to_SELF
      End If
      EmReadScreen PERS_panel_check, 4, 2, 47
      If PERS_panel_check <> "PERS" Then 
        script_end_procedure_with_error_report("Script was unable to navigate to PERS search. Script will now end")
      End If

      EMWriteScreen pmi_number, 15, 36
      transmit

      'To do - add handling for no PMI match when trying to search

      'Reset the pers search result string 
      selected_PERS_search_results_string = ""

      'Reset the MTCH_row
      MTCH_row = 8

      'Open the PMI match to gather details
      'Read the data from the corresponding MTCH row
      EmReadScreen SSN_MTCH_panel, 11, MTCH_row, 7
      SSN_MTCH_panel = trim(SSN_MTCH_panel)
      
      EmReadScreen last_name_MTCH_panel, 20, MTCH_row, 21
      last_name_MTCH_panel = trim(last_name_MTCH_panel)
      
      EmReadScreen first_name_MTCH_panel, 12, MTCH_row, 42
      first_name_MTCH_panel = trim(first_name_MTCH_panel)
      
      EmReadScreen gender_MTCH_panel, 1, MTCH_row, 58
      gender_MTCH_panel = trim(gender_MTCH_panel)
      
      EmReadScreen dob_MTCH_panel, 10, MTCH_row, 60
      dob_MTCH_panel = trim(dob_MTCH_panel)
      
      EmReadScreen pmi_MTCH_panel, 10, MTCH_row, 71
      pmi_MTCH_panel = trim(pmi_MTCH_panel)
      ' msgbox "first_name_MTCH_panel > " & first_name_MTCH_panel & vbcr & "last_name_MTCH_panel >" & last_name_MTCH_panel & vbcr & "gender_MTCH_panel >" & gender_MTCH_panel & vbcr & "dob_MTCH_panel >" & dob_MTCH_panel & vbcr & "SSN_MTCH_panel >" & vbcr & "pmi_MTCH_panel >" & pmi_MTCH_panel & vbcr & "SSN_MTCH_panel >" & SSN_MTCH_panel

      'Validate the PMI number. Script will only display a potential match if the PMI number exists
      CALL write_value_and_transmit("X", MTCH_row, 5)
      EMReadScreen PMI_exists_check, 24, 24, 2
      If Instr(PMI_exists_check, "PMI NBR ASSIGNED") = 0 Then
        ' Read all of the case numbers and add to array
        DSPL_row = 10
        DSPL_case_number_string = "*"
        Do 
          EmReadScreen DSPL_case_number, 12, DSPL_row, 6
          DSPL_case_number = trim(DSPL_case_number)
          If Instr(DSPL_case_number_string, DSPL_case_number) = 0 Then DSPL_case_number_string = DSPL_case_number_string & DSPL_case_number & "*"  
          DSPL_row = DSPL_row + 1
          EmReadScreen blank_case_number_check, 12, DSPL_row, 6
          If trim(blank_case_number_check) = "" then Exit Do
          If DSPL_row = 20 then 
            EMReadScreen more_check, 9, 20, 3
            more_check = trim(more_check)
            If more_check = "" or more_check = "More: -" Then Exit Do
            If more_check = "More: +" OR more_check = "More: +/-" Then 
              PF8
              DSPL_row = 10
            End If
          End If
        Loop
        selected_PERS_search_results_string = selected_PERS_search_results_string & first_name_MTCH_panel & " " & last_name_MTCH_panel & " " & "(DOB: " & dob_MTCH_panel & "; SSN: " & SSN_MTCH_panel & "; PMI: " & pmi_MTCH_panel & "; Gender: " & gender_MTCH_panel & ")" & DSPL_case_number_string & "#"
        ' msgbox "selected_PERS_search_results_string >" & selected_PERS_search_results_string
        PF3   'Back to MTCH panel
      Else
        'Clear the X
        script_end_procedure_with_error_report("1573 The script failed to open the PMI match. The PMI does not exist in MAXIS.")
        EMWriteScreen "_", MTCH_row, 5
      End If

      ' Set PMI number in array and set exists to True
      household_members(MEMBER_PMI, member) = pmi_number
      household_members(MEMBER_EXISTS_MAXIS, member) = True

      'Pull out the PMI number from the PERS_search_results_string
      PMI_number = mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, "PMI: ") + 5, instr(selected_PERS_search_results_string, "; Gender: ") - instr(selected_PERS_search_results_string, "PMI: ") - 5)
  
      'Pull out the case numbers identified for the PMI # and remove the leading and trailing *
      selected_PERS_search_case_numbers = mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, "*"), len(selected_PERS_search_results_string) - instr(selected_PERS_search_results_string, "*"))
      selected_PERS_search_case_numbers = right(selected_PERS_search_case_numbers, len(selected_PERS_search_case_numbers) - 1)
      selected_PERS_search_case_numbers = left(selected_PERS_search_case_numbers, len(selected_PERS_search_case_numbers) - 1)

      selected_PERS_search_case_numbers_array = split(selected_PERS_search_case_numbers, "*")

      If Ubound(selected_PERS_search_case_numbers_array) = 0 then case_selection_list_options = "Select one:" + chr(9) + "Create new case" + chr(9) + selected_PERS_search_case_numbers_array(0)
      If Ubound(selected_PERS_search_case_numbers_array) = 1 then case_selection_list_options = "Select one:" + chr(9) + "Create new case" + chr(9) + selected_PERS_search_case_numbers_array(0) + chr(9) + selected_PERS_search_case_numbers_array(1)
      If Ubound(selected_PERS_search_case_numbers_array) = 2 then case_selection_list_options = "Select one:" + chr(9) + "Create new case" + chr(9) + selected_PERS_search_case_numbers_array(0) + chr(9) + selected_PERS_search_case_numbers_array(1) + chr(9) + selected_PERS_search_case_numbers_array(2)
      If Ubound(selected_PERS_search_case_numbers_array) = 3 then case_selection_list_options = "Select one:" + chr(9) + "Create new case" + chr(9) + selected_PERS_search_case_numbers_array(0) + chr(9) + selected_PERS_search_case_numbers_array(1) + chr(9) + selected_PERS_search_case_numbers_array(2) + chr(9) + selected_PERS_search_case_numbers_array(3)
      If Ubound(selected_PERS_search_case_numbers_array) = 4 then case_selection_list_options = "Select one:" + chr(9) + "Create new case" + chr(9) + selected_PERS_search_case_numbers_array(0) + chr(9) + selected_PERS_search_case_numbers_array(1) + chr(9) + selected_PERS_search_case_numbers_array(2) + chr(9) + selected_PERS_search_case_numbers_array(3) + chr(9) + selected_PERS_search_case_numbers_array(4)
      
      'Only need to evaluate case numbers if we are on the applicant
      If member = 0 Then
        For case_number = 0 to Ubound(selected_PERS_search_case_numbers_array)
          'Navigate to STAT/MEMB to pull the household member number and relationship code
          Call back_to_SELF
          'Write the MAXIS case code
          EMWriteScreen "________", 18, 43
          ' msgbox selected_PERS_search_case_numbers_array(case_number)
          EMWriteScreen selected_PERS_search_case_numbers_array(case_number), 18, 43 
          EMWriteScreen "STAT", 16, 43 
          CALL write_value_and_transmit("MEMB", 21, 70)
          'Script will read each PMI number to see if there is a match
          Do
            ref_nbr = ""
            rel_code = ""
            pmi_number_check = ""
            last_memb_check = ""
    
            EmReadScreen pmi_number_check, 11, 4, 46
            pmi_number_check = trim(pmi_number_check)
            If pmi_number_check = pmi_number then
              EmReadScreen ref_nbr, 2, 4, 33
              EmReadScreen rel_code, 18, 10, 42
              Call back_to_SELF
              Exit Do
            End If
            transmit
            EmReadScreen last_memb_check, 5, 24, 2
            If last_memb_check = "ENTER" then Exit Do
          Loop
          'Update the array item with the details from STAT/MEMB
          If ref_nbr <> "" Then
            selected_PERS_search_case_numbers_array(case_number) = selected_PERS_search_case_numbers_array(case_number) & "&" & ref_nbr & "                         " & rel_code 
          Else
            selected_PERS_search_case_numbers_array(case_number) = selected_PERS_search_case_numbers_array(case_number) & "&" & "Person does not appear on case" 
          End If
        Next 
      End If
      'Set the variables to resizing the dialog
      case_num_btn_y = 60
      ref_rel_y = 60

      ' msgbox "selected_PERS_search_results_string " & selected_PERS_search_results_string
      If UBound(selected_PERS_search_case_numbers_array) > 4 Then msgbox "There are more case numbers than can be displayed. There are " & ubound(selected_PERS_search_case_numbers_array) + 1 & " cases."

      'Display dialog with details from MAXIS compared to details from XML
      Dialog1 = "" 'Blanking out previous dialog detail
      BeginDialog Dialog1, 0, 0, 270, 385, "Verify MNBenefits XML Details - Household Members"
        Text 5, 5, 250, 20, "Please review and verify the household member details for each household member pulled from the XML file below. Make any updates as needed."
        If member = 0 Then
          GroupBox 10, 30, 255, 85, "Review the Cases below for PERS Match"
          Text 20, 40, 185, 10, "Nav to case                  Ref nbr                 Rel code"
          ButtonGroup ButtonPressed
            PushButton 15, 50, 55, 10, mid(selected_PERS_search_case_numbers_array(0), 1, instr(selected_PERS_search_case_numbers_array(0), "&") - 1), case_number_nav_0_button
            Text 95, 50, 165, 10, mid(selected_PERS_search_case_numbers_array(0), instr(selected_PERS_search_case_numbers_array(0), "&") + 1)
    
            If UBound(selected_PERS_search_case_numbers_array) > 0 Then
              ' If selected_PERS_search_case_numbers_array(1) <> "" Then
                PushButton 15, case_num_btn_y, 55, 10, mid(selected_PERS_search_case_numbers_array(1), 1, instr(selected_PERS_search_case_numbers_array(1), "&") - 1), case_number_nav_1_button
                case_num_btn_y = case_num_btn_y + 10
                Text 95, ref_rel_y, 165, 10, mid(selected_PERS_search_case_numbers_array(1), instr(selected_PERS_search_case_numbers_array(1), "&") + 1)
                ref_rel_y = ref_rel_y + 10
              ' End If
            End If
            If UBound(selected_PERS_search_case_numbers_array) > 1 Then
              ' If selected_PERS_search_case_numbers_array(2) <> "" Then
                PushButton 15, case_num_btn_y, 55, 10, mid(selected_PERS_search_case_numbers_array(2), 1, instr(selected_PERS_search_case_numbers_array(2), "&") - 1), case_number_nav_2_button
                case_num_btn_y = case_num_btn_y + 10
                Text 95, ref_rel_y, 165, ref_rel_y, mid(selected_PERS_search_case_numbers_array(2), instr(selected_PERS_search_case_numbers_array(2), "&") + 1)
                ref_rel_y = ref_rel_y + 10
              ' End If
            End If
            If UBound(selected_PERS_search_case_numbers_array) > 2 Then
              ' If selected_PERS_search_case_numbers_array(3) <> "" Then
                PushButton 15, case_num_btn_y, 55, 10, mid(selected_PERS_search_case_numbers_array(3), 1, instr(selected_PERS_search_case_numbers_array(3), "&") - 1), case_number_nav_3_button
                case_num_btn_y = case_num_btn_y + 10
                Text 95, ref_rel_y, 165, ref_rel_y, mid(selected_PERS_search_case_numbers_array(3), instr(selected_PERS_search_case_numbers_array(3), "&") + 1)
                ref_rel_y = ref_rel_y + 10
              ' End If
            End If
            If UBound(selected_PERS_search_case_numbers_array) > 3 Then
              ' If selected_PERS_search_case_numbers_array(4) <> "" Then
                PushButton 15, case_num_btn_y, 55, 10, mid(selected_PERS_search_case_numbers_array(4), 1, instr(selected_PERS_search_case_numbers_array(4), "&") - 1), case_number_nav_4_button
                case_num_btn_y = case_num_btn_y + 10
                Text 95, ref_rel_y, 165, ref_rel_y, mid(selected_PERS_search_case_numbers_array(4), instr(selected_PERS_search_case_numbers_array(4), "&") + 1)
                ref_rel_y = ref_rel_y + 10
              ' End If
            End If
          GroupBox 10, 120, 255, 80, "PERS Search Result Details"
          Text 15, 135, 40, 10, "First name:"
          Text 65, 135, 190, 10, mid(selected_PERS_search_results_string, 1, instr(selected_PERS_search_results_string, " ") - 1)
          Text 15, 145, 40, 10, "Last name:"
          Text 65, 145, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, " ") + 1, instr(selected_PERS_search_results_string, " (DOB:") - instr(selected_PERS_search_results_string, " "))
          Text 15, 155, 45, 10, "Date of birth:"
          Text 65, 155, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, " (DOB:") + 7, instr(selected_PERS_search_results_string, "; SSN:") - instr(selected_PERS_search_results_string, " (DOB:") - 7) 
          Text 15, 165, 20, 10, "SSN:"
          Text 65, 165, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, "; SSN:") + 7, 11)
          Text 15, 175, 20, 10, "PMI:"
          Text 65, 175, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, "; PMI:") + 7, instr(selected_PERS_search_results_string, "; Gen") - instr(selected_PERS_search_results_string, "; PMI:") - 7)
          Text 15, 185, 30, 10, "Gender:"
          Text 65, 185, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, "; Gender: ") + 10, 1)
          GroupBox 10, 205, 255, 140, "Verify the XML details below (update as needed):"
          Text 15, 225, 40, 10, "First name:"
          EditBox 70, 220, 100, 15, household_members(MEMBER_FIRST_NAME, member)
          Text 15, 240, 40, 10, "Last name:"
          EditBox 70, 235, 100, 15, household_members(MEMBER_LAST_NAME, member)
          Text 15, 255, 30, 10, "Gender:"
          DropListBox 70, 250, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, member)
          Text 15, 270, 50, 10, "Marital status:"
          DropListBox 70, 265, 120, 20, "Select one:"+chr(9)+"Never Married"+chr(9)+"Married Living w/Spouse"+chr(9)+"Divorced"+chr(9)+"Separated (Married but living apart)", household_members(MEMBER_MARITAL_STATUS, member) 
          Text 15, 285, 45, 10, "Date of birth:"
          EditBox 70, 280, 100, 15, household_members(MEMBER_DOB, member)
          Text 15, 300, 20, 10, "SSN:"
          EditBox 70, 295, 100, 15, household_members(MEMBER_SSN, member)
          Text 15, 315, 45, 10, "Citizenship:"
          DropListBox 70, 310, 60, 15, "Select one:"+chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, member)
          Text 15, 330, 45, 10, "Relationship:"
          DropListBox 70, 325, 60, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, member)
          Text 10, 350, 165, 10, "Select existing case number or indicate new case:"
          DropListBox 180, 350, 80, 15, case_selection_list_options, case_selection_list
        Else
          GroupBox 10, 25, 255, 80, "PERS Search Result Details"
          Text 15, 40, 40, 10, "First name:"
          Text 65, 40, 190, 10, mid(selected_PERS_search_results_string, 1, instr(selected_PERS_search_results_string, " ") - 1)
          Text 15, 50, 40, 10, "Last name:"
          Text 65, 50, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, " ") + 1, instr(selected_PERS_search_results_string, " (DOB:") - instr(selected_PERS_search_results_string, " "))
          Text 15, 60, 45, 10, "Date of birth:"
          Text 65, 60, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, " (DOB:") + 7, instr(selected_PERS_search_results_string, "; SSN:") - instr(selected_PERS_search_results_string, " (DOB:") - 7) 
          Text 15, 70, 20, 10, "SSN:"
          Text 65, 70, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, "; SSN:") + 7, 11)
          Text 15, 80, 20, 10, "PMI:"
          Text 65, 80, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, "; PMI:") + 7, instr(selected_PERS_search_results_string, "; Gen") - instr(selected_PERS_search_results_string, "; PMI:") - 7)
          Text 15, 90, 30, 10, "Gender:"
          Text 65, 90, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, "; Gender: ") + 10, 1)
          GroupBox 10, 110, 255, 140, "Verify the XML details below (update as needed):"
          Text 15, 130, 40, 10, "First name:"
          EditBox 70, 125, 100, 15, household_members(MEMBER_FIRST_NAME, member)
          Text 15, 145, 40, 10, "Last name:"
          EditBox 70, 140, 100, 15, household_members(MEMBER_LAST_NAME, member)
          Text 15, 160, 30, 10, "Gender:"
          DropListBox 70, 155, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, member)
          Text 15, 175, 50, 10, "Marital status:"
          DropListBox 70, 170, 120, 20, "Select one:"+chr(9)+"Never Married"+chr(9)+"Married Living w/Spouse"+chr(9)+"Divorced"+chr(9)+"Separated (Married but living apart)", household_members(MEMBER_MARITAL_STATUS, member) 
          Text 15, 190, 45, 10, "Date of birth:"
          EditBox 70, 185, 100, 15, household_members(MEMBER_DOB, member)
          Text 15, 205, 20, 10, "SSN:"
          EditBox 70, 200, 100, 15, household_members(MEMBER_SSN, member)
          Text 15, 220, 45, 10, "Citizenship:"
          DropListBox 70, 215, 60, 15, "Select one:"+chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, member)
          Text 15, 235, 45, 10, "Relationship:"
          DropListBox 70, 230, 60, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, member)
        End If
        ButtonGroup ButtonPressed
        PushButton 175, 365, 45, 15, "Next", next_hh_memb_btn
        CancelButton 220, 365, 45, 15
      EndDialog
  
      DO
        DO
          err_msg = ""					'establishing value of variable, this is necessary for the Do...LOOP
          dialog Dialog1				'main dialog
          cancel_without_confirmation
          'to do - add error handling

          If member = 0 Then
            If ButtonPressed = case_number_nav_0_button Then
              Call back_to_SELF
              EMWriteScreen "________", 18, 43 
              EMWriteScreen mid(selected_PERS_search_case_numbers_array(0), 1, instr(selected_PERS_search_case_numbers_array(0), "&") - 1), 18, 43
              EMWriteScreen "STAT", 16, 43 
              CALL write_value_and_transmit("MEMB", 21, 70)
              err_msg = "LOOP"
            End If  
            If ButtonPressed = case_number_nav_1_button Then
              Call back_to_SELF
              EMWriteScreen "________", 18, 43 
              EMWriteScreen mid(selected_PERS_search_case_numbers_array(1), 1, instr(selected_PERS_search_case_numbers_array(1), "&") - 1), 18, 43
              EMWriteScreen "STAT", 16, 43 
              CALL write_value_and_transmit("MEMB", 21, 70)
              err_msg = "LOOP"
            End If
            If ButtonPressed = case_number_nav_2_button Then
              Call back_to_SELF
              EMWriteScreen "________", 18, 43 
              EMWriteScreen mid(selected_PERS_search_case_numbers_array(2), 1, instr(selected_PERS_search_case_numbers_array(2), "&") - 1), 18, 43
              EMWriteScreen "STAT", 16, 43 
              CALL write_value_and_transmit("MEMB", 21, 70)
              err_msg = "LOOP"
            End If
            If ButtonPressed = case_number_nav_3_button Then
              Call back_to_SELF
              EMWriteScreen "________", 18, 43 
              EMWriteScreen mid(selected_PERS_search_case_numbers_array(3), 1, instr(selected_PERS_search_case_numbers_array(3), "&") - 1), 18, 43
              EMWriteScreen "STAT", 16, 43 
              CALL write_value_and_transmit("MEMB", 21, 70)
              err_msg = "LOOP"
            End If
            If ButtonPressed = case_number_nav_4_button Then
              Call back_to_SELF
              EMWriteScreen "________", 18, 43 
              EMWriteScreen mid(selected_PERS_search_case_numbers_array(4), 1, instr(selected_PERS_search_case_numbers_array(4), "&") - 1), 18, 43
              EMWriteScreen "STAT", 16, 43 
              CALL write_value_and_transmit("MEMB", 21, 70)
              err_msg = "LOOP"
            End If

            If case_selection_list = "Create new case" and err_msg = "" Then
              proceed_new_case = MsgBox("You selected the 'Create new case' option despite the household member appearing on one or more existing cases. If you want to proceed with this selection then press 'Yes', otherwise you can return to previous dialog by pressing 'No'.", vbYesNo)
              If proceed_new_case = vbNo then err_msg = "LOOP"
            End If
          End If

          If err_msg <> "" and err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""									'loops until all errors are resolved
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
      Loop until are_we_passworded_out = false					'loops until user passwords back in
    End If

    If manual_no_match_found_checkbox = 1 Then
      ' Set exists to False
      household_members(MEMBER_EXISTS_MAXIS, member) = False

      Dialog1 = "" 'Blanking out previous dialog detail
      BeginDialog Dialog1, 0, 0, 271, 375, "Verify MNBenefits XML Details - Household Member"
        Text 5, 5, 250, 20, "Please review and verify the household member details for the household member pulled from the XML file below. Make any updates as needed."
        GroupBox 5, 30, 255, 140, "Verify the XML details below (update as needed):"
        Text 10, 50, 40, 10, "First name:"
        EditBox 65, 45, 100, 15, household_members(MEMBER_FIRST_NAME, member)
        Text 10, 65, 40, 10, "Last name:"
        EditBox 65, 60, 100, 15, household_members(MEMBER_LAST_NAME, member)
        Text 10, 80, 30, 10, "Gender:"
        DropListBox 65, 75, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, member)
        Text 10, 95, 50, 10, "Marital status:"
        DropListBox 65, 90, 100, 20, "Select one:"+chr(9)+"Never Married"+chr(9)+"Married Living w/Spouse"+chr(9)+"Divorced"+chr(9)+"Separated (Married but living apart)", household_members(MEMBER_MARITAL_STATUS, member)
        Text 10, 110, 45, 10, "Date of birth:"
        EditBox 65, 105, 100, 15, household_members(MEMBER_DOB, member)
        Text 10, 125, 20, 10, "SSN:"
        EditBox 65, 120, 100, 15, household_members(MEMBER_SSN, member)
        Text 10, 140, 45, 10, "Citizenship:"
        DropListBox 65, 135, 60, 15, "Select one:" +chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, member)
        Text 10, 155, 45, 10, "Relationship:"
        DropListBox 65, 150, 60, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, member)
        If member = 0 Then
          Text 5, 175, 170, 10, "Enter an existing case number or indicate new case:"
          Text 185, 175, 75, 15, "New case"
        End If
        ButtonGroup ButtonPressed
          PushButton 175, 355, 45, 15, "Next", next_hh_memb_btn
          CancelButton 220, 355, 45, 15
      EndDialog
    
  
      DO
        DO
          err_msg = ""					'establishing value of variable, this is necessary for the Do...LOOP
          dialog Dialog1				'main dialog
          cancel_without_confirmation
          'To do - add error handling
  
          If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""									'loops until all errors are resolved
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
      Loop until are_we_passworded_out = false					'loops until user passwords back in
    End If

  Else
    'Determine which result selected
    If pers_search_results_0 = 1 Then selected_PERS_search_results_string = PERS_search_results_string_array(0)
    If pers_search_results_1 = 1 Then selected_PERS_search_results_string = PERS_search_results_string_array(1)
    If pers_search_results_2 = 1 Then selected_PERS_search_results_string = PERS_search_results_string_array(2)
    If pers_search_results_3 = 1 Then selected_PERS_search_results_string = PERS_search_results_string_array(3)
    If pers_search_results_4 = 1 Then selected_PERS_search_results_string = PERS_search_results_string_array(4)
    If pers_search_results_5 = 1 Then selected_PERS_search_results_string = PERS_search_results_string_array(5)
    If pers_search_results_6 = 1 Then selected_PERS_search_results_string = PERS_search_results_string_array(6)
    If pers_search_results_7 = 1 Then selected_PERS_search_results_string = PERS_search_results_string_array(7)
    If pers_search_results_8 = 1 Then selected_PERS_search_results_string = PERS_search_results_string_array(8)
    If pers_search_results_9 = 1 Then selected_PERS_search_results_string = PERS_search_results_string_array(9)

    ' Set exists to True
    household_members(MEMBER_EXISTS_MAXIS, member) = True

    ' Split out the case numbers from the selected PERS search result

    'Pull out the PMI number
    PMI_number = mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, "PMI: ") + 5, instr(selected_PERS_search_results_string, "; Gender: ") - instr(selected_PERS_search_results_string, "PMI: ") - 5)

    'Pull out the case numbers
    selected_PERS_search_case_numbers = mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, "*"), len(selected_PERS_search_results_string) - instr(selected_PERS_search_results_string, "*"))
    selected_PERS_search_case_numbers = right(selected_PERS_search_case_numbers, len(selected_PERS_search_case_numbers) - 1)
    'to do - remove leading and ending asterisk

    selected_PERS_search_case_numbers_array = split(selected_PERS_search_case_numbers, "*")

    If Ubound(selected_PERS_search_case_numbers_array) = 0 then case_selection_list_options = "Select one:" + chr(9) + "Create new case" + chr(9) + selected_PERS_search_case_numbers_array(0)
    If Ubound(selected_PERS_search_case_numbers_array) = 1 then case_selection_list_options = "Select one:" + chr(9) + "Create new case" + chr(9) + selected_PERS_search_case_numbers_array(0) + chr(9) + selected_PERS_search_case_numbers_array(1)
    If Ubound(selected_PERS_search_case_numbers_array) = 2 then case_selection_list_options = "Select one:" + chr(9) + "Create new case" + chr(9) + selected_PERS_search_case_numbers_array(0) + chr(9) + selected_PERS_search_case_numbers_array(1) + chr(9) + selected_PERS_search_case_numbers_array(2)
    If Ubound(selected_PERS_search_case_numbers_array) = 3 then case_selection_list_options = "Select one:" + chr(9) + "Create new case" + chr(9) + selected_PERS_search_case_numbers_array(0) + chr(9) + selected_PERS_search_case_numbers_array(1) + chr(9) + selected_PERS_search_case_numbers_array(2) + chr(9) + selected_PERS_search_case_numbers_array(3)
    If Ubound(selected_PERS_search_case_numbers_array) = 4 then case_selection_list_options = "Select one:" + chr(9) + "Create new case" + chr(9) + selected_PERS_search_case_numbers_array(0) + chr(9) + selected_PERS_search_case_numbers_array(1) + chr(9) + selected_PERS_search_case_numbers_array(2) + chr(9) + selected_PERS_search_case_numbers_array(3) + chr(9) + selected_PERS_search_case_numbers_array(4)


    For case_number = 0 to Ubound(selected_PERS_search_case_numbers_array)
      'Navigate to STAT/MEMB to pull the household member number and relationship code
      Call back_to_SELF
      'Write the MAXIS case code
      EMWriteScreen "________", 18, 43
      EMWriteScreen selected_PERS_search_case_numbers_array(case_number), 18, 43 
      EMWriteScreen "STAT", 16, 43 
      CALL write_value_and_transmit("MEMB", 21, 70)
      'Script will read each PMI number to see if there is a match
      Do
        ref_nbr = ""
        rel_code = ""
        pmi_number_check = ""
        last_memb_check = ""

        EmReadScreen pmi_number_check, 11, 4, 46
        pmi_number_check = trim(pmi_number_check)
        If pmi_number_check = pmi_number then
          EmReadScreen ref_nbr, 2, 4, 33
          EmReadScreen rel_code, 18, 10, 42
          Call back_to_SELF
          Exit Do
        End If
        transmit
        EmReadScreen last_memb_check, 5, 24, 2
        If last_memb_check = "ENTER" then Exit Do
      Loop
      'Update the array item with the details from STAT/MEMB
      If ref_nbr <> "" Then
        selected_PERS_search_case_numbers_array(case_number) = selected_PERS_search_case_numbers_array(case_number) & "&" & ref_nbr & "                         " & rel_code 
      Else
        selected_PERS_search_case_numbers_array(case_number) = selected_PERS_search_case_numbers_array(case_number) & "&" & "Person does not appear on case" 
      End If
    Next 


    'Set the variables to resizing the dialog
    case_num_btn_y = 60
    ref_rel_y = 60

    If UBound(selected_PERS_search_case_numbers_array) > 4 Then msgbox "There are more case numbers than can be displayed. There are " & ubound(selected_PERS_search_case_numbers_array) + 1 & " cases."

    'Display dialog with details from MAXIS compared to details from XML
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 270, 385, "Verify MNBenefits XML Details - Household Members"
      Text 5, 5, 250, 20, "Please review and verify the household member details for each household member pulled from the XML file below. Make any updates as needed."
      If member = 0 Then
        GroupBox 10, 30, 255, 85, "Review the Cases below for PERS Match"
        Text 20, 40, 185, 10, "Nav to case                  Ref nbr                 Rel code"
        ButtonGroup ButtonPressed
          PushButton 15, 50, 55, 10, mid(selected_PERS_search_case_numbers_array(0), 1, instr(selected_PERS_search_case_numbers_array(0), "&") - 1), case_number_nav_0_button
          Text 95, 50, 165, 10, mid(selected_PERS_search_case_numbers_array(0), instr(selected_PERS_search_case_numbers_array(0), "&") + 1)
          If UBound(selected_PERS_search_case_numbers_array) > 0 Then
            PushButton 15, case_num_btn_y, 55, 10, mid(selected_PERS_search_case_numbers_array(1), 1, instr(selected_PERS_search_case_numbers_array(1), "&") - 1), case_number_nav_1_button
              case_num_btn_y = case_num_btn_y + 10
              Text 95, ref_rel_y, 165, 10, mid(selected_PERS_search_case_numbers_array(1), instr(selected_PERS_search_case_numbers_array(1), "&") + 1)
            ref_rel_y = ref_rel_y + 10
          End If
          If UBound(selected_PERS_search_case_numbers_array) > 1 Then
            PushButton 15, case_num_btn_y, 55, 10, mid(selected_PERS_search_case_numbers_array(2), 1, instr(selected_PERS_search_case_numbers_array(2), "&") - 1), case_number_nav_2_button
            case_num_btn_y = case_num_btn_y + 10
            Text 95, ref_rel_y, 165, ref_rel_y, mid(selected_PERS_search_case_numbers_array(2), instr(selected_PERS_search_case_numbers_array(2), "&") + 1)
            ref_rel_y = ref_rel_y + 10
          End If
          If UBound(selected_PERS_search_case_numbers_array) > 2 Then
            PushButton 15, case_num_btn_y, 55, 10, mid(selected_PERS_search_case_numbers_array(3), 1, instr(selected_PERS_search_case_numbers_array(3), "&") - 1), case_number_nav_3_button
            case_num_btn_y = case_num_btn_y + 10
            Text 95, ref_rel_y, 165, ref_rel_y, mid(selected_PERS_search_case_numbers_array(3), instr(selected_PERS_search_case_numbers_array(3), "&") + 1)
            ref_rel_y = ref_rel_y + 10
          End If
          If UBound(selected_PERS_search_case_numbers_array) > 3 Then
            PushButton 15, case_num_btn_y, 55, 10, mid(selected_PERS_search_case_numbers_array(4), 1, instr(selected_PERS_search_case_numbers_array(4), "&") - 1), case_number_nav_4_button
            case_num_btn_y = case_num_btn_y + 10
            Text 95, ref_rel_y, 165, ref_rel_y, mid(selected_PERS_search_case_numbers_array(4), instr(selected_PERS_search_case_numbers_array(4), "&") + 1)
            ref_rel_y = ref_rel_y + 10
          End If
        GroupBox 10, 120, 255, 80, "PERS Search Result Details"
        Text 15, 135, 40, 10, "First name:"
        Text 65, 135, 190, 10, mid(selected_PERS_search_results_string, 1, instr(selected_PERS_search_results_string, " ") - 1)
        Text 15, 145, 40, 10, "Last name:"
        Text 65, 145, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, " ") + 1, instr(selected_PERS_search_results_string, " (DOB:") - instr(selected_PERS_search_results_string, " "))
        Text 15, 155, 45, 10, "Date of birth:"
        Text 65, 155, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, " (DOB:") + 7, instr(selected_PERS_search_results_string, "; SSN:") - instr(selected_PERS_search_results_string, " (DOB:") - 7) 
        Text 15, 165, 20, 10, "SSN:"
        Text 65, 165, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, "; SSN:") + 7, 11)
        Text 15, 175, 20, 10, "PMI:"
        Text 65, 175, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, "; PMI:") + 7, instr(selected_PERS_search_results_string, "; Gen") - instr(selected_PERS_search_results_string, "; PMI:") - 7)
        Text 15, 185, 30, 10, "Gender:"
        Text 65, 185, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, "; Gender: ") + 10, 1)
        GroupBox 10, 205, 255, 140, "Verify the XML details below (update as needed):"
        Text 15, 225, 40, 10, "First name:"
        EditBox 70, 220, 100, 15, household_members(MEMBER_FIRST_NAME, member)
        Text 15, 240, 40, 10, "Last name:"
        EditBox 70, 235, 100, 15, household_members(MEMBER_LAST_NAME, member)
        Text 15, 255, 30, 10, "Gender:"
        DropListBox 70, 250, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, member)
        Text 15, 270, 50, 10, "Marital status:"
        DropListBox 70, 265, 120, 20, "Select one:"+chr(9)+"Never Married"+chr(9)+"Married Living w/Spouse"+chr(9)+"Divorced"+chr(9)+"Separated (Married but living apart)", household_members(MEMBER_MARITAL_STATUS, member) 
        Text 15, 285, 45, 10, "Date of birth:"
        EditBox 70, 280, 100, 15, household_members(MEMBER_DOB, member)
        Text 15, 300, 20, 10, "SSN:"
        EditBox 70, 295, 100, 15, household_members(MEMBER_SSN, member)
        Text 15, 315, 45, 10, "Citizenship:"
        DropListBox 70, 310, 60, 15, "Select one:"+chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, member)
        Text 15, 330, 45, 10, "Relationship:"
        DropListBox 70, 325, 60, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, member)
        Text 10, 350, 165, 10, "Select existing case number or indicate new case:"
        DropListBox 180, 350, 80, 15, case_selection_list_options, case_selection_list
      Else
        GroupBox 10, 25, 255, 80, "PERS Search Result Details"
        Text 15, 40, 40, 10, "First name:"
        Text 65, 40, 190, 10, mid(selected_PERS_search_results_string, 1, instr(selected_PERS_search_results_string, " ") - 1)
        Text 15, 50, 40, 10, "Last name:"
        Text 65, 50, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, " ") + 1, instr(selected_PERS_search_results_string, " (DOB:") - instr(selected_PERS_search_results_string, " "))
        Text 15, 60, 45, 10, "Date of birth:"
        Text 65, 60, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, " (DOB:") + 7, instr(selected_PERS_search_results_string, "; SSN:") - instr(selected_PERS_search_results_string, " (DOB:") - 7) 
        Text 15, 70, 20, 10, "SSN:"
        Text 65, 70, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, "; SSN:") + 7, 11)
        Text 15, 80, 20, 10, "PMI:"
        Text 65, 80, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, "; PMI:") + 7, instr(selected_PERS_search_results_string, "; Gen") - instr(selected_PERS_search_results_string, "; PMI:") - 7)
        Text 15, 90, 30, 10, "Gender:"
        Text 65, 90, 190, 10, mid(selected_PERS_search_results_string, instr(selected_PERS_search_results_string, "; Gender: ") + 10, 1)
        GroupBox 10, 110, 255, 140, "Verify the XML details below (update as needed):"
        Text 15, 130, 40, 10, "First name:"
        EditBox 70, 125, 100, 15, household_members(MEMBER_FIRST_NAME, member)
        Text 15, 145, 40, 10, "Last name:"
        EditBox 70, 140, 100, 15, household_members(MEMBER_LAST_NAME, member)
        Text 15, 160, 30, 10, "Gender:"
        DropListBox 70, 155, 60, 10, "Select one:"+chr(9)+"Male"+chr(9)+"Female"+chr(9)+"Other", household_members(MEMBER_GENDER, member)
        Text 15, 175, 50, 10, "Marital status:"
        DropListBox 70, 170, 120, 20, "Select one:"+chr(9)+"Never Married"+chr(9)+"Married Living w/Spouse"+chr(9)+"Divorced"+chr(9)+"Separated (Married but living apart)", household_members(MEMBER_MARITAL_STATUS, member) 
        Text 15, 190, 45, 10, "Date of birth:"
        EditBox 70, 185, 100, 15, household_members(MEMBER_DOB, member)
        Text 15, 205, 20, 10, "SSN:"
        EditBox 70, 200, 100, 15, household_members(MEMBER_SSN, member)
        Text 15, 220, 45, 10, "Citizenship:"
        DropListBox 70, 215, 60, 15, "Select one:"+chr(9)+"Yes"+chr(9)+"No", household_members(MEMBER_CITIZENSHIP, member)
        Text 15, 235, 45, 10, "Relationship:"
        DropListBox 70, 230, 60, 10, "Select one:"+chr(9)+"Self"+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Step Child"+chr(9)+"Parent"+chr(9)+"Sibling"+chr(9)+"Other Relative"+chr(9)+"Other", household_members(MEMBER_RELATIONSHIP, member)
      End If  
      ButtonGroup ButtonPressed
        PushButton 175, 365, 45, 15, "Next", next_hh_memb_btn
        CancelButton 220, 365, 45, 15
    EndDialog

    DO
      DO
        err_msg = ""					'establishing value of variable, this is necessary for the Do...LOOP
        dialog Dialog1				'main dialog
        cancel_without_confirmation
        If member = 0 Then
          If ButtonPressed = case_number_nav_0_button Then
            Call back_to_SELF
            EMWriteScreen "________", 18, 43 
            EMWriteScreen mid(selected_PERS_search_case_numbers_array(0), 1, instr(selected_PERS_search_case_numbers_array(0), "&") - 1), 18, 43
            EMWriteScreen "STAT", 16, 43 
            CALL write_value_and_transmit("MEMB", 21, 70)
            err_msg = "LOOP"
          End If  
          If ButtonPressed = case_number_nav_1_button Then
            Call back_to_SELF
            EMWriteScreen "________", 18, 43 
            EMWriteScreen mid(selected_PERS_search_case_numbers_array(1), 1, instr(selected_PERS_search_case_numbers_array(1), "&") - 1), 18, 43
            EMWriteScreen "STAT", 16, 43 
            CALL write_value_and_transmit("MEMB", 21, 70)
            err_msg = "LOOP"
          End If
          If ButtonPressed = case_number_nav_2_button Then
            Call back_to_SELF
            EMWriteScreen "________", 18, 43 
            EMWriteScreen mid(selected_PERS_search_case_numbers_array(2), 1, instr(selected_PERS_search_case_numbers_array(2), "&") - 1), 18, 43
            EMWriteScreen "STAT", 16, 43 
            CALL write_value_and_transmit("MEMB", 21, 70)
            err_msg = "LOOP"
          End If
          If ButtonPressed = case_number_nav_3_button Then
            Call back_to_SELF
            EMWriteScreen "________", 18, 43 
            EMWriteScreen mid(selected_PERS_search_case_numbers_array(3), 1, instr(selected_PERS_search_case_numbers_array(3), "&") - 1), 18, 43
            EMWriteScreen "STAT", 16, 43 
            CALL write_value_and_transmit("MEMB", 21, 70)
            err_msg = "LOOP"
          End If
          If ButtonPressed = case_number_nav_4_button Then
            Call back_to_SELF
            EMWriteScreen "________", 18, 43 
            EMWriteScreen mid(selected_PERS_search_case_numbers_array(4), 1, instr(selected_PERS_search_case_numbers_array(4), "&") - 1), 18, 43
            EMWriteScreen "STAT", 16, 43 
            CALL write_value_and_transmit("MEMB", 21, 70)
            err_msg = "LOOP"
          End If

          'To do - add error handling here

          If case_selection_list = "Create new case" and err_msg = "" Then
            proceed_new_case = MsgBox("You selected the 'Create new case' option despite the household member appearing on one or more existing cases. If you want to proceed with this selection then press 'Yes', otherwise you can return to previous dialog by pressing 'No'.", vbYesNo)
            If proceed_new_case = vbNo then err_msg = "LOOP"
          End If
        End If

        If err_msg <> "" and err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
      LOOP UNTIL err_msg = ""									'loops until all errors are resolved
      CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
  End If

  'Set the case number based on selection
  If member = 0 Then
    household_members(APPL_CASE_NUMBER, member) = case_selection_list
  End If 

  'Reset all dialog fields that may carry forward
  pers_search_results_0           = 0
  pers_search_results_1           = 0
  pers_search_results_2           = 0
  pers_search_results_3           = 0
  pers_search_results_4           = 0
  pers_search_results_5           = 0
  pers_search_results_6           = 0
  pers_search_results_7           = 0
  pers_search_results_8           = 0
  pers_search_results_9           = 0
  no_match_search_manually        = 0
  manual_match_found_checkbox     = 0
  PMI_number                      = ""
  manual_no_match_found_checkbox  = 0
Next

'To do - include dialog to review address details
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 256, 265, "Process MNBenefits Application"
  Text 10, 5, 240, 20, "Please verify the application and address details pulled from the XML file below. Make updates as needed."
  Text 15, 30, 150, 10, "Adjust date to correct business day, if needed"
  Text 15, 45, 60, 10, "Application Date: "
  EditBox 80, 40, 50, 15, formatted_app_date
  GroupBox 10, 60, 175, 105, "Household Address"
  Text 15, 75, 35, 10, "Address:"
  EditBox 70, 70, 100, 15, household_address
  Text 15, 90, 25, 10, "City:"
  EditBox 70, 85, 100, 15, household_city
  Text 15, 105, 30, 10, "State:"
  EditBox 70, 100, 20, 15, household_state
  Text 15, 120, 20, 10, "Zip:"
  EditBox 70, 115, 45, 15, household_zip
  Text 15, 135, 55, 10, "Phone number:"
  EditBox 70, 130, 100, 15, household_phone_number
  Text 15, 150, 30, 10, "County:"
  EditBox 70, 145, 100, 15, household_county
  GroupBox 10, 165, 175, 75, "Mailing Address"
  Text 15, 180, 35, 10, "Address:"
  EditBox 70, 175, 100, 15, mailing_address
  Text 15, 195, 25, 10, "City:"
  EditBox 70, 190, 100, 15, mailing_city
  Text 15, 210, 30, 10, "State:"
  EditBox 70, 205, 20, 15, mailing_state
  Text 15, 225, 20, 10, "Zip:"
  EditBox 70, 220, 45, 15, mailing_zip
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
    'To do - verify which fields are REQUIRED
    If trim(formatted_app_date) = "" OR IsDate(formatted_app_date) = False OR Len(trim(formatted_app_date)) <> 10 then err_msg = err_msg & vbCr & "* You must enter the application date in the format MM/DD/YYYY."
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

'After PERS search complete, update the XML file
member_array_index = 0

For Each objMemberNode In objHouseholdMemberNodes
  Set objFirstNameNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:Person/ns4:FirstName")
  Set objLastNameNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:Person/ns4:LastName")
  Set objSSNNode = objMemberNode.selectSingleNode("ns4:CitizenshipInfo/ns4:SSNInfo/ns4:SSN")
  Set objDOBNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:DOB")
  Set objRelationshipNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:Relationship") 
  Set objMaritalStatusNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:MaritalStatus")
  Set objCitizenshipNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:CitizenshipInfo")
  Set objGenderNode = objMemberNode.selectSingleNode("ns4:PersonalInfo/ns4:Gender")

  If household_members(MEMBER_FIRST_NAME, member_array_index) <> "" Then objFirstNameNode.Text = household_members(MEMBER_FIRST_NAME, member_array_index)

  If household_members(MEMBER_LAST_NAME, member_array_index) <> "" Then objLastNameNode.Text = household_members(MEMBER_LAST_NAME, member_array_index)

  If household_members(MEMBER_DOB, member_array_index) <> "" Then objDOBNode.Text = household_members(MEMBER_DOB, member_array_index)

  If household_members(MEMBER_SSN, member_array_index) <> "" Then objSSNNode.Text = household_members(MEMBER_SSN, member_array_index)

  If household_members(MEMBER_RELATIONSHIP, member_array_index) <> "" Then objRelationshipNode.Text = household_members(MEMBER_RELATIONSHIP, member_array_index)

  If household_members(MEMBER_MARITAL_STATUS, member_array_index) <> "" Then objMaritalStatusNode.Text = household_members(MEMBER_MARITAL_STATUS, member_array_index)

  If household_members(MEMBER_GENDER, member_array_index) <> "" Then objGenderNode.Text = household_members(MEMBER_GENDER, member_array_index)

  If household_members(MEMBER_FIRST_NAME, member_array_index) = "" Then Exit For

  member_array_index = member_array_index + 1
Next

'Display the details before proceeding to APPL case or update existing case
dialog_member_count = 0

'XML File Confirmation Dialog
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 256, 245, "Verify MNBenefits XML Details - Household Members"
  Text 5, 5, 250, 25, "The script will now proceed to make updates in MAXIS to APPL the case or update an existing case based on the information entered. Please review the information below and then press 'Continue' to proceed. "
  GroupBox 10, 50, 270, 155, "Household members listed on MNBenefits Application"
  Text 15, 60, 75, 10, "Confirmation number:"
  Text 100, 45, 50, 10, confirmation_number
  Text 15, 70, 60, 10, "Application Date:"
  Text 100, 70, 60, 10, formatted_app_date
  Text 15, 80, 75, 10, "Household Member 1:"
  Text 100, 80, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"
  dialog_member_count = dialog_member_count + 1
  If member_count > 1 Then
    Text 15, 90, 75, 10, "Household Member 2:"
    Text 100, 90, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 2 Then
    Text 15, 100, 75, 10, "Household Member 3:"
    Text 100, 100, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 3 Then
    Text 15, 110, 75, 10, "Household Member 4:"
    Text 100, 110, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 4 Then
    Text 15, 120, 75, 10, "Household Member 5:"
    Text 100, 120, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 5 Then
    Text 15, 130, 75, 10, "Household Member 6:"
    Text 100, 130, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 6 Then
    Text 15, 140, 75, 10, "Household Member 7:"
    Text 100, 140, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 7 Then
    Text 15, 150, 75, 10, "Household Member 8:"
    Text 100, 150, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 8 Then
    Text 15, 160, 75, 10, "Household Member 9:"
    Text 100, 160, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 9 Then
    Text 15, 170, 80, 10, "Household Member 10:"
    Text 100, 170, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 10 Then
    Text 15, 180, 80, 10, "Household Member 11:"
    Text 100, 180, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"
    dialog_member_count = dialog_member_count + 1
  End If
  If member_count > 11 Then
    Text 15, 190, 80, 10, "Household Member 12:"
    Text 100, 190, 175, 10, left(household_members(MEMBER_LAST_NAME, dialog_member_count) & ", " & household_members(MEMBER_FIRST_NAME, dialog_member_count), 25) & " (" & household_members(MEMBER_DOB, dialog_member_count) & "; " & household_members(MEMBER_SSN, dialog_member_count) & ")"
    dialog_member_count = dialog_member_count + 1
  End If
    ButtonGroup ButtonPressed
      PushButton 235, 225, 45, 15, "Continue", continue_button
  EndDialog

DO
  dialog Dialog1
  cancel_without_confirmation
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'---Update XML - add to very end
'Replace the application date
'Format 2025-11-26

' current_XML_app_date = left(objApplicationDate.Text, 10)
' 'Convert the formatted_app_date to XML format and replace
' updated_XML_app_date = right(formatted_app_date, 4) & "-" & left(formatted_app_date, 2) & "-" & mid(formatted_app_date, 4, 2)
' objApplicationDate.Text       = replace(objApplicationDate.Text, current_XML_app_date, updated_XML_app_date)
' objHouseholdAddress.Text      = household_address
' objHouseholdCity.Text         = household_city
' objHouseholdState.Text        = household_state
' objHouseholdZip.Text          = household_zip
' objPhoneNumber.Text           = household_phone_number
' objCounty.Text                = household_county
' objMailingAddress.Text        = mailing_address
' objMailingCity.Text           = mailing_city
' objMailingState.Text          = mailing_state
' objMailingZip.Text            = mailing_zip

' ' Save the updated XML to a file
' ' To do - update with actual file path once done testing
' xmlDoc.Save "C:\Users\mari001\OneDrive - Hennepin County\Desktop\New XML Files\new xml file success.xml"
' xmlDoc.Save replace(XML_file_path, confirmation_number, confirmation_number & "_" & "processed")

' ' Save the XML document with 'processed' in file name
' xmlDoc.Save replace(XML_file_path, confirmation_number, confirmation_number & "_" & "processed")

' ' To do - uncomment after testing, this is where file is saved and moved

' On Error Resume Next

' ' Attempt to save the XML document
' Dim XML_file_path_processed
' XML_file_path_processed = Replace(XML_file_path, confirmation_number, confirmation_number & "_processed")
' xmlDoc.Save XML_file_path_processed

' ' Check for errors
' If Err.Number <> 0 Then
'   WScript.Echo "Error saving file: " & Err.Description
'   ' Optional: log the error or take corrective action
'   script_end_procedure_with_error_report("Script failed to save the processed XML file. The script will now end.")
' Else
'   msgbox "Success!"
' End If

' On Error GoTo 0 ' Reset error handling

' Set fso = CreateObject("Scripting.FileSystemObject")

' If fso.FileExists(XML_file_path) Then
'   fso.DeleteFile XML_file_path
' Else
'   script_end_procedure_with_error_report("Script failed to delete XML file.")
' End If

' ' Clean up
' Set objMemberNode           = Nothing
' Set objFirstNameNode        = Nothing
' Set objLastNameNode         = Nothing
' Set objDOBNode              = Nothing
' Set objSSNNode              = Nothing
' Set objRelationshipNode     = Nothing
' Set objMaritalStatusNode    = Nothing
' Set objGenderNode           = Nothing
' Set objCitizenshipNode      = Nothing
' Set objHouseholdMemberNodes = Nothing
' Set objHouseholdMemberNode  = Nothing
' Set xmlDoc                  = Nothing

' MsgBox "XML file saved and updated successfully from array."
'---update XML code
