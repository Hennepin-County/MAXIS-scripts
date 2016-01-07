'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - PROOF OF RELATIONSHIP.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
'END FUNCTIONS LIBRARY BLOCK============================================================================================================================

'Required for statistical purposes==========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'Dialog Box for Case Number.
BeginDialog Case_Number_Dialog, 0, 0, 146, 65, "Case Number"
  EditBox 60, 5, 60, 15, case_number
  ButtonGroup ButtonPressed
    OkButton 15, 30, 50, 15
    CancelButton 70, 30, 50, 15
  Text 10, 10, 50, 10, "Case Number: "
EndDialog

'Connecting to BlueZone.
EMConnect ""
EMFocus
call Check_for_MAXIS(True)

Dialog Case_Number_Dialog

'Drop down list for available household members to put into the drop down list.
Call HH_member_custom_dialog(HH_Member_Array)
Call convert_array_to_droplist_items(HH_Member_Array, hh_member_dropdown)

'Dialog Box to list members and documentation received.
BeginDialog Proof_of_Relationship_Dialog, 0, 0, 296, 210, "Proof of Relationship"
  DropListBox 5, 25, 70, 15, hh_member_dropdown, received_for
  DropListBox 90, 25, 110, 15, "Select one..."+chr(9)+"Is Another Relative of"+chr(9)+"Is the Child of"+chr(9)+"Is the Foster Child of"+chr(9)+"Is the Grandchild of"+chr(9)+"Is the Guardian of"+chr(9)+"Is the Nephew of"+chr(9)+"Is the Niece of"+chr(9)+"Is the Parent of"+chr(9)+"Is the Sibling of"+chr(9)+"Is the Spouse of"+chr(9)+"Is the Step Child of"+chr(9)+"Is Unrelated to"+chr(9)+"Other", relationship_type
  EditBox 90, 60, 110, 15, other_relationship_list
  DropListBox 215, 25, 70, 15, hh_member_dropdown, relationship_to
  EditBox 85, 100, 115, 15, documents_received
  EditBox 60, 125, 140, 20, other_notes
  CheckBox 215, 105, 50, 10, "MEMB", memb
  CheckBox 215, 115, 50, 10, "PARE", pare
  CheckBox 215, 125, 50, 10, "ABPS", abps
  CheckBox 215, 135, 30, 10, "SIBL", sibl
  CheckBox 215, 145, 30, 10, "Other:", other
  EditBox 215, 155, 70, 15, other_option
  EditBox 75, 160, 125, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 150, 185, 50, 15
    CancelButton 205, 185, 50, 15
  Text 90, 5, 50, 10, "Relationship: "
  Text 10, 165, 60, 10, "Worker signature:"
  Text 215, 5, 50, 10, "Related to:"
  Text 5, 5, 75, 15, "Member received for: "
  Text 90, 50, 90, 10, "If Other, list relationship:"
  Text 5, 105, 75, 10, "Document(s) received: "
  Text 10, 135, 45, 10, "Other Notes:"
  Text 215, 90, 60, 10, "Panel(s) updated: "
EndDialog
Do
	Do
		Do
			Do
				Dialog Proof_of_relationship_dialog
				cancel_confirmation
				IF relationship_type = "Select one..." THEN MsgBox "You must select a relationship type."
			Loop until relationship_type <> "Select one..."
			IF relationship_type = "Other" AND other_relationship_list = "" THEN MsgBox "You must list a relationship if Other is selected."
		Loop until (relationship_type = "Other" AND other_relationship_list <> "") OR relationship_type <> "Other"
		IF other = 1 AND other_option = "" THEN MsgBox "You must list a panel if Other is selected."
	Loop until (other = 1 AND other_option <> "") OR other = 0
	IF worker_signature = "" THEN MsgBox "You must enter a worker signature."
Loop until worker_signature <> ""

Call navigate_to_MAXIS_screen("CASE","NOTE")

PF9

'Statements needed for the check boxes for panels updated, defined further in case notes below.
IF memb = 1 THEN 
	memb = "MEMB " 
ELSE
	memb = ""
END IF


IF pare = 1 THEN
	pare = "PARE "
ELSE
	pare = ""
END IF


IF abps = 1 THEN
	abps = "ABPS "
ELSE
	abps = ""
END IF


IF sibl = 1 THEN
	sibl = "SIBL "
ELSE
	sibl = ""
END IF


IF other = 1 THEN
	other = "- Other: "
ELSE
	other = ""
END IF
	
'Information for the case note.

Call write_variable_in_CASE_NOTE("Documentation Received: Proof of Relationship")
Call write_new_line_in_CASE_NOTE("* Proof of relationship received for: Memb " & received_for)
Call write_new_line_in_CASE_NOTE("* Verifies relationship to: Memb " & relationship_to)
Call write_new_line_in_CASE_NOTE("* Relationship: Memb " & received_for & " --- " & relationship_type & " --- Memb " & relationship_to)
Call write_bullet_and_variable_in_CASE_NOTE("Other Relationship", other_relationship_list)
Call write_bullet_and_variable_in_CASE_NOTE("Document(s) received", documents_received)
Call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
Call write_bullet_and_variable_in_CASE_NOTE("Panel(s) updated", memb & pare & abps & sibl & other & other_option)
Call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
