'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - CITIZENSHIP-IDENTITY VERIFIED.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" OR default_directory = "" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
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


'DIALOGS-------------------------------------------------------------------

BeginDialog cit_ID_dialog, 0, 0, 346, 222, "CIT-ID dialog"
  Text 5, 10, 50, 10, "Case number:"
  EditBox 60, 5, 75, 15, case_number
  Text 20, 25, 45, 10, "HH member"
  Text 85, 25, 55, 10, "Exempt reason"
  Text 200, 25, 35, 10, "Cit proof"
  Text 290, 25, 35, 10, "ID proof"
  EditBox 5, 40, 65, 15, HH_memb_01
  ComboBox 80, 40, 85, 10, "(select or type here)"+chr(9)+"MEDI enrollee"+chr(9)+"SSI/RSDI recip."+chr(9)+"foster care"+chr(9)+"adoption assist."+chr(9)+"auto newborn", exempt_reason_01
  ComboBox 170, 40, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Birth Certificate"+chr(9)+"Nat. papers"+chr(9)+"US passport", cit_proof_01
  ComboBox 260, 40, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Drivers License"+chr(9)+"State ID"+chr(9)+"School ID"+chr(9)+"Parent Signature"+chr(9)+"US passport", ID_proof_01
  EditBox 5, 60, 65, 15, HH_memb_02
  ComboBox 80, 60, 85, 10, "(select or type here)"+chr(9)+"MEDI enrollee"+chr(9)+"SSI/RSDI recip."+chr(9)+"foster care"+chr(9)+"adoption assist."+chr(9)+"auto newborn", exempt_reason_02
  ComboBox 170, 60, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Birth Certificate"+chr(9)+"Nat. papers"+chr(9)+"US passport", cit_proof_02
  ComboBox 260, 60, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Drivers License"+chr(9)+"State ID"+chr(9)+"School ID"+chr(9)+"Parent Signature"+chr(9)+"US passport", ID_proof_02
  EditBox 5, 80, 65, 15, HH_memb_03
  ComboBox 80, 80, 85, 10, "(select or type here)"+chr(9)+"MEDI enrollee"+chr(9)+"SSI/RSDI recip."+chr(9)+"foster care"+chr(9)+"adoption assist."+chr(9)+"auto newborn", exempt_reason_03
  ComboBox 170, 80, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Birth Certificate"+chr(9)+"Nat. papers"+chr(9)+"US passport", cit_proof_03
  ComboBox 260, 80, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Drivers License"+chr(9)+"State ID"+chr(9)+"School ID"+chr(9)+"Parent Signature"+chr(9)+"US passport", ID_proof_03
  EditBox 5, 100, 65, 15, HH_memb_04
  ComboBox 80, 100, 85, 10, "(select or type here)"+chr(9)+"MEDI enrollee"+chr(9)+"SSI/RSDI recip."+chr(9)+"foster care"+chr(9)+"adoption assist."+chr(9)+"auto newborn", exempt_reason_04
  ComboBox 170, 100, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Birth Certificate"+chr(9)+"Nat. papers"+chr(9)+"US passport", cit_proof_04
  ComboBox 260, 100, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Drivers License"+chr(9)+"State ID"+chr(9)+"School ID"+chr(9)+"Parent Signature"+chr(9)+"US passport", ID_proof_04
  EditBox 5, 120, 65, 15, HH_memb_05
  ComboBox 80, 120, 85, 10, "(select or type here)"+chr(9)+"MEDI enrollee"+chr(9)+"SSI/RSDI recip."+chr(9)+"foster care"+chr(9)+"adoption assist."+chr(9)+"auto newborn", exempt_reason_05
  ComboBox 170, 120, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Birth Certificate"+chr(9)+"Nat. papers"+chr(9)+"US passport", cit_proof_05
  ComboBox 260, 120, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Drivers License"+chr(9)+"State ID"+chr(9)+"School ID"+chr(9)+"Parent Signature"+chr(9)+"US passport", ID_proof_05
  EditBox 5, 140, 65, 15, HH_memb_06
  ComboBox 80, 140, 85, 10, "(select or type here)"+chr(9)+"MEDI enrollee"+chr(9)+"SSI/RSDI recip."+chr(9)+"foster care"+chr(9)+"adoption assist."+chr(9)+"auto newborn", exempt_reason_06
  ComboBox 170, 140, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Birth Certificate"+chr(9)+"Nat. papers"+chr(9)+"US passport", cit_proof_06
  ComboBox 260, 140, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Drivers License"+chr(9)+"State ID"+chr(9)+"School ID"+chr(9)+"Parent Signature"+chr(9)+"US passport", ID_proof_06
  EditBox 5, 160, 65, 15, HH_memb_07
  ComboBox 80, 160, 85, 10, "(select or type here)"+chr(9)+"MEDI enrollee"+chr(9)+"SSI/RSDI recip."+chr(9)+"foster care"+chr(9)+"adoption assist."+chr(9)+"auto newborn", exempt_reason_07
  ComboBox 170, 160, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Birth Certificate"+chr(9)+"Nat. papers"+chr(9)+"US passport", cit_proof_07
  ComboBox 260, 160, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Drivers License"+chr(9)+"State ID"+chr(9)+"School ID"+chr(9)+"Parent Signature"+chr(9)+"US passport", ID_proof_07
  EditBox 5, 180, 65, 15, HH_memb_08
  ComboBox 80, 180, 85, 10, "(select or type here)"+chr(9)+"MEDI enrollee"+chr(9)+"SSI/RSDI recip."+chr(9)+"foster care"+chr(9)+"adoption assist."+chr(9)+"auto newborn", exempt_reason_08
  ComboBox 170, 180, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Birth Certificate"+chr(9)+"Nat. papers"+chr(9)+"US passport", cit_proof_08
  ComboBox 260, 180, 85, 15, "(select or type here)"+chr(9)+"Elect. verif."+chr(9)+"Drivers License"+chr(9)+"State ID"+chr(9)+"School ID"+chr(9)+"Parent Signature"+chr(9)+"US passport", ID_proof_08
  Text 5, 205, 65, 10, "Sign the case note:"
  EditBox 75, 200, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 195, 200, 50, 15
    CancelButton 250, 200, 50, 15
EndDialog


'THE SCRIPT------------------------------------------------------------------------------------------
'Connecting to BlueZone
EMConnect ""
'Checking for MAXIS
call check_for_MAXIS(True)

'Searching for case number
call MAXIS_case_number_finder(case_number)

'Show the dialog, determine if it's filled out correctly (at least one line must be filled out), then navigating to a blank case note.
Do
	Dialog cit_ID_dialog
	cancel_confirmation
	If (HH_memb_01 <> "" and (exempt_reason_01 = "(select or type here)" and (cit_proof_01 = "(select or type here)" or ID_proof_01 = "(select or type here)"))) or _
	   (HH_memb_02 <> "" and (exempt_reason_02 = "(select or type here)" and (cit_proof_02 = "(select or type here)" or ID_proof_02 = "(select or type here)"))) or _
	   (HH_memb_03 <> "" and (exempt_reason_03 = "(select or type here)" and (cit_proof_03 = "(select or type here)" or ID_proof_03 = "(select or type here)"))) or _
	   (HH_memb_04 <> "" and (exempt_reason_04 = "(select or type here)" and (cit_proof_04 = "(select or type here)" or ID_proof_04 = "(select or type here)"))) or _
	   (HH_memb_05 <> "" and (exempt_reason_05 = "(select or type here)" and (cit_proof_05 = "(select or type here)" or ID_proof_05 = "(select or type here)"))) or _
	   (HH_memb_06 <> "" and (exempt_reason_06 = "(select or type here)" and (cit_proof_06 = "(select or type here)" or ID_proof_06 = "(select or type here)"))) or _
	   (HH_memb_07 <> "" and (exempt_reason_07 = "(select or type here)" and (cit_proof_07 = "(select or type here)" or ID_proof_07 = "(select or type here)"))) or _
	   (HH_memb_08 <> "" and (exempt_reason_08 = "(select or type here)" and (cit_proof_08 = "(select or type here)" or ID_proof_08 = "(select or type here)"))) then
		can_move_on = False
    Else
		can_move_on = True
	End if
	If can_move_on = False then MsgBox "You must select a CIT and ID proof for each client whose name you've typed."
Loop until can_move_on = True

'checking for an active MAXIS session
Call check_for_MAXIS(False)

'Case noting----------------------------------------------------------------------------------------------------
Call start_a_blank_CASE_NOTE
'body of the case note
Call write_variable_in_CASE_NOTE("***CITIZENSHIP/IDENTITY***")
Call write_variable_in_CASE_NOTE("--------------------------------------------------------------------------------")
Call write_variable_in_CASE_NOTE("    HH MEMB         EXEMPT REASON            CIT PROOF         ID PROOF")
If HH_memb_01 <> "" then 
	Call write_variable_in_CASE_NOTE("                                                                            ")
	Call write_variable_in_CASE_NOTE(HH_memb_01)
	IF exempt_reason_01 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(exempt_reason_01)
	IF cit_proof_01 <> "(select or type here)" then call write_variable_in_CASE_NOTE (cit_proof_01)
	IF ID_proof_01 <> "(select or type here)" then call write_variable_in_CASE_NOTE (ID_proof_01)
End if
If HH_memb_02 <> "" then
	Call write_variable_in_CASE_NOTE("                                                                            ")
	Call write_variable_in_CASE_NOTE(HH_memb_02)
	IF exempt_reason_02 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(exempt_reason_02)
	IF cit_proof_02 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(cit_proof_02)
	IF ID_proof_02 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(ID_proof_02)
End if
If HH_memb_03 <> "" then
	Call write_variable_in_CASE_NOTE("                                                                            ")
	Call write_variable_in_CASE_NOTE(HH_memb_03)
	IF exempt_reason_03 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(exempt_reason_03)
	IF cit_proof_03 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(cit_proof_03)
	IF ID_proof_03 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(ID_proof_03)
If HH_memb_04 <> "" then
	Call write_variable_in_CASE_NOTE("                                                                            ")
	Call write_variable_in_CASE_NOTE(HH_memb_04)
	IF exempt_reason_04 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(exempt_reason_04)
	IF cit_proof_04 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(cit_proof_04)
	IF ID_proof_04 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(ID_proof_04)
End if
If HH_memb_05 <> "" then
	Call write_variable_in_CASE_NOTE("                                                                            ")
	Call write_variable_in_CASE_NOTE(HH_memb_05)
	IF exempt_reason_05 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(exempt_reason_05)
	IF cit_proof_05 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(cit_proof_05)
	IF ID_proof_05 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(ID_proof_05)
End if
If HH_memb_06 <> "" then
	Call write_variable_in_CASE_NOTE("                                                                            ")
	Call write_variable_in_CASE_NOTE(HH_memb_06)	
	IF exempt_reason_06 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(exempt_reason_06)
	IF cit_proof_06 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(cit_proof_06)
	IF ID_proof_06 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(ID_proof_06)
End if
If HH_memb_07 <> "" then
	Call write_variable_in_CASE_NOTE("                                                                            ")
	Call write_variable_in_CASE_NOTE(HH_memb_07)
	IF exempt_reason_07 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(exempt_reason_07)
	IF cit_proof_07 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(cit_proof_07)
	IF ID_proof_07 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(ID_proof_07)
End if
If HH_memb_08 <> "" then
	Call write_variable_in_CASE_NOTE("                                                                            ")
	Call write_variable_in_CASE_NOTE(HH_memb_08)
	IF exempt_reason_08 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(exempt_reason_08)
	IF cit_proof_08 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(cit_proof_08)
	IF ID_proof_08 <> "(select or type here)" then Call write_variable_in_CASE_NOTE(ID_proof_08)
End if
Call write_variable_in_CASE_NOTE("--------------------------------------------------------------------------------")
Call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")