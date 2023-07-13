'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - REPAIR EX PARTE PHASE 2.vbs"
start_time = timer

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
call changelog_update("07/07/2023", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
SQL_Case_Number = right("00000000" & MAXIS_case_number, 8)

admin_run = False
If user_ID_for_validation = "CALO001" Then admin_run = True

user_id_to_update = user_ID_for_validation

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 211, 175, "Ex Parte Phase 2 Completion Repair"
  If admin_run = True Then EditBox 150, 5, 50, 15, user_id_to_update
  If admin_run = False Then Text 150, 10, 50, 10, user_id_to_update
  DropListBox 10, 40, 170, 45, "Yes - HC is Eligibile"+chr(9)+"No - HC has been Closed"+chr(9)+"Case already Transferred out of County", elig_or_not
  ButtonGroup ButtonPressed
    OkButton 95, 150, 50, 15
    CancelButton 150, 150, 50, 15
  Text 10, 10, 100, 10, "Case Number: " & MAXIS_case_number
  Text 115, 10, 30, 10, "User ID"
  Text 10, 30, 125, 10, "Is HC Approved as Eligible ongoing?"
  Text 10, 65, 185, 25, "This script is only to record that Ex Parte has been completed on a Phase 2 Case if the Eligibility Summary functionality did not operate as expected."
  Text 10, 95, 40, 10, "Script will:"
  Text 15, 110, 180, 10, "- Enter Ex Parte Completed CASE/NOTE (Eligible Only)"
  Text 15, 120, 165, 10, "- Enter Standard Policy CASE/NOTE (Eligible Only)"
  Text 15, 130, 145, 10, "- Record Completion in Ex parte Data List"
EndDialog

Do
	err_msg = ""

	dialog Dialog1
	cancel_confirmation

Loop until err_msg = ""

If elig_or_not = "Case already Transferred out of County" Then
	' MsgBox "STOP - YOU ARE GOING TO UPDATE"
	objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET Phase2HSR = '" & user_id_to_update & "', ExParteAfterPhase2 = 'Case not in 27' WHERE CaseNumber = '" & SQL_Case_Number & "'"

	'Creating objects for Access
	Set objUpdateConnection = CreateObject("ADODB.Connection")
	Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
End If

If elig_or_not = "No - HC has been Closed" Then
	' MsgBox "STOP - YOU ARE GOING TO UPDATE"
	objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET Phase2HSR = '" & user_id_to_update & "', ExParteAfterPhase2 = 'Closed HC' WHERE CaseNumber = '" & SQL_Case_Number & "'"

	'Creating objects for Access
	Set objUpdateConnection = CreateObject("ADODB.Connection")
	Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
End If

If elig_or_not = "Yes - HC is Eligibile" Then
	' MsgBox "STOP - YOU ARE GOING TO UPDATE"
	objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET Phase2HSR = '" & user_id_to_update & "', ExParteAfterPhase2 = 'Approved as Ex Parte' WHERE CaseNumber = '" & SQL_Case_Number & "'"

	'Creating objects for Access
	Set objUpdateConnection = CreateObject("ADODB.Connection")
	Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection


	Call start_a_blank_CASE_NOTE

	Call write_variable_in_CASE_NOTE(CM_plus_1_mo & "/" & CM_plus_1_yr & " Ex Parte Renewal Complete - HEALTH CARE")
	Call write_variable_in_CASE_NOTE("Approved HC for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " renewal.")
	Call write_variable_in_CASE_NOTE("Renewal was completed using the Ex Parte process.")
	Call write_variable_in_CASE_NOTE("   - This is also known as an 'Auto Renewal'.")
	Call write_variable_in_CASE_NOTE("-------------------------------------------------")
	Call write_variable_in_CASE_NOTE("All eligibility details are in a previous NOTE.")
	If MSP_approvals_only = True and MSP_memo_success = True Then
		Call write_variable_in_CASE_NOTE("MEMO sent to resident with Approval Information.")
		Call write_variable_in_CASE_NOTE("     (Manual MEMO required for MSP only case.)")
	End If
	' Call write_variable_in_CASE_NOTE("")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)

	If developer_mode = True Then
		MsgBox "Ex Parte NOTE REVIEW"			'TESTING OPTION'
		PF10
		MsgBox "Ex Parte note Gone?"
	End If
	PF3

	Next_REPT_year = CM_plus_1_yr				'We only need this for the CASE/NOTE returning HC to standard policy - the asset part.
	Next_REPT_year = Next_REPT_year*1
	Next_REPT_year = Next_REPT_year + 1
	Next_REPT_year = Next_REPT_year & ""

	Call start_a_blank_CASE_NOTE

	Call write_variable_in_CASE_NOTE("~*~*~ MA STANDARD POLICY APPLIES TO THIS CASE ~*~*~")
	Call write_variable_in_CASE_NOTE("Case has completed a Health Care Eligibility Review (Annual Renewal)")
	Call write_variable_in_CASE_NOTE("Review completed for " & CM_plus_1_mo & "/" & CM_plus_1_yr & ")")
	Call write_variable_in_CASE_NOTE("**************************************************************************")
	Call write_variable_in_CASE_NOTE("Any future changes or CICs reported can be acted on,")
	Call write_variable_in_CASE_NOTE("even if they result in negative action for Health Care eligibility.")
	Call write_variable_in_CASE_NOTE("- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -")
	Call write_variable_in_CASE_NOTE("Continuous Coverage no longer applies to this case.")
	Call write_variable_in_CASE_NOTE("**************************************************************************")
	Call write_variable_in_CASE_NOTE("If enrollees on this case have an asset limit:")
	Call write_variable_in_CASE_NOTE("Assets will NOT be counted until after " & CM_plus_1_mo & "/01/" & Next_REPT_year & ".")
	Call write_variable_in_CASE_NOTE("Asset panels should reflect known information.")
	Call write_variable_in_CASE_NOTE("Review other CASE/NOTEs for detail on if the DHS-8445 was sent.")
	Call write_variable_in_CASE_NOTE("- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -")
	Call write_variable_in_CASE_NOTE("Details about this determination can be found in")
	Call write_variable_in_CASE_NOTE("        ONESource in the COVID-19 Page.")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
	If developer_mode = True Then
		MsgBox "Standard Polity NOTE REVIEW"			'TESTING OPTION'
		PF10
		MsgBox "Standard Policy Note Gone?"
	End If

	PF3


End If

Call script_end_procedure_with_error_report("Phase 2 case updated and work recorded in Data List.")