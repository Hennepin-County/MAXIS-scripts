'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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

EMConnect ""
'About Arrays


const Case_Nbr_const		= 0
const Name_const			= 1
const Revw_Date_const		= 2
const Cash_1_prog_const		= 3
const Cash_1_stat_const		= 4
const Cash_2_prog_const		= 5
const Cash_2_stat_const		= 6
const FS_const				= 7
const HC_const				= 8
const EA_const				= 9
const GR_const				= 10
const IVE_const				= 11
const FIAT_const			= 12
const CC_const				= 13
const appl_date_const 		= 14
const the_last_const		= 15


DIM REPT_ACVT_ARRAY()
ReDim REPT_ACVT_ARRAY(the_last_const, 0)

Call back_to_SELF
Call navigate_to_MAXIS_screen("REPT", "ACTV")
Call write_value_and_transmit("X127EZ2", 21, 13)

row = 7
array_counter = 0
Do
	EMReadScreen case_numb_temp, 7, row, 13
	case_numb_temp = trim(case_numb_temp)

	if case_numb_temp <> "" Then
		ReDim Preserve REPT_ACVT_ARRAY(the_last_const, array_counter)

		REPT_ACVT_ARRAY(Case_Nbr_const, array_counter) = case_numb_temp

		EMReadScreen name_temp, 	20, row, 21
		EMReadScreen revw_date_temp, 8, row, 42
		EMReadScreen cash_1p_temp, 	 2, row, 51
		EMReadScreen cash_1s_temp, 	 1, row, 54
		EMReadScreen cash_2p_temp, 	 2, row, 56
		EMReadScreen cash_2s_temp, 	 1, row, 59
		EMReadScreen fs_temp, 		 1, row, 61
		EMReadScreen hc_temp, 		 1, row, 64
		EMReadScreen ea_temp, 		 1, row, 67
		EMReadScreen grh_temp,		 1, row, 70
		EMReadScreen ive_temp, 		 1, row, 73
		EMReadScreen fiat_temp, 	 1, row, 77
		EMReadScreen cc_temp, 		 1, row, 80



		REPT_ACVT_ARRAY(Name_const, array_counter) = trim(name_temp)
		REPT_ACVT_ARRAY(Revw_Date_const, array_counter) = replace(revw_date_temp, " ", "/")
		REPT_ACVT_ARRAY(Revw_Date_const, array_counter) = DateAdd("d", 0, REPT_ACVT_ARRAY(Revw_Date_const, array_counter))
		REPT_ACVT_ARRAY(Cash_1_prog_const, array_counter) = cash_1p_temp
		REPT_ACVT_ARRAY(Cash_1_stat_const, array_counter) = cash_1s_temp
		REPT_ACVT_ARRAY(Cash_2_prog_const, array_counter) = cash_2p_temp
		REPT_ACVT_ARRAY(Cash_2_stat_const, array_counter) = cash_2s_temp
		REPT_ACVT_ARRAY(FS_const		, array_counter) = fs_temp
		REPT_ACVT_ARRAY(HC_const		, array_counter) = hc_temp
		REPT_ACVT_ARRAY(EA_const		, array_counter) = ea_temp
		REPT_ACVT_ARRAY(GR_const		, array_counter) = grh_temp
		REPT_ACVT_ARRAY(IVE_const		, array_counter) = ive_temp
		REPT_ACVT_ARRAY(FIAT_const		, array_counter) = fiat_temp
		REPT_ACVT_ARRAY(CC_const		, array_counter) = cc_temp

		array_counter = array_counter + 1
	end if

	row = row + 1
	If row = 19 Then
		PF8
		row = 7
	End if
Loop until case_numb_temp = ""

' MsgBox "array_counter - " & array_counter
' array_counter = array_counter-1

for each_thing = 0 to UBound(REPT_ACVT_ARRAY, 2)
	' MsgBox "each_thing - " & each_thing
	If REPT_ACVT_ARRAY(Cash_1_stat_const, each_thing) = "A" Then
		MsgBox REPT_ACVT_ARRAY(Case_Nbr_const, each_thing) & " is on " & REPT_ACVT_ARRAY(Cash_1_prog_const, each_thing)
	End If

Next

MsgBox "UBound 1 - " & UBound(REPT_ACVT_ARRAY) & vbCr & "UBound 2 - " & UBound(REPT_ACVT_ARRAY, 2)






























call script_end_procedure("that is the end of Multi Dimensional Array detail")
'This is all about arrays.

'OUR_TEAM = Array("Ilse", "Casey", "Mark", "Megan")
' team_persons = InputBox("Who is on your team?" & vbCr & "Separate names by commas")

' OUR_TEAM = split(team_persons, ",")
Dim OUR_TEAM()
ReDim OUR_TEAM(0)
' ReDim OUR_TEAM(0)


person_count = 0
Do
	' MsgBox "my counter (person_count) is at " & person_count
	team_person = InputBox("Enter a team member name." & " (when all have been entered, type done.)")
	If team_person <> "done" Then
		ReDim Preserve OUR_TEAM(person_count)
		OUR_TEAM(person_count) = team_person
		person_count = person_count + 1
	End If
Loop until team_person = "done"
' MsgBox "person_count - " & person_count & vbCr & "UBOUND - " & UBound(OUR_TEAM)
' person_count = ""
team_person = ""

our_team_name = "Automation and Integration Team"



Do
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 191, 200, "Dialog"
		Text 10, 10, 170, 10, "Our team is the " & our_team_name
		Text 10, 20, 150, 10, "On our team we have:"
		y_pos = 35
		For each person in OUR_TEAM
			Text 15, y_pos, 20, 10, person
			y_pos = y_pos + 10
		Next
		ButtonGroup ButtonPressed
			OkButton 130, y_pos+5, 50, 15
	EndDialog


	dialog Dialog1

	' If ButtonPressed = add_another_form_button Then
	' 	'increment your array of forms - YOU MUST KEEP A COUNTER THAT HAS THE RIGHT INFORMATION or
	' 	next_index = UBOUND(OUR_TEAM) + 1
	' End If

Loop until err_msg = ""
'MsgBox "Our team is the " & our_team_name & " and we have:" & vbCr & join(OUR_TEAM, ", ")

' For each person in OUR_TEAM
' 	MsgBox person
' Next

' For each word in our_team_name
' 	MsgBox word
' Next

'MsgBox "This is person at the 1th instance - " & OUR_TEAM(0)

' For each person in OUR_TEAM
' 	person = trim(person)
' 	MsgBox person
' Next

' For pers_index = 0 to UBound(OUR_TEAM)
' 	MsgBox pers_index & vbCr & OUR_TEAM(pers_index)
' Next

' pers_index = 0
' Do
' 	MsgBox pers_index & vbCr & OUR_TEAM(pers_index)
' 	pers_index = pers_index + 1
' Loop until pers_index > UBound(OUR_TEAM)

