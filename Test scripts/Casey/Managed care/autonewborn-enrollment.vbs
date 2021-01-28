'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "Autonewborn Enrollment.vbs"
start_time = timer
STATS_counter = 1			 'sets the stats counter at one
STATS_manualtime = 60			 'manual run time in seconds
STATS_denomination = "C"		 'M is for Member
'END OF stats block==============================================================================================

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
call changelog_update("04/24/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'END DIALOGS===========================================================================================================
EMConnect ""

check_for_MAXIS(True)
'TESTING CASE 265463'

call MAXIS_case_number_finder(MAXIS_case_number)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 136, 50, "Dialog"
  EditBox 60, 5, 70, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 25, 30, 50, 15
    CancelButton 80, 30, 50, 15
  Text 10, 10, 50, 10, "Case Number:"
EndDialog

Do
    err_msg = ""

    Dialog Dialog1
    If buttonpressed = cancel then stopscript

    If trim(MAXIS_case_number) = "" Then err_msg = err_msg & vbNewLine * "* Enter a case number."
    If IsNumeric(MAXIS_case_number) = False Then err_msg = err_msg & vbNewLine & "* The case number entered is not a number, please check again."
    If len(MAXIS_case_number) > 8 Then err_msg = err_msg & vbNewLine & "* The case number is too long, please check again."
Loop until err_msg = ""

MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
list_of_newborns = ""

Call navigate_to_MAXIS_screen("STAT", "MEMB")

Do
    EMReadScreen clt_age, 2, 8, 76
    clt_age = trim(clt_age)

    If clt_age = "" Then
        EMReadScreen ref_number, 2, 4, 33
        EMReadScreen pmi_number, 8, 4, 46
        EMReadScreen clt_dob,   10, 8, 42

        pmi_number = trim(pmi_number)
        clt_dob = replace(clt_dob, " ", "/")

        If list_of_newborns = "" Then
            list_of_newborns = ref_number & "|" & pmi_number & "|" & clt_dob
        Else
            list_of_newborns = list_of_newborns & "~" & ref_number & "|" & pmi_number & "|" & clt_dob
        End If
    End If

    transmit
    EMReadScreen last_member_check, 13, 24, 2
Loop until last_member_check = "ENTER A VALID"

If list_of_newborns = "" Then script_end_procedure("There is no child under 1 year in MAXIS. Add the new baby first, approve MA-11x in MAXIS, then run the script again.")
If InStr(list_of_newborns, "~") <> 0 Then
    ARRAY_OF_NEWBORNS = split(list_of_newborns, "~")
    number_of_newborns = Ubound(ARRAY_OF_NEWBORNS) + 1
Else
    number_of_newborns = 1
End If

Call navigate_to_MAXIS_screen ("STAT", "PARE")

pare_ref = ""
stat_row = 5
Do
    EMReadScreen clt_ref, 2, stat_row, 3
    EmWriteScreen clt_ref, 20, 76
    transmit

    EMReadScreen panel_check, 1, 2, 73
    If panel_check = "1" Then
        pare_row = 8
        Do
            EMReadScreen child_ref, 2, pare_row, 24
            For newborn = 1 to number_of_newborns
                If number_of_newborns = 1 Then
                    If child_ref = left(list_of_newborns, 2) Then
                        if pare_ref = "" Then
                            pare_ref = clt_ref
                        else
                            pare_ref = pare_ref & "~" & clt_ref
                        end if
                    End If
                Else
                    If child_ref = left(ARRAY_OF_NEWBORNS(newborn-1), 2) Then
                        if pare_ref = "" Then
                            pare_ref = clt_ref
                        else
                            pare_ref = pare_ref & "~" & clt_ref
                        end if
                    End If
                End If
            Next
        Loop until child_ref = "__"
    End If

    stat_row = stat_row + 1
    EMReadScreen next_clt_ref, 2, stat_row, 3
Loop until next_clt_ref = "  "
