'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - ELIGIBILITY SUMMARY.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 150                	'manual run time in seconds
STATS_denomination = "C"       			'C is for each Case
'END OF stats block=========================================================================================================

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
call changelog_update("06/06/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
EMConnect ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 186, 110, "Dialog"
  EditBox 55, 10, 125, 15, worker_editbox
  CheckBox 10, 35, 50, 10, "DWP", dwp_checkbox
  CheckBox 10, 50, 50, 10, "MFIP", MFIP_checkbox
  CheckBox 10, 65, 50, 10, "MSA", MSA_checkbox
  CheckBox 10, 80, 50, 10, "GA", GA_checkbox
  CheckBox 10, 95, 50, 10, "Cash Denial", cash_denial_checkbox
  CheckBox 70, 35, 50, 10, "GRH", GRH_checkbox
  CheckBox 70, 50, 50, 10, "IVE", IVE_checkbox
  CheckBox 70, 65, 50, 10, "EMER", EMER_checkbox
  CheckBox 70, 80, 50, 10, "SNAP", SNAP_checbox
  CheckBox 70, 95, 50, 10, "HC", HC_checkbox
  ButtonGroup ButtonPressed
    OkButton 130, 90, 50, 15
  Text 10, 15, 40, 10, "X-Numbers"
EndDialog

dialog Dialog1

x_numbers_array = split(worker_editbox, ", ")

For each basket_number in x_numbers_array

    call navigate_to_MAXIS_screen("REPT", "ACTV")

    EMWriteScreen basket_number, 21, 13
    transmit
    PF5

    actv_row = 7

    Do

        EMReadScreen MAXIS_case_number, 8, actv_row, 12
        MAXIS_case_number = trim(MAXIS_case_number)
        'MsgBox "Case: " & MAXIS_case_number
        If MAXIS_case_number <> "" Then
            EMWriteScreen "E", actv_row, 3
            transmit
            EMWriteScreen "SUMM", 20, 71
            transmit

            EMReadScreen elig_summ_check, 4, 2, 54
            If elig_summ_check <> "SUMM" Then
                MsgBox "Case: " & MAXIS_case_number & vbNewLine & "Basket: " & basket_number
            End If

            case_found = TRUE

            If dwp_checkbox = checked Then
                EMReadScreen DWP_app_date, 8, 7, 48

                If DWP_app_date = "        " Then
                    case_found = FALSE
                Else
                    If DateDiff("d", DWP_app_date, date) <> 0 Then
                        case_found = FALSE
                    Else
                        EMWriteScreen "DWP", 19, 71
                        transmit
                        EMReadScreen version_status, 8, 3, 3
                        If version_status <> "APPROVED" Then case_found = FALSE
                        EMWriteScreen "SUMM", 20, 71
                        transmit
                    End If
                End If
            End If

            If MFIP_checkbox = checked Then
                EMReadScreen MFIP_app_date, 8, 8, 48

                If MFIP_app_date = "        " Then
                    case_found = FALSE
                Else
                    If DateDiff("d", MFIP_app_date, date) <> 0 Then
                        case_found = FALSE
                    Else
                        EMWriteScreen "MFIP", 19, 71
                        transmit
                        EMReadScreen version_status, 8, 3, 3
                        If version_status <> "APPROVED" Then case_found = FALSE
                        EMWriteScreen "SUMM", 20, 71
                        transmit
                    End If
                End If
            End If

            If MSA_checkbox = checked Then
                EMReadScreen MSA_app_date, 8, 11, 48

                If MSA_app_date = "        " Then
                    case_found = FALSE
                Else
                    If DateDiff("d", MSA_app_date, date) <> 0 Then
                        case_found = FALSE
                    Else
                        EMWriteScreen "MSA", 19, 71
                        transmit
                        EMReadScreen version_status, 8, 3, 3
                        If version_status <> "APPROVED" Then case_found = FALSE
                        EMWriteScreen "SUMM", 20, 71
                        transmit
                    End If
                End If
            End If

            If GA_checkbox = checked Then
                EMReadScreen GA_app_date, 8, 12, 48

                If GA_app_date = "        " Then
                    case_found = FALSE
                Else
                    'MsgBox GA_app_date
                    If DateDiff("d", GA_app_date, date) <> 0 Then
                        case_found = FALSE
                    Else
                        EMWriteScreen "GA", 19, 71
                        transmit
                        EMReadScreen version_status, 8, 3, 3
                        If version_status <> "APPROVED" Then case_found = FALSE
                        EMWriteScreen "SUMM", 20, 70
                        transmit
                    End If
                End If
            End If

            If cash_denial_checkbox = checked Then
                EMReadScreen cash_denial_app_date, 8, 13, 48

                If cash_denial_app_date = "        " Then
                    case_found = FALSE
                Else
                    If DateDiff("d", cash_denial_app_date, date) <> 0 Then
                        case_found = FALSE
                    Else
                        EMWriteScreen "DENY", 19, 71
                        transmit
                        EMReadScreen version_status, 8, 3, 3
                        If version_status <> "APPROVED" Then case_found = FALSE
                        EMWriteScreen "SUMM", 19, 70
                        transmit
                    End If
                End If
            End If

            If GRH_checkbox = checked Then
                EMReadScreen GRH_app_date, 8, 14, 48

                If GRH_app_date = "        " Then
                    case_found = FALSE
                Else
                    If DateDiff("d", GRH_app_date, date) <> 0 Then
                        case_found = FALSE
                    Else
                        EMWriteScreen "GRH", 19, 71
                        transmit
                        EMReadScreen version_status, 8, 3, 3
                        If version_status <> "APPROVED" Then case_found = FALSE
                        EMWriteScreen "SUMM", 20, 71
                        transmit
                    End If
                End If
            End If

            If IVE_checkbox = checked Then
                EMReadScreen IVE_app_date, 8, 15, 48

                If IVE_app_date = "        " Then
                    case_found = FALSE
                Else
                    If DateDiff("d", IVE_app_date, date) <> 0 Then
                        case_found = FALSE
                    Else
                        EMWriteScreen "IVE", 19, 71
                        transmit
                        EMReadScreen version_status, 8, 3, 3
                        If version_status <> "APPROVED" Then case_found = FALSE
                        EMWriteScreen "SUMM", 20, 70
                        transmit
                    End If
                End If
            End If

            If EMER_checkbox = checked Then
                EMReadScreen EMER_app_date, 8, 16, 48

                If EMER_app_date = "        " Then
                    case_found = FALSE
                Else
                    If DateDiff("d", EMER_app_date, date) <> 0 Then
                        case_found = FALSE
                    Else
                        EMWriteScreen "EMER", 19, 71
                        transmit
                        EMReadScreen version_status, 8, 3, 3
                        If version_status <> "APPROVED" Then case_found = FALSE
                        EMWriteScreen "SUMM", 20, 70
                        transmit
                    End If
                End If
            End If

            If SNAP_checbox = checked Then
                EMReadScreen SNAP_app_date, 8, 17, 48

                If SNAP_app_date = "        " Then
                    case_found = FALSE
                Else
                    If DateDiff("d", SNAP_app_date, date) <> 0 Then
                        case_found = FALSE
                    Else
                        EMWriteScreen "FS", 19, 71
                        transmit
                        EMReadScreen version_status, 8, 3, 3
                        If version_status <> "APPROVED" Then case_found = FALSE
                        EMWriteScreen "SUMM", 19, 70
                        transmit
                    End If
                End If
            End If

            If HC_checkbox = checked Then
                EMReadScreen HC_app_date, 8, 18, 48

                If HC_app_date = "        " Then
                    case_found = FALSE
                Else
                    If DateDiff("d", HC_app_date, date) <> 0 Then
                        case_found = FALSE
                    Else
                        EMWriteScreen "HC", 19, 71
                        transmit
                        EMReadScreen version_status, 8, 3, 3
                        If version_status <> "APPROVED" Then case_found = FALSE
                        EMWriteScreen "SUMM", 20, 70
                        transmit
                    End If
                End If
            End If

            PF3
            PF3

        End If

        If case_found = TRUE Then Exit Do

        'actv_row = actv_row + 1
        If actv_row = 7 then actv_row = 8

        EMReadScreen end_check, 8, actv_row, 3

        If actv_row = 19 Then
            PF8
            actv_row = 7
        End If
        MAXIS_case_number = ""
        'MsgBox "END Check: " & end_check
    Loop until end_check = "More:  -" OR end_check = "_       "
    If case_found = TRUE Then
        Exit For
    Else
        Call back_to_SELF
    End If
Next

If case_found = TRUE Then
    Call back_to_SELF

    Call navigate_to_MAXIS_screen("ELIG", "SUMM")

    end_msg = "Case found with programs approved today  -  " & MAXIS_case_number
Else
    end_msg = "No case could be found with programs approved today. Try again with more baskets OR later in the day"
End If



script_end_procedure(end_msg)
