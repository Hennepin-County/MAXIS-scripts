run_locally = true
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
function assess_button_pressed()
    If ButtonPressed = dlg_one_button Then
        pass_one = false
        pass_two = False
        pass_three = false
        pass_four = false
        pass_five = false
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = true
        show_two = true
        show_three = true
        show_four = true
        show_five = true
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true
    End If
    If ButtonPressed = dlg_two_button Then
        pass_one = true
        pass_two = False
        pass_three = false
        pass_four = false
        pass_five = false
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = true
        show_three = true
        show_four = true
        show_five = true
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true
    End If
    If ButtonPressed = dlg_three_button Then
        pass_one = true
        pass_two = true
        pass_three = false
        pass_four = false
        pass_five = false
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = true
        show_four = true
        show_five = true
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true
    End If
    If ButtonPressed = dlg_four_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = false
        pass_five = false
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = true
        show_five = true
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true
    End If
    If ButtonPressed = dlg_five_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = false
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = true
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true
    End If
    If ButtonPressed = dlg_six_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = true
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = false
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true
    End If
    If ButtonPressed = dlg_seven_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = true
        pass_six = true
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = false
        show_six = false
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true
    End If
    If ButtonPressed = dlg_eight_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = true
        pass_six = true
        pass_seven = true
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = false
        show_six = false
        show_seven = false
        show_eight = true
        show_nine = true
        show_ten = true
    End If
    If ButtonPressed = dlg_nine_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = true
        pass_six = true
        pass_seven = true
        pass_eight = true
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = false
        show_six = false
        show_seven = false
        show_eight = false
        show_nine = true
        show_ten = true
    End If
    If ButtonPressed = dlg_ten_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = true
        pass_six = true
        pass_seven = true
        pass_eight = true
        pass_nine = true
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = false
        show_six = false
        show_seven = false
        show_eight = false
        show_nine = false
        show_ten = true
    End If
end function
'use these to move between the loops ... somehow'
pass_one = False
pass_two = False
pass_three = False
pass_four = False
pass_five = False
pass_six = False
pass_seven = False
pass_eight = False
pass_nine = False
pass_ten = False

show_one = True
show_two = True
show_three = True
show_four = True
show_five = True
show_six = True
show_seven = True
show_eight = True
show_nine = True
show_ten = True
Do
    full_err_msg = ""
    Do
        Do
            Do
                Do
                    Do
                        Do
                            If show_one = true Then
                                Dialog1 = ""
                                BeginDialog Dialog1, 0, 0, 466, 310, "Dialog 1 - Personal"
                                  ButtonGroup ButtonPressed
                                    CancelButton 410, 290, 50, 15
                                    Text 55, 295, 45, 10, "1 - Personal"
                                    PushButton 105, 295, 35, 10, "2 - JOBS", dlg_two_button
                                    PushButton 145, 295, 35, 10, "3 - BUSI", dlg_three_button
                                    PushButton 185, 295, 35, 10, "4 - CSES", dlg_four_button
                                    PushButton 225, 295, 35, 10, "5 - UNEA", dlg_five_button
                                    PushButton 265, 295, 35, 10, "6 - Other", dlg_six_button
                                    PushButton 305, 295, 50, 10, "7 - Interview", dlg_seven_button
                                    PushButton 370, 290, 35, 15, "NEXT", go_to_next_page
                                EndDialog
                                Dialog Dialog_One
                                cancel_without_confirmation

                                Call assess_button_pressed
                                If ButtonPressed = go_to_next_page Then pass_one = true
                            End If
                        Loop Until pass_one = TRUE
                        If show_two = true Then
                            Dialog1 = ""
                            BeginDialog Dialog1, 0, 0, 466, 310, "Dialog 2 - JOBS"
                              ButtonGroup ButtonPressed
                                CancelButton 410, 290, 50, 15
                                PushButton 55, 295, 45, 10, "1 - Personal", dlg_one_button
                                Text 105, 295, 35, 10, "2 - JOBS"
                                PushButton 145, 295, 35, 10, "3 - BUSI", dlg_three_button
                                PushButton 185, 295, 35, 10, "4 - CSES", dlg_four_button
                                PushButton 225, 295, 35, 10, "5 - UNEA", dlg_five_button
                                PushButton 265, 295, 35, 10, "6 - Other", dlg_six_button
                                PushButton 305, 295, 50, 10, "7 - Interview", dlg_seven_button
                                PushButton 370, 290, 35, 15, "NEXT", go_to_next_page
                            EndDialog
                            Dialog Dialog_two
                            cancel_without_confirmation

                            Call assess_button_pressed
                            If ButtonPressed = go_to_next_page Then pass_two = true
                        End If
                    Loop Until pass_two = true
                    If show_three = true Then
                        Dialog1 = ""
                        BeginDialog Dialog1, 0, 0, 466, 310, "Dialog 3 - BUSI"
                          ButtonGroup ButtonPressed
                            CancelButton 410, 290, 50, 15
                            PushButton 55, 295, 45, 10, "1 - Personal", dlg_one_button
                            PushButton 105, 295, 35, 10, "2 - JOBS", dlg_two_button
                            Text 145, 295, 35, 10, "3 - BUSI"
                            PushButton 185, 295, 35, 10, "4 - CSES", dlg_four_button
                            PushButton 225, 295, 35, 10, "5 - UNEA", dlg_five_button
                            PushButton 265, 295, 35, 10, "6 - Other", dlg_six_button
                            PushButton 305, 295, 50, 10, "7 - Interview", dlg_seven_button
                            PushButton 370, 290, 35, 15, "NEXT", go_to_next_page
                        EndDialog
                        Dialog Dialog_three
                        cancel_without_confirmation

                        Call assess_button_pressed
                        If ButtonPressed = go_to_next_page Then pass_three = true
                    End If
                Loop Until pass_three = true
                If show_four = true Then
                    Dialog1 = ""
                    BeginDialog Dialog1, 0, 0, 466, 310, "Dialog 4 - CSES"
                      ButtonGroup ButtonPressed
                        CancelButton 410, 290, 50, 15
                        PushButton 55, 295, 45, 10, "1 - Personal", dlg_one_button
                        PushButton 105, 295, 35, 10, "2 - JOBS", dlg_two_button
                        PushButton 145, 295, 35, 10, "3 - BUSI", dlg_three_button
                        Text 185, 295, 35, 10, "4 - CSES"
                        PushButton 225, 295, 35, 10, "5 - UNEA", dlg_five_button
                        PushButton 265, 295, 35, 10, "6 - Other", dlg_six_button
                        PushButton 305, 295, 50, 10, "7 - Interview", dlg_seven_button
                        PushButton 370, 290, 35, 15, "NEXT", go_to_next_page
                    EndDialog
                    Dialog Dialog_four
                    cancel_without_confirmation

                    Call assess_button_pressed
                    If ButtonPressed = go_to_next_page Then pass_four = true
                End If
            Loop Until pass_four = true
            If show_five = true Then
                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 466, 310, "Dialog 5 - UNEA"
                  ButtonGroup ButtonPressed
                    CancelButton 410, 290, 50, 15
                    PushButton 55, 295, 45, 10, "1 - Personal", dlg_one_button
                    PushButton 105, 295, 35, 10, "2 - JOBS", dlg_two_button
                    PushButton 145, 295, 35, 10, "3 - BUSI", dlg_three_button
                    PushButton 185, 295, 35, 10, "4 - CSES", dlg_four_button
                    Text 225, 295, 35, 10, "5 - UNEA"
                    PushButton 265, 295, 35, 10, "6 - Other", dlg_six_button
                    PushButton 305, 295, 50, 10, "7 - Interview", dlg_seven_button
                    PushButton 370, 290, 35, 15, "NEXT", go_to_next_page
                EndDialog
                Dialog Dialog_five
                cancel_without_confirmation

                Call assess_button_pressed
                If ButtonPressed = go_to_next_page Then pass_five = true
            End If
        Loop Until pass_five = true
        If show_six = true Then
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 466, 310, "Dialog 6 - Other"
              ButtonGroup ButtonPressed
                CancelButton 410, 290, 50, 15
                PushButton 55, 295, 45, 10, "1 - Personal", dlg_one_button
                PushButton 105, 295, 35, 10, "2 - JOBS", dlg_two_button
                PushButton 145, 295, 35, 10, "3 - BUSI", dlg_three_button
                PushButton 185, 295, 35, 10, "4 - CSES", dlg_four_button
                PushButton 225, 295, 35, 10, "5 - UNEA", dlg_five_button
                Text 265, 295, 35, 10, "6 - Other"
                PushButton 305, 295, 50, 10, "7 - Interview", dlg_seven_button
                PushButton 370, 290, 35, 15, "NEXT", go_to_next_page
            EndDialog
            Dialog Dialog_six
            cancel_without_confirmation

            Call assess_button_pressed
            If ButtonPressed = go_to_next_page Then pass_six = true
        End If
    Loop Until pass_six = true
    If show_seven = true Then
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 466, 310, "Dialog 7 - Interview"
          ButtonGroup ButtonPressed
            CancelButton 410, 290, 50, 15
            PushButton 55, 295, 45, 10, "1 - Personal", dlg_one_button
            PushButton 105, 295, 35, 10, "2 - JOBS", dlg_two_button
            PushButton 145, 295, 35, 10, "3 - BUSI", dlg_three_button
            PushButton 185, 295, 35, 10, "4 - CSES", dlg_four_button
            PushButton 225, 295, 35, 10, "5 - UNEA", dlg_five_button
            PushButton 265, 295, 35, 10, "6 - Other", dlg_six_button
            Text 305, 295, 50, 10, "7 - Interview"
            PushButton 370, 290, 35, 15, "NEXT", go_to_next_page
        EndDialog
        Dialog Dialog_seven
        cancel_without_confirmation

        Call assess_button_pressed
        If ButtonPressed = go_to_next_page Then pass_seven = true
    End If
Loop Until pass_seven = true
