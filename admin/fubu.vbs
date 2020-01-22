'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - Review All the Scripts.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "I"       		'C is for each CASE
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
function add_page_buttons_to_dialog(page_variable, items_per_page, total_items, dlg_vert)

    If total_items > items_per_page AND total_items < items_per_page*2+1 Then
        If page = 1 Then
            Text 12, dlg_vert + 2, 10, 10, "1"
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
        Else
            PushButton 5, dlg_vert, 10, 10, "1", page_one_btn
            Text 22, dlg_vert + 2, 10, 10, "2"
        End If
    ElseIf total_items > items_per_page*2 AND total_items < items_per_page*3+1 Then
        If page = 1 Then
            Text 12, dlg_vert + 2, 10, 10, "1"
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
        ElseIf page = 2 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            Text 22, dlg_vert + 2, 10, 10, "2"
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
        Else
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            Text 32, dlg_vert + 2, 10, 10, "3"
        End If
    ElseIf total_items > items_per_page*3 AND total_items < items_per_page*4+1 Then
        If page = 1 Then
            Text 12, dlg_vert + 2, 10, 10, "1"
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
        ElseIf page = 2 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            Text 22, dlg_vert + 2, 10, 10, "2"
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
        ElseIf page = 3 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            Text 32, dlg_vert + 2, 10, 10, "3"
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
        Else
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            Text 42, dlg_vert + 2, 10, 10, "4"
        End If
    ElseIf total_items > items_per_page*4 AND total_items < items_per_page*5+1 Then
        If page = 1 Then
            Text 12, dlg_vert + 2, 10, 10, "1"
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
        ElseIf page = 2 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            Text 22, dlg_vert + 2, 10, 10, "2"
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
        ElseIf page = 3 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            Text 32, dlg_vert + 2, 10, 10, "3"
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
        ElseIf page = 4 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            Text 42, dlg_vert + 2, 10, 10, "4"
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
        Else
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            Text 52, dlg_vert + 2, 10, 10, "5"
        End If
    ElseIf total_items > items_per_page*5 AND total_items < items_per_page*6+1 Then       'SIX Buttons
        If page = 1 Then
            Text 12, dlg_vert + 2, 10, 10, "1"
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
        ElseIf page = 2 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            Text 22, dlg_vert + 2, 10, 10, "2"
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
        ElseIf page = 3 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            Text 32, dlg_vert + 2, 10, 10, "3"
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
        ElseIf page = 4 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            Text 42, dlg_vert + 2, 10, 10, "4"
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
        ElseIf page = 5 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            Text 52, dlg_vert + 2, 10, 10, "5"
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
        ElseIf page = 6 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            Text 62, dlg_vert + 2, 10, 10, "6"
        End If
    ElseIf total_items > items_per_page*6 AND total_items < items_per_page*7+1 Then       'SEVEN Buttons
        If page = 1 Then
            Text 12, dlg_vert + 2, 10, 10, "1"
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
        ElseIf page = 2 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            Text 22, dlg_vert + 2, 10, 10, "2"
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
        ElseIf page = 3 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            Text 32, dlg_vert + 2, 10, 10, "3"
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
        ElseIf page = 4 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            Text 42, dlg_vert + 2, 10, 10, "4"
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
        ElseIf page = 5 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            Text 52, dlg_vert + 2, 10, 10, "5"
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
        ElseIf page = 6 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            Text 62, dlg_vert + 2, 10, 10, "6"
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
        ElseIf page = 7 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            Text 72, dlg_vert + 2, 10, 10, "7"
        End If
    ElseIf total_items > items_per_page*7 AND total_items < items_per_page*8+1 Then      'EIGHT Buttons
        If page = 1 Then
            Text 12, dlg_vert + 2, 10, 10, "1"
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
        ElseIf page = 2 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            Text 22, dlg_vert + 2, 10, 10, "2"
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
        ElseIf page = 3 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            Text 32, dlg_vert + 2, 10, 10, "3"
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
        ElseIf page = 4 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            Text 42, dlg_vert + 2, 10, 10, "4"
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
        ElseIf page = 5 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            Text 52, dlg_vert + 2, 10, 10, "5"
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
        ElseIf page = 6 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            Text 62, dlg_vert + 2, 10, 10, "6"
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
        ElseIf page = 7 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            Text 72, dlg_vert + 2, 10, 10, "7"
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
        ElseIf page = 8 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            Text 82, dlg_vert + 2, 10, 10, "8"
        End If
    ElseIf total_items > items_per_page*8 AND total_items < items_per_page*9+1 Then      'NINE Buttons
        If page = 1 Then
            Text 12, dlg_vert + 2, 10, 10, "1"
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
        ElseIf page = 2 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            Text 22, dlg_vert + 2, 10, 10, "2"
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
        ElseIf page = 3 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            Text 32, dlg_vert + 2, 10, 10, "3"
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
        ElseIf page = 4 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            Text 42, dlg_vert + 2, 10, 10, "4"
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
        ElseIf page = 5 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            Text 52, dlg_vert + 2, 10, 10, "5"
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
        ElseIf page = 6 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            Text 62, dlg_vert + 2, 10, 10, "6"
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
        ElseIf page = 7 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            Text 72, dlg_vert + 2, 10, 10, "7"
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
        ElseIf page = 8 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            Text 82, dlg_vert + 2, 10, 10, "8"
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
        ElseIf page = 9 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            Text 92, dlg_vert + 2, 10, 10, "9"
        End If
    ElseIf total_items > items_per_page*9 AND total_items < items_per_page*10+1 Then      'TEN Buttons
        If page = 1 Then
            Text 12, dlg_vert + 2, 10, 10, "1"
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
        ElseIf page = 2 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            Text 22, dlg_vert + 2, 10, 10, "2"
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
        ElseIf page = 3 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            Text 32, dlg_vert + 2, 10, 10, "3"
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
        ElseIf page = 4 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            Text 42, dlg_vert + 2, 10, 10, "4"
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
        ElseIf page = 5 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            Text 52, dlg_vert + 2, 10, 10, "5"
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
        ElseIf page = 6 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            Text 62, dlg_vert + 2, 10, 10, "6"
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
        ElseIf page = 7 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            Text 72, dlg_vert + 2, 10, 10, "7"
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
        ElseIf page = 8 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            Text 82, dlg_vert + 2, 10, 10, "8"
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
        ElseIf page = 9 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            Text 92, dlg_vert + 2, 10, 10, "9"
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
        ElseIf page = 10 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            Text 102, dlg_vert + 2, 10, 10, "10"
        End If
    ElseIf total_items > items_per_page*10 AND total_items < items_per_page*11+1 Then      'ELEVEN Buttons'
        If page = 1 Then
            Text 12, dlg_vert + 2, 10, 10, "1"
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
        ElseIf page = 2 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            Text 22, dlg_vert + 2, 10, 10, "2"
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
        ElseIf page = 3 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            Text 32, dlg_vert + 2, 10, 10, "3"
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
        ElseIf page = 4 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            Text 42, dlg_vert + 2, 10, 10, "4"
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
        ElseIf page = 5 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            Text 52, dlg_vert + 2, 10, 10, "5"
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
        ElseIf page = 6 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            Text 62, dlg_vert + 2, 10, 10, "6"
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
        ElseIf page = 7 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            Text 72, dlg_vert + 2, 10, 10, "7"
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
        ElseIf page = 8 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            Text 82, dlg_vert + 2, 10, 10, "8"
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
        ElseIf page = 9 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            Text 92, dlg_vert + 2, 10, 10, "9"
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
        ElseIf page = 10 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            Text 102, dlg_vert + 2, 10, 10, "10"
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
        ElseIf page = 11 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            Text 112, dlg_vert + 2, 10, 10, "11"
        End If
    ElseIf total_items > items_per_page*11 AND total_items < items_per_page*12+1 Then      'TWELVE Buttons
        If page = 1 Then
            Text 12, dlg_vert + 2, 10, 10, "1"
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
        ElseIf page = 2 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            Text 22, dlg_vert + 2, 10, 10, "2"
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
        ElseIf page = 3 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            Text 32, dlg_vert + 2, 10, 10, "3"
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
        ElseIf page = 4 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            Text 42, dlg_vert + 2, 10, 10, "4"
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
        ElseIf page = 5 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            Text 52, dlg_vert + 2, 10, 10, "5"
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
        ElseIf page = 6 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            Text 62, dlg_vert + 2, 10, 10, "6"
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
        ElseIf page = 7 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            Text 72, dlg_vert + 2, 10, 10, "7"
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
        ElseIf page = 8 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            Text 82, dlg_vert + 2, 10, 10, "8"
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
        ElseIf page = 9 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            Text 92, dlg_vert + 2, 10, 10, "9"
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
        ElseIf page = 10 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            Text 102, dlg_vert + 2, 10, 10, "10"
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
        ElseIf page = 11 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            Text 112, dlg_vert + 2, 10, 10, "11"
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
        ElseIf page = 12 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            Text 122, dlg_vert + 2, 10, 10, "12"
        End If
    ElseIf total_items > items_per_page*12 AND total_items < items_per_page*13+1 Then      'THIRTEEN Buttons
        If page = 1 Then
            Text 12, dlg_vert + 2, 10, 10, "1"
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
        ElseIf page = 2 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            Text 22, dlg_vert + 2, 10, 10, "2"
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
        ElseIf page = 3 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            Text 32, dlg_vert + 2, 10, 10, "3"
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
        ElseIf page = 4 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            Text 42, dlg_vert + 2, 10, 10, "4"
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
        ElseIf page = 5 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            Text 52, dlg_vert + 2, 10, 10, "5"
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
        ElseIf page = 6 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            Text 62, dlg_vert + 2, 10, 10, "6"
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
        ElseIf page = 7 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            Text 72, dlg_vert + 2, 10, 10, "7"
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
        ElseIf page = 8 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            Text 82, dlg_vert + 2, 10, 10, "8"
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
        ElseIf page = 9 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            Text 92, dlg_vert + 2, 10, 10, "9"
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
        ElseIf page = 10 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            Text 102, dlg_vert + 2, 10, 10, "10"
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
        ElseIf page = 11 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            Text 112, dlg_vert + 2, 10, 10, "11"
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
        ElseIf page = 12 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            Text 122, dlg_vert + 2, 10, 10, "12"
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
        ElseIf page = 13 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            Text 132, dlg_vert + 2, 10, 10, "13"
        End If
    ElseIf total_items > items_per_page*13 Then                             'FOURTEEN Buttons
        If page = 1 Then
            Text 12, dlg_vert + 2, 10, 10, "1"
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
            PushButton 140, dlg_vert, 10, 10, "14", page_fourteen_btn
        ElseIf page = 2 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            Text 22, dlg_vert + 2, 10, 10, "2"
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
            PushButton 140, dlg_vert, 10, 10, "14", page_fourteen_btn
        ElseIf page = 3 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            Text 32, dlg_vert + 2, 10, 10, "3"
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
            PushButton 140, dlg_vert, 10, 10, "14", page_fourteen_btn
        ElseIf page = 4 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            Text 42, dlg_vert + 2, 10, 10, "4"
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
            PushButton 140, dlg_vert, 10, 10, "14", page_fourteen_btn
        ElseIf page = 5 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            Text 52, dlg_vert + 2, 10, 10, "5"
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
            PushButton 140, dlg_vert, 10, 10, "14", page_fourteen_btn
        ElseIf page = 6 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            Text 62, dlg_vert + 2, 10, 10, "6"
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
            PushButton 140, dlg_vert, 10, 10, "14", page_fourteen_btn
        ElseIf page = 7 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            Text 72, dlg_vert + 2, 10, 10, "7"
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
            PushButton 140, dlg_vert, 10, 10, "14", page_fourteen_btn
        ElseIf page = 8 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            Text 82, dlg_vert + 2, 10, 10, "8"
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
            PushButton 140, dlg_vert, 10, 10, "14", page_fourteen_btn
        ElseIf page = 9 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            Text 92, dlg_vert + 2, 10, 10, "9"
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
            PushButton 140, dlg_vert, 10, 10, "14", page_fourteen_btn
        ElseIf page = 10 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            Text 102, dlg_vert + 2, 10, 10, "10"
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
            PushButton 140, dlg_vert, 10, 10, "14", page_fourteen_btn
        ElseIf page = 11 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            Text 112, dlg_vert + 2, 10, 10, "11"
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
            PushButton 140, dlg_vert, 10, 10, "14", page_fourteen_btn
        ElseIf page = 12 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            Text 122, dlg_vert + 2, 10, 10, "12"
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
            PushButton 140, dlg_vert, 10, 10, "14", page_fourteen_btn
        ElseIf page = 13 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            Text 132, dlg_vert + 2, 10, 10, "13"
            PushButton 140, dlg_vert, 10, 10, "14", page_fourteen_btn
        ElseIf page = 14 Then
            PushButton 10, dlg_vert, 10, 10, "1", page_one_btn
            PushButton 20, dlg_vert, 10, 10, "2", page_two_btn
            PushButton 30, dlg_vert, 10, 10, "3", page_three_btn
            PushButton 40, dlg_vert, 10, 10, "4", page_four_btn
            PushButton 50, dlg_vert, 10, 10, "5", page_five_btn
            PushButton 60, dlg_vert, 10, 10, "6", page_six_btn
            PushButton 70, dlg_vert, 10, 10, "7", page_seven_btn
            PushButton 80, dlg_vert, 10, 10, "8", page_eight_btn
            PushButton 90, dlg_vert, 10, 10, "9", page_nine_btn
            PushButton 100, dlg_vert, 10, 10, "10", page_ten_btn
            PushButton 110, dlg_vert, 10, 10, "11", page_eleven_btn
            PushButton 120, dlg_vert, 10, 10, "12", page_twelve_btn
            PushButton 130, dlg_vert, 10, 10, "13", page_thirteen_btn
            Text 142, dlg_vert + 2, 10, 10, "14"
        End If
    End If
end function

Dim page_one_btn, page_two_btn, page_three_btn, page_four_btn, page_five_btn, page_six_btn, page_seven_btn, page_eight_btn, page_nine_btn, page_ten_btn, page_eleven_btn, page_twelve_btn, page_thirteen_btn, page_fourteen_btn
excel_created = FALSE

script_repository = "C:\MAXIS-Scripts\"         'TODO - change this to not look in my test scripts when we move the COMPLETE LIST OF SCRIPTS to it's FINAL HOME.
script_list_URL = script_repository & "Test scripts\Casey\Tabs\COMPLETE LIST OF SCRIPTS.vbs"
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile(script_list_URL)
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

script_selection = "Select One..."

script_selection = "Select One..."+chr(9)+"All in Testing"



done_btn = 5000
page = 1
Do
    dlg_len = 80
    total_scripts = 0
    dlg_width = 815
    button_pos = 625
    If script_selection = "All in Testing" Then
        dlg_width = 1000
        button_pos = 810
    End If

    old_detail = detail_edit
    total_scripts = 0
    script_counter = 0

    ' If script_selection = "Release Before" OR script_selection = "Release After" Then
    '     If IsDate(detail_edit) = FALSE Then
    '         MsgBox "You have selected 'Release Before' or 'Release After' but ahve not provided a date to compare." & vbNewLine & vbNewLine & "The script has defaulted to 'ALL' and you can re-enter the selection and detail. If using a date specific selection be sure to enter a valid date."
    '         script_selection = "All"
    '         detail_edit = ""
    '     End If
    ' End If

    detail_operator = ""                    'Maybe we want to be able to select and or or when listing options. Discussion with MiKayla and Ilse'
    If Instr(detail_edit, ",") <> 0 Then
        detail_array = split(detail_edit, ",")
    ' ElseIf Instr(detail_edit, "AND") <> 0 Then
    '     detail_array = split(detail_edit, "AND")
    '     detail_operator = "AND"
    ' ElseIf Instr(detail_edit, "OR") <> 0 Then
    '     detail_array = split(detail_edit, "OR")
    '     detail_operator = "OR"
    Else
        detail_array = ARRAY(detail_edit)
    End If

    For each script_item in script_array
        script_item.show_script = FALSE
        Select Case script_selection
            Case "All"
                dlg_len = dlg_len + 20
                total_scripts = total_scripts + 1
                script_item.show_script = TRUE
            Case "All in Testing"
                If script_item.in_testing = TRUE Then
                    dlg_len = dlg_len + 20
                    total_scripts = total_scripts + 1
                    script_item.show_script = TRUE
                End If
            Case "Tags"
                For each tag_to_see in detail_array
                    tag_to_see = trim(tag_to_see)
                    tag_to_see = UCase(tag_to_see)
                    For each script_tag in script_item.tags
                        script_tag = trim(script_tag)
                        script_tag = UCase(script_tag)
                        ' MsgBox script_item.script_name & vbNewLine & "Detail Edit - " & detail_edit & vbNewLine & "Tag to see - " & tag_to_see & vbNewLine & "Script tag - " & script_tag
                        If script_tag = tag_to_see Then
                            dlg_len = dlg_len + 20
                            total_scripts = total_scripts + 1
                            script_item.show_script = TRUE
                        End If
                    Next
                Next
            Case "Key Codes"
                For each key_code_to_see in detail_array
                    key_code_to_see = trim(key_code_to_see)
                    key_code_to_see = UCase(key_code_to_see)
                    For each script_key_code in script_item.dlg_keys
                        script_key_code = trim(script_key_code)
                        script_key_code = UCase(script_key_code)
                        ' MsgBox script_item.script_name & vbNewLine & "Detail Edit - " & detail_edit & vbNewLine & "Tag to see - " & tag_to_see & vbNewLine & "Script tag - " & script_tag
                        If script_key_code = key_code_to_see Then
                            dlg_len = dlg_len + 20
                            total_scripts = total_scripts + 1
                            script_item.show_script = TRUE
                        End If
                    Next
                Next
            Case "Category"
                For each category_to_see in detail_array
                    category_to_see = trim(category_to_see)
                    category_to_see = UCase(category_to_see)
                    If category_to_see = script_item.category Then
                        dlg_len = dlg_len + 20
                        total_scripts = total_scripts + 1
                        script_item.show_script = TRUE
                    End If
                Next
            Case "Subcategory"
                For each subcategory_to_see in detail_array
                    subcategory_to_see = trim(subcategory_to_see)
                    subcategory_to_see = UCase(subcategory_to_see)
                    For each script_subcategory in script_item.subcategory
                        script_subcategory = trim(script_subcategory)
                        script_subcategory = UCase(script_subcategory)
                        ' MsgBox script_item.script_name & vbNewLine & "Detail Edit - " & detail_edit & vbNewLine & "Tag to see - " & tag_to_see & vbNewLine & "Script tag - " & script_tag
                        If script_subcategory = subcategory_to_see Then
                            dlg_len = dlg_len + 20
                            total_scripts = total_scripts + 1
                            script_item.show_script = TRUE
                        End If
                    Next
                Next
            Case "Release Before"
                If IsDate(script_item.release_date) = TRUE Then
                    If DateDiff("d", detail_edit, script_item.release_date) < 0 Then
                        dlg_len = dlg_len + 20
                        total_scripts = total_scripts + 1
                        script_item.show_script = TRUE
                    End If
                End If
            Case "Release After"
                If IsDate(script_item.release_date) = TRUE Then
                    If DateDiff("d", detail_edit, script_item.release_date) > 0 Then
                        dlg_len = dlg_len + 20
                        total_scripts = total_scripts + 1
                        script_item.show_script = TRUE
                    End If
                End If
        End Select
    Next

    If dlg_len > 385 Then dlg_len = 385
    If dlg_len = 80 Then dlg_len = 100

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, dlg_width, dlg_len, "Detailed Script Information"
      Text 10, 15, 170, 10, "Select what information you want to reviw/gather."
      Text 190, 15, 55, 10, "Script Selection:"
      DropListBox 260, 10, 130, 45, "Select One..."+chr(9)+"All"+chr(9)+"All in Testing"+chr(9)+"Tags"+chr(9)+"Key Codes"+chr(9)+"Category"+chr(9)+"Subcategory"+chr(9)+"Release Before"+chr(9)+"Release After", script_selection
      Text 400, 15, 30, 10, "which is:"
      EditBox 440, 10, 145, 15, detail_edit
      Text 600, 15, 110, 10, "Scripts Found: " & total_scripts
      Text 445, 30, 95, 10, "If a list, separate by commas"

      Text 10, 50, 45, 10, "Script Name"
      Text 185, 50, 40, 10, "Description"
      Text 390, 50, 20, 10, "Tags"
      Text 535, 50, 40, 10, "Key Codes"
      Text 590, 50, 50, 10, "Subcategories"
      GroupBox 665, 40, 145, 25, "Dates"
      Text 670, 50, 30, 10, "Release"
      Text 715, 50, 35, 10, "Hot Topic"
      Text 760, 50, 50, 10, "Retirement"
      Text 815, 50, 50, 10, "Keywords"
      If script_selection = "All in Testing" Then Text 875, 50, 100, 10, "Testing Type and criteria"

      y_pos = 65
      For each script_item in script_array
          skip_this_script = FALSE
          If page = 2 and script_counter < 15 Then skip_this_script = TRUE
          If page = 3 and script_counter < 30 Then skip_this_script = TRUE
          If page = 4 and script_counter < 46 Then skip_this_script = TRUE
          If page = 5 and script_counter < 60 Then skip_this_script = TRUE
          If page = 6 and script_counter < 75 Then skip_this_script = TRUE
          If page = 7 and script_counter < 90 Then skip_this_script = TRUE
          If page = 8 and script_counter < 105 Then skip_this_script = TRUE
          If page = 9 and script_counter < 120 Then skip_this_script = TRUE
          If page = 10 and script_counter < 135 Then skip_this_script = TRUE
          If page = 11 and script_counter < 150 Then skip_this_script = TRUE
          If page = 12 and script_counter < 165 Then skip_this_script = TRUE
          If page = 13 and script_counter < 180 Then skip_this_script = TRUE
          If page = 14 and script_counter < 195 Then skip_this_script = TRUE

          If script_item.show_script = TRUE Then
              If skip_this_script = TRUE Then
                script_counter = script_counter + 1
              Else
                  ' MsgBox "BEFORE" & vbNewLine & "Page - " & page & vbNewLine & "Script COunter - " & script_counter
                  If script_item.in_testing = TRUE Then
                      Text 10, y_pos, 160, 15, "TESTING " & script_item.category & " - " & script_item.script_name
                  Else
                      Text 10, y_pos, 160, 15, script_item.category & " - " & script_item.script_name
                  End If
                  Text 185, y_pos, 195, 15, script_item.description
                  all_the_tags = join(script_item.tags, ", ")
                  Text 390, y_pos, 140, 15, all_the_tags

                  all_the_keys = join(script_item.dlg_keys, ", ")
                  Text 535, y_pos, 50, 10, all_the_keys

                  all_the_subcats = join(script_item.subcategory, ", ")
                  Text 590, y_pos, 75, 15, all_the_subcats
                  Text 670, y_pos, 40, 10, script_item.release_date
                  Text 715, y_pos, 40, 10, script_item.hot_topic_date
                  Text 760, y_pos, 40, 10, script_item.retirement_date

                  ' all_the_keywords = join(script_item.keywords , ", ")                'This isn't in the complete list yet but when it is - we are ready
                  Text 815, y_pos, 50, 15, all_the_keywords

                  If script_selection = "All in Testing" Then

                      If IsArray(script_item.testing_criteria) = TRUE Then
                        all_the_test_criteria = join(script_item.testing_criteria, ", ")
                      Else
                        all_the_test_criteria = ""
                      End If
                      Text 875, y_pos, 100, 10, script_item.testing_category & " - " & all_the_test_criteria
                      ' Text 850, y_pos, 50, 10, all_the_test_criteria

                  End If
                  script_counter = script_counter + 1
                  y_pos = y_pos + 20
              End If
          End If

          If page = 1 and script_counter = 15 Then Exit For
          If page = 2 and script_counter = 30 Then Exit For
          If page = 3 and script_counter = 45 Then Exit For
          If page = 4 and script_counter = 60 Then Exit For
          If page = 5 and script_counter = 75 Then Exit For
          If page = 6 and script_counter = 90 Then Exit For
          If page = 7 and script_counter = 105 Then Exit For
          If page = 8 and script_counter = 120 Then Exit For
          If page = 9 and script_counter = 135 Then Exit For
          If page = 10 and script_counter = 150 Then Exit For
          If page = 11 and script_counter = 165 Then Exit For
          If page = 12 and script_counter = 180 Then Exit For
          If page = 13 and script_counter = 195 Then Exit For

      Next
      ' MsgBox "AFTER" & vbNewLine & "Page - " & page & vbNewLine & "Script COunter - " & script_counter

      ' Text 10, 85, 110, 10, "UTILITIES - POLI TEMP to WORD"
      ' Text 135, 85, 195, 15, "this is all the words and all the description because there is description madness."
      ' Text 340, 85, 90, 15, "TAG1, TAG2, TAG3"
      ' Text 435, 85, 50, 10, "U, Oe, M, C"
      ' Text 490, 85, 75, 15, "Subcategories"
      ' Text 570, 85, 40, 10, "12/30/2019"
      ' Text 615, 85, 40, 10, "12/30/2019"
      ' Text 660, 85, 40, 10, "12/30/2019"
      ' Text 715, 85, 50, 15, "Keywords here"
      If y_pos = 65 Then y_pos = 75
      ButtonGroup ButtonPressed
        call add_page_buttons_to_dialog(page, 15, total_scripts, y_pos)

        PushButton button_pos, y_pos, 70, 15, "Export to EXCEL", export_btn
        PushButton button_pos + 75, y_pos, 50, 15, "Search", search_btn
        PushButton button_pos + 130, y_pos, 50, 15, "DONE", done_btn
    EndDialog

    Dialog Dialog1

    page = 1
    If ButtonPressed = page_one_btn Then page = 1
    If ButtonPressed = page_two_btn Then page = 2
    If ButtonPressed = page_three_btn Then page = 3
    If ButtonPressed = page_four_btn Then page = 4
    If ButtonPressed = page_five_btn Then page = 5
    If ButtonPressed = page_six_btn Then page = 6
    If ButtonPressed = page_seven_btn Then page = 7
    If ButtonPressed = page_eight_btn Then page = 8
    If ButtonPressed = page_nine_btn Then page = 9
    If ButtonPressed = page_ten_btn Then page = 10
    If ButtonPressed = page_eleven_btn Then page = 11
    If ButtonPressed = page_twelve_btn Then page = 12
    If ButtonPressed = page_thirteen_btn Then page = 13
    If ButtonPressed = page_fourteen_btn Then page = 14


    If ButtonPressed = 0 Then ButtonPressed = done_btn
    If ButtonPressed = -1 Then ButtonPressed = search_btn

    If old_detail <> detail_edit Then page = 1

    ' MsgBox "The button pressed was - " & ButtonPressed
    If script_selection = "Release Before" OR script_selection = "Release After" Then
        If IsDate(detail_edit) = FALSE Then
            MsgBox "You have selected 'Release Before' or 'Release After' but ahve not provided a date to compare." & vbNewLine & vbNewLine & "The script has defaulted to 'ALL' and you can re-enter the selection and detail. If using a date specific selection be sure to enter a valid date."
            script_selection = "All"
            detail_edit = ""
            ButtonPressed = search_btn
        End If
    End If
    If ButtonPressed = export_btn Then

        detail_operator = ""                    'Maybe we want to be able to select and or or when listing options. Discussion with MiKayla and Ilse'
        If Instr(detail_edit, ",") <> 0 Then
            detail_array = split(detail_edit, ",")
        ' ElseIf Instr(detail_edit, "AND") <> 0 Then
        '     detail_array = split(detail_edit, "AND")
        '     detail_operator = "AND"
        ' ElseIf Instr(detail_edit, "OR") <> 0 Then
        '     detail_array = split(detail_edit, "OR")
        '     detail_operator = "OR"
        Else
            detail_array = ARRAY(detail_edit)
        End If

        For each script_item in script_array
            script_item.show_script = FALSE
            Select Case script_selection
                Case "All"
                    dlg_len = dlg_len + 20
                    total_scripts = total_scripts + 1
                    script_item.show_script = TRUE
                Case "All in Testing"
                    If script_item.in_testing = TRUE Then
                        dlg_len = dlg_len + 20
                        total_scripts = total_scripts + 1
                        script_item.show_script = TRUE
                    End If
                Case "Tags"
                    For each tag_to_see in detail_array
                        tag_to_see = trim(tag_to_see)
                        tag_to_see = UCase(tag_to_see)
                        For each script_tag in script_item.tags
                            script_tag = trim(script_tag)
                            script_tag = UCase(script_tag)
                            ' MsgBox script_item.script_name & vbNewLine & "Detail Edit - " & detail_edit & vbNewLine & "Tag to see - " & tag_to_see & vbNewLine & "Script tag - " & script_tag
                            If script_tag = tag_to_see Then
                                dlg_len = dlg_len + 20
                                total_scripts = total_scripts + 1
                                script_item.show_script = TRUE
                            End If
                        Next
                    Next
                Case "Key Codes"
                    For each key_code_to_see in detail_array
                        key_code_to_see = trim(key_code_to_see)
                        key_code_to_see = UCase(key_code_to_see)
                        For each script_key_code in script_item.dlg_keys
                            script_key_code = trim(script_key_code)
                            script_key_code = UCase(script_key_code)
                            ' MsgBox script_item.script_name & vbNewLine & "Detail Edit - " & detail_edit & vbNewLine & "Tag to see - " & tag_to_see & vbNewLine & "Script tag - " & script_tag
                            If script_key_code = key_code_to_see Then
                                dlg_len = dlg_len + 20
                                total_scripts = total_scripts + 1
                                script_item.show_script = TRUE
                            End If
                        Next
                    Next
                Case "Category"
                    For each category_to_see in detail_array
                        category_to_see = trim(category_to_see)
                        category_to_see = UCase(category_to_see)
                        If category_to_see = script_item.category Then
                            dlg_len = dlg_len + 20
                            total_scripts = total_scripts + 1
                            script_item.show_script = TRUE
                        End If
                    Next
                Case "Subcategory"
                    For each subcategory_to_see in detail_array
                        subcategory_to_see = trim(subcategory_to_see)
                        subcategory_to_see = UCase(subcategory_to_see)
                        For each script_subcategory in script_item.subcategory
                            script_subcategory = trim(script_subcategory)
                            script_subcategory = UCase(script_subcategory)
                            ' MsgBox script_item.script_name & vbNewLine & "Detail Edit - " & detail_edit & vbNewLine & "Tag to see - " & tag_to_see & vbNewLine & "Script tag - " & script_tag
                            If script_subcategory = subcategory_to_see Then
                                dlg_len = dlg_len + 20
                                total_scripts = total_scripts + 1
                                script_item.show_script = TRUE
                            End If
                        Next
                    Next
                Case "Release Before"
                    If IsDate(script_item.release_date) = TRUE Then
                        If DateDiff("d", detail_edit, script_item.release_date) < 0 Then
                            dlg_len = dlg_len + 20
                            total_scripts = total_scripts + 1
                            script_item.show_script = TRUE
                        End If
                    End If
                Case "Release After"
                    If IsDate(script_item.release_date) = TRUE Then
                        If DateDiff("d", detail_edit, script_item.release_date) > 0 Then
                            dlg_len = dlg_len + 20
                            total_scripts = total_scripts + 1
                            script_item.show_script = TRUE
                        End If
                    End If
            End Select
        Next

        sheet_title = "SCRIPTS sorted " & script_selection
        If excel_created = FALSE Then
            'Opening a new Excel file
            Set ObjExcel = CreateObject("Excel.Application")
            ObjExcel.Visible = True
            Set objWorkbook = ObjExcel.Workbooks.Add()
            ObjExcel.DisplayAlerts = True

            excel_created = TRUE
        Else
            ObjExcel.Worksheets.Add().Name = sheet_title
        End If

        ObjExcel.ActiveSheet.Name = sheet_title

        ObjExcel.Cells(1, 1).Value = "Script Category"
        ObjExcel.Cells(1, 2).Value = "Script Name"
        ObjExcel.Cells(1, 3).Value = "Description"
        ObjExcel.Cells(1, 4).Value = "Tags"
        ObjExcel.Cells(1, 5).Value = "Key Codes"
        ObjExcel.Cells(1, 6).Value = "Subcategory"
        ObjExcel.Cells(1, 7).Value = "Keywords"
        ObjExcel.Cells(1, 8).Value = "Release Date"
        ObjExcel.Cells(1, 9).Value = "Hot Topic Date"
        ObjExcel.Cells(1, 10).Value = "Retired Date"
        ObjExcel.Cells(1, 11).Value = "In Testing"
        ObjExcel.Cells(1, 12).Value = "Testing Category"
        ObjExcel.Cells(1, 13).Value = "Testing Criteria"
        'ADD MORE PROPERTIES HERE

        ObjExcel.Rows(1).Font.Bold = TRUE

        ' If testers_options = "All" Then detail_edit = ""
        ' ' If testers_options = "Confirmed Only" Then detail_edit = ""
        ' If InStr(detail_edit, ",") <> 0 Then
        '     detail_array = Split(detail_edit, ",")
        ' Else
        '     detail_array = array(detail_edit)
        ' End If

        row_to_use = 2

        For each script_item in script_array
            If script_item.show_script = TRUE Then
                ObjExcel.Cells(row_to_use, 1).Value = script_item.category
                ObjExcel.Cells(row_to_use, 2).Value = script_item.script_name
                ObjExcel.Cells(row_to_use, 3).Value = script_item.description
                ObjExcel.Cells(row_to_use, 4).Value = join(script_item.tags, ", ")
                ObjExcel.Cells(row_to_use, 5).Value = join(script_item.dlg_keys, ", ")
                ObjExcel.Cells(row_to_use, 6).Value = join(script_item.subcategory, ", ")
                ' ObjExcel.Cells(row_to_use, 7).Value = join(script_item.keywords, ", ")
                ObjExcel.Cells(row_to_use, 8).Value = script_item.release_date
                ObjExcel.Cells(row_to_use, 9).Value = script_item.hot_topic_date
                ObjExcel.Cells(row_to_use, 10).Value = script_item.retirement_date
                ObjExcel.Cells(row_to_use, 11).Value = script_item.in_testing
                ObjExcel.Cells(row_to_use, 12).Value = script_item.testing_category
                If IsArray(script_item.testing_criteria) = TRUE Then ObjExcel.Cells(row_to_use, 13).Value = join(script_item.testing_criteria, ", ")

                row_to_use = row_to_use + 1
            End If
        Next
            'Autofitting columns
        For col_to_autofit = 1 to 13
            ObjExcel.columns(col_to_autofit).AutoFit()
        Next

    End If

Loop until ButtonPressed = done_btn


Call script_end_procedure("")
