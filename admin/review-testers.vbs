'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - Review Testers.vbs"
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

testers_options = "Select One..."
page = 1
Do
    dlg_len = 65
    total_testers = 0
    old_detail = detail_edit
    If testers_options <> "Select One..." Then
        dlg_len = dlg_len + 20
        If testers_options = "All" Then detail_edit = ""
        ' If testers_options = "Confirmed Only" Then detail_edit = ""
        If InStr(detail_edit, ",") <> 0 Then
            detail_array = Split(detail_edit, ",")
        Else
            detail_array = array(detail_edit)
        End If

        For each tester in tester_array                         'looping through all of the testers
            show_this_tester = FALSE
            For each detail in detail_array
                detail = trim(detail)
                Select Case testers_options
                    Case "All"
                        dlg_len = dlg_len + 15
                        total_testers = total_testers + 1
                    Case "By Group"
                        For each group in tester.tester_groups
                            If UCase(detail) = UCase(group) Then
                                dlg_len = dlg_len + 15
                                total_testers = total_testers + 1
                            End If
                        Next
                    Case "By Population"
                        If UCase(detail) = UCase(tester.tester_population) Then
                            dlg_len = dlg_len + 15
                            total_testers = total_testers + 1
                        End If
                    Case "By Region"
                        If UCase(detail) = UCase(tester.tester_region) Then
                            dlg_len = dlg_len + 15
                            total_testers = total_testers + 1
                        End If
                    Case "Confirmed Only"
                        If UCase(detail_edit) = "NOT" Then
                            If tester.tester_confirmed = FALSE Then
                                dlg_len = dlg_len + 15
                                total_testers = total_testers + 1
                            End If
                        Else
                            If tester.tester_confirmed = TRUE Then
                                dlg_len = dlg_len + 15
                                total_testers = total_testers + 1
                            End If
                        End If
                    Case "By Supervisor"
                        If UCase(detail) = UCase(tester.tester_supervisor_name) Then
                            dlg_len = dlg_len + 15
                            total_testers = total_testers + 1
                        End If
                    Case "By Script"
                        For each script_name in tester.tester_scripts
                            If script_name <> "" Then
                                script_shortened = script_name
                                script_shortened = replace(script_shortened, ".vbs", "")
                                If InStr(detail, "-") = 0 Then
                                  position_of_dash = InStr(script_shortened, "-")
                                  script_shortened = right(script_shortened, (len(script_shortened) - (position_of_dash + 1)))
                                End If
                                If UCase(detail) = UCase(script_shortened) Then
                                    dlg_len = dlg_len + 15
                                    total_testers = total_testers + 1
                                End If
                            End If
                        Next
                End Select
                If tester.tester_id_number = "" Then show_this_tester = FALSE
            Next
        Next
    End If
    ' msgBox dlg_len
    If dlg_len > 385 Then dlg_len = 385

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 776, dlg_len, "Detailed Tester Information"
      Text 10, 15, 170, 10, "Select what information you want to review/gather."
      Text 190, 15, 55, 10, "Testers options:"
      DropListBox 260, 10, 130, 45, "Select One... "+chr(9)+"All"+chr(9)+"By Group"+chr(9)+"By Population"+chr(9)+"By Region"+chr(9)+"Confirmed Only"+chr(9)+"By Supervisor"+chr(9)+"By Script", testers_options
      Text 400, 15, 30, 10, "which is:"
      EditBox 440, 10, 145, 15, detail_edit
      Text 600, 15, 110, 10, "Testers Found: " & total_testers
      Text 445, 30, 95, 10, "If a list, separate by commas"

      If testers_options <> "Select One..." Then
          y_pos = 45
          Text 5, 45, 25, 10, "Name:"
          Text 120, 45, 40, 10, "Supervisor:"
          Text 215, 45, 40, 10, "Population:"
          Text 295, 45, 30, 10, "Region:"
          Text 365, 45, 30, 10, "Groups:"
          Text 535, 45, 30, 10, "Scripts:"
          Text 710, 45, 40, 10, "Confirmed:"
          y_pos = y_pos + 20

          tester_counter = 0

          For each tester in tester_array                         'looping through all of the testers
              show_this_tester = FALSE
              For each detail in detail_array
                  detail = trim(detail)
                  Select Case testers_options
                      Case "All"
                          show_this_tester = TRUE
                          tester_counter = tester_counter + 1
                      Case "By Group"
                          For each group in tester.tester_groups
                              If UCase(detail) = UCase(group) Then
                                show_this_tester = TRUE
                                tester_counter = tester_counter + 1
                              End If
                          Next
                      Case "By Population"
                          If UCase(detail) = UCase(tester.tester_population) Then
                            show_this_tester = TRUE
                            tester_counter = tester_counter + 1
                          End If
                      Case "By Region"
                          If UCase(detail) = UCase(tester.tester_region) Then
                            show_this_tester = TRUE
                            tester_counter = tester_counter + 1
                          End If
                      Case "Confirmed Only"
                          If UCase(detail_edit) = "NOT" Then
                              If tester.tester_confirmed = FALSE Then
                                show_this_tester = TRUE
                                tester_counter = tester_counter + 1
                              End If
                          Else
                              If tester.tester_confirmed = TRUE Then
                                show_this_tester = TRUE
                                tester_counter = tester_counter + 1
                              End If
                          End If
                      Case "By Supervisor"
                          If UCase(detail) = UCase(tester.tester_supervisor_name) Then
                            show_this_tester = TRUE
                            tester_counter = tester_counter + 1
                          End If
                      Case "By Script"
                          For each script_name in tester.tester_scripts
                              If script_name <> "" Then
                                  script_shortened = script_name
                                  script_shortened = replace(script_shortened, ".vbs", "")
                                  If InStr(detail, "-") = 0 Then
                                    position_of_dash = InStr(script_shortened, "-")
                                    script_shortened = right(script_shortened, len(script_shortened) - (position_of_dash + 1))
                                  End If
                                  If UCase(detail) = UCase(script_shortened) Then
                                    show_this_tester = TRUE
                                    tester_counter = tester_counter + 1
                                  End If
                              End If
                          Next
                  End Select
                  If tester.tester_id_number = "" Then show_this_tester = FALSE
              Next
              If page = 1 and tester_counter > 20 Then show_this_tester = FALSE
              If page = 2 Then
                If tester_counter < 21 Then show_this_tester = FALSE
                If tester_counter > 40 Then show_this_tester = FALSE
              End If
              If page = 3 Then
                If tester_counter < 41 Then show_this_tester = FALSE
                If tester_counter > 60 Then show_this_tester = FALSE
              End If
			  If page = 4 Then
                If tester_counter < 61 Then show_this_tester = FALSE
                If tester_counter > 80 Then show_this_tester = FALSE
              End If
			  If page = 5 Then
				  If tester_counter < 81 Then show_this_tester = FALSE
				  If tester_counter > 100 Then show_this_tester = FALSE
				End If
              If page = 6 Then if tester_counter < 101 Then show_this_tester = FALSE
              If show_this_tester = TRUE Then
                  Text 5, y_pos, 95, 10, tester.tester_full_name
                  Text 120, y_pos, 85, 10, tester.tester_supervisor_name
                  Text 215, y_pos, 60, 10, tester.tester_population
                  Text 295, y_pos, 55, 10, tester.tester_region
                  Text 365, y_pos, 165, 10, join(tester.tester_groups, ", ")
                  Text 535, y_pos, 165, 10, join(tester.tester_scripts, ", ")
                  Text 710, y_pos, 50, 10, tester.tester_confirmed & ""
                  y_pos = y_pos + 15
              End If
          Next

          ' y_pos = y_pos + 5
      Else
        y_pos = y_pos + 40
      End If
      ButtonGroup ButtonPressed
        If total_testers > 20 AND total_testers < 41 Then
            If page = 1 Then
                Text 12, y_pos + 2, 10, 10, "1"
                PushButton 20, y_pos, 10, 10, "2", page_two_btn
            Else
                PushButton 5, y_pos, 10, 10, "1", page_one_btn
                Text 22, y_pos + 2, 10, 10, "2"
            End If
        ElseIf total_testers > 40 AND total_testers < 61 Then
            If page = 1 Then
                Text 12, y_pos + 2, 10, 10, "1"
                PushButton 20, y_pos, 10, 10, "2", page_two_btn
                PushButton 30, y_pos, 10, 10, "3", page_three_btn
            ElseIf page = 2 Then
                PushButton 10, y_pos, 10, 10, "1", page_one_btn
                Text 22, y_pos + 2, 10, 10, "2"
                PushButton 30, y_pos, 10, 10, "3", page_three_btn
            Else
                PushButton 10, y_pos, 10, 10, "1", page_one_btn
                PushButton 20, y_pos, 10, 10, "2", page_two_btn
                Text 32, y_pos + 2, 10, 10, "3"
            End If
        ElseIf total_testers > 60 AND total_testers < 81 Then
            If page = 1 Then
                Text 12, y_pos + 2, 10, 10, "1"
                PushButton 20, y_pos, 10, 10, "2", page_two_btn
                PushButton 30, y_pos, 10, 10, "3", page_three_btn
                PushButton 40, y_pos, 10, 10, "4", page_four_btn
            ElseIf page = 2 Then
                PushButton 10, y_pos, 10, 10, "1", page_one_btn
                Text 22, y_pos + 2, 10, 10, "2"
                PushButton 40, y_pos, 10, 10, "4", page_four_btn
                PushButton 30, y_pos, 10, 10, "3", page_three_btn
            ElseIf page = 3 Then
                PushButton 10, y_pos, 10, 10, "1", page_one_btn
                PushButton 20, y_pos, 10, 10, "2", page_two_btn
                Text 32, y_pos + 2, 10, 10, "3"
                PushButton 40, y_pos, 10, 10, "4", page_four_btn
            Else
                PushButton 10, y_pos, 10, 10, "1", page_one_btn
                PushButton 20, y_pos, 10, 10, "2", page_two_btn
                PushButton 30, y_pos, 10, 10, "3", page_three_btn
                Text 42, y_pos + 2, 10, 10, "4"
            End If
		ElseIf total_testers > 80 AND total_testers < 101 Then
            If page = 1 Then
                Text 12, y_pos + 2, 10, 10, "1"
                PushButton 20, y_pos, 10, 10, "2", page_two_btn
                PushButton 30, y_pos, 10, 10, "3", page_three_btn
				PushButton 40, y_pos, 10, 10, "4", page_four_btn
                PushButton 50, y_pos, 10, 10, "5", page_five_btn
            ElseIf page = 2 Then
                PushButton 10, y_pos, 10, 10, "1", page_one_btn
                Text 22, y_pos + 2, 10, 10, "2"
				PushButton 30, y_pos, 10, 10, "3", page_three_btn
				PushButton 40, y_pos, 10, 10, "4", page_four_btn
                PushButton 50, y_pos, 10, 10, "5", page_five_btn
            ElseIf page = 3 Then
                PushButton 10, y_pos, 10, 10, "1", page_one_btn
                PushButton 20, y_pos, 10, 10, "2", page_two_btn
                Text 32, y_pos + 2, 10, 10, "3"
				PushButton 40, y_pos, 10, 10, "4", page_four_btn
                PushButton 50, y_pos, 10, 10, "5", page_five_btn
			ElseIf page = 4 Then
                PushButton 10, y_pos, 10, 10, "1", page_one_btn
                PushButton 20, y_pos, 10, 10, "2", page_two_btn
				PushButton 30, y_pos, 10, 10, "3", page_three_btn
                Text 42, y_pos + 2, 10, 10, "4"
                PushButton 50, y_pos, 10, 10, "5", page_five_btn

            Else
                PushButton 10, y_pos, 10, 10, "1", page_one_btn
                PushButton 20, y_pos, 10, 10, "2", page_two_btn
                PushButton 30, y_pos, 10, 10, "3", page_three_btn
				PushButton 40, y_pos, 10, 10, "4", page_four_btn
                Text 52, y_pos + 2, 10, 10, "5"
            End If
		ElseIf total_testers > 100 Then
			If page = 1 Then
				Text 12, y_pos + 2, 10, 10, "1"
				PushButton 20, y_pos, 10, 10, "2", page_two_btn
				PushButton 30, y_pos, 10, 10, "3", page_three_btn
				PushButton 40, y_pos, 10, 10, "4", page_four_btn
				PushButton 50, y_pos, 10, 10, "5", page_five_btn
				PushButton 60, y_pos, 10, 10, "6", page_six_btn
			ElseIf page = 2 Then
				PushButton 10, y_pos, 10, 10, "1", page_one_btn
				Text 22, y_pos + 2, 10, 10, "2"
				PushButton 30, y_pos, 10, 10, "3", page_three_btn
				PushButton 40, y_pos, 10, 10, "4", page_four_btn
				PushButton 50, y_pos, 10, 10, "5", page_five_btn
				PushButton 60, y_pos, 10, 10, "6", page_six_btn
			ElseIf page = 3 Then
				PushButton 10, y_pos, 10, 10, "1", page_one_btn
				PushButton 20, y_pos, 10, 10, "2", page_two_btn
				Text 32, y_pos + 2, 10, 10, "3"
				PushButton 40, y_pos, 10, 10, "4", page_four_btn
				PushButton 50, y_pos, 10, 10, "5", page_five_btn
				PushButton 60, y_pos, 10, 10, "6", page_six_btn
			ElseIf page = 4 Then
				PushButton 10, y_pos, 10, 10, "1", page_one_btn
				PushButton 20, y_pos, 10, 10, "2", page_two_btn
				PushButton 30, y_pos, 10, 10, "3", page_three_btn
				Text 42, y_pos + 2, 10, 10, "4"
				PushButton 50, y_pos, 10, 10, "5", page_five_btn
				PushButton 60, y_pos, 10, 10, "6", page_six_btn

			ElseIf page = 5 Then
				PushButton 10, y_pos, 10, 10, "1", page_one_btn
				PushButton 20, y_pos, 10, 10, "2", page_two_btn
				PushButton 30, y_pos, 10, 10, "3", page_three_btn
				PushButton 40, y_pos, 10, 10, "4", page_four_btn
				Text 52, y_pos + 2, 10, 10, "5"
				PushButton 60, y_pos, 10, 10, "6", page_six_btn
			Else
				PushButton 10, y_pos, 10, 10, "1", page_one_btn
				PushButton 20, y_pos, 10, 10, "2", page_two_btn
				PushButton 30, y_pos, 10, 10, "3", page_three_btn
				PushButton 40, y_pos, 10, 10, "4", page_four_btn
				PushButton 50, y_pos, 10, 10, "6", page_five_btn
				Text 62, y_pos + 2, 10, 10, "6"
			End If
        End If
        PushButton 585, y_pos, 70, 15, "Export to EXCEL", export_btn
        PushButton 660, y_pos, 50, 15, "Search", search_btn
        PushButton 715, y_pos, 50, 15, "DONE", done_btn
    EndDialog

    Dialog Dialog1

    If ButtonPressed = page_one_btn Then page = 1
    If ButtonPressed = page_two_btn Then page = 2
    If ButtonPressed = page_three_btn Then page = 3

    If ButtonPressed = page_four_btn Then page = 4
	If ButtonPressed = page_five_btn Then page = 5
	If ButtonPressed = page_six_btn Then page = 6
    If ButtonPressed = 0 Then ButtonPressed = done_btn
    If ButtonPressed = -1 Then ButtonPressed = search_btn

    If old_detail <> detail_edit Then page = 1

    If ButtonPressed = export_btn Then
        sheet_title = "Testers sorted " & testers_options
        'Opening a new Excel file
        Set objExcel = CreateObject("Excel.Application")
        objExcel.Visible = True
        Set objWorkbook = objExcel.Workbooks.Add()
        objExcel.DisplayAlerts = True

        ObjExcel.ActiveSheet.Name = sheet_title

        ObjExcel.Cells(1, 1).Value = "Tester First Name"
        ObjExcel.Cells(1, 2).Value = "Tester Last Name"
        ObjExcel.Cells(1, 3).Value = "Tester Email"
        ObjExcel.Cells(1, 4).Value = "Supervisor Name"
        ObjExcel.Cells(1, 5).Value = "Supervisor Email"
        ObjExcel.Cells(1, 6).Value = "Population"
        ObjExcel.Cells(1, 7).Value = "Region"
        ObjExcel.Cells(1, 8).Value = "Groups"
        ObjExcel.Cells(1, 9).Value = "Scripts"
        ObjExcel.Cells(1, 10).Value = "Confirmed"

        ObjExcel.Rows(1).Font.Bold = TRUE

        If testers_options = "All" Then detail_edit = ""
        ' If testers_options = "Confirmed Only" Then detail_edit = ""
        If InStr(detail_edit, ",") <> 0 Then
            detail_array = Split(detail_edit, ",")
        Else
            detail_array = array(detail_edit)
        End If

        row_to_use = 2
        For each tester in tester_array                         'looping through all of the testers
            show_this_tester = FALSE
            For each detail in detail_array
                Select Case testers_options
                    Case "All"
                        show_this_tester = TRUE
                    Case "By Group"
                        For each group in tester.tester_groups
                            If UCase(detail) = UCase(group) Then show_this_tester = TRUE
                        Next
                    Case "By Population"
                        If UCase(detail) = UCase(tester.tester_population) Then show_this_tester = TRUE
                    Case "By Region"
                        If UCase(detail) = UCase(tester.tester_region) Then  show_this_tester = TRUE
                    Case "Confirmed Only"
                        If UCase(detail_edit) = "NOT" Then
                            If tester.tester_confirmed = FALSE Then
                              show_this_tester = TRUE
                              tester_counter = tester_counter + 1
                            End If
                        Else
                            If tester.tester_confirmed = TRUE Then
                              show_this_tester = TRUE
                              tester_counter = tester_counter + 1
                            End If
                        End If
                    Case "By Supervisor"
                        If UCase(detail) = UCase(tester.tester_supervisor_name) Then show_this_tester = TRUE
                    Case "By Script"
                        For each script_name in tester.tester_scripts
                            If script_name <> "" Then
                                script_shortened = script_name
                                script_shortened = replace(script_shortened, ".vbs", "")
                                If InStr(detail, "-") = 0 Then
                                  position_of_dash = InStr(script_shortened, "-")
                                  script_shortened = right(script_shortened, len(script_shortened) - (position_of_dash + 1))
                                End If
                                If UCase(detail) = UCase(script_shortened) Then show_this_tester = TRUE
                            End If
                        Next
                End Select
                If tester.tester_id_number = "" Then show_this_tester = FALSE
            Next
            If show_this_tester = TRUE Then
                ObjExcel.Cells(row_to_use, 1).Value = tester.tester_first_name
                ObjExcel.Cells(row_to_use, 2).Value = tester.tester_last_name
                ObjExcel.Cells(row_to_use, 3).Value = tester.tester_email
                ObjExcel.Cells(row_to_use, 4).Value = tester.tester_supervisor_name
                ObjExcel.Cells(row_to_use, 5).Value = tester.tester_supervisor_email
                ObjExcel.Cells(row_to_use, 6).Value = tester.tester_population
                ObjExcel.Cells(row_to_use, 7).Value = tester.tester_region
                ObjExcel.Cells(row_to_use, 8).Value = join(tester.tester_groups, ",")
                ObjExcel.Cells(row_to_use, 9).Value = join(tester.tester_scripts, ",")
                ObjExcel.Cells(row_to_use, 10).Value = tester.tester_confirmed & ""
                row_to_use = row_to_use + 1
            End If
            'Autofitting columns
            For col_to_autofit = 1 to 10
            	ObjExcel.columns(col_to_autofit).AutoFit()
            Next
        Next
    End If

Loop until ButtonPressed = done_btn

Call script_end_procedure("")
