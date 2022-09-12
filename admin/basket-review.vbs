'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - Basket Review.vbs"
start_time = timer
STATS_counter = 0                          'sets the stats counter at one
STATS_manualtime = 60                      'manual run time in seconds
STATS_denomination = "I"       			   'C is for each CASE
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

' 'Reading Locally held FuncLib in leiu of issues with connecting to GitHub
' Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
' Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs")
' text_from_the_other_script = fso_command.ReadAll
' fso_command.Close
' Execute text_from_the_other_script

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("04/15/2020", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'THE SCRIPT =================================================================================================================

'Timing for 229 baskets - 0:05:50
'Timing for all of Hennepin - 0:45:40
EMConnect ""
check_for_MAXIS(TRUE)

get_county_code

basket_list = "X127EE1~X127EE2~X127EE3~X127EE4~X127EE5~X127EE6~X127EE7~X127EL2~X127EL3~X127EL4~X127EL5~X127EL6~X127EL7~X127EL8~X127EL9~X127EN1~X127EN2~X127EN3~X127EN4~X127EQ1~X127EQ4~X127EQ5~X127EQ8~X127EQ9~X127EN5~X127EG4~X127ED8~X127EH8~X127EQ3~X127EQ2~X127EG5~X127EJ1~X127EH9~X127EM2~X127FE6~X127F3D~X127ES1~X127ES2~X127ES4~X127ES5~X127ES6~X127ES7~X127ES9~X127ET1~X127ET2~X127ET3~X127ET4~X127ET5~X127ET6~X127ET7~X127ET9~X127ET8~X127ES8~X127ES3~X127F3H~X127F4E~X127EZ2~X127EZ7~X127FB7~X127EZ5~X127EZ6~X127EZ9~X127EZ4~X127EZ3~X127EZ1~X127EZ8~X127EJ6~X127FE5~X127EK1~X127EK2~X127EJ7~X127EJ8~X127EK3~X127EH6~X127EM1~X127FI7~X127EK4~X127EK5~X127FH5~X127EN7~X127EK6~X127EK9~X127FI2~X127FG3~X127EM8~X127EM9~X127EJ4~X127EJ5~X127EF8~X127EF9~X127EG9~X127EG0~X127EH1~X127EH7~X127EH2~X127EH3~X127EN6~X127FH4~X127EP3~X127EP4~X127EP5~X127EP9~X127F3U~X127F3V~X127FE7~X127FE8~X127FE9~X127CCR~X127CCA~X127FA5~X127FA6~X127FA7~X127FA8~X127FB1~X127F3S~X127FA9~X127F4A~X127F4B~X127FI1~X127FI3~X127EX4~X127EX5~X127FF1~X127FF2~X127FH3~X127FI6~X127EN8~X127EN9~X127EQ6~X127EQ7~X127EP1~X127EP2~X127FE2~X127FE3~X127FG5~X127FG9~X127F3E~X127F3J~X127F3N~X127LE1~X127SH1~X127AN1~X127EHD~X127EA0~X127EAK~X127FF6~X127FF7~X127ER7~X127FF8~X127FF9~X127FG6~X127FG7~X127EM3~X127EM4~X127EW7~X127EW8~X127NP0~X127NPC~X127FF4~X127FF5~X127FG1~X127EW6~X1274EC~X127FG2~X127EW4~X127F3F~X127F3K~X127F3P~X127EH4~X127EH5~X127EK7~X127EK8~X127EP6~X127EP7~X127EP8~X127EX2~X127FJ2~X127FF3~X127EX3~X127EM5~X127EM6~X127EX1~X127FD4~X127FD5~X127FH6~X127FD6~X127FD7~X127EZ0~X127EU5~X127EX7~X127EU6~X127EY1~X127FJ5~X127EY2~X127F3W~X127FA1~X127EU8~X127FA4~X127F3T~X127F3X~X127FA2~X127EU7~X127F3R~X127EX8~X127EX9~X127F3Z~X127EV1~X127FC2~X127EL1~X127EV2~X127EV4~X127EV3~X127FB8~X127ER8~X127F3B~X127FB4~X127F3A~X127F4C~X127F4F~X127F4D~X127FB3~X127ER9~X127EW2~X127EW3~X127EU1~X127EU3~X127EU2~X127EY8~X127EY9"
' basket_list = "X127EE2~X127EN5~X127EG4~X127ED8~X127EH8~X127EQ3~X127EQ2"
' basket_list = "X127EQ3~X127EQ2"
' basket_list = "X127ET9~X127ET8~X127ES8~X127ES3"
basket_list = "X127EE2~X127EN5~X127EG4~X127ED8~X127EH8~X127EQ3~X127EQ2~X127ET9~X127ET8~X127ES8~X127ES3~X127EP6~X127EP7~X127EP8"

basket_list = "X127EQ3~X127EH8~X127EK9~X127ES8~X127ES3~X127EG4~X127EN5~X127ED8~X127EQ2~X127ET9~X127ET8~X127EK4~X127EN6~X127EF8~X127EL5~X127FA5~X127EF9~X127EA0~X127FE6~X127EE2~X127EM4~X127FE7~X127FF5~X127ET4~X127ES1~X127EL2~X127EL4~X127EL8~X127FF6~X127FF7~X127ET5~X127EP6~X127EP7~X127EP8"

basket_array = split(basket_list, "~")

' call create_array_of_all_active_x_numbers_in_county(basket_array, two_digit_county_code)

Call back_to_SELF
basket_count = Ubound(basket_array) + 1

Do
    Do
        err_msg = ""
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 176, 60, "Confirm Basket Review Run"
          ButtonGroup ButtonPressed
            OkButton 65, 40, 50, 15
            CancelButton 120, 40, 50, 15
            PushButton 10, 45, 50, 10, "Basket List", basket_list_button
          Text 10, 10, 130, 25, "This script will review " & basket_count & " baskets for PND2 and PND1 counts and general information and output it to a Report"
        EndDialog

        dialog Dialog1
        cancel_without_confirmation

        If ButtonPressed = basket_list_button Then
            Dialog1 = ""
            x_pos = 10
            y_pos = 10
            BeginDialog Dialog1, 0, 0, 640, 300, "Confirm Basket Review Run"
              ButtonGroup ButtonPressed
                OkButton 580, 275, 50, 15
              For each basket in basket_array
                Text x_pos, y_pos, 40, 10, basket
                x_pos = x_pos + 45
                If x_pos = 640 Then
                    x_pos = 10
                    y_pos = y_pos + 15
                End If
              Next
            EndDialog

            dialog Dialog1
            cancel_without_confirmation

            err_msg = "LOOP"

        End If
    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

basket_sheet_name = "Basket Report on " & date
basket_sheet_name = replace(basket_sheet_name, "/", "-")

Set ObjExcel = CreateObject("Excel.Application")

ObjExcel.Visible = True
Set objWorkbook = ObjExcel.Workbooks.Add()
ObjExcel.DisplayAlerts = True

ObjExcel.ActiveSheet.Name = basket_sheet_name

ObjExcel.Cells(1,1).Value = "Basket"
ObjExcel.Cells(1,2).Value = "Basket Name"
ObjExcel.Cells(1,3).Value = "PND2 Pages"
ObjExcel.Cells(1,4).Value = "PND2 Case Count"
ObjExcel.Cells(1,5).Value = "PND1 Case Count"
ObjExcel.Rows(1).Font.Bold = TRUE

excel_row = 2
For each basket in basket_array
    Call back_to_SELF
    STATS_counter = STATS_counter + 1

    Call navigate_to_MAXIS_screen("REPT", "PND2")
    EMWriteScreen basket, 21, 13
    transmit

    row = 1
    col = 1
    EMSearch "The REPT:PND2 Display Limit Has Been Reached.", row, col
    If row <> 0 Then transmit

    EMReadScreen basket_total_pages, 2, 3, 79
    basket_total_pages = trim(basket_total_pages)
    basket_total_pages = basket_total_pages * 1
    EMReadScreen basket_name, 40, 3, 11
    pnd2_row = 7
    pnd2_count = 0
    Do
        EMReadScreen case_nbr, 8, pnd2_row, 5
        EMReadScreen case_name, 10, pnd2_row, 16
        case_name = trim(case_name)
        ' MsgBox "Before - Case Number: ~" & case_nbr & "~" & vbNewLine & "Count: " & pnd2_count
        If case_nbr <> prev_case_nbr Then pnd2_count = pnd2_count + 1
        ' MsgBox "Case Number: ~" & case_nbr & "~" & vbNewLine & "Count: " & pnd2_count

        If case_name = "" Then
            PF8
            EMReadScreen end_of_list, 9, 24, 14
            If end_of_list = "         " Then
                case_name = "something"
            Else
                PF7
            End If
        End If
        pnd2_row = pnd2_row + 1
        If pnd2_row = 19 Then
            ' MsgBox pnd2_count
            PF8
            EMReadScreen end_of_list, 9, 24, 14
            If end_of_list = "LAST PAGE" Then Exit Do
            pnd2_row = 7
        End If
        prev_case_nbr = case_nbr
    Loop until case_name = ""

    ObjExcel.Cells(excel_row, 1).Value = basket
    ObjExcel.Cells(excel_row, 2).Value = trim(basket_name)
    ObjExcel.Cells(excel_row, 3).Value = basket_total_pages
    If basket_total_pages > 29 Then ObjExcel.Range(ObjExcel.Cells(excel_row, 1), ObjExcel.Cells(excel_row, 5)).Interior.ColorIndex = 6
    If basket_total_pages > 34 Then ObjExcel.Range(ObjExcel.Cells(excel_row, 1), ObjExcel.Cells(excel_row, 5)).Interior.ColorIndex = 45
    If basket_total_pages > 39 Then ObjExcel.Range(ObjExcel.Cells(excel_row, 1), ObjExcel.Cells(excel_row, 5)).Interior.ColorIndex = 3
    ObjExcel.Cells(excel_row, 4).Value = pnd2_count

    Call navigate_to_MAXIS_screen("REPT", "PND1")
    EMWriteScreen basket, 21, 13
    transmit

    pnd1_row = 7
    pnd1_count = 0
    Do
        EMReadScreen case_nbr, 8, pnd1_row, 3
        case_nbr = trim(case_nbr)

        If case_nbr <> "" Then pnd1_count = pnd1_count + 1

        pnd1_row = pnd1_row + 1
        If pnd1_row = 19 Then
            PF8
            EMReadScreen end_of_list, 9, 24, 14
            If end_of_list = "LAST PAGE" Then Exit Do
            pnd1_row = 7
        End If
    Loop until case_nbr = ""
    ObjExcel.Cells(excel_row, 5).Value = pnd1_count

    excel_row = excel_row + 1
Next

For col_to_autofit = 1 to 5
    ObjExcel.columns(col_to_autofit).AutoFit()
Next

call script_end_procedure("Report generated")
'============================================================================================================================
