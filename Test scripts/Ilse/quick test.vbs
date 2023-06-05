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

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone
'next_revw_date = "01/01/19"
'
'last_day_of_revw = dateadd("d", -1, next_revw_date) & "" 	'blank space added to make vorianble to make a string
'
'revw_start_date = dateadd("M", - 6, next_revw_date)	'blank space added to make vorianble to make a string
''revw_start_date = right("0" & DatePart("YYYY", next_revw_date), 2)
'
'msgbox last_day_of_revw & vbcr & revw_start_date
cash_review_year = "19"
cash_review_year = abs(cash_review_year) - 1
MsgBox cash_review_year

'start_date = "02/01/2018"   'start and end service agreement dates
'end_date = "06/18/2018"
'total_units = datediff("D", start_date, end_date)
'msgbox total_units

'MsgBox(client_age("08/18/1963"))
'Function client_age(client_DOB)
'    Dim CurrentDate, Years, ThisYear, Months, ThisMonth, Days
'    CurrentDate = CDate(client_DOB)
'    Years = DateDiff("yyyy", CurrentDate, Date)
'    ThisYear = DateAdd("yyyy", Years, CurrentDate)
'    Months = DateDiff("m", ThisYear, Date)
'    ThisMonth = DateAdd("m", Months, ThisYear)
'    Days = DateDiff("d", ThisMonth, Date)
'
'    Do While (Days < 0) Or (Months < 0)
'        If Days < 0 Then
'            Months = Months - 1
'            ThisMonth = DateAdd("m", Months, ThisYear)
'            Days = DateDiff("d", ThisMonth, Date)
'        End If
'        If Months < 0 Then
'            Years = Years - 1
'            ThisYear = DateAdd("yyyy", Years, CurrentDate)
'            Months = DateDiff("m", ThisYear, Date)
'            ThisMonth = DateAdd("m", Months, ThisYear)
'            Days = DateDiff("d", ThisMonth, Date)
'        End If
'    Loop
'    client_age = Years & "y/" & Months & "m/" & Days
'End Function


stopscript