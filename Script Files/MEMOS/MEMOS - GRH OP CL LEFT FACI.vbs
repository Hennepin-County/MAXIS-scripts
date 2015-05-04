OPTION EXPLICIT

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
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

DIM
DIM
DIM

DIALOG----------------------------------------------------------------------------------------------------
BeginDialog GRH_OP_LEAVING_FACI_dialog, 0, 0, 306, 360, "GRH overpayment due to leaving facility dialog"
  EditBox 55, 5, 40, 15, case_number
  EditBox 200, 5, 55, 15, 
  EditBox 70, 40, 230, 15, facility_address_line_01
  EditBox 70, 60, 230, 15, facility_address_line_02
  EditBox 70, 80, 80, 15, facility_city
  EditBox 155, 80, 25, 15, facility_state
  EditBox 185, 80, 45, 15, facility_zip
  EditBox 90, 115, 210, 15, Edit14
  EditBox 60, 135, 45, 15, discovery_date
  EditBox 180, 135, 40, 15, established_date
  EditBox 45, 160, 45, 15, overpayment_date_01
  EditBox 110, 160, 30, 15, 
  EditBox 185, 160, 45, 15, 
  EditBox 255, 160, 45, 15, Edit21
  EditBox 45, 180, 45, 15, 
  EditBox 110, 180, 30, 15, 
  EditBox 185, 180, 45, 15, 
  EditBox 255, 180, 45, 15, 
  EditBox 45, 200, 45, 15, 
  EditBox 110, 200, 30, 15, 
  EditBox 185, 200, 45, 15, 
  EditBox 255, 200, 45, 15, 
  EditBox 65, 250, 235, 15, address_line_01
  EditBox 65, 270, 235, 15, address_line_02
  EditBox 65, 290, 80, 15, address_city
  EditBox 150, 290, 25, 15, address_state
  EditBox 180, 290, 45, 15, address_zip
  CheckBox 40, 315, 95, 10, "Send overpayment to DHS", send_OP_to _DHS_check
  CheckBox 155, 315, 125, 10, "Set TIKL to recheck case in 30 days", set_TIKL_check
  EditBox 95, 330, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 195, 330, 50, 15
    CancelButton 250, 330, 50, 15
  Text 40, 30, 265, 10, "**FACILITY ADDRESS WHERE THE OVERPAYMENT MEMO WILL BE SENT**"
  Text 120, 140, 60, 10, "Established date:"
  Text 65, 240, 235, 10, "**COUNTY ADDRESS WHERE THE OVERPAYMENT WILL BE SENT**"
  Text 35, 335, 60, 10, "Worker signature:"
  Text 5, 165, 40, 10, "Date of OP:"
  Text 5, 275, 55, 10, "Address Line 2:"
  Text 95, 165, 15, 10, "Amt:"
  Text 5, 255, 55, 10, "Address Line 1:"
  Text 145, 165, 40, 10, "Date of OP:"
  Text 5, 315, 25, 10, "**OR**"
  Text 235, 165, 15, 10, "Amt:"
  Text 10, 295, 50, 10, "City/State/Zip:"
  Text 5, 185, 40, 10, "Date of OP:"
  Text 20, 85, 50, 10, "City/State/Zip:"
  Text 95, 185, 15, 10, "Amt:"
  Text 5, 120, 85, 10, "Reason for overpayment:"
  Text 145, 185, 40, 10, "Date of OP:"
  Text 5, 10, 45, 10, "Case number:"
  Text 235, 185, 15, 10, "Amt:"
  Text 5, 140, 55, 10, "Discovery date:"
  Text 5, 205, 40, 10, "Date of OP:"
  Text 105, 10, 90, 10, "Total overpayment amount:"
  Text 95, 205, 15, 10, "Amt:"
  Text 5, 65, 65, 10, "FACI ADDR Line 2:"
  Text 145, 205, 40, 10, "Date of OP:"
  Text 5, 45, 60, 10, "FACI ADDR line 1:"
  Text 235, 205, 15, 10, "Amt:"
  GroupBox 0, 0, 305, 100, ""
  GroupBox 0, 110, 305, 110, ""
  GroupBox 0, 230, 305, 120, ""
EndDialog
