'Required for statistical purposes===============================================================================
name_of_script = "UTILITIES - OPEN INTERVIEW PDF.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 120                      'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
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
call changelog_update("07/29/2021", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""            'connect to MAXIS



Set objINTVWFolder = objFSO.GetFolder("T:\Eligibility Support\Assignments\Interview Notes for ECF")										'Creates an oject of the whole my documents folder
' Set objINTVWFolder = objFSO.GetFolder("T:\Eligibility Support\Assignments\Interview Notes for ECF\Archive\TRAINING REGION Interviews - NOT for ECF")										'Creates an oject of the whole my documents folder

Set colINTVWFiles = objINTVWFolder.Files																'Creates an array/collection of all the files in the folder
pdf_found = False
For Each objFile in colINTVWFiles																'looping through each file
	If objFile.Type = "Adobe Acrobat Document" Then pdf_found = True
Next
If pdf_found = False Then Call script_end_procedure("The folder to store the Interview PDFs does not have any PDF files. This means all files generated by the NOTES - Interview script have been added to ECF. Please check ECF for the document." & vbCr & vbCr & "The script will now end.")

developer_mode = False
CALL MAXIS_case_number_finder(MAXIS_case_number)                    'autofilling MAXIS case number and footer month/year
date_of_script_run = date & ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 241, 120, "Interview PDF Reopen"
  EditBox 75, 30, 60, 15, MAXIS_case_number
  EditBox 75, 50, 45, 15, date_of_script_run
  ButtonGroup ButtonPressed
    OkButton 150, 100, 40, 15
    CancelButton 195, 100, 40, 15
  Text 25, 35, 50, 10, "Case Number:"
  Text 10, 55, 65, 10, "Date of Script Run:"
  Text 10, 10, 230, 20, "To reopen a PDF from a previous run of the NOTES - Interview script, we need the case number and date of script run."
  Text 90, 75, 150, 20, "If the script cannot find the PDF, check ECF - the file may already be added."
EndDialog

'displaying the dialog to confirm or set the case number and footer month/year

Do
    err_msg = ""
    dialog Dialog1
    cancel_without_confirmation                         'power the cancel button
    CALL validate_MAXIS_case_number(err_msg, "*")       'making sure the case number is present and valid
    If IsDate(date_of_script_run) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the date you ran the script."
	If DateDiff("d", date, date_of_script_run) > 0 Then err_msg = err_msg & vbNewLine & "* The date the script was run cannot be in the future."
    If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg      'showing the error message
Loop until err_msg = ""

date_of_script_run = DateAdd("d", 0, date_of_script_run)
file_safe_date = replace(date_of_script_run, "/", "-")

pdf_doc_path = t_drive & "\Eligibility Support\Assignments\Interview Notes for ECF\Interview - " & MAXIS_case_number & " on " & file_safe_date & ".pdf"
If developer_mode = True Then pdf_doc_path = t_drive & "\Eligibility Support\Assignments\Interview Notes for ECF\Archive\TRAINING REGION Interviews - NOT for ECF\Interview - " & MAXIS_case_number & " on " & file_safe_date & ".pdf"

If objFSO.FileExists(pdf_doc_path) Then
	run_path = chr(34) & pdf_doc_path & chr(34)
	wshshell.Run run_path
	end_msg = "The PDF has been opened."
Else
	end_msg = "This PDF for Case " & MAXIS_case_number & " created by the NOTES - Interview script run on " & date_of_script_run & " could not be found. This file may have already been added to ECF."
End If

script_end_procedure(end_msg)            'all done
