'Required for statistical purposes==========================================================================================
name_of_script = "CA - PERSON SEARCH.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 10                      'manual run time in seconds
STATS_denomination = "I"                   'I is for Item - based on search criteria
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
'
'
'END CHANGELOG BLOCK =======================================================================================================
Dialog1 = ""												'updated the dialog name because the dialog functionality requires this
BeginDialog Dialog1, 0, 0, 146, 80, "Select Search Type"
  DropListBox 10, 14, 110, 5, "Select"+chr(9)+"SSN"+chr(9)+"Name and DOB"+chr(9)+"Last Name and First Name", Search_type
  ButtonGroup ButtonPressed
    OkButton 10, 50, 45, 20
  ButtonGroup ButtonPressed
    CancelButton 60, 50, 45, 20
EndDialog


DO
			err_msg = ""
			Dialog Dialog1
			If ButtonPressed = 0 then Stopscript
			If Search_type = "Select" THEN err_msg = err_msg & vbCr & "Please select SSN or Name and DOB."
			If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr
LOOP UNTIL err_msg = ""

CALL file_selection_system_dialog(xmlPath, ".xml")

'Grabbing some information from the xml file for testing report
file_name = replace(xmlPath, "T:\Eligibility Support\EA_ADAD\EA_ADAD_Common\CASE ASSIGNMENT\MNB_XML_files\", "")
running_error = ""

Set ObjFso = CreateObject("Scripting.FileSystemObject")
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.async = False

' Load the XML file
xmlDoc.load(xmlPath)

Dim FirstName, LastName, SSN, DOB

' Check if the XML file is loaded successfully
If xmlDoc.parseError.errorCode <> 0 Then
   MsgBox "Error in XML: " & xmlDoc.parseError.reason
   running_error = running_error & vbCr & "Error in XML: " & xmlDoc.parseError.reason
Else
   Set firstNameNode = xmlDoc.selectSingleNode("//ns4:FirstName")
   Set lastNameNode = xmlDoc.selectSingleNode("//ns4:LastName")
   Set ssnNode = xmlDoc.selectSingleNode("//ns4:SSN")
   Set dobNode = xmlDoc.selectSingleNode("//ns4:DOB")

   Set caseNumberNode = xmlDoc.selectSingleNode("//ns4:CaseNumber")		'ADDING THIS INFORMATION FOR TESTING TO CAPTURE SOME DETAILS FOR TESTING REPORT
   Set applicationIdNode = xmlDoc.selectSingleNode("//io4:ApplicationID")
   Set submitDateNode = xmlDoc.selectSingleNode("//io4:SubmitDate")

   If Not caseNumberNode Is Nothing Then
		case_number_from_form = caseNumberNode.Text
   Else
		case_number_from_form = ""
   End If

   If Not applicationIdNode Is Nothing Then
		confrimation_number_from_form = applicationIdNode.Text
   Else
		confrimation_number_from_form = ""
   End If

   If Not submitDateNode Is Nothing Then
		appl_date_from_form = submitDateNode.Text
		appl_date_from_form = replace(appl_date_from_form, "T", " at ")
		appl_date_from_form = replace(appl_date_from_form, "Z", "")
   Else
		appl_date_from_form = ""
   End If																'----------------------------------------------------------------------------

   If Not firstNameNode Is Nothing Then
       firstName = firstNameNode.Text
   Else
       firstName = "Not found"
	   running_error = running_error & vbCr & "First Name - NOT FOUND"
   End If

   If Not lastNameNode Is Nothing Then
       lastName = lastNameNode.Text
   Else
       lastName = "Not found"
	   running_error = running_error & vbCr & "Last Name - NOT FOUND"
   End If

   If Not dobNode Is Nothing Then
       dob = dobNode.Text
   Else
       dob = "Not found"
	   running_error = running_error & vbCr & "DOB - NOT FOUND"
   End If

   If Not ssnNode Is Nothing Then
       ssn = ssnNode.Text
   Else
       ssn = "Not found"
	   running_error = running_error & vbCr & "SSN - NOT FOUND"
   End If


'   MsgBox "First Name: " & firstName & vbCrLf & _
'          "Last Name: " & lastName & vbCrLf & _
'          "SSN: " & ssn
End If

' Release the XML DOM object when you're done
Set xmlDoc = Nothing



'CONNECTING TO MAXIS, STOPPING THE CASE NUMBER FROM CARRYING THROUGH
EMConnect ""
MAXIS_case_number = "________"

'NAVIGATING TO THE SCREEN---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'checking for active MAXIS session
call check_for_MAXIS(True)
call navigate_to_MAXIS_screen("pers", "____")

IF Search_type = "Name and DOB" Then
   EMWriteScreen lastName, 04, 36
   EMWriteScreen firstName, 10, 36
   EMWriteScreen left(dob, 2), 11, 53
   EMWriteScreen mid(dob, 4, 2), 11, 56
   EMWriteScreen right(dob, 4), 11, 59
End If
IF Search_type = "SSN" Then
   EMWriteScreen left(ssn, 3), 14, 36
   EMWriteScreen mid(ssn, 5, 2), 14, 40
   EMWriteScreen right(ssn, 4), 14, 43
End IF
IF Search_type = "Last Name and First Name" Then
   EMWriteScreen lastName, 04, 36
   EMWriteScreen firstName, 10, 36
End If
Transmit


'THIS PART HAS BEEN ADDED TO ASSIST WITH GATHERING INFORMATION FOR TESTING AND CREATING A TESTING REPORT
search_time = timer-start_time
EMReadScreen panel_title, 78, 2, 2
panel_title = trim(panel_title)


Dialog1 = ""
BeginDialog Dialog1, 0, 0, 411, 200, "Search Script Testing"
  DropListBox 150, 60, 60, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No", search_found
  EditBox 15, 90, 380, 15, notes_about_search
  EditBox 15, 135, 380, 15, testing_report
  If case_number_from_form <> "" Then DropListBox 210, 175, 85, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Unsure", case_number_on_form_correct
  ButtonGroup ButtonPressed
    OkButton 350, 175, 50, 15
  Text 10, 10, 335, 10, "Thank you for testing the new MNIT script for assisting with searching from MNBenefits applications."
  Text 10, 25, 335, 10, "You can use this dialog to record any notes, comments, issues, or ideas from running this new script."
  GroupBox 10, 45, 390, 70, "Did the script function well?"
  Text 15, 65, 135, 10, "Was the person found with this search?"
  Text 15, 80, 75, 10, "Notes about search:"
  Text 15, 125, 200, 10, "Detail any issues with the search or during the script run:"
  If case_number_from_form <> "" Then
	Text 15, 160, 240, 10, "Case Number provided by resident on MN Benefits form: " & case_number_from_form
	Text 15, 180, 195, 10, "Does this appear to be the Case Number for this resident?"
  End If
EndDialog

dialog Dialog1


txt_file_name = "mnit_pers_search_script_test_" & replace(Search_type, " ", "_") & "_" & confrimation_number_from_form & "_" & replace(replace(replace(now, "/", "_"),":", "_")," ", "_") & ".txt"
script_test_info_file_path = t_drive &"\Eligibility Support\Assignments\Script Testing Logs\"  & txt_file_name

Call find_user_name(script_run_worker)

'CREATING THE TESTING REPORT
With (CreateObject("Scripting.FileSystemObject"))
	'Creating an object for the stream of text which we'll use frequently
	Dim objTextStream

	Set objTextStream = .OpenTextFile(script_test_info_file_path, ForWriting, true)

	objTextStream.WriteLine "SCRIPT Run Date and Time: " & now
	objTextStream.WriteLine "Script run be: " & script_run_worker
	objTextStream.WriteLine "File Name Selected: " & file_name
	objTextStream.WriteLine "Confirmation Number: " & confrimation_number_from_form
	objTextStream.WriteLine "APPL Date: " & appl_date_from_form
	objTextStream.WriteLine "Length of script run: " & search_time
	objTextStream.WriteLine "Panel Title at the End: " & panel_title
	objTextStream.WriteLine "-------------------------------------------------"
	objTextStream.WriteLine "Search Type: " & Search_type
	objTextStream.WriteLine "Was the search found: " & search_found
	objTextStream.WriteLine "Search Notes: " &  notes_about_search
	objTextStream.WriteLine "-------------------------------------------------"
	If case_number_from_form <> "" Then
		objTextStream.WriteLine "Case Number from Form: " & case_number_from_form
		objTextStream.WriteLine "Does this Case Number appear to be accurate: " & case_number_on_form_correct
	Else
		objTextStream.WriteLine "No Case Number was found on the XML."
	End If
	' objTextStream.WriteLine ": " &
	objTextStream.WriteLine "-------------------------------------------------"
	objTextStream.WriteLine "Testing Information: " & testing_report
	objTextStream.WriteLine "Running Error: " & running_error
	objTextStream.WriteLine "-------------------------------------------------"

	objTextStream.Close
End With
'END OF TESTING REPORT

script_end_procedure("")
