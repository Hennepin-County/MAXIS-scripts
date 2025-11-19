'STATS GATHERING=============================================================================================================
name_of_script = "ACTIONS - PROCESS MNBENEFITS APPLICATION.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 30            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the anctual manualtime based on time study
'To do - update manual time calculations
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomonation appliable to your script.
'END OF stats block==========================================================================================================

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

'CHANGELOG BLOCK===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
CALL changelog_update("11/18/25", "Initial version.", "Mark Riegle, Hennepin County") 'REPLACE with release date and your name.

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DEFINING CONSTANTS, VARIABLES, ARRAYS, AND BUTTONS===========================================================================

'Buttons Defined
'--Navigation buttons

'--Other buttons
' instructions_btn
' file_selection_button


'Defining variables


'Dimming variables
Dim folderPath, application_ID, fso, folder, fileList, file, xml_file_path, script_testing

'Initialize variables
script_testing = false



'DEFINING FUNCTIONS===========================================================================
' function dialog()

' end function
' Dim 


'THE SCRIPT=================================================================================================================
EMConnect "" 'Connects to BlueZone

'Initial Dialog - Instructions
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 281, 220, "Process MNBenefits Application"
  Text 10, 5, 245, 20, "Script Purpose: This script performs a PERS search, APPLs the case using the MNBenefits XML file details, and then moves the case to PND2 status."
  ButtonGroup ButtonPressed
    PushButton 10, 30, 65, 15, "Script Instructions", instructions_btn
  GroupBox 5, 50, 270, 100, "Choose one option"
  CheckBox 15, 65, 250, 10, "Enter 10-digit application ID for the XML file. Then press the Search button.", application_ID_checkbox
  EditBox 20, 80, 55, 15, application_ID
  ButtonGroup ButtonPressed
    PushButton 80, 80, 40, 15, "Search", search_button
  CheckBox 15, 100, 215, 10, "Press button below to locate XML file using Windows Explorer", manual_file_select_checkbox
  ButtonGroup ButtonPressed
    PushButton 25, 115, 85, 15, "Open Windows Explorer", file_selection_button
  CheckBox 15, 135, 205, 10, "Enter application details manually to perform PERS search", enter_app_manually_checkbox
  Text 10, 160, 95, 10, "XML File Path (if applicable):"
  EditBox 10, 170, 265, 15, XML_file_path
  ButtonGroup ButtonPressed
    OkButton 185, 200, 45, 15
    CancelButton 230, 200, 45, 15
EndDialog

DO
	DO
		err_msg = ""					'establishing value of variable, this is necessary for the Do...LOOP
		dialog Dialog1				'main dialog
		cancel_without_confirmation
    If ButtonPressed = file_selection_button then 
      call file_selection_system_dialog(XML_file_path, ".xml")
      err_msg = "LOOP"
    End If
    If ButtonPressed = instructions_btn Then
      'To do - update with instructions 
      Call open_URL_in_browser("https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/") 
      err_msg = "LOOP"
    End IF 

    If trim(application_ID) <> "" and len(application_ID) = 10 and IsNumeric(application_ID) then
      If ButtonPressed = search_button Then

        If script_testing = false Then

          folderPath = "T:\Eligibility Support\EA_ADAD\EA_ADAD_Common\CASE ASSIGNMENT\MNB_XML_files\"

          Set fso = CreateObject("Scripting.FileSystemObject")
          Set folder = fso.GetFolder(folderPath)
          XML_file_found = False

          For Each file In folder.Files
            If InStr(1, file.Name, "_" & application_ID & "_", vbTextCompare) > 0 Then
              msgbox "Found: " & file.Path
              XML_file_path = file.Path
              XML_file_found = True
              err_msg = "LOOP"
              Exit For
            End If
          Next
          If XML_file_found = False Then
            err_msg = err_msg & vbCr & "* The script was unable to locate a MNBenefits XML file with the application ID you provided. You must click the 'Select File' button and select the XML file or manually enter the file path in the field."
          End If
        Else
          startTime = Timer
          folderPath = "C:\Users\mari001\OneDrive - Hennepin County\Desktop\XML Files"

          Set fso = CreateObject("Scripting.FileSystemObject")
          Set folder = fso.GetFolder(folderPath)
          XML_file_found = False
          file_count = 0

          For Each file In folder.Files
            If InStr(1, file.Name, "_" & application_ID & "_", vbTextCompare) > 0 Then
              msgbox "Found: " & file.Path
              XML_file_path = file.Path
              XML_file_found = True
              Exit For
            End If
            file_count = file_count + 1
          Next
          If XML_file_found = False Then
            err_msg = err_msg & vbCr & "* The script was unable to locate a MNBenefits XML file with the application ID you provided. You must click the 'Select File' button and select the XML file or manually enter the file path in the field."
          End If
          endTime = Timer
          duration = endTime - startTime
          msgbox "Search took " & duration & " seconds. It evaluated " & file_count & " files."
        End If
      End If
    End If
    If application_ID_checkbox + manual_file_select_checkbox + enter_app_manually_checkbox = 0 then err_msg = err_msg & vbCr & "* You must check the box for one of the options to locate the MNBenefits XML File or enter application details manually."
    If application_ID_checkbox + manual_file_select_checkbox + enter_app_manually_checkbox > 1 then err_msg = err_msg & vbCr & "* You must only check ONE checkbox."
    If application_ID_checkbox = 1 and (trim(application_ID) = "" OR len(application_ID) <> 10 OR IsNumeric(application_ID) = false) then err_msg = err_msg & vbCr & "* You must enter the 10-digit application ID."
    If manual_file_select_checkbox = 1 and trim(file_path) = "" then err_msg = err_msg & vbCr & "* You must click the 'Select File' button and select the XML file or manually enter the file path in the field."
		If err_msg <> "" and err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in