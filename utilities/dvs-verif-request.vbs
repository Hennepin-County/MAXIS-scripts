'STATS GATHERING=============================================================================================================
name_of_script = "DAIL - DVS Verif Request.vbs"
start_time = timer
STATS_counter = 1                 'sets the stats counter at one
STATS_manualtime = 120            'manual run time in seconds
STATS_denomination = "C"          'C is for each case; I is for Instance, M is for member
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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County
CALL changelog_update("10/30/25", "Initial version.", "Mark Riegel, Hennepin County") 'REPLACE with release date and your name.

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT==================================================================================================================
EMConnect "" 'Connects to BlueZone
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Initial dialog to gather case number and provide script information
Dialog1 = "" 'blanking out dialog name
BeginDialog Dialog1, 0, 0, 351, 65, "DVS Verification Request"
  Text 10, 5, 270, 20, "Script Purpose: Submits a DVS verification request email. The script will pull details from MAXIS and allow user entry to add additional details for the request."
  Text 10, 30, 50, 10, "Case Number:"
  EditBox 75, 25, 55, 15, MAXIS_case_number
  Text 10, 50, 60, 10, "Worker Signature:"
  EditBox 75, 45, 140, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 260, 45, 45, 15
    CancelButton 305, 45, 45, 15
    PushButton 285, 5, 65, 15, "Script Instructions", instructions_btn
    PushButton 285, 20, 65, 15, "HSR Manual", hsr_manual_btn
EndDialog

DO
  Do
    err_msg = ""    'This is the error message handling
    Dialog Dialog1
    cancel_without_confirmation
    Call validate_MAXIS_case_number(err_msg, "*")
		If ButtonPressed = instructions_btn Then 
      'to do - update with script instructions
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/UTILITIES/UTILITIES%20-%20DVS%20VERIFICATION%20REQUEST.docx"
			err_msg = "LOOP"
		End If
		If ButtonPressed = hsr_manual_btn Then 
      'to do - update with script instructions
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Vehicles.aspx"
			err_msg = "LOOP"
		End If
  Loop until err_msg = ""
  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'Generate a list of HH members for the case number so the user can select
Call Generate_Client_List(HH_Memb_DropDown, "Select One:")

'Dialog to select HH member
Dialog1 = "" 'blanking out dialog name
BeginDialog Dialog1, 0, 0, 220, 70, "Select Household Member"
  Text 10, 5, 200, 20, "Select the household member that you want to submit the DVS verification request for:"
  DropListBox 10, 30, 200, 15, HH_Memb_DropDown, hh_memb
  ButtonGroup ButtonPressed
    OkButton 120, 50, 45, 15
    CancelButton 165, 50, 45, 15
EndDialog

DO
  Do
    err_msg = ""    'This is the error message handling
    Dialog Dialog1
    cancel_without_confirmation
    If hh_memb = "Select One:" Then err_msg = err_msg & vbCr & "* Select the household member you want to submit the DVS verification for."
  Loop until err_msg = ""
  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'Gather the birthdate for the HH memb
hh_memb_number = left(hh_memb, 2)
Call navigate_to_MAXIS_screen("STAT", "MEMB")
Call write_value_and_transmit(hh_memb_number, 20, 76)
EMReadScreen hh_memb_dob, 10, 8, 42
hh_memb_dob = replace(hh_memb_dob, " ","/")

'Starting variables for dialog
total_vehicles = 1
dialog_height = 110
vehicle_groupbox_height = 40
ok_cancel_y = 90
add_vehicle_btn_y = 80

DO
  Do
    err_msg = ""    'This is the error message handling
    Dialog1 = "" 'blanking out dialog name
    BeginDialog Dialog1, 0, 0, 470, dialog_height, "DVS Verification Request Details"
      Text 10, 5, 50, 10, "Case number:"
      Text 70, 5, 90, 10, MAXIS_case_number
      Text 10, 15, 55, 10, "Owner's Name:"
      Text 70, 15, 250, 10, hh_memb
      Text 10, 25, 55, 10, "Owner's DOB:"
      Text 70, 25, 90, 10, hh_memb_dob
      Text 10, 65, 75, 10, "Vehicle looking for:"
      EditBox 90, 60, 50, 15, vehicle_1
      GroupBox 145, 40, 320, vehicle_groupbox_height, "Optional Information (if available):"
      Text 155, 50, 300, 10, "         Alias                               VIN                        Plate #              Title #               Owner DLN"
      EditBox 150, 60, 70, 15, alias_1
      EditBox 225, 60, 75, 15, vin_1
      EditBox 305, 60, 40, 15, plate_1
      EditBox 350, 60, 50, 15, title_1
      EditBox 405, 60, 55, 15, owner_dln_1
      vehicle_fields_y = 60
      vehicle_text_y = 65
      vehicle_field_y = 60 
      If total_vehicles > 1 Then
        vehicle_text_y = vehicle_text_y + 20
        vehicle_field_y = vehicle_field_y + 20
        Text 10, vehicle_text_y, 75, 10, "Vehicle looking for:"
        EditBox 90, vehicle_field_y, 50, 15, vehicle_2
        vehicle_fields_y = vehicle_fields_y + 20
        EditBox 150, vehicle_fields_y, 70, 15, alias_2
        EditBox 225, vehicle_fields_y, 75, 15, vin_2
        EditBox 305, vehicle_fields_y, 40, 15, plate_2
        EditBox 350, vehicle_fields_y, 50, 15, title_2
        EditBox 405, vehicle_fields_y, 55, 15, owner_dln_2
      End If
      If total_vehicles > 2 Then
        vehicle_text_y = vehicle_text_y + 20
        vehicle_field_y = vehicle_field_y + 20
        Text 10, vehicle_text_y, 75, 10, "Vehicle looking for:"
        EditBox 90, vehicle_field_y, 50, 15, vehicle_3
        vehicle_fields_y = vehicle_fields_y + 20
        EditBox 150, vehicle_fields_y, 70, 15, alias_3
        EditBox 225, vehicle_fields_y, 75, 15, vin_3
        EditBox 305, vehicle_fields_y, 40, 15, plate_3
        EditBox 350, vehicle_fields_y, 50, 15, title_3
        EditBox 405, vehicle_fields_y, 55, 15, owner_dln_3
      End If
      If total_vehicles > 3 Then
        vehicle_text_y = vehicle_text_y + 20
        vehicle_field_y = vehicle_field_y + 20
        Text 10, vehicle_text_y, 75, 10, "Vehicle looking for:"
        EditBox 90, vehicle_field_y, 50, 15, vehicle_4
        vehicle_fields_y = vehicle_fields_y + 20
        EditBox 150, vehicle_fields_y, 70, 15, alias_4
        EditBox 225, vehicle_fields_y, 75, 15, vin_4
        EditBox 305, vehicle_fields_y, 40, 15, plate_4
        EditBox 350, vehicle_fields_y, 50, 15, title_4
        EditBox 405, vehicle_fields_y, 55, 15, owner_dln_4
      End If
      If total_vehicles > 4 Then
        vehicle_text_y = vehicle_text_y + 20
        vehicle_field_y = vehicle_field_y + 20
        Text 10, vehicle_text_y, 75, 10, "Vehicle looking for:"
        EditBox 90, vehicle_field_y, 50, 15, vehicle_5
        vehicle_fields_y = vehicle_fields_y + 20
        EditBox 150, vehicle_fields_y, 70, 15, alias_5
        EditBox 225, vehicle_fields_y, 75, 15, vin_5
        EditBox 305, vehicle_fields_y, 40, 15, plate_5
        EditBox 350, vehicle_fields_y, 50, 15, title_5
        EditBox 405, vehicle_fields_y, 55, 15, owner_dln_5
      End If
      ButtonGroup ButtonPressed
        OkButton 375, ok_cancel_y, 45, 15
        CancelButton 420, ok_cancel_y, 45, 15
        If total_vehicles < 5 Then
          PushButton 10, add_vehicle_btn_y, 50, 15, "Add Vehicle", add_vehicle_btn
        End If
        If total_vehicles > 1 and total_vehicles <> 5 Then
          PushButton 60, add_vehicle_btn_y, 60, 15, "Remove Vehicle", remove_vehicle_btn
        ElseIf total_vehicles = 5 Then
          PushButton 10, add_vehicle_btn_y, 60, 15, "Remove Vehicle", remove_vehicle_btn
        End If
    EndDialog

    Dialog Dialog1
    cancel_without_confirmation
    If total_vehicles = 1 Then 
      If trim(vehicle_1) = "" THEN err_msg = err_msg & vbCr & "* You must fill out the 'Vehicle looking for' field."
    ElseIf total_vehicles = 2 Then 
      If trim(vehicle_1) = "" OR trim(vehicle_2) = "" THEN err_msg = err_msg & vbCr & "* You must fill out each of the 'Vehicle looking for' fields."
    ElseIf total_vehicles = 3 Then
      If trim(vehicle_1) = "" OR trim(vehicle_2) = "" OR trim(vehicle_3) = "" THEN err_msg = err_msg & vbCr & "* You must fill out each of the 'Vehicle looking for' fields."
    ElseIf total_vehicles = 4 Then
      If trim(vehicle_1) = "" OR trim(vehicle_2) = "" OR trim(vehicle_3) = "" OR trim(vehicle_4) = "" THEN err_msg = err_msg & vbCr & "* You must fill out each of the 'Vehicle looking for' fields."
    ElseIf total_vehicles = 5 Then
      If trim(vehicle_1) = "" OR trim(vehicle_2) = "" OR trim(vehicle_3) = "" OR trim(vehicle_4) = "" OR trim(vehicle_5) = "" THEN err_msg = err_msg & vbCr & "* You must fill out each of the 'Vehicle looking for' fields."
    End If
    If ButtonPressed = add_vehicle_btn Then 
      If total_vehicles < 5 Then
        total_vehicles = total_vehicles + 1
        dialog_height = dialog_height + 20
        vehicle_groupbox_height = vehicle_groupbox_height + 20
        ok_cancel_y = ok_cancel_y + 20
        add_vehicle_btn_y = add_vehicle_btn_y + 20
        err_msg = "LOOP"
      End If
    End If
    If ButtonPressed = remove_vehicle_btn Then 
      If total_vehicles > 1 Then
        total_vehicles = total_vehicles - 1
        dialog_height = dialog_height - 20
        vehicle_groupbox_height = vehicle_groupbox_height - 20
        ok_cancel_y = ok_cancel_y - 20
        add_vehicle_btn_y = add_vehicle_btn_y - 20
        err_msg = "LOOP"
      End If
    End If
    IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
  Loop until err_msg = ""
  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

email_body = "Case #: " & MAXIS_case_number & vBcr & "Owner's Name: " & hh_memb & vbcr & "Owner's DOB: " & hh_memb_dob & vbcr & vBcr & "---" & vBcr & "Vehicle looking for: " & vehicle_1 & vbcr & "Alias: " & alias_1 & vbcr & "VIN(s): " & vin_1 & vbcr & "Plate #(s): " & plate_1 & vbcr & "Title #: " & title_1 & vbcr & "Owner DLN: " & owner_dln_1 & vbCR
If total_vehicles > 1 Then email_body = email_body & vbCR & "---" & vBcr & "Vehicle looking for: " & vehicle_2 & vbcr & "Alias: " & alias_2 & vbcr & "VIN(s): " & vin_2 & vbcr & "Plate #(s): " & plate_2 & vbcr & "Title #: " & title_2 & vbcr & "Owner DLN: " & owner_dln_2 & vbCR
If total_vehicles > 2 Then email_body = email_body & vbCR & "---" & vBcr & "Vehicle looking for: " & vehicle_3 & vbcr & "Alias: " & alias_3 & vbcr & "VIN(s): " & vin_3 & vbcr & "Plate #(s): " & plate_3 & vbcr & "Title #: " & title_3 & vbcr & "Owner DLN: " & owner_dln_3 & vbCR
If total_vehicles > 3 Then email_body = email_body & vbCR & "---" & vBcr & "Vehicle looking for: " & vehicle_4 & vbcr &  "Alias: " & alias_4 & vbcr & "VIN(s): " & vin_4 & vbcr & "Plate #(s): " & plate_4 & vbcr & "Title #: " & title_4 & vbcr & "Owner DLN: " & owner_dln_4 & vbCR
If total_vehicles > 4 Then email_body = email_body & vbCR & "---" & vBcr & "Vehicle looking for: " & vehicle_5 & vbcr & "Alias: " & alias_5 & vbcr & "VIN(s): " & vin_5 & vbcr & "Plate #(s): " & plate_5 & vbcr & "Title #: " & title_5 & vbcr & "Owner DLN: " & owner_dln_5 & vbCR

Call create_outlook_email("", "hsph.es.dvs@hennepin.us", "", "", "DVS Verification Request for MAXIS Case # " & MAXIS_case_number, 1, False, "", "", False, "", email_body, False, "", True)

'End the script.
script_end_procedure("Success! DVS Verification Request email sent to hsph.es.dvs@hennepin.us.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------10/30/2025
'--Tab orders reviewed & confirmed----------------------------------------------10/30/2025
'--Mandatory fields all present & Reviewed--------------------------------------10/30/2025
'--All variables in dialog match mandatory fields-------------------------------10/30/2025
'Review dialog names for content and content fit in dialog----------------------10/30/2025
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------10/30/2025
'--Include script category and name somewhere on first dialog-------------------10/30/2025
'--Create a button to reference instructions------------------------------------10/30/2025
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function: confirm proper punctuation is used-----N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------10/30/2025
'--MAXIS_background_check reviewed (if applicable)------------------------------10/30/2025
'--PRIV Case handling reviewed -------------------------------------------------10/30/2025
'--Out-of-County handling reviewed----------------------------------------------10/30/2025
'--script_end_procedures (w/ or w/o error messaging)----------------------------10/30/2025
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------10/30/2025
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------10/30/2025
'--Incrementors reviewed (if necessary)-----------------------------------------10/30/2025
'--Denomination reviewed -------------------------------------------------------10/30/2025
'--Script name reviewed---------------------------------------------------------10/30/2025
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------10/30/2025
'--comment Code-----------------------------------------------------------------10/30/2025
'--Update Changelog for release/update------------------------------------------10/30/2025
'--Remove testing message boxes-------------------------------------------------10/30/2025
'--Remove testing code/unnecessary code-----------------------------------------10/30/2025
'--Review/update SharePoint instructions----------------------------------------10/30/2025
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------10/30/2025
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------10/30/2025
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------10/30/2025
'--Complete misc. documentation (if applicable)---------------------------------10/30/2025
'--Update project team/issue contact (if applicable)----------------------------10/30/2025