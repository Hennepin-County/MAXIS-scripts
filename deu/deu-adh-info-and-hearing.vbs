'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "DEU-ADH INFO & HEARING.vbs" 'BULK script that creates a list of cases that require an interview, and the contact phone numbers'
start_time = timer
STATS_counter = 1               'sets the stats counter at one 
STATS_manualtime = 120           'manual run time in seconds 
STATS_denomination = "C"        'C is for each case 
 'END OF stats block========================================================================================================= 

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib IF it already loaded once
	IF run_loCALLy = FALSE or run_loCALLy = "" THEN	   'IF the scripts are set to run loCALLy, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'IF the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End IF
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
'Example: CALL changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("11/08/2017", "Updated to add first name of memb to casenote.", "MiKayla Handley, Hennepin County")
CALL changelog_update("7/07/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'----------------------------------------------------------------------------------------------------The script
EMCONNECT ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
memb_number = "01" 

'------------------------------------------------------------------------------------------------------------Initial dialog 
BeginDialog , 0, 0, 166, 75, "ADH INFORMATION"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  EditBox 135, 5, 25, 15, memb_number
  DropListBox 80, 30, 80, 15, "Select one..."+chr(9)+"ADH waiver signed"+chr(9)+"Hearing Held", ADH_option
  ButtonGroup ButtonPressed
    OkButton 55, 55, 50, 15
    CancelButton 110, 55, 50, 15
  Text 5, 35, 75, 10, "Select an ADH option:"
  Text 5, 10, 45, 10, "Case number:"
  Text 105, 10, 30, 10, "Memb#:"
EndDialog

Do
	Do
        err_msg = "" 
		Dialog
		IF ButtonPressed = 0 then StopScript
		IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		IF IsNumeric(memb_number) = false or len(memb_number) <> 2 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid two-digit member number."
		IF ADH_option = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Please select an ADH action."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine		
    Loop until err_msg = ""	
 	CALL check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

'------------------------------------------------------------------------------------getting the case name
CALL navigate_to_MAXIS_screen("STAT", "MEMB")
EMwritescreen memb_number, 20, 76
transmit

EMReadscreen memb_name, 12, 6, 63
memb_name = replace(memb_name, "_", "") 

EMReadScreen edit_error, 2, 24, 2
edit_error = trim(edit_error)
IF edit_error <> "" THEN script_end_procedure("No Memb # matches and/ or could not access MEMB.")


		
IF ADH_option = "ADH waiver signed" then
	BeginDialog , 0, 0, 291, 140, "ADH waiver signed"
  		EditBox 110, 5, 50, 15, date_waiver_signed
  		DropListBox 210, 5, 70, 15, "Select Programs:"+chr(9)+"CASH"+chr(9)+"SNAP"+chr(9)+"CASH & SNAP", Program_droplist
  		EditBox 75, 35, 50, 15, start_date
  		EditBox 75, 55, 50, 15, end_date
  		EditBox 215, 35, 20, 15, months_disq		
  		EditBox 215, 55, 50, 15, DISQ_begin_date
  		EditBox 80, 90, 50, 15, fraud_case_number
  		DropListBox 80, 110, 75, 15, "Select one..."+chr(9)+"Unknown"+chr(9)+"Judy Grandel"+chr(9)+"Chris Gormley"+chr(9)+"Keyatta Hill"+chr(9)+"Amanda Lange"+chr(9)+"Kimberly Littlejohn"+chr(9)+"Jonathan Martin"+chr(9)+"Ryan Swanson"+chr(9)+"Scott Benedict", Fraud_investigator
  		ButtonGroup ButtonPressed
    		OkButton 175, 115, 50, 15
    		CancelButton 230, 115, 50, 15
  		Text 10, 110, 65, 10, "Fraud Investigator:"
  		Text 170, 10, 35, 10, "Programs:"
  		Text 55, 60, 15, 10, "End:"
  		Text 155, 40, 55, 10, "Months of DISQ:"
  		Text 5, 10, 100, 10, "Date client signed ADH waiver:"
  		Text 55, 40, 15, 10, "Start:"
  		Text 155, 55, 60, 10, "DISQ Begin Date:"
  		Text 10, 95, 70, 10, "Fraud Claim Number"
  		GroupBox 5, 25, 275, 50, "Period of offense:"
  		GroupBox 5, 80, 155, 50, "Fraud Information"
	EndDialog
DO		
	Do 
		err_msg = "" 
		Dialog
		cancel_confirmation 
		IF isdate(date_waiver_signed) = false THEN err_msg = err_msg & vbNewLine & "* Please enter date waiver was signed."
		IF program_droplist = "Select Programs:" THEN err_msg = err_msg & vbNewLine & "* Please select the program."
		IF isdate(start_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter start date."
		IF isdate(end_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter end date."
		IF IsNumeric(months_disq) = false THEN err_msg = err_msg & vbNewLine & "* Please enter the amount of disqualification months."
		IF isdate(DISQ_begin_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter DISQ beign date."
		IF trim(fraud_case_number) = "" and Fraud_investigator <> "Select one..." then err_msg = err_msg & vbNewLine & "* Enter both the fraud case number AND the Fraud Investigator's name, or clear the non-applicable info."
		IF trim(fraud_case_number) <> "" and Fraud_investigator = "Select one..."  THEN err_msg = err_msg & vbNewLine & "* Enter both the fraud case number AND the Fraud Investigator's name, or clear the non-applicable info."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine		
	Loop until err_msg = ""	
	CALL check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

DISQ_end_date = DateAdd("M", months_disq, DISQ_begin_date)
IF Fraud_investigator = "Select one..."  THEN Fraud_investigator = ""
IF Fraud_investigator = "" 						then fraud_email = ""
IF Fraud_investigator = "Judy Grandel" 			then fraud_email = "Judith.Grandel@hennepin.us"
IF Fraud_investigator = "Chris Gormley"		 	then fraud_email = "Chris.Gormley@hennepin.us"
IF Fraud_investigator = "Keyatta Hill" 	 		then fraud_email = "Keyatta.Hill@hennepin.us"
IF Fraud_investigator = "Amanda Lange" 			then fraud_email = "Amanda.Lange@hennepin.us"
IF Fraud_investigator = "Kimberly Littlejohn"	then fraud_email = "Kimberly.Littlejohn@Hennepin.us"
IF Fraud_investigator = "Jonathan Martin" 		then fraud_email = "Jonathan.Martin@Hennepin.us"
IF Fraud_investigator = "Ryan Swanson" 			then fraud_email = "Ryan.Swanson@hennepin.us"
IF Fraud_investigator = "Scott Benedict" 		then fraud_email = "Scott.Benedict@hennepin.us"
	
'The 1st case note-------------------------------------------------------------------------------------------------
 	start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode		 
	 CALL write_variable_in_CASE_NOTE("-----1st Fraud DISQ/Claims (" & memb_name & ") ADH Waiver Signed-----")
	 CALL write_variable_in_CASE_NOTE("Client signed ADH waiver on: " & date_waiver_signed & " waiving his/her right to an Administrative Disqualification Hearing for wrongfully obtaining public assistance. This disqualification is not for any other household member and does not affect MA eligibility.")
	 CALL write_variable_in_CASE_NOTE("* Programs: " & program_droplist)
	 CALL write_variable_in_CASE_NOTE("* Period of Offense: " & start_date & " - " & end_date)
	 CALL write_variable_in_CASE_NOTE("* Client is subject to a " & months_disq & " month DISQ from " & DISQ_begin_date & "-" & DISQ_end_date & ".")
	 IF program_droplist <> "SNAP"  THEN CALL write_variable_in_CASE_NOTE("* Other Notes: Because member " & memb_number & " is DQ'd from MFIP, client is also barred from FS for that same period of time.")
	 IF fraud_case_number <> "" THEN 
		 CALL write_variable_in_CASE_NOTE("----- ----- -----")
		 CALL write_bullet_and_variable_in_CASE_NOTE("Fraud claim number", fraud_case_number) 
		 CALL write_bullet_and_variable_in_CASE_NOTE("Fraud Investigator", Fraud_investigator)
	 END IF
	 CALL write_variable_in_CASE_NOTE("* Email sent to team: L. Bloomquist, TTL, and FSS")
	 CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
 	 CALL write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1") 
	 'Drafting an email. Does not send the email!!!!
	 'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
	 CALL create_outlook_email("Lea.Bloomquist@hennepin.us", "HSPH.ES.TEAM.TTL@hennepin.us;" & "HSPH.FSSDataTeam@hennepin.us;" & fraud_email, "1st Fraud DISQ/Claims--ADH Waiver Signed for #" &  MAXIS_case_number, "Member #: " & memb_number & vbcr & "Client signed ADH waiver on: " & date_waiver_signed & " waiving his/her right to an Administrative Disqualification Hearing for wrongfully obtaining public assistance." & vbcr & "Programs: " & program_droplist & vbcr & "Period of Offense: " & start_date & " - " & end_date & vbcr & "See case notes for further details.", "", False)
END IF 		

IF ADH_option = "Hearing Held" then	
    BeginDialog , 0, 0, 276, 150, "ADH Hearing Held"
      EditBox 70, 5, 50, 15, hearing_date
      DropListBox 195, 5, 75, 15, "Select Programs:"+chr(9)+"CASH"+chr(9)+"CASH & SNAP"+chr(9)+"SNAP", Program_droplist
      EditBox 70, 25, 50, 15, date_order_signed
      EditBox 70, 50, 50, 15, start_date
      EditBox 70, 70, 50, 15, end_date
      EditBox 210, 50, 20, 15, months_disq
      EditBox 210, 70, 50, 15, DISQ_begin_date
      EditBox 80, 105, 50, 15, fraud_case_number
      DropListBox 80, 125, 75, 15, "Select one..."+chr(9)+"Unknown"+chr(9)+"Judy Grandel"+chr(9)+"Chris Gormley"+chr(9)+"Keyatta Hill"+chr(9)+"Amanda Lange"+chr(9)+"Kimberly Littlejohn"+chr(9)+"Jonathan Martin"+chr(9)+"Ryan Swanson"+chr(9)+"Scott Benedict", Fraud_investigator
      ButtonGroup ButtonPressed
        OkButton 165, 125, 50, 15
        CancelButton 220, 125, 50, 15
      Text 10, 130, 65, 10, "Fraud Investigator:"
      Text 50, 75, 15, 10, "End:"
      Text 150, 55, 55, 10, "Months of DISQ:"
      Text 5, 25, 50, 10, "Order Signed:"
      Text 50, 55, 15, 10, "Start:"
      Text 150, 75, 60, 10, "DISQ Begin Date:"
      Text 10, 110, 70, 10, "Fraud Claim Number:"
      GroupBox 5, 40, 265, 50, "Period of offense:"
      GroupBox 5, 95, 155, 50, "Fraud Information"
      Text 5, 10, 65, 10, "ADH Hearing Held:"
      Text 155, 10, 35, 10, "Programs:"
    EndDialog
	DO		
		Do 
			err_msg = "" 
			Dialog
			cancel_confirmation 
			IF isdate(hearing_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter the hearing date."
			IF program_droplist = "Select Programs:" THEN err_msg = err_msg & vbNewLine & "* Please select the program."
			IF isdate(date_order_signed) = false THEN err_msg = err_msg & vbNewLine & "* Please enter date order was signed"
			IF isdate(start_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter start date"
			IF isdate(end_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter end date"
			IF IsNumeric(months_disq) = false THEN err_msg = err_msg & vbNewLine & "* Please enter the amount of disqualification months."
			IF isdate(DISQ_begin_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter date DISQ begins"
			IF trim(fraud_case_number) = "" and Fraud_investigator <> "Select one..." THEN err_msg = err_msg & vbNewLine & "* Enter both the fraud case number AND the Fraud Investigator's name, or clear the non-applicable info."
			IF trim(fraud_case_number) <> "" and Fraud_investigator = "Select one..."  THEN err_msg = err_msg & vbNewLine & "* Enter both the fraud case number AND the Fraud Investigator's name, or clear the non-applicable info."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine		
		Loop until err_msg = ""	
		CALL check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False
	
	DISQ_end_date = DateAdd("M", months_disq, DISQ_begin_date)
	IF Fraud_investigator = "Unknown" THEN Fraud_investigator = ""	
	IF Fraud_investigator = "" 						then fraud_email = ""
	IF Fraud_investigator = "Judy Grandel" 			then fraud_email = "Judith.Grandel@hennepin.us"
	IF Fraud_investigator = "Chris Gormley"		 	then fraud_email = "Chris.Gormley@hennepin.us"
	IF Fraud_investigator = "Keyatta Hill" 	 		then fraud_email = "Keyatta.Hill@hennepin.us"
	IF Fraud_investigator = "Amanda Lange" 			then fraud_email = "Amanda.Lange@hennepin.us"
	IF Fraud_investigator = "Kimberly Littlejohn"	then fraud_email = "Kimberly.Littlejohn@Hennepin.us"
	IF Fraud_investigator = "Jonathan Martin" 		then fraud_email = "Jonathan.Martin@Hennepin.us"
	IF Fraud_investigator = "Ryan Swanson" 			then fraud_email = "Ryan.Swanson@hennepin.us"
	IF Fraud_investigator = "Scott Benedict" 		then fraud_email = "Scott.Benedict@hennepin.us"
	
'The 2nd case note-------------------------------------------------------------------------------------------------
	start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode		 
	 CALL write_variable_in_CASE_NOTE("-----1st Fraud DISQ/Claim (" & memb_name & ")  ADH Hearing Held-----")
	 CALL write_variable_in_CASE_NOTE("Administrative Disqualification Hearing for Wrongfully Obtaining Public Assistance was held on: " & hearing_date & " " & "Order was signed: " & date_order_signed & "This disqualification is not for any other household member and does not affect MA eligibility.")
	 CALL write_variable_in_CASE_NOTE("----- ----- -----")
	 CALL write_variable_in_CASE_NOTE("* Programs: " & program_droplist)
	 CALL write_variable_in_CASE_NOTE("* Period of Offense: " & start_date & " - " & end_date)
	 CALL write_variable_in_CASE_NOTE("* Client is subject to a " & months_disq & " month DISQ from " & DISQ_begin_date & "-" & DISQ_end_date)
	 IF program_droplist <> "SNAP"  then CALL write_variable_in_CASE_NOTE("* Other Notes: Because member " & memb_number & " is DQ'd from MFIP, client is also barred from FS for that same period of time.")
	 IF fraud_case_number <> "" THEN 
		 CALL write_variable_in_CASE_NOTE("----- ----- -----")
		 CALL write_bullet_and_variable_in_CASE_NOTE("Fraud claim number", fraud_case_number) 
		 CALL write_bullet_and_variable_in_CASE_NOTE("Fraud Investigator", Fraud_investigator)
	 END IF
	 CALL write_variable_in_CASE_NOTE("* Email sent to team: L. Bloomquist, TTL, and FSS")
	 CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
	 CALL write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1") 
	 'Drafting an email. Does not send the email!!!!
	 'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
	 CALL create_outlook_email("Lea.Bloomquist@hennepin.us", "HSPH.ES.TEAM.TTL@hennepin.us;" & "HSPH.FSSDataTeam@hennepin.us;" & fraud_email, "1st Fraud DISQ/Claims--ADH Hearing Held for #" &  MAXIS_case_number, "Member #: " & memb_number & vbcr & "Administrative Disqualification Hearing for Wrongfully Obtaining Public Assistance was held on: " & hearing_date & vbcr & "Order was signed: " & date_order_signed & vbcr & "Programs: " & program_droplist & vbcr & "Period of Offense: " & start_date & " - " & end_date & vbcr & "See case notes for further details.", "", False)
END IF 		
script_end_procedure("Please select the applicable team in the drafted email, any additional notes required and send the email regarding ADH information.") 