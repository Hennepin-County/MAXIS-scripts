'GATHERING STATS===========================================================================================
name_of_script = "NOTES - VENDOR.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 120
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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
CALL changelog_update("09/20/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("10/23/2019", "New functionality added to the Vendor script to pull vendor information from MAXIS.##~##", "Casey Love, Hennepin County")
call changelog_update("10/21/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Declarations
const vnds_nbr          = 0
const vnds_type         = 1
const vnds_name         = 2
const vnds_address      = 3
const vnds_phone        = 4
const vnds_grh          = 5
const vnds_stat         = 6
const clt_account_numb  = 7
const vnds_amount       = 8

Dim VENDOR_INFO_ARRAY()
ReDim VENDOR_INFO_ARRAY(vnds_amount, 0)

'The script ----------------------------------------------------------------------------------------------------
EMConnect ""                    'Connect to MAXIS
Call MAXIS_case_number_finder(MAXIS_case_number)            'Pulling the case number from MAXIS if it can be found

Do
    Do
        err_msg = ""
        'Case number and Vendor Number Dialog - finding the information needed to autofill
        'This has to be defined within the loop because there is another dialog in this loop - the search function
		'-------------------------------------------------------------------------------------------------DIALOG
		Dialog1 = "" 'Blanking out previous dialog detail
        BeginDialog Dialog1, 0, 0, 331, 105, "Vendor Numbers to NOTE"
          EditBox 275, 15, 50, 15, MAXIS_case_number
          EditBox 110, 35, 215, 15, vendor_number_list
          ButtonGroup ButtonPressed
            PushButton 265, 70, 60, 10, "Vendor Search", vendor_search_button
            OkButton 220, 85, 50, 15
            CancelButton 275, 85, 50, 15
          Text 10, 10, 195, 20, "This script will enter a CASE/NOTE with Vendor Information for a case. Mutiple vendors can be noted at a time."
          Text 220, 20, 50, 10, "Case Number:"
          Text 10, 40, 100, 10, "Established Vendor Numbers:"
          Text 115, 55, 210, 10, "*Enter vendor numbers separated by commas if more than one."
          Text 10, 75, 140, 25, "There will be a place to add vendor information for a vendor that does not yet have a vendor number in the next dialog."
        EndDialog

		DO
		    DO
		    	err_msg = ""
		    		DIALOG Dialog1
		    		cancel_without_confirmation
		    		IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND len(MAXIS_case_number) > 8) OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
					IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		    	LOOP UNTIL err_msg = ""
		    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
			LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
		CALL check_for_MAXIS(False)

                'If the search button is pressed, the functionality for searching vendors here will start
        If ButtonPressed = vendor_search_button Then
            err_msg = "LOOP" & err_msg              'this makes sure we loop back to the first dialog
            vendor_search_county = "27 Hennepin"    'defaulting the parameters of the search
            vendor_search_status = "Any"
            vendor_search_name = ""

            'this is the search dialog
            BeginDialog Dialog1, 0, 0, 266, 85, "Vendor Search Criteria"
              EditBox 70, 25, 190, 15, vendor_search_name
              DropListBox 70, 45, 70, 45, "Any"+chr(9)+"Active"+chr(9)+"Merged"+chr(9)+"Pending"+chr(9)+"Terminated", vendor_search_status
              DropListBox 70, 65, 80, 45, "Any"+chr(9)+county_list, vendor_search_county
              ButtonGroup ButtonPressed
                PushButton 220, 65, 40, 15, "SEARCH", execute_search
              Text 10, 10, 160, 10, "Enter vendor information you wish to search by:"
              Text 15, 30, 50, 10, "Vendor Name:"
              Text 15, 50, 55, 10, "Vendor Status:"
              Text 15, 70, 50, 10, "Vendor County:"
            EndDialog

			DO
				DO
					err_msg = ""
					DIALOG Dialog1
						cancel_without_confirmation
						IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
					LOOP UNTIL err_msg = ""
					CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
				LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
			CALL check_for_MAXIS(False)

            Call navigate_to_MAXIS_screen("MONY", "VNDS")           'going to vendor search in MAXIS

            EMWriteScreen vendor_search_name, 4, 15                 'entering the information from the dialog into the search in MAXIS
            If vendor_search_status <> "Any" Then EMWriteScreen left(vendor_search_status, 1), 5, 21
            If vendor_search_county <> "Any" Then EMWriteScreen vendor_search_county, 5, 10
            EMWriteScreen "        ", 4, 59                         'blanking the vendor number
            transmit                                              'submitting the search
        End If
        If err_msg <> "" AND left(err_msg, 4) <> "LOOP" Then MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg       'showing the error message if anything is missing from the initial dialog
	Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)                  'password handling
Loop until check_for_password(are_we_passworded_out) = False


vendor_number_list = trim(vendor_number_list)           'formatting information from the main dialog
If InStr(vendor_number_list, ",") = 0 Then              'making the list of vendor numbers an array
    vendor_array = ARRAY(vendor_number_list)
Else
    vendor_array = split(vendor_number_list, ",")
End If

vendor_counter = 0      'setting the incrementer for the large array
For each vendor_number in vendor_array                  'cycling through the array of numbers to gather information and fill the large array
    vendor_number = trim(vendor_number)                 'getting rid of spaces
    If vendor_number <> "" Then                         'making sure we don't have a blank in the array
        ReDim Preserve VENDOR_INFO_ARRAY(vnds_amount, vendor_counter)       'resizing the large array of vendor detail information
        VENDOR_INFO_ARRAY(vnds_nbr, vendor_counter) = vendor_number         'setting the vendor number into the large array
        Call navigate_to_MAXIS_screen("MONY", "VNDS")                       'going to look for more vendor detail in MAXIS
        EMWriteScreen vendor_number, 4, 59                                  'enter the vendor number into VNDS
        transmit
        EMReadScreen check_for_vndm, 4, 2, 54                               'making sure we got to the vendor file maintenance - VNDM
        If check_for_vndm = "VNDM" Then
            EMReadScreen vndm_name, 30, 3, 15                               'reading all the information from VNDM
            EMReadScreen vndm_grh_yn, 1, 4, 57
            EMREadScreen vndm_street_one, 22, 5, 15
            EMReadScreen vndm_street_two, 22, 6, 15
            EMReadScreen vndm_city, 15, 7, 15
            EMReadScreen vndm_state, 2, 7, 36
            EMReadScreen vndm_zip, 5, 7, 46
            EMReadScreen vndm_phone, 18, 6, 52
            EMReadScreen vndm_status, 1, 16, 15
            VENDOR_INFO_ARRAY(vnds_name, vendor_counter) = replace(vndm_name, "_", "")      'fromatting the name of the vendor and save it to the large array
            vndm_phone = replace(vndm_phone, " )  ", ")")                       'formatting the phone number and adding it to the large array
            vndm_phone = replace(vndm_phone, "  ", "-")
            vndm_phone = replace(vndm_phone, " ", "")
            If vndm_phone = "(___)___-____" Then vndm_phone = ""
            VENDOR_INFO_ARRAY(vnds_phone, vendor_counter) = vndm_phone
            vndm_street_one = replace(vndm_street_one, "_", "")                 'formatting the address and compiling the pieces and adding it to the large array
            vndm_street_two = replace(vndm_street_two, "_", "")
            vndm_city = replace(vndm_city, "_", "")
            VENDOR_INFO_ARRAY(vnds_address, vendor_counter) = vndm_street_one & " " & vndm_street_two & " " & vndm_city & ", " & vndm_state & " " & vndm_zip
            VENDOR_INFO_ARRAY(vnds_grh, vendor_counter) = vndm_grh_yn           'saving the grh information to the large array
            If vndm_status = "A" Then VENDOR_INFO_ARRAY(vnds_stat, vendor_counter) = "Active"           'Setting the vendor status in the large array with actual words
            If vndm_status = "P" Then VENDOR_INFO_ARRAY(vnds_stat, vendor_counter) = "Pending"
            If vndm_status = "M" Then VENDOR_INFO_ARRAY(vnds_stat, vendor_counter) = "Merged"
            If vndm_status = "T" Then VENDOR_INFO_ARRAY(vnds_stat, vendor_counter) = "Terminated"
        End If

        Call back_to_SELF       'after gathering the vendor information, go back to self to reset all the things
        vendor_counter = vendor_counter + 1     'incrementing the counter for the next vendor number
    End If
Next

Do
    Do
        err_msg = ""            'resetting the error message for the next dialog
        y_pos = 45              'variables for height of dialog as this is dynamic
        dlg_len = 65 + ( 25 * (UBound(VENDOR_INFO_ARRAY, 2) + 1))

		'-------------------------------------------------------------------------------------------------DIALOG
		Dialog1 = "" 'Blanking out previous dialog detail
        BeginDialog Dialog1, 0, 0, 680, dlg_len, "Vendor Detail Information"
          ButtonGroup ButtonPressed
            PushButton 565, 10, 105, 10, "ADD ANOTHER VENDOR", add_another_button
          Text 10, 10, 340, 10, "Enter detail about the vendor that should be entered into the case note detailing the vendor information."
          Text 10, 30, 30, 10, "Number"
          Text 60, 30, 25, 10, "Name"
          Text 150, 30, 30, 10, "Address"
          Text 290, 30, 25, 10, "Phone"
          Text 355, 30, 25, 10, "GRH?"
          Text 395, 30, 25, 10, "Status"
          Text 450, 30, 20, 10, "Type"
          Text 565, 30, 30, 10, "Amount"
          Text 605, 30, 60, 10, "Client Account #"
          For each_vendor = 0 to UBound(VENDOR_INFO_ARRAY, 2)
              EditBox 10, y_pos, 35, 15, VENDOR_INFO_ARRAY(vnds_nbr, each_vendor)
              EditBox 60, y_pos, 85, 15, VENDOR_INFO_ARRAY(vnds_name, each_vendor)
              EditBox 150, y_pos, 135, 15, VENDOR_INFO_ARRAY(vnds_address, each_vendor)
              EditBox 290, y_pos, 55, 15, VENDOR_INFO_ARRAY(vnds_phone, each_vendor)
              DropListBox 355, y_pos, 30, 45, ""+chr(9)+"Yes"+chr(9)+"No", VENDOR_INFO_ARRAY(vnds_grh, each_vendor)
              DropListBox 395, y_pos, 45, 45, ""+chr(9)+"Active"+chr(9)+"Pending"+chr(9)+"Merged"+chr(9)+"Terminated", VENDOR_INFO_ARRAY(vnds_stat, each_vendor)
              ComboBox 450, y_pos, 105, 45, "Select or Type"+chr(9)+"Mandatory Shelter"+chr(9)+"Mandatory Utility"+chr(9)+"Voluntary Shelter"+chr(9)+"Voluntary Utility"+chr(9)+"GRH"+chr(9)+VENDOR_INFO_ARRAY(vnds_type, each_vendor), VENDOR_INFO_ARRAY(vnds_type, each_vendor)
              EditBox 565, y_pos, 35, 15, VENDOR_INFO_ARRAY(vnds_amount, each_vendor)
              EditBox 605, y_pos, 70, 15, VENDOR_INFO_ARRAY(clt_account_numb, each_vendor)
              y_pos = y_pos + 25
          Next
          Text 10, y_pos + 5, 25, 10, "Notes:"
          EditBox 40, y_pos, 340, 15, other_notes
          Text 395, y_pos + 5, 60, 10, "Worker Signature:"
          EditBox 460, y_pos, 90, 15, worker_signature
          ButtonGroup ButtonPressed
            OkButton 570, y_pos, 50, 15
            CancelButton 625, y_pos, 50, 15
        EndDialog

        dialog Dialog1                  'showing the dialog
        cancel_confirmation             'cancel the run

        If ButtonPressed = add_another_button Then      'This increases the large array and adds another line to the dialog for manual entry of vendor information if no number was listed
            err_msg = "LOOP" & err_msg                  'making sure the dialog reappears'\
            vendor_counter = Ubound(VENDOR_INFO_ARRAY, 2) + 1
            ReDim Preserve VENDOR_INFO_ARRAY(vnds_amount, vendor_counter)
        End If

        For each_vendor = 0 to UBound(VENDOR_INFO_ARRAY, 2)     'Looking for mandatory fields
            If trim(VENDOR_INFO_ARRAY(vnds_name, each_vendor)) = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the vendor."          'VENDOR NAME
            If trim(VENDOR_INFO_ARRAY(vnds_type, each_vendor)) = "" OR trim(VENDOR_INFO_ARRAY(vnds_type, each_vendor)) = "Select or Type" Then err_msg = err_msg & vbNewLine & "* Enter the type of vendor"     'VENDOR TYPE
        Next
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
        If err_msg <> "" AND left(err_msg, 4) <> "LOOP" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg     'Showing the error message
    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)          'password handling
Loop until are_we_passworded_out = False

'The case note---------------------------------------------------------------------------------------
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("---Vendor information---")

For each_vendor = 0 to UBound(VENDOR_INFO_ARRAY, 2)                     'This adds each of the vendor peices to the case note in a formatted way
    If trim(VENDOR_INFO_ARRAY(vnds_nbr, each_vendor)) = "" Then
        Call write_variable_in_CASE_NOTE("Vendor: " & VENDOR_INFO_ARRAY(vnds_name, each_vendor))
    Else
        Call write_variable_in_CASE_NOTE("Vendor #" & VENDOR_INFO_ARRAY(vnds_nbr, each_vendor) & " : " & VENDOR_INFO_ARRAY(vnds_name, each_vendor))
    End If
    Call write_bullet_and_variable_in_CASE_NOTE("Address", VENDOR_INFO_ARRAY(vnds_address, each_vendor))
    Call write_bullet_and_variable_in_CASE_NOTE("Phone", VENDOR_INFO_ARRAY(vnds_phone, each_vendor))
    Call write_bullet_and_variable_in_CASE_NOTE("Status", VENDOR_INFO_ARRAY(vnds_stat, each_vendor))
    If VENDOR_INFO_ARRAY(vnds_grh, each_vendor) = "Yes" Then Call write_variable_in_CASE_NOTE("* This Vendor is a GRH.")
    Call write_variable_in_CASE_NOTE("* On this case:")
    Call write_variable_with_indent_in_CASE_NOTE("Type: " & VENDOR_INFO_ARRAY(vnds_type, each_vendor))
    If trim(VENDOR_INFO_ARRAY(vnds_amount, each_vendor)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Amount: $" & VENDOR_INFO_ARRAY(vnds_amount, each_vendor))
    If trim(VENDOR_INFO_ARRAY(clt_account_numb, each_vendor)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Account Number: " & VENDOR_INFO_ARRAY(clt_account_numb, each_vendor))
Next

 Call write_variable_in_CASE_NOTE ("---")
 Call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
 Call write_variable_in_CASE_NOTE (worker_signature)

 Call script_end_procedure_with_error_report("Vendor information entered into CASE/NOTE")
