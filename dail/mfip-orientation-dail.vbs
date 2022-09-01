'Required for statistical purposes===============================================================================
name_of_script = "DAIL - MFIP ORIENTATION.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 300          'manual run time in seconds
STATS_denomination = "M"       'M is for Member
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
call changelog_update("09/01/2022", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'NO CHANGES ERE MADE TO THIS FUNCTION FROM NOTES - INTERVIEW - ONCE WE ARE READY TO PULL THIS INTO THE FUNCTIONS LIBRARY WE SHOULD NOT NEED TO MAKE ANY COMPARRISON BETWEEN THE TWO SCRIPTS.
function complete_MFIP_orientation(CAREGIVER_ARRAY, memb_ref_numb_const, memb_name_const, memb_age_const, memb_is_caregiver, cash_request_const, hours_per_week_const, exempt_from_ed_const, comply_with_ed_const, orientation_needed_const, orientation_done_const, orientation_exempt_const, exemption_reason_const, emps_exemption_code_const, choice_form_done_const, orientation_notes, family_cash_program)
'DO NOT CHANGE THIS FUNCTION - IT IS DUPLICATED IN AANOTHER SCRIPT AND WE DO NOT WANT TO HAVE TO COMPARE

	'first - assess if caregiver meets an exemption
		'- Single parent household employed at least 35 hours per week
		'- 2 Parent household where the 1st parent is employed at least 35 hours per week
		'- 2 Parened household where the 2nd parent is employed at least 20 hours per week and the 1st is employed 35
		'- Pregnant or parenting minor under 20 who is coplying with the educational requirements
		'- Caregiver is not receiving MFIP

	'Identify the caregivers
	'Identify if they are requesting Cash
	'Indicate if this will be DWP or MFIP
	'Identify if the caregiver is a minor
	'List the hours employed for each caregiver
	'
	person_list = "Select One..."+chr(9)+"No Caregiver"
	second_person_list = "Select One..."+chr(9)+"No Second Caregiver"

	For person = 0 to UBound(CAREGIVER_ARRAY, 2)
		person_list = person_list+chr(9)+CAREGIVER_ARRAY(memb_name_const, person)
		second_person_list = second_person_list+chr(9)+CAREGIVER_ARRAY(memb_name_const, person)
	Next
	caregiver_one = CAREGIVER_ARRAY(memb_name_const, 0)

	Do
		err_msg = ""
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 551, 150, "Assess for Caregiver MFIP Orientation Requirement"
		  DropListBox 185, 10, 60, 45, "MFIP"+chr(9)+"DWP", family_cash_program
		  EditBox 110, 30, 430, 15, famliy_cash_notes
		  DropListBox 65, 65, 140, 45, person_list, caregiver_one
		  DropListBox 330, 65, 45, 45, "Yes"+chr(9)+"No"+chr(9)+"Not Elig", caregiver_one_req_cash
		  EditBox 430, 65, 30, 15, caregiver_one_hours_per_week
		  DropListBox 65, 85, 140, 45, second_person_list, caregiver_two
		  DropListBox 330, 85, 45, 45, "Yes"+chr(9)+"No"+chr(9)+"Not Elig", caregiver_two_req_cash
		  EditBox 430, 85, 30, 15, caregiver_two_hours_per_week
		  Text 15, 125, 450, 20, "These questions will identify if these caregivers need an MFIP orientation. See CM 05.12.12.06   to see the reasons that a caregiver would not need an MFIP Orientation. The script will use this information to determine if the MFIP Orientation Functionality should be run."
		  ButtonGroup ButtonPressed
			OkButton 490, 125, 50, 15
			PushButton 420, 10, 120, 15, "MFIP Orientation Script Instructions", msg_mfip_orientation_btn
            PushButton 260, 123, 55, 10, "CM05.12.12.06", cm_05_12_12_06_btn
		  Text 10, 15, 170, 10, "Which Family Cash Program is this Application for?"
		  Text 10, 35, 100, 10, "Notes on Program Selection:"
		  GroupBox 10, 50, 530, 55, "Who are the Caregivers"
		  Text 20, 70, 40, 10, "Caregiver:"
		  Text 215, 70, 115, 10, "Is this caregiver requesting cash?"
		  Text 385, 70, 40, 10, "Employed: "
		  Text 465, 70, 50, 10, "hours/week"
		  Text 20, 90, 40, 10, "Caregiver:"
		  Text 215, 90, 115, 10, "Is this caregiver requesting cash?"
		  Text 385, 90, 40, 10, "Employed: "
		  Text 465, 90, 50, 10, "hours/week"
		  Text 15, 110, 100, 10, "Why is this being asked?"
		EndDialog

		dialog Dialog1
		cancel_confirmation

		If caregiver_one = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate the First Caregiver or clarify that there is no caregiver"
		If caregiver_two = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate the Second Caregiver or clarify that there is no second caregiver"
		If caregiver_one = caregiver_two Then err_msg = err_msg & vbCr & "* Select two different caregivers"
		If IsNumeric(caregiver_one_hours_per_week) = False AND trim(caregiver_one_hours_per_week) <> "" Then err_msg = err_msg & vbCr & "* Hours per week should be a number, or left blank."
		If IsNumeric(caregiver_two_hours_per_week) = False AND trim(caregiver_two_hours_per_week) <> "" Then err_msg = err_msg & vbCr & "* Hours per week should be a number, or left blank."

		If family_cash_program = "DWP" Then err_msg = ""

		If ButtonPressed <> -1 Then err_msg = "LOOP"
		If err_msg <> "" And ButtonPressed = -1 Then MsgBox err_msg

        If ButtonPressed = cm_05_12_12_06_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_0005121206"
		If ButtonPressed = msg_mfip_orientation_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW%20-%20MFIP%20ORIENTATION.docx"

	Loop until err_msg = ""

	If family_cash_program = "MFIP" Then
		If IsNumeric(caregiver_one_hours_per_week) = True Then caregiver_one_hours_per_week = caregiver_one_hours_per_week * 1
		If trim(caregiver_one_hours_per_week) = "" Then caregiver_one_hours_per_week = 0

		If IsNumeric(caregiver_two_hours_per_week) = True Then caregiver_two_hours_per_week = caregiver_two_hours_per_week * 1
		If trim(caregiver_two_hours_per_week) = "" Then caregiver_two_hours_per_week = 0

		minor_caregiver_on_case = 0

		For person = 0 to UBound(CAREGIVER_ARRAY, 2)
			If CAREGIVER_ARRAY(memb_name_const, person) = caregiver_one Then
				CAREGIVER_ARRAY(memb_is_caregiver, person) = True
				CAREGIVER_ARRAY(orientation_needed_const, person) = True

				If caregiver_one_req_cash = "Yes" Then CAREGIVER_ARRAY(cash_request_const, person) = True
				If caregiver_one_req_cash <> "Yes" Then
					CAREGIVER_ARRAY(cash_request_const, person) = False
					CAREGIVER_ARRAY(orientation_needed_const, person) = False
					CAREGIVER_ARRAY(orientation_exempt_const, person) = True
					CAREGIVER_ARRAY(exemption_reason_const, person) = "Caregiver Not on MFIP"
					CAREGIVER_ARRAY(emps_exemption_code_const, person) = "NO"
				End If
				CAREGIVER_ARRAY(hours_per_week_const, person) = caregiver_one_hours_per_week

				If CAREGIVER_ARRAY(hours_per_week_const, person) > 34 Then
					CAREGIVER_ARRAY(orientation_needed_const, person) = False
					CAREGIVER_ARRAY(orientation_exempt_const, person) = True
					CAREGIVER_ARRAY(exemption_reason_const, person) = "Employed 35+ hours per week"
					CAREGIVER_ARRAY(emps_exemption_code_const, person) = "20"
				ElseIf CAREGIVER_ARRAY(hours_per_week_const, person) > 19 Then
					If caregiver_two <> "No Second Caregiver" AND caregiver_two_req_cash = "Yes" AND caregiver_two_hours_per_week > 34 Then
						CAREGIVER_ARRAY(orientation_needed_const, person) = False
						CAREGIVER_ARRAY(orientation_exempt_const, person) = True
						CAREGIVER_ARRAY(exemption_reason_const, person) = "2nd Caregiver Employed 20+ hours per week"
						CAREGIVER_ARRAY(emps_exemption_code_const, person) = "21"
					End If
				End If
				If CAREGIVER_ARRAY(memb_age_const, person) < 20 Then
					minor_caregiver_on_case = minor_caregiver_on_case + 1
					CAREGIVER_ARRAY(exempt_from_ed_const, person) = "No"
					CAREGIVER_ARRAY(comply_with_ed_const, person) = "Yes"
				End If

			End If

			If CAREGIVER_ARRAY(memb_name_const, person) = caregiver_two Then
				CAREGIVER_ARRAY(memb_is_caregiver, person) = True
				CAREGIVER_ARRAY(orientation_needed_const, person) = True

				If caregiver_two_req_cash = "Yes" Then CAREGIVER_ARRAY(cash_request_const, person) = True
				If caregiver_two_req_cash <> "Yes" Then
					CAREGIVER_ARRAY(cash_request_const, person) = False
					CAREGIVER_ARRAY(orientation_needed_const, person) = False
					CAREGIVER_ARRAY(orientation_exempt_const, person) = True
					CAREGIVER_ARRAY(exemption_reason_const, person) = "Caregiver Not on MFIP"
					CAREGIVER_ARRAY(emps_exemption_code_const, person) = "NO"
				End If
				CAREGIVER_ARRAY(hours_per_week_const, person) = caregiver_two_hours_per_week

				If CAREGIVER_ARRAY(hours_per_week_const, person) > 34 Then
					CAREGIVER_ARRAY(orientation_needed_const, person) = False
					CAREGIVER_ARRAY(orientation_exempt_const, person) = True
					CAREGIVER_ARRAY(exemption_reason_const, person) = "Employed 35+ hours per week"
					CAREGIVER_ARRAY(emps_exemption_code_const, person) = "20"
				ElseIf CAREGIVER_ARRAY(hours_per_week_const, person) > 19 Then
					If caregiver_one <> "No Second Caregiver" AND caregiver_one_req_cash = "Yes" AND caregiver_one_hours_per_week > 34 Then
						CAREGIVER_ARRAY(orientation_needed_const, person) = False
						CAREGIVER_ARRAY(orientation_exempt_const, person) = True
						CAREGIVER_ARRAY(exemption_reason_const, person) = "2nd Caregiver Employed 20+ hours per week"
						CAREGIVER_ARRAY(emps_exemption_code_const, person) = "21"
					End If
				End If
				If CAREGIVER_ARRAY(memb_age_const, person) < 20 Then
					minor_caregiver_on_case = minor_caregiver_on_case + 1
					CAREGIVER_ARRAY(exempt_from_ed_const, person) = "No"
					CAREGIVER_ARRAY(comply_with_ed_const, person) = "Yes"
				End If
			End If



		Next

		'IF A MINOR IS FOUND
		If minor_caregiver_on_case > 0 Then
			Do
				err_msg = ""
				dlg_len = 210
				If minor_caregiver_on_case = 2 Then dlg_len = 290

				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 551, dlg_len, "Assess for Caregiver MFIP Orientation Requirement"
				  Text 10, 15, 200, 10, "Which Family Cash Program is this Application for? " & family_cash_program
				  Text 10, 25, 500, 20, "Notes on Program Selection: " & famliy_cash_notes
				  GroupBox 10, 50, 530, 40, "Who are the Caregivers"
				  Text 20, 60, 190, 10, "Caregiver: " & caregiver_one
				  Text 215, 60, 165, 10, "Is this caregiver requesting cash? " & caregiver_one_req_cash
				  Text 385, 60, 90, 10, "Employed: " & caregiver_one_hours_per_week
				  Text 465, 60, 50, 10, "hours/week"
				  Text 20, 75, 190, 10, "Caregiver: " & caregiver_two
				  Text 215, 75, 165, 10, "Is this caregiver requesting cash? " & caregiver_two_req_cash
				  Text 385, 75, 90, 10, "Employed: " & caregiver_two_hours_per_week
				  Text 465, 75, 50, 10, "hours/week"
				  y_pos = 30
				  For caregiver = 0 to UBound(CAREGIVER_ARRAY, 2)
					  If CAREGIVER_ARRAY(memb_is_caregiver, caregiver) = True and CAREGIVER_ARRAY(memb_age_const, caregiver) < 20 Then
						  y_pos = y_pos + 70
						  GroupBox 10, y_pos, 530, 65, CAREGIVER_ARRAY(memb_name_const, caregiver)
						  Text 20, y_pos + 10, 270, 10, "This caregiver appears to be a minor by MFIP program rules (under 20 years old)."
						  Text 20, y_pos + 30, 195, 10, "Is this caregiver exempt from the Educational Requirement?"
						  DropListBox 230, y_pos + 25, 40, 45, "No"+chr(9)+"Yes", CAREGIVER_ARRAY(exempt_from_ed_const, caregiver)
						  Text 20, y_pos + 50, 205, 10, "Is this caregiver complying with the Educational Requirement?"
						  DropListBox 230, y_pos + 45, 40, 45, "No"+chr(9)+"Yes", CAREGIVER_ARRAY(comply_with_ed_const, caregiver)
					  End If
				  Next
				  Text 15, y_pos + 90, 450, 20, "These questions will identify if these caregivers need an MFIP orientation. See CM 05.12.12.06 to see the reasons that a caregiver would not need an MFIP Orientation. The script will use this information to determine if the MFIP Orientation Functionality should be run."
				  ButtonGroup ButtonPressed
					OkButton 490, y_pos + 90, 50, 15
					PushButton 485, y_pos + 45, 50, 15, "CM 28.12", cm_28_12_btn
					PushButton 260, y_pos + 87, 55, 10, "CM05.12.12.06", cm_05_12_12_06_btn
				  Text 355, y_pos + 45, 125, 20, "See details about the educational requirement in the Combined Manual "
				  Text 15, y_pos + 75, 100, 10, "Why is this being asked?"
				EndDialog

				dialog Dialog1
				cancel_confirmation

				If err_msg <> "" Then MsgBox err_msg

			Loop until err_msg = ""

			For caregiver = 0 to UBound(CAREGIVER_ARRAY, 2)
				If CAREGIVER_ARRAY(memb_is_caregiver, caregiver) = True and CAREGIVER_ARRAY(memb_age_const, caregiver) < 20 Then
					If CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = "No" Then CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = False
					If CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = "Yes" Then CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = True
					If CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = "No" Then CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = False
					If CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = "Yes" Then CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = True

					If CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = False and CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = True Then
						CAREGIVER_ARRAY(orientation_needed_const, caregiver) = False
						CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True
						CAREGIVER_ARRAY(exemption_reason_const, caregiver) = "Minor Caregiver meeting Educational Requirements"
						CAREGIVER_ARRAY(emps_exemption_code_const, caregiver) = "22"
					End If
				Else
					CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = False
					CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = False
				End If
			Next

		End If

		const mf_step_rights_resp 	= 1
		const mf_step_time_limits	= 2
		const mf_step_extension		= 3
		const mf_step_dv			= 4
		const mf_step_expectations	= 5
		const mf_step_esp			= 6
		const mf_step_compliance	= 7
		const mf_step_ep			= 8
		const mf_step_ccap			= 9
		const mf_step_incentives	= 10
		const mf_step_hc			= 11
		const mf_completion			= 12

		' mf_step_rights_resp_viewed = False
		' mf_step_time_limits_viewed = False
		' mf_step_extension_viewed = False
		' mf_step_dv_viewed = False
		' mf_step_expectations_viewed = False
		' mf_step_esp_viewed = False
		' mf_step_compliance_viewed = False
		' mf_step_ep_viewed = False
		' mf_step_ccap_viewed = False
		' mf_step_incentives_viewed = False
		' mf_step_hc_viewed = False
		' mf_completion_viewed = False
		'
		' orientation_script_document_viewed = False
		'
		'FIRST - Participant Responsibilities and Rights'
		'SECOND - MFIP Time Limits'
		'THIRD - MFIp Extension Eligibility'
		'FOURTH - Family Violence'
		'FIFTH - Expectations'
		'SIXTH - Choosing ESP'
		'SEVENTH - Assignment and Compliance'
		'EIGHTH - Developing an EP'
		'NINTH - CCAP'
		'TENTH - Incentives'
		'ELEVENTH - Health Care'

		' all_mfip_orientation_info_viewed = False
		For caregiver = 0 to UBound(CAREGIVER_ARRAY, 2)

			If CAREGIVER_ARRAY(orientation_needed_const, caregiver) = True Then
                Call Navigate_to_MAXIS_screen("STAT", "EMPS")
    			If CAREGIVER_ARRAY(memb_ref_numb_const, caregiver) <> "" Then
    				EMWriteScreen CAREGIVER_ARRAY(memb_ref_numb_const, caregiver), 20, 76
    				transmit
    			End If

                MFIP_orientation_step = mf_step_rights_resp

				mf_step_rights_resp_viewed = False
				mf_step_time_limits_viewed = False
				mf_step_extension_viewed = False
				mf_step_dv_viewed = False
				mf_step_expectations_viewed = False
				mf_step_esp_viewed = False
				mf_step_compliance_viewed = False
				mf_step_ep_viewed = False
				mf_step_ccap_viewed = False
				mf_step_incentives_viewed = False
				mf_step_hc_viewed = False
				mf_completion_viewed = False

				orientation_script_document_viewed = False

				all_mfip_orientation_info_viewed = False

				Do
					err_msg = ""

					Dialog1 = ""
					BeginDialog Dialog1, 0, 0, 551, 385, "MFIP Orientation"
					  ' GroupBox 10, 10, 450, 45, "Group1"
					  ButtonGroup ButtonPressed
					  	If MFIP_orientation_step <> mf_completion Then PushButton 495, 365, 50, 15, "NEXT", next_btn

						'FIRST - Participant Responsibilities and Rights'
						If MFIP_orientation_step = mf_step_rights_resp Then
						  Text 10, 10, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  GroupBox 10, 30, 450, 130, "Participant Responsibilities and Rights"
						  Text 20, 45, 370, 10, "As a program participant you have responsibilities and rights that were discussed during your intake interview."
						  Text 20, 60, 430, 10, "Please keep a copy of the Client Responsibilities and Rights (DHS-4163) for your reference. Let us know if you have any questions."
						  Text 20, 80, 335, 10, "Please remember it's important to report ANY changes that could affect your eligibility within 10 days."
						  GroupBox 10, 75, 450, 15, ""
						  Text 20, 100, 420, 20, "If your income decreases by at least 50% contact your financial worker right away!  You may be eligible for a significant change meaning a recalculation of your income which may result in an increase of your cash and/or food benefits."
						  Text 20, 125, 335, 20, "If you do not meet program eligibility such as cash assistance, your financial worker will assess other program eligibility such as SNAP."
						  PushButton 385, 160, 75, 15, "DHS - 4163", open_dhs_4163_btn

						  mf_step_rights_resp_viewed = True
						  'ADD BUTTON DHS 4163'
						End If

						'SECOND - MFIP Time Limits'
						If MFIP_orientation_step = mf_step_time_limits Then
						  GroupBox 10, 10, 450, 160, "MFIP Time-Limits"
						  Text 20, 25, 430, 30, "The MFIP program is available to you for up to 60 months in your lifetime.  If you have used cash assistance in another state those months must be reported and may count toward your lifetime limit. There are some instances the months you use may be exempt, meaning the months do not count towards the 60-month lifetime limit."
						  Text 20, 55, 55, 10, "These Include:"
						  Text 30, 70, 125, 10, "1. Months you are over 60 years old"
						  Text 30, 80, 310, 10, "2. Months you are living on a reservation where at least 50% of the adults were not employed"
						  Text 30, 90, 360, 10, "3. Months when you are a victim of family violence AND have an approved family violence waiver plan"
						  Text 30, 100, 335, 10, "4. Months you don't receive the cash portion of MFIP (*talk to your financial worker for more details)"
						  Text 30, 110, 350, 10, "5. Months you are a parent under 18 years of age and complying with your school or social service plan"
						  Text 30, 120, 395, 10, "6. Months you are 18 or 19 years old and do not have a high school diploma/GED AND complying with a school plan"
						  Text 40, 135, 355, 25, "Note: If you are eligible for an exemption but you are not complying with program requirements and do not meet a good cause reason, those months will count toward the lifetime limit. If you have questions about possible good cause reasons, talk to a worker."
							  Text 10, 175, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)

						  mf_step_time_limits_viewed = True
						End If

						'THIRD - MFIp Extension Eligibility'
						If MFIP_orientation_step = mf_step_extension Then
						  GroupBox 10, 10, 450, 295, "MFIP Extension Eligibility"
						  Text 20, 30, 165, 10, "You may be eligible for an MFIP Extension if:"
						  Text 30, 45, 375, 10, "- You are a single or two parent household working the required number of hours that meet extension eligibility"
						  Text 30, 55, 365, 10, "- Your health care provider states you are only able to work 20 hours per week due to an illness or disability"
						  Text 20, 75, 380, 20, "A qualified professional verifies you have one or more of the conditions below that severely limits your ability to obtain or maintain suitable employment for 20 or more hours per week:"
						  Text 30, 100, 165, 10, "- Developmentally Disabled or Mentally Ill"
						  Text 30, 110, 95, 10, "- Learning Disability"
						  Text 30, 120, 60, 10, "- IQ Below 80"
						  Text 30, 130, 260, 10, "- You are ill/injured or incapacitated that's expected to last more than 30 days"
						  Text 20, 145, 125, 10, "A qualified professional verifies:"
						  Text 35, 160, 280, 15, "You are needed in the home to provide care for a family member or foster child in the household that is expected to continue for more than 30 days "
						  Text 35, 185, 285, 35, "A child or adult in the home meets the Special Medical Criteria for home care services or a home and community-based waiver services program, severe emotional disturbance (SED diagnosed child) or serious and persistent mental illness (SPMI diagnosed adult)"
						  Text 35, 225, 275, 20, "You have significant barriers to employment and determined Unemployable by a vocational specialist or other qualified professional designated by the county"
						  Text 35, 250, 165, 10, "You are a victim of family violence"
						  Text 20, 265, 415, 30, "If you believe you meet any of the criteria's above it's important to discuss with your financial worker AND your employment counselor. You may qualify for a modified employment plan prior to reaching your 60-month as well as receive an extension of your cash benefits."
							  Text 10, 310, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)

						  mf_step_extension_viewed = True
						End If

						'FOURTH - Family Violence'
						If MFIP_orientation_step = mf_step_dv Then
						  GroupBox 10, 10, 450, 75, "Family Violence Resources/Supports"
						  Text 20, 30, 390, 10, "Your financial worker discussed and provided information regarding resources if you are a victim of family violence."
						  Text 20, 45, 375, 35, " Please review that brochure if you need assistance with shelter and/or supports Domestic Violence Information (DHS 3477) and Family Violence Referral (DHS 3323). If you are a victim of domestic violence, you may choose to work with your assigned Employment Counselor to determine if you are eligible for a Family Violence Waiver to allow your family time and flexibility to focus on safety issues."
							  Text 10, 90, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  PushButton 385, 90, 75, 15, "DHS - 3477", open_dhs_3477_btn
						  PushButton 385, 105, 75, 15, "DHS - 3323", open_dhs_3323_btn

						  mf_step_dv_viewed = True
						  'ADD BUTTON DHS 3477
						  'ADD BUTTON DHS 3323
						End If

						'FIFTH - Expectations'
						If MFIP_orientation_step = mf_step_expectations Then
						  GroupBox 10, 10, 450, 110, "Expectations of Participants Approved for the MFIP Program"
						  Text 20, 30, 360, 20, "MFIP services focus on putting you on the most direct path to employment and other related steps that will support long-term economic stability."
						  Text 20, 55, 375, 20, "While you are expected to work, look for work, or participate in activities to prepare for work, the steps toward economic stability look different for all families and participants."
						  Text 20, 80, 405, 20, "Employment Services have a variety of tools to address the unique needs of each family. You will hear more about these tools and resources during your Employment Services Overview."
							  Text 10, 125, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)

						  mf_step_expectations_viewed = True
						End If

						'SIXTH - Choosing ESP'
						If MFIP_orientation_step = mf_step_esp Then
						  GroupBox 10, 10, 450, 155, "Choosing an MFIP Employment Service Provider (ESP)"
						  Text 25, 25, 345, 10, "As part of the MFIP program you are required to work with an MFIP Employment Service Provider (ESP)."
						  Text 25, 40, 410, 20, "There's a variety of providers available to help support your employment goals. On the MFIP ESP Choice Sheet, choose the top three providers you'd like to work with listing your top three choices with 1 being the provider you most want to work with."
						  Text 25, 65, 330, 10, "We will do our best to refer you to one of your top three choices depending on available openings."
						  Text 25, 85, 195, 10, "There are a few exceptions in choosing your provider:"
						  Text 40, 100, 345, 10, "If you have worked with an MFIP ESP in the past ninety (90) days, you may be referred to that provider."
						  Text 40, 115, 350, 20, "If you are under 18 and do not have a HS diploma/GED, you will be referred to Minnesota Visiting Nurse Association to discuss your education and employment options"
						  Text 40, 140, 345, 20, "If you have used 60 months or more of your TANF time limit and granted an extension under a specific category you will be referred to an agency that specializes in that type of extension."
							  Text 10, 170, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  ' PushButton 385, 170, 75, 15, "Choice Sheet", open_choice_sheet_btn

						  mf_step_esp_viewed = True
						  'ADD BUTTON - CHOICE SHEET ???'
						End If

						'SEVENTH - Assignment and Compliance'
						If MFIP_orientation_step = mf_step_compliance Then
						  GroupBox 10, 10, 450, 110, "Assignment and Compliance with MFIP Employment Services"
						  Text 25, 30, 295, 10, "Once you are approved for MFIP you will be referred to an Employment Service Provider."
						  Text 25, 45, 375, 20, "In Hennepin County, many of the Employment Services Providers are community based nonprofit organizations who partner with Hennepin County to deliver services."
						  Text 25, 70, 410, 20, "The provider will send you a notice to attend an MFIP Employment Service Overview. You are required to attend the overview and work with your assigned employment service counselor."
						  Text 25, 95, 400, 20, "If you choose not to comply with program requirements, your case may be sanctioned resulting in a reduction of your cash and/or food benefits."
							  Text 10, 125, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)

						  mf_step_compliance_viewed = True
						End If

						'EIGHTH - Developing an EP'
						If MFIP_orientation_step = mf_step_ep Then
						  GroupBox 10, 10, 450, 270, "Developing an Employment Plan (EP) with your MFIP Employment Counselor"
						  Text 20, 25, 430, 25, "Program participants will work with their assigned Employment Counselor to develop an Employment Plan. Your Employment Plan will be based on your goals and will include activities that are intended to lead to employment and financial stability. On the path to stable employment, many different types of activities are available."
						  Text 20, 55, 140, 10, "Some of the allowable activities include:"
						  Text 30, 70, 260, 10, "- Job search (including participation in job clubs, workshops, and hiring events)"
						  Text 30, 80, 260, 10, "- Employment"
						  Text 30, 90, 260, 10, "- Self-employment"
						  Text 30, 100, 260, 10, "- Community work experience and/or volunteer work"
						  Text 30, 110, 260, 10, "- On the job training"
						  Text 30, 120, 260, 10, "- English Language Learning (ELL and ESL) or Functional Work Literacy (FWL)"
						  Text 30, 130, 260, 10, "- Adult Basic Education, GED preparation and Adult High School Diploma"
						  Text 30, 140, 260, 10, "- Job skills training directly related to employment"
						  Text 30, 150, 260, 10, "- Post-Secondary Training and Education"
						  Text 30, 160, 415, 10, "- Other activities that are critical to your family's success in reaching your employment goals such as chemical dependency"
						  Text 35, 170, 260, 10, "treatment, mental health services, social services, and parenting education."
						  Text 20, 190, 430, 25, "You are required to follow through with the activities in your employment plan. If you are unable to complete the activities, contact your Employment Counselor right away to determine if your plan need to be updated. Good communication with your employment counselor can help prevent reduction in your grant (sanctions)."
						  Text 20, 220, 425, 30, "Your Employment Counselor may conduct assessments with you to support you in selecting an education and training path that creates opportunities for long term economic stability. If you have more questions about education and training options, you can also see the Education and Training Brochure (DHS 3366)."
						  Text 20, 255, 420, 20, "Work study programs under the higher education systems may also be available.  Your assigned employment counselor will discuss this opportunity in more detail."
							  Text 10, 285, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  PushButton 385, 285, 75, 15, "DHS - 3366", open_dhs_3366_btn

						  mf_step_ep_viewed = True
						  'ADD BUTTON DHS 3366'
						ENd If

						'NINTH - CCAP'
						If MFIP_orientation_step = mf_step_ccap Then
						  GroupBox 10, 10, 450, 150, "Availability of Childcare Assistance"
						  Text 25, 25, 425, 20, "There are several Childcare Assistance programs (CCAP) available to support your participation in employment, pre-employment activities, training, and/or educational programs"
						  Text 40, 45, 120, 10, "- MFIP/DWP Childcare assistance"
						  Text 40, 55, 120, 10, "- Transition Year Childcare assistance"
						  Text 60, 65, 385, 25, "Many families continue to be eligible for childcare assistance when their MFIP case closes.  It's highly recommended that you speak to your assigned childcare worker to discuss eligibility details specific to your continued needs for assistance when MFIP closes"
						  Text 40, 90, 170, 10, "- Transition Year Extension Childcare assistance"
						  Text 40, 100, 170, 10, "- Basic sliding fee Childcare assistance"
						  Text 60, 110, 215, 10, "If funds are not available, you may be put on a waiting list"
						  Text 25, 125, 430, 10, "Contact your assigned Employment Counselor or Childcare Assistance Worker to discuss eligibility requirements in more detail."
						  Text 25, 140, 395, 10, "If you need help locating childcare provider options, here's a great resource to contact Think Small or (651-641-0332)"
						  GroupBox 10, 165, 450, 65, "Who to Contact about Childcare Assistance?"
						  Text 25, 180, 420, 20, "If you are receiving MFIP your assigned Employment Counselor will work with you to determine how many childcare hours need to be approved based on the activities in your Employment Plan"
						  Text 25, 205, 420, 20, "If you are receiving MFIP but have not been assigned to an Employment Counselor or if your MFIP has closed contact the childcare assistance line directly at 612-348-5937"
						  GroupBox 10, 235, 450, 115, "Program Compliance and Unavailability of Childcare Assistance"
						  Text 25, 250, 425, 20, "The county may NOT impose a sanction for failure to comply with program requirements if you have good cause because of the unavailability of childcare. The inability to obtain childcare does not exempt or extend your TANF time limit."
						  Text 25, 275, 105, 10, "Some good cause reasons are:"
						  Text 35, 285, 135, 10, "- Unavailability of appropriate childcare"
						  Text 35, 295, 135, 10, "- Unreasonable distance to childcare provider"
						  Text 35, 305, 235, 10, "- Provider does not meet health and safety standards for the child(ren)"
						  Text 35, 315, 275, 10, "- The provider charges an excess amount above the maximum the county can pay"
						  Text 25, 330, 335, 10, "Your Childcare Worker or Employment Counselor can discuss good cause reasons in more detail"
							  Text 10, 365, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)

						  mf_step_ccap_viewed = True
						End If

						'TENTH - Incentives'
						If MFIP_orientation_step = mf_step_incentives Then
						  GroupBox 10, 10, 450, 150, "Incentives and Tax Credits"
						  Text 25, 30, 230, 10, "The MFIP program is designed to benefit you when you are working."
						  Text 25, 45, 420, 35, "For example, your financial worker will not budget all your earned income when they calculate the amount of cash and food benefits you are eligible for. When determining your benefit amount, they will not count the first $65 of income you earn AND beyond that, they will only count half of your remaining gross earned income. Here is a link to explain how this works: Bulletin 21-11-01 - DHS Reissues 'Work Will Always Pay ... With MFIP'"
						  Text 25, 85, 425, 10, "If you are working, when you file your taxes apply for the Earned Income Credit and the Minnesota Working Family Credit."
						  Text 25, 100, 225, 10, "Getting a tax refund will NOT affect your eligibility for MFIP."
						  Text 25, 115, 425, 35, "Have your taxes done for FREE! For a list of free tax preparation sites call the Minnesota Department of Revenue at 651-296-3781 or 1-800-652-9094. Neighborhood Volunteer Income Tax Assistance (VITA) sites are available throughout the state. They are open from February 1 through April 15. Some sites are open year around to help you file back taxes. Search for free tax preparation sites at Department of Revenue."
							  Text 10, 165, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  PushButton 385, 165, 75, 15, "DHS Bulletin 21-11-01", open_dhs_bulletin_21_11_01_btn

						  mf_step_incentives_viewed = True
						  'ADD BUTTON BULLETIN 21-11-01'
						End If

						'ELEVENTH - Health Care'
						If MFIP_orientation_step = mf_step_hc Then
						  GroupBox 10, 10, 450, 90, "Health Care"
						  Text 25, 30, 230, 10, "You may qualify for Minnesota Health Care programs."
						  Text 25, 45, 410, 20, "You can apply for health care online at www.mnsure.org (for assistance completing an online application call 1-855-366-7873) or we can mail you a paper application (DHS 6696)."
						  Text 25, 70, 425, 20, "For help with age-appropriate preventive health services check out the Child and Teen Checkup program at: http://edocs.dhs.state.mn.us/lfserver/public/DHS-1826-ENG"
							  Text 10, 105, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  PushButton 385, 105, 75, 15, "DHS - 1826", open_dhs_1826_btn

						  mf_step_hc_viewed = True
						  'ADD BUTTON DHS 1826'
						End If

						If MFIP_orientation_step = mf_completion Then
						  GroupBox 10, 10, 450, 140, "Document MFIP Orientation Completion"
						  Text 20, 30, 135, 10, "For " & CAREGIVER_ARRAY(memb_name_const, caregiver) & ":"
						  Text 25, 50, 215, 10, "Did you verbally review all information in the MFIP Oreientation?"
						  DropListBox 240, 45, 210, 45, "Select One..."+chr(9)+"Yes - all information has been reviewed"+chr(9)+"No - could not complete", CAREGIVER_ARRAY(orientation_done_const, caregiver)
						  Text 25, 65, 240, 10, "Notes from any questions/conversation during the MFIP Orientation:"
						  EditBox 25, 75, 425, 15, CAREGIVER_ARRAY(orientation_notes, caregiver)
						  Text 25, 105, 125, 10, "IF COMPLETE - OPEN ECF NOW"
						  Text 35, 120, 220, 10, "Complete the ESP Choice Sheet (D387) with the resident now."
						  Text 35, 135, 175, 10, "Confirm Choice Sheet Completed and saved to ECF:"
						  DropListBox 205, 130, 140, 45, "Select One..."+chr(9)+"Yes - Choice Sheet Saved to ECF"+chr(9)+"No - could not complete", CAREGIVER_ARRAY(choice_form_done_const, caregiver)
						  Text 210, 155, 250, 25, "MFIP Orientation is now complete for this resident. If this case has a second caregiver that requires the MFIP Orientation, this dialog will reappear for the next caregiver as this is a person based process."
						  PushButton 385, 180, 75, 15, "HSR Manual", open_hsr_manual_btn
						  PushButton 385, 195, 75, 15, "CM 05.12.12.06", cm_05_12_12_06_btn

						  mf_completion_viewed = True
						End If

						Text 470, 5, 80, 10, "MFIP Orientation Topics"
						Text 10, 360, 190, 20, "The entire MFIP Orientation to Financial Serviews Script can be viewed on Sharepoint - Open Word Document here:"

						If MFIP_orientation_step = mf_step_rights_resp Then 	Text 500, 18, 55, 10, "Rights / Resp"
						If MFIP_orientation_step = mf_step_time_limits Then 	Text 504, 33, 55, 10, "Time Limits"
						If MFIP_orientation_step = mf_step_extension Then 	Text 509, 48, 55, 10, "Extention"
						If MFIP_orientation_step = mf_step_dv Then 			Text 497, 63, 55, 10, "Family Violence"
						If MFIP_orientation_step = mf_step_expectations Then 	Text 503, 78, 55, 10, "Expectations"
						If MFIP_orientation_step = mf_step_esp Then 			Text 508, 93, 55, 10, "MFIP ESP"
						If MFIP_orientation_step = mf_step_compliance Then 	Text 497, 108, 55, 10, "ES Compliance"
						If MFIP_orientation_step = mf_step_ep Then 			Text 505, 123, 55, 10, "Emplmt Plan"
						If MFIP_orientation_step = mf_step_ccap Then 			Text 512, 138, 55, 10, "CCAP"
						If MFIP_orientation_step = mf_step_incentives Then 	Text 506, 153, 55, 10, "Incentives"
						If MFIP_orientation_step = mf_step_hc Then 			Text 505, 168, 55, 10, "Health Care"
						If MFIP_orientation_step = mf_completion Then 		Text 502, 188, 55, 10, "Confirmation"


					    If MFIP_orientation_step = mf_completion Then PushButton 495, 365, 50, 15, "DONE", done_btn


					    If MFIP_orientation_step <> mf_step_rights_resp Then 	PushButton 495, 15, 55, 15, "Rights / Resp", button_one
					    If MFIP_orientation_step <> mf_step_time_limits Then 	PushButton 495, 30, 55, 15, "Time Limits", button_two
						If MFIP_orientation_step <> mf_step_extension Then 		PushButton 495, 45, 55, 15, "Extention", button_three
					    If MFIP_orientation_step <> mf_step_dv Then 			PushButton 495, 60, 55, 15, "Family Violence", button_four
					    If MFIP_orientation_step <> mf_step_expectations Then 	PushButton 495, 75, 55, 15, "Expectations", button_five
					    If MFIP_orientation_step <> mf_step_esp Then 			PushButton 495, 90, 55, 15, "MFIP ESP", button_six
					    If MFIP_orientation_step <> mf_step_compliance Then 	PushButton 495, 105, 55, 15, "ES Compliance", button_seven
					    If MFIP_orientation_step <> mf_step_ep Then 			PushButton 495, 120, 55, 15, "Emplmt Plan", button_eight
					    If MFIP_orientation_step <> mf_step_ccap Then 			PushButton 495, 135, 55, 15, "CCAP", button_nine
					    If MFIP_orientation_step <> mf_step_incentives Then 	PushButton 495, 150, 55, 15, "Incentives", button_ten
					    If MFIP_orientation_step <> mf_step_hc Then 			PushButton 495, 165, 55, 15, "Health Care", button_eleven
					    If MFIP_orientation_step <> mf_completion Then 			PushButton 495, 185, 55, 15, "Confirmation", button_twelve

					    ' PushButton 495, 195, 55, 15, "Button 2", Button13
					    ' PushButton 495, 210, 55, 15, "Button 2", Button14
					    ' PushButton 495, 225, 55, 15, "Button 2", Button15
					    ' PushButton 495, 240, 55, 15, "Button 2", Button16
						PushButton 205, 360, 135, 15, "MFIP Oriendation Document", mfip_orientation_word_doc_btn
						' OkButton 495, 365, 50, 15

					EndDialog

					dialog Dialog1
					cancel_confirmation

					err_msg = ""

					If ButtonPressed = next_btn Then MFIP_orientation_step = MFIP_orientation_step + 1
					If ButtonPressed = button_one Then MFIP_orientation_step = mf_step_rights_resp
					If ButtonPressed = button_two Then MFIP_orientation_step = mf_step_time_limits
					If ButtonPressed = button_three Then MFIP_orientation_step = mf_step_extension
					If ButtonPressed = button_four Then MFIP_orientation_step = mf_step_dv
					If ButtonPressed = button_five Then MFIP_orientation_step = mf_step_expectations
					If ButtonPressed = button_six Then MFIP_orientation_step = mf_step_esp
					If ButtonPressed = button_seven Then MFIP_orientation_step = mf_step_compliance
					If ButtonPressed = button_eight Then MFIP_orientation_step = mf_step_ep
					If ButtonPressed = button_nine Then MFIP_orientation_step = mf_step_ccap
					If ButtonPressed = button_ten Then MFIP_orientation_step = mf_step_incentives
					If ButtonPressed = button_eleven Then MFIP_orientation_step = mf_step_hc
					If ButtonPressed = button_twelve Then MFIP_orientation_step = mf_completion


					If ButtonPressed = mfip_orientation_word_doc_btn Then
						run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-es-manual/_layouts/15/Doc.aspx?sourcedoc=%7BCB2C8281-95F1-45EE-84D8-B2DF617AA62C%7D&file=MFIP%20Orientation%20to%20Financial%20Services.docx"
						MFIP_orientation_step = mf_completion
						orientation_script_document_viewed = True
					End If
					If ButtonPressed = open_dhs_4163_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-4163-ENG"
					If ButtonPressed = open_dhs_3477_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG"
					If ButtonPressed = open_dhs_3323_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3323-ENG"
					If ButtonPressed = open_dhs_3366_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3366-ENG"
					If ButtonPressed = open_dhs_bulletin_21_11_01_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_FILE&RevisionSelectionMethod=LatestReleased&Rendition=Primary&allowInterrupt=1&noSaveAs=1&dDocName=dhs-328254"
					If ButtonPressed = open_dhs_1826_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-1826-ENG"

					If ButtonPressed = open_hsr_manual_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/MFIP_Orientation.aspx"
					If ButtonPressed = cm_05_12_12_06_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_0005121206"
					' If ButtonPressed = cm_28_12_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_002812"




					If mf_step_rights_resp_viewed = True and mf_step_time_limits_viewed = True and mf_step_extension_viewed = True and mf_step_dv_viewed = True and mf_step_expectations_viewed = True and mf_step_esp_viewed = True and mf_step_compliance_viewed = True and mf_step_ep_viewed = True and mf_step_ccap_viewed = True and mf_step_incentives_viewed = True and mf_step_hc_viewed = True and mf_completion_viewed = True Then all_mfip_orientation_info_viewed = True
					If orientation_script_document_viewed = True and mf_completion_viewed = True Then all_mfip_orientation_info_viewed = True


					' MsgBox "DONE? - " & CAREGIVER_ARRAY(orientation_done_const, caregiver) & vbCr & "CHOICE SHEET? - " & CAREGIVER_ARRAY(choice_form_done_const, caregiver)
					If all_mfip_orientation_info_viewed = False and CAREGIVER_ARRAY(orientation_done_const, caregiver) = "No - could not complete" Then err_msg = err_msg & vbCr & "* You must review the entire MFIP Orientation before continuing."
					If CAREGIVER_ARRAY(orientation_done_const, caregiver) = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate if the MFIP Orientation has been completed."
					If CAREGIVER_ARRAY(orientation_done_const, caregiver) = "Yes - all information has been reviewed" and CAREGIVER_ARRAY(choice_form_done_const, caregiver) = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate if the MFIP ESP Choice Sheet has been completed in ECF."

					If ButtonPressed = done_btn and err_msg <> "" Then MsgBox err_msg
					' If ButtonPressed = done_btn Then MsgBox err_msg
					If ButtonPressed <> done_btn Then err_msg = "HOLD"

				Loop Until all_mfip_orientation_info_viewed = True and err_msg = ""
			End If
			If CAREGIVER_ARRAY(orientation_done_const, caregiver) = "Yes - all information has been reviewed" Then CAREGIVER_ARRAY(orientation_done_const, caregiver) = True
			If CAREGIVER_ARRAY(orientation_done_const, caregiver) = "No - could not complete" Then CAREGIVER_ARRAY(orientation_done_const, caregiver) = False
			If CAREGIVER_ARRAY(choice_form_done_const, caregiver) = "Yes - Choice Sheet Saved to ECF" Then CAREGIVER_ARRAY(choice_form_done_const, caregiver) = True
			If CAREGIVER_ARRAY(choice_form_done_const, caregiver) = "No - could not complete" Then CAREGIVER_ARRAY(choice_form_done_const, caregiver) = False

			'HERE WE HAVE A DIALOG TO GO TO EMPS AND GIVE INSTRUCTION ON HOW TO COMPLETE IT
			If (CAREGIVER_ARRAY(orientation_needed_const, caregiver) = True and CAREGIVER_ARRAY(orientation_done_const, caregiver) = True) or CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True Then
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 281, 185, "Update EMPS Panel"
				  ButtonGroup ButtonPressed
				    PushButton 125, 135, 145, 15, "The EMPS Panel Update is Complete", emps_update_complete_btn
				  Text 15, 10, 125, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
				  If CAREGIVER_ARRAY(orientation_needed_const, caregiver) = True Then Text 35, 20, 205, 10, "NEEDS an MFIP Orientation"
				  If CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True Then Text 35, 20, 205, 10, "Is Exempt from having an MFIP Orientation"
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = True Then Text 15, 35, 255, 10, "The MFIP Orientation to Financial Services is Completed"
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = False Then Text 15, 35, 255, 10, "The MFIP Orientation to Financial Services is NOT Completed"
				  If CAREGIVER_ARRAY(choice_form_done_const, caregiver) = True Then Text 15, 45, 255, 10, "The ESP Choice Sheet in ECF is Completed"
				  If CAREGIVER_ARRAY(choice_form_done_const, caregiver) = False Then Text 15, 45, 255, 10, "The ESP Choice Sheet in ECF is NOT Completed"

				  Text 15, 65, 260, 10, "This person has met the requirement for the MFIP Orientation."
				  If CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True Then Text 20, 75, 260, 10, "Exemption Reason: " & CAREGIVER_ARRAY(exemption_reason_const, caregiver)
				  GroupBox 15, 90, 255, 65, "Update EMPS Panel Now"
				  Text 25, 105, 210, 10, "Update panel: EMPS for " & CAREGIVER_ARRAY(memb_name_const, caregiver)
				  Text 30, 115, 45, 10, "Fin Orient Dt: "
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = True Then Text 85, 115, 40, 10, date
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = False Then Text 85, 115, 40, 10, "__ __ __"
				  Text 45, 125, 35, 10, "Attended: "
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = True Then Text 85, 125, 20, 10, "Y"
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = False Then Text 85, 125, 20, 10, "N"
				  Text 30, 135, 45, 10, "Good Cause:"
				  If CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = False Then Text 85, 135, 20, 10, "__"
				  If CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True Then Text 85, 135, 20, 10, CAREGIVER_ARRAY(emps_exemption_code_const, caregiver)
				EndDialog

				dialog Dialog1

				Call start_a_blank_CASE_NOTE        'QUESTION - I believe we are going to lose the tie to the DAIL here - do we care?'

				If CAREGIVER_ARRAY(orientation_done_const, caregiver) = True Then

					Call write_variable_in_CASE_NOTE("MFIP Orientation completed with " & CAREGIVER_ARRAY(memb_name_const, caregiver))
					Call write_bullet_and_variable_in_CASE_NOTE("Orientation Completed on", date)
					Call write_bullet_and_variable_in_CASE_NOTE("Orientation Notes", CAREGIVER_ARRAY(orientation_notes, caregiver))
					If CAREGIVER_ARRAY(choice_form_done_const, caregiver) = True Then Call write_variable_in_CASE_NOTE("* ESP Choice Sheet: Completed in Case File ")
					Call write_variable_in_CASE_NOTE("---")
					Call write_variable_in_CASE_NOTE(CAREGIVER_ARRAY(memb_name_const, caregiver) & " did not meet an exemption from completing an MFIP Orientation")
					Call write_variable_in_CASE_NOTE("---")
					Call write_variable_in_CASE_NOTE(worker_signature)

				ElseIf CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True Then

					Call write_variable_in_CASE_NOTE(CAREGIVER_ARRAY(memb_name_const, caregiver) & " is Exempt from MFIP Orientation")
					Call write_bullet_and_variable_in_CASE_NOTE("Assessment Completed", date)
					Call write_bullet_and_variable_in_CASE_NOTE("Exemption Reason", CAREGIVER_ARRAY(exemption_reason_const, caregiver))
					Call write_variable_in_CASE_NOTE("---")
					Call write_variable_in_CASE_NOTE(worker_signature)

				End If
				PF3

                call back_to_SELF

			End If
			' MsgBox CAREGIVER_ARRAY(memb_name_const, caregiver) & " - DONE"
		Next
	End If
	' MsgBox "STOP HERE"
end function
'===========================================================================================================================

'DECLARATIONS================================================================================================================
'constants for the HH_MEMB_ARRAY array
const ref_number				= 0
const full_name_const			= 1
const age						= 2
const memb_is_caregiver			= 3
const cash_request_const		= 4
const hours_per_week_const		= 5
const exempt_from_ed_const		= 6
const comply_with_ed_const		= 7
const orientation_needed_const	= 8
const orientation_done_const	= 9
const orientation_exempt_const	= 10
const exemption_reason_const	= 11
const emps_exemption_code_const	= 12
const choice_form_done_const	= 13
const orientation_notes			= 14
const last_const				= 15

Dim HH_MEMB_ARRAY(last_const, 0)        'This is set up like an array but only works for the person the DAIL is for.
'============================================================================================================================


'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

EMReadScreen name_for_dail, 57, 5, 5			'Reading the name of the client
'This next block will determine the name of the client the message is for
'If the message is for someone other than M01 - the name is writen next to the name of M01
other_person = InStr(name_for_dail, "--(")	'This determines if it for someone other than M01
'This is for if the message is for M01'
If other_person = 0 Then
	comma_loc = InStr(name_for_dail, ",")  	'Determines the end of the last name
	dash_loc = InStr(name_for_dail, "-")	'Determines the end of the name
	EMReadscreen last_name, comma_loc - 1, 5, 5									'Reading the last name
	EMReadscreen middle_exists, 1, 5, 5 + (dash_loc - 2)						'Determines if clt's middle initial is listed
	If middle_exists = " " Then 												'If not - reads first name
		EMReadscreen first_name, dash_loc - comma_loc - 5, 5, comma_loc + 5
	Else 																		'If so - reads first name
		EMReadScreen first_name, dash_loc - comma_loc - 3, 5, comma_loc + 5
	End If
'This is for if the message is for a different HH Member
Else
	end_other = InStr(name_for_dail, ")--")
	comma_loc = InStr(other_person, name_for_dail, ",")
	EMReadscreen last_name, comma_loc - other_person - 3, 5, other_person + 7
	EMReadscreen middle_exists, 1, 5, end_other + 2
	If middle_exists = " " Then
		EMReadscreen first_name, end_other - comma_loc - 3, 5, comma_loc + 5
	Else
		EMReadScreen first_name, end_other - comma_loc - 1, 5, comma_loc + 5
	End If
    HH_MEMB_ARRAY(ref_number, 0) = "01"
End If
HH_MEMB_ARRAY(full_name_const, 0) = first_name & " " & last_name		'putting the name into one string

'Goes to STAT
EMSendKey "S"
transmit

Call EMWriteScreen "MEMB", 20, 71
transmit
EMWriteScreen "01", 20, 76

If HH_MEMB_ARRAY(ref_number, 0) = "01" Then
    EMReadScreen HH_MEMB_ARRAY(age, 0), 3, 8, 76					'Reading the name and age if there was not 'Access Denied' issue
Else
    Do
        EMReadscreen memb_last_name_const, 25, 6, 30
        EMReadscreen memb_first_name_const, 12, 6, 63
        memb_last_name_const = trim(replace(memb_last_name_const, "_", ""))
        memb_first_name_const = trim(replace(memb_first_name_const, "_", ""))

        If memb_first_name_const & " " & memb_last_name_const = HH_MEMB_ARRAY(full_name_const, 0) Then
            EMReadScreen HH_MEMB_ARRAY(age, 0), 3, 8, 76					'Reading the name and age if there was not 'Access Denied' issue
            EMReadScreen HH_MEMB_ARRAY(ref_number, 0), 2, 4, 33
            Exit Do
        End If
        transmit      'Going to the next MEMB panel
    	Emreadscreen edit_check, 7, 24, 2 'looking to see if we are at the last member
    LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

End If

HH_MEMB_ARRAY(age, clt_count) = trim(HH_MEMB_ARRAY(age, clt_count))			'formatting the age and name information.
If HH_MEMB_ARRAY(age, clt_count) = "" Then HH_MEMB_ARRAY(age, clt_count) = 0
HH_MEMB_ARRAY(age, clt_count) = HH_MEMB_ARRAY(age, clt_count) * 1


family_cash_program = "MFIP"			'defaulting to MFIP as the program selection.

'this iswhere the main functionality of this script is called.
'We are using a function because this needs to match the experiance in other scripts.
'This function will call dialogs and enter CASE/NOTEs - eventually it may update EMPS panels
Call complete_MFIP_orientation(HH_MEMB_ARRAY, ref_number, full_name_const, age, memb_is_caregiver, cash_request_const, hours_per_week_const, exempt_from_ed_const, comply_with_ed_const, orientation_needed_const, orientation_done_const, orientation_exempt_const, exemption_reason_const, emps_exemption_code_const, choice_form_done_const, orientation_notes, family_cash_program)

'Now that the CASE/NOTES are completed the script will gather information for the end_msg report out MsgBox
'This next block is ONLY to fill the end_msg
If family_cash_program = "DWP" Then
	STATS_manualtime = 60		'if DWP - the manual time is changed becuase we didn't complete an orientation
	end_msg = "The NOTES - MFIP Orienation script has completed without taking any action." & vbCr
	end_msg = end_msg & "You have indicated that the family cash program is DWP." & vbCr & vbCr
	end_msg = end_msg & "This script does not have support for the financial orientation and information on DWP cases. This functionality is built to specifically support MFIP cases and MFIP caregivers."
Else
	end_msg = "NOTES - MFIP Orientation script run completed." & vbCr

	If HH_MEMB_ARRAY(memb_is_caregiver, 0) = True Then
		caregiver_detail = HH_MEMB_ARRAY(full_name_const, 0) & " is a caregiver on this case." & vbCr
		If HH_MEMB_ARRAY(orientation_needed_const, 0) = True Then caregiver_detail = caregiver_detail & " - An MFIP Orientation is needed for this caregiver. " & vbCr
		If HH_MEMB_ARRAY(orientation_needed_const, 0) = False Then caregiver_detail = caregiver_detail & " - An MFIP Orientation is NOT needed for this caregiver." & vbCr
		If HH_MEMB_ARRAY(orientation_exempt_const, 0) = True Then
			caregiver_detail = caregiver_detail & " - This caregiver is exemmpt from needing an MFIP Orientation." & vbCr
			caregiver_detail = caregiver_detail & "   Exemption Reason: " & HH_MEMB_ARRAY(exemption_reason_const, 0) & vbCr
		End If
		If HH_MEMB_ARRAY(orientation_done_const, 0) = True Then  caregiver_detail = caregiver_detail & " * The orientation was completed during this script run and a CASE/NOTE has been entered." & vbCr
		If HH_MEMB_ARRAY(orientation_done_const, 0) = False Then  caregiver_detail = caregiver_detail & " * MFIP ORIENTATION NOT COMPLETED AND STILL NEEDED FOR " & HH_MEMB_ARRAY(full_name_const, 0) & "." & vbCr

		end_msg = end_msg & vbCr & caregiver_detail
	End If
End If
end_msg = end_msg & vbCr & "CASE/NOTEs have been made by the script. Updates to EMPS should have been completed manually during the script run. If that is still needed, go back and update STAT/EMPS now."

'End the script.
script_end_procedure_with_error_report(end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------08/31/2022
'--Tab orders reviewed & confirmed----------------------------------------------08/31/2022
'--Mandatory fields all present & Reviewed--------------------------------------08/31/2022
'--All variables in dialog match mandatory fields-------------------------------08/31/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------08/31/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------08/31/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------08/31/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------08/31/2022
'--PRIV Case handling reviewed -------------------------------------------------08/31/2022
'--Out-of-County handling reviewed----------------------------------------------08/31/2022
'--script_end_procedures (w/ or w/o error messaging)----------------------------08/31/2022
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---08/31/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------08/31/2022
'--Incrementors reviewed (if necessary)-----------------------------------------08/31/2022
'--Denomination reviewed -------------------------------------------------------08/31/2022
'--Script name reviewed---------------------------------------------------------08/31/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------08/31/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------08/31/2022
'--comment Code-----------------------------------------------------------------08/31/2022
'--Update Changelog for release/update------------------------------------------08/31/2022
'--Remove testing message boxes-------------------------------------------------08/31/2022
'--Remove testing code/unnecessary code-----------------------------------------08/31/2022					There is still some testing code in the function - this will behandled when moved to FuncLib
'--Review/update SharePoint instructions----------------------------------------08/31/2022					QUESTION - the instructions are connected to interview - should I make 2? or remove the Interview specific identifier
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------							TODO - Once initial testing is done - add feedback to add the script to the HSR manual page
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------TODO
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A
