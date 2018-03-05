'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - BURIAL ASSETS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 600                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'SECTION 01 -- Dialogs
BeginDialog opening_dialog_01, 0, 0, 311, 420, "LTC Burial Assets"
  EditBox 95, 25, 60, 15, MAXIS_case_number
  EditBox 225, 25, 30, 15, hh_member
  DropListBox 165, 45, 90, 15, "Select one..."+chr(9)+"GA"+chr(9)+"Health Care"+chr(9)+"MFIP/DWP"+chr(9)+"MSA/GRH", programs
  EditBox 135, 65, 120, 15, worker_signature
  DropListBox 110, 105, 60, 15, "None"+chr(9)+"CD"+chr(9)+"Money Market"+chr(9)+"Stock"+chr(9)+"Bond", type_of_designated_account
  EditBox 240, 105, 60, 15, account_identifier
  EditBox 180, 130, 120, 15, why_not_seperated
  EditBox 90, 155, 65, 15, account_create_date
  EditBox 225, 155, 75, 15, counted_value_designated
  EditBox 70, 180, 230, 15, BFE_information_designated
  EditBox 65, 230, 80, 15, insurance_policy_number
  EditBox 220, 230, 80, 15, insurance_create_date
  EditBox 80, 255, 220, 15, insurance_company
  EditBox 105, 280, 60, 15, insurance_csv
  EditBox 235, 280, 65, 15, insurance_counted_value
  EditBox 75, 305, 225, 15, insurance_BFE_steps_info
  ButtonGroup ButtonPressed
    PushButton 195, 375, 50, 15, "Next", open_dialog_next_button
    CancelButton 250, 375, 50, 15
  Text 40, 30, 50, 10, "Case Number:"
  Text 175, 30, 45, 10, "HH Member:"
  Text 40, 70, 95, 10, "Please sign your case note:"
  GroupBox 5, 90, 300, 110, "Designated Account Information"
  Text 10, 110, 95, 10, "Type of designated account:"
  Text 180, 110, 60, 10, "Account Identifier:"
  Text 10, 135, 170, 10, "Reason funds could not be separated as applicable:"
  Text 10, 160, 75, 10, "Date Account Created:"
  Text 170, 160, 50, 10, "Counted value:"
  Text 10, 185, 55, 10, "BFE information:"
  GroupBox 5, 215, 300, 110, "Non-Term Life Insurance Information"
  Text 10, 235, 50, 10, "Policy Number:"
  Text 10, 260, 70, 10, "Insurance Company:"
  Text 150, 235, 70, 10, "Date Policy Created:"
  Text 10, 285, 90, 10, "CSV/FV Designated to BFE:"
  Text 180, 285, 50, 10, "Counted Value:"
  Text 10, 310, 65, 10, "Info/Steps on BFE:"
  Text 40, 50, 125, 10, "Program asset is being evaluated for"
  GroupBox 35, 5, 230, 80, "Case and Worker Information"
  Text 25, 350, 260, 20, "Please refer to CM 0015.21 (burial funds) and CM 0015.24 (burial contracts) for information on how to evaluate burial assets for each program."
  GroupBox 5, 335, 300, 65, "Each program evaluates burial assets differently"
EndDialog

'Burial Agreement Dialogs----------------------------------------------------------------------------------------------------
BeginDialog burial_assets_dialog_01, 0, 0, 301, 210, "Burial assets dialog (01)"
  CheckBox 5, 25, 160, 10, "Applied $1500 of burial services to BFE?", applied_BFE_check
  ComboBox 95, 70, 65, 15, "Select One..."+chr(9)+"None"+chr(9)+"AFB"+chr(9)+"CSA"+chr(9)+"IBA"+chr(9)+"IFB"+chr(9)+"RBA", type_of_burial_agreement
  EditBox 220, 70, 65, 15, purchase_date
  EditBox 60, 90, 125, 15, issuer_name
  EditBox 230, 90, 55, 15, policy_number
  EditBox 60, 110, 55, 15, face_value
  EditBox 170, 110, 115, 15, funeral_home
  CheckBox 10, 130, 280, 10, "Primary beneficiary is : Any funeral provider whose interest may appear irrevocably", Primary_benficiary_check
  CheckBox 10, 145, 175, 10, "Contingent Beneficiary is: The estate of the insured ", Contingent_benficiary_check
  CheckBox 10, 160, 215, 10, "Policy's CSV is irrevocably designated to the funeral provider", policy_CSV_check
  ButtonGroup ButtonPressed
    PushButton 95, 185, 50, 15, "Next", next_to_02_button
    CancelButton 155, 185, 50, 15
  Text 10, 75, 85, 10, "Type of burial agreement:"
  Text 165, 75, 50, 10, "Purchase date:"
  Text 10, 95, 50, 10, "Issuer name:"
  Text 200, 95, 30, 10, "Policy #:"
  Text 10, 115, 40, 10, "Face value:"
  Text 120, 115, 50, 10, "Funeral home:"
  GroupBox 0, 5, 290, 170, "Burial agreements"
  Text 35, 40, 205, 25, "NOTE: You can mark specific items/services as applied to the BFE on the following panels. These will be calculated and case noted"
EndDialog




'SECTION 2: Functions/dimming array----------------------------------------------------------------------------------------------------
function new_BS_BSI_heading
  EMGetCursor MAXIS_row, MAXIS_col
  If MAXIS_row = 4 then
    EMSendKey "--------BURIAL SPACE/ITEMS---------------AMOUNT----------STATUS--------------" & "<newline>"
    MAXIS_row = 5
  end if
End function

function new_CAI_heading
  EMGetCursor MAXIS_row, MAXIS_col
  If MAXIS_row = 4 then
    EMSendKey "--------CASH ADVANCE ITEMS---------------AMOUNT----------STATUS--------------" & "<newline>"
    MAXIS_row = 5
  end if
End function

function new_service_heading
  EMGetCursor MAXIS_service_row, MAXIS_service_col
  If MAXIS_service_row = 4 then
    EMSendKey "--------------SERVICE--------------------AMOUNT----------STATUS--------------" & "<newline>"
    MAXIS_service_row = 5
  end if
End function

function case_note_page_four 'check for 4th page of case note
  line_one_for_part_one = "**BURIAL ASSETS (1 of 2) -- Memb: " + hh_member
  line_one_for_part_two = "**BURIAL ASSETS (2 of 2) -- Memb: " + hh_member
  EMReadScreen page_four, 20, 24, 2
  IF page_four = "A MAXIMUM OF 4 PAGES" THEN
    PF7
    PF7
    PF7
    EMsetcursor 4, 3
    EMSendKey line_one_for_part_one
    PF3
    PF9
    EMSendKey line_one_for_part_two
    EMsetcursor 5, 3
  END IF
END function

'Dimming array
DIM calc_array(42, 4)

'Array Map
' 0, 0 - name of asset
' 0, 1 - value of asset
' 0, 2 - status of asset (counted, excluded, not counted, applied to bfe etc)
' 0, 3 - type of asset (service, bs/bse, cash advance items)


calc_array(0, 0) = "Basic Service Funeral Director"
calc_array(0, 3) = "Service"

calc_array(1, 0) = "Embalming"
calc_array(1, 3) = "Service"

calc_array(2, 0) = "Other preperation to body"
calc_array(2, 3) = "Service"

calc_array(3, 0) = "Visitation at funeral chapel"
calc_array(3, 3) = "Service"

calc_array(4, 0) = "Visitation at other facility"
calc_array(4, 3) = "Service"

calc_array(5, 0) = "Funeral serv at funeral chapel"
calc_array(5, 3) = "Service"

calc_array(6, 0) = "Funeral serv at other facility"
calc_array(6, 3) = "Service"


calc_array(7, 0) = "Memorial serv at funeral chapel"
calc_array(7, 3) = "Service"

calc_array(8, 0) = "Memorial serv at other facility"
calc_array(8, 3) = "Service"

calc_array(9, 0) = "Graveside service"
calc_array(9, 3) = "Service"

calc_array(10, 0) = "Transfer remains to funeral home"
calc_array(10, 3) = "Service"

calc_array(11, 0) = "Funeral coach"
calc_array(11, 3) = "Service"

calc_array(12, 0) = "Funderal sedan/limo"
calc_array(12, 3) = "Service"

calc_array(13, 0) = "Service vehicle"
calc_array(13, 3) = "Service"

calc_array(14, 0) = "Forwarding of remains"
calc_array(14, 3) = "Service"

calc_array(15, 0) = "Receiving of remains"
calc_array(15, 3) = "Service"

calc_array(16, 0) = "Direct Cremation"
calc_array(16, 3) = "Service"

calc_array(17, 0) = "Markers/Headstone"
calc_array(17, 3) = "Burial Space/Item"

calc_array(18, 0) = "Engraving"
calc_array(18, 3) = "Burial Space/Item"

calc_array(19, 0) = "Opening/Closing of space"
calc_array(19, 3) = "Burial Space/Item"

calc_array(20, 0) = "Perpetual Care"
calc_array(20, 3) = "Burial Space/Item"

calc_array(21, 0) = "Casket"
calc_array(21, 3) = "Burial Space/Item"

calc_array(22, 0) = "Vault"
calc_array(22, 3) = "Burial Space/Item"

calc_array(23, 0) = "Cemetery plot"
calc_array(23, 3) = "Burial Space/Item"

calc_array(24, 0) = "Crypt"
calc_array(24, 3) = "Burial Space/Item"

calc_array(25, 0) = "Mausoleum"
calc_array(25, 3) = "Burial Space/Item"

calc_array(26, 0) = "Urns"
calc_array(26, 3) = "Burial Space/Item"

calc_array(27, 0) = "Niches"
calc_array(27, 3) = "Burial Space/Item"

calc_array(28, 0) = "Alternative Container"
calc_array(28, 3) = "Burial Space/Item"

calc_array(29, 0) = "Certified death certificate"
calc_array(29, 3) = "Cash Advance Item"

calc_array(30, 0) = "Motor Escort"
calc_array(30, 3) = "Cash Advance Item"

calc_array(31, 0) = "Clergy honorarium"
calc_array(31, 3) = "Cash Advance Item"

calc_array(32, 0) = "Music Honorarium"
calc_array(32, 3) = "Cash Advance Item"

calc_array(33, 0) = "Flowers"
calc_array(33, 3) = "Cash Advance Item"

calc_array(34, 0) = "Obituary notice"
calc_array(34, 3) = "Cash Advance Item"

calc_array(35, 0) = "Crematory charges"
calc_array(35, 3) = "Cash Advance Item"

calc_array(36, 0) = "Acknowledgement card"
calc_array(36, 3) = "Cash Advance Item"

calc_array(37, 0) = "Register book"
calc_array(37, 3) = "Cash Advance Item"

calc_array(38, 0) = "Service folder/prayer cards"
calc_array(38, 3) = "Cash Advance Item"

calc_array(39, 0) = "Luncheon"
calc_array(39, 3) = "Cash Advance Item"

calc_array(40, 0) = "Medical Exam Fee"
calc_array(40, 3) = "Cash Advance Item"

calc_array(41, 0) = ""							'these two options allows users to enter extra assets that aren't previously listed
calc_array(41, 3) = other_status				'these will need to be specially accounted for when generating totals later.

calc_array(42, 0) = ""
calc_array(42, 3) = other_status

'This sets the total to 0 for the calculation later, it also defines the starting dialog since it will be determined dynamically to cut down on dialog count.
running_total = 0
current_dialog = "services"

'function to create dynamic dialog for counting/listing assets. This contains the calculations to build totals
'vairiables to pull through are: the calculation array, the total applied to the BFE, the total of BS/BSI items, the counted asset total, the unavailable asset total
FUNCTION build_dynamic_burial_dialog(calc_array, BFE_total, BS_BSI_total, final_counted_total, final_unavailable_total)
	DO
		DO
			FOR i = 0 TO 42											'defining each value as blank to make sure calculations work properly.
				calc_array(i, 1) = calc_array(i, 1) & ""
			NEXT

			'resetting locations for items in dialog
			row_height = 0
			'defining err_msg as blank
			err_msg = ""

			'dialog to be build dynamically
			BeginDialog Dialog1, 0, 0, 306, 410, "Dialog"
			Text 10, 5, 60, 10, "BFE Total: "												'here the running totals will be displayed for the end user to keep track.
			Text 100, 5, 60, 10, FormatCurrency(running_total)
			Text 10, 15, 60, 10, "Counted Total: "
			Text 100, 15, 60, 10, FormatCurrency(counted_total)
			Text 10, 25, 60, 10, "Unavailable Total: "
			Text 100, 25, 60, 10, FormatCurrency(unavailable_total)
			ButtonGroup ButtonPressed														'the buttons for calculating, and navigating between the other dialogs
				PushButton 215, 10, 50, 15, "Calculate", calc_button
				PushButton 10, 390, 50, 15, "Services", service_button
				PushButton 65, 390, 50, 15, "BS/BSI", BSI_button
				PushButton 120, 390, 50, 15, "CAI", CAI_button
				OkButton 195, 390, 50, 15
				CancelButton 250, 390, 50, 15
			IF current_dialog = "services" THEN Text 25, 35, 240, 10, "BURIAL SERVICES                       VALUE                          STATUS"
			IF current_dialog = "bsi" THEN Text 25, 35, 240, 10, "BURIAL SPACE ITEMS                VALUE                            STATUS"
			IF current_dialog = "cai" THEN Text 25, 35, 240, 10, "CASH ADVANCE ITEMS                  VALUE                            STATUS"
			'Here is where the dialog differs, if a certain dialog is clicked on the asset fields will be generated based on the calc_array
			IF current_dialog = "services" THEN							'values 0-16 are the services assets
				FOR i = 0 TO 16
					Text 10, 50 + (20 * row_height), 110, 10, calc_array(i, 0)
					EditBox 120, 50 + (20 * row_height), 60, 15, calc_array(i, 1)
					DropListBox 195, 50 + (20 * row_height), 95, 15, ""+chr(9)+"counted"+chr(9)+"excluded"+chr(9)+"unavailable"+chr(9)+"applied to BFE"+chr(9)+"applied to BFE/counted"+chr(9)+"applied to BFE/unavailable", calc_array(i, 2)
					row_height = row_height + 1
				NEXT
			END IF
			IF current_dialog = "bsi" THEN								'values 17-28 are the Burial space/burial space items assets
				FOR i = 17 TO 28
					Text 10, 50 + (20 * row_height), 110, 10, calc_array(i, 0)
					EditBox 120, 50 + (20 * row_height), 60, 15, calc_array(i, 1)
					DropListBox 195, 50 + (20 * row_height), 95, 15, ""+chr(9)+"counted"+chr(9)+"excluded", calc_array(i, 2)
					row_height = row_height + 1
				NEXT
			END IF
			IF current_dialog = "cai" THEN								'values 29-40 are the cash advance items assets
				FOR i = 29 TO 40
					Text 10, 50 + (20 * row_height), 110, 10, calc_array(i, 0)
					EditBox 120, 50 + (20 * row_height), 60, 15, calc_array(i, 1)
					DropListBox 195, 50 + (20 * row_height), 95, 15, ""+chr(9)+"counted"+chr(9)+"unavailable", calc_array(i, 2)
					row_height = row_height + 1
				NEXT
				FOR i = 41 TO 41										'this is the extra spot to allow users to enter their own asset, value, and type.
					Text 10, 50 + (20 * row_height), 40, 10, "Other 1:"
					DropListBox 50, 50 + (20 * row_height), 55, 15, "Select Type"+chr(9)+"Services"+chr(9)+"BS/BSI"+chr(9)+"CAI", calc_array(i, 3)
					EditBox 120, 50 + (20 * row_height), 60, 15, calc_array(i, 1)
					DropListBox 195, 50 + (20 * row_height), 95, 15, ""+chr(9)+"counted"+chr(9)+"excluded"+chr(9)+"unavailable"+chr(9)+"applied to BFE"+chr(9)+"applied to BFE/counted"+chr(9)+"applied to BFE/unavailable", calc_array(i, 2)
					row_height = row_height + 1
					Text 15, 50 + (20 * row_height), 40, 10, "This is a:"
					EditBox 55, 50 + (20 * row_height), 60, 15, calc_array(i, 0)
				NEXT
				FOR i = 42 to 42										'this is the extra spot to allow users to enter their own asset, value, and type.
					Text 10, 70 + (20 * row_height), 40, 10, "Other 2:"
					DropListBox 50, 70 + (20 * row_height), 55, 15, "Select Type"+chr(9)+"Services"+chr(9)+"BS/BSI"+chr(9)+"CAI", calc_array(i, 3)
					EditBox 120, 70 + (20 * row_height), 60, 15, calc_array(i, 1)
					DropListBox 195, 70 + (20 * row_height), 95, 15, ""+chr(9)+"counted"+chr(9)+"excluded"+chr(9)+"unavailable"+chr(9)+"applied to BFE"+chr(9)+"applied to BFE/counted"+chr(9)+"applied to BFE/unavailable", calc_array(i, 2)
					row_height = row_height + 1
					Text 15, 70 + (20 * row_height), 40, 10, "This is a:"
					EditBox 55, 70 + (20 * row_height), 60, 15, calc_array(i, 0)
				NEXT
			END IF
			EndDialog

			'calling the dynamic dialog
			Dialog Dialog1
			cancel_confirmation
			'if the calc button or any of the other dialogs buttons are pressed this will activate the calculation function thus generating new totals
			IF ButtonPressed = calc_button or ButtonPressed = service_button  or ButtonPressed = BSI_button or ButtonPressed = CAI_button or ButtonPressed = -1 THEN
				running_total = 0					'at the start of each calculation phase the totals need to be reset.
				counted_total = 0
				unavailable_total = 0
				BFE_full = FALSE					'this is re-defined each time because if the BFE is true the calc function needs to act a different way.
				FOR i = 0 TO 42
					'first the array will single out the items applied to the BFE that are not blank and start adding them up.
					IF calc_array(i, 1) <> "" AND IsNumeric(calc_array(i, 1)) = TRUE AND (calc_array(i, 2) = "applied to BFE" OR calc_array(i, 2) = "applied to BFE/counted" OR calc_array(i, 2) = "applied to BFE/unavailable") THEN
						calc_array(i, 1) = calc_array(i, 1) * 1								'converting what was entered to a number so we can manipulate it.
						running_total = running_total + calc_array(i, 1)					'a running total based on the values entered for the BFE related status
						IF running_total > 1500 THEN 										'if we hit the limit of the BFE at $1500 we have to check if one of the assets needs to be re-defined
							IF BFE_full <> TRUE THEN										'if we have not filled the BFE for this calc phase
								IF calc_array(i, 2) <> "applied to BFE/unavailable" THEN 	'if the current asset is not listed as applied to bfe/unavailable then we change it to counted.
									calc_array(i, 2) = "applied to BFE/counted"
									msgbox "BFE limit of $1500 has been met. " & vbCr & calc_array(i, 0) & " was relabeled as Applied to BFE/counted and the amount was split between the BFE and Counted Totals. Please review status for this change for accuracy."
								END IF
								BFE_full = TRUE												'if we've gone over 1500 we are over the BFE thus we need to change it to true.
							ELSE
								IF calc_array(i, 2) = "applied to BFE" THEN 				'otherwise if someone has left items that are applied to the BFE after we've already filled it we change those to counted.
									msgbox "BFE limit of $1500 has been met. " & vbCr & calc_array(i, 0) & " was relabeled as counted. Please review status for this change for accuracy."
									calc_array(i, 2) = "counted"
								END IF
							END IF
							remainder_total = running_total - 1500							'this is what builds out remainder total so we can accurately determine the running totals included the split bfe/other values.
							IF calc_array(i, 2) = "applied to BFE/counted" OR calc_array(i, 2) = "counted" THEN counted_total = remainder_total + counted_total
							IF calc_array(i, 2) = "applied to BFE/unavailable" OR calc_array(i, 2) = "unavailable" THEN unavailable_total = remainder_total + unavailable_total
							running_total = 1500											'since we've exceeded/maxed the BFE then we can redefine it as $1500
						ELSEIF running_total < 1500 THEN 									'if the running total for the BFE is under $1500 we can keep BFE_full as false
								BFE_full = FALSE
						ELSEIF running_total = 1500 THEN									'if the running total for the BFE is EXACTLY 1500 then the current item that was added is changed to applied
								calc_array(i, 2) = "applied to BFE"
								BFE_full = TRUE												'we also need to mark the BFE as being full.
						END IF
					ELSE																	'if the current item isn't related to the BFE then we just add that value to the other totals.
						IF calc_array(i, 1) <> "" AND IsNumeric(calc_array(i, 1)) = TRUE AND calc_array(i, 2) = "counted" THEN
							calc_array(i, 1) = calc_array(i, 1) * 1
							counted_total = counted_total + calc_array(i, 1)
						END IF
						IF calc_array(i, 1) <> "" AND IsNumeric(calc_array(i, 1)) = TRUE AND calc_array(i, 2) = "unavailable" THEN
							calc_array(i, 1) = calc_array(i, 1) * 1
							unavailable_total = unavailable_total + calc_array(i, 1)
						END IF
					END IF
				NEXT
			END IF
			'after the calculation is completed if the user hit a button to nav to another dialog the value to determine the dialog is redefined.
			IF ButtonPressed = service_button THEN current_dialog = "services"
			IF ButtonPressed = BSI_button THEN current_dialog = "bsi"
			IF ButtonPressed = CAI_button THEN current_dialog = "cai"
			'error proofing to ensure that people are entering the correct information into the fields.
			FOR i = 0 TO 42
				IF calc_array(i, 1) <> "" AND IsNumeric(calc_array(i, 1)) = FALSE THEN err_msg = err_msg & "The value entered for " & calc_array(i, 0) & " is not a number." & vbCr
				IF calc_array(i, 1) <> "" AND calc_array (i, 2) = "" THEN err_msg = err_msg & "You entered a value entered for " & calc_array(i, 0) & " but did not select a status." & vbCr
				IF calc_array(i, 1) = "" AND calc_array (i, 2) <> "" THEN err_msg = err_msg & "You entered a status entered for " & calc_array(i, 0) & " but did not select a value." & vbCr
			NEXT
			'additional error proofing is needed for array positions 41 and 42 since the type is needed for these write
			IF calc_array(41, 1) <> "" AND calc_array(41, 3) = "Select Type" THEN err_msg = err_msg & "You entered a value entered for " & calc_array(41, 0) & " but did not select a type." & vbCr
			IF calc_array(42, 1) <> "" AND calc_array(42, 3) = "Select Type" THEN err_msg = err_msg & "You entered a value entered for " & calc_array(42, 0) & " but did not select a type." & vbCr
			IF calc_array(41, 1) <> "" AND calc_array(41, 0) = "" THEN err_msg = err_msg & "You entered a value entered for " & calc_array(41, 0) & " but did enter what the value is for." & vbCr
			IF calc_array(42, 1) <> "" AND calc_array(42, 0) = "" THEN err_msg = err_msg & "You entered a value entered for " & calc_array(42, 0) & " but did enter what the value is for." & vbCr
			IF err_msg <> "" THEN msgbox "Please resolve these issues:" & vbCr & vbCr & err_msg
		LOOP until ButtonPressed = -1								'keep looping this calculation and error checking until the worker clicks ok

		'Here is where the final totals are being build.
		'Totalling BFE and defining starting point as 0.
		remainder_counted_total = 0
		remainder_unavailable_total = 0
		BFE_total = 0
		BFE_running_total = 0
		FOR i = 0 to 42								'since the BFE can be applied to all of the assets we have to check each and every asset with a value marked as applied to BFE and or combination
			IF calc_array(i, 1) <> "" AND IsNumeric(calc_array(i, 1)) = TRUE AND calc_array(i, 2) = "applied to BFE" THEN  'if the item has a status of applied to BFE we need to add it to the total
				calc_array(i, 1) = calc_array(i, 1) * 1
				BFE_running_total = BFE_running_total + calc_array(i, 1)
				IF BFE_running_total > 1500 THEN 											'however we must still keep track of the limit of the bfe total.
					BFE_total = 1500
				ELSEIF BFE_running_total < 1500 THEN
					BFE_total = BFE_running_total
				ELSEIF BFE_running_total = 1500 THEN
					BFE_total = 1500
				END IF
			ELSEIF calc_array(i, 1) <> "" AND IsNumeric(calc_array(i, 1)) = TRUE AND calc_array(i, 2) = "applied to BFE/counted" THEN  'here we determine the remainders of the split status options to count everything correctly
				BFE_running_total = BFE_running_total + calc_array(i, 1)
				remainder_counted_total = BFE_running_total - 1500
				final_counted_total = remainder_counted_total
				BFE_total = 1500
			ELSEIF calc_array(i, 1) <> "" AND IsNumeric(calc_array(i, 1)) = TRUE AND calc_array(i, 2) = "applied to BFE/unavailable" THEN
				BFE_running_total = BFE_running_total + calc_array(i, 1)
				remainder_unavailable_total = BFE_running_total - 1500
				final_unavailable_total = remainder_unavailable_total
				BFE_total = 1500
			END IF
		NEXT

		'Totalling BS/BSI Excluded Amount
		BS_BSI_total = 0
		FOR i = 17 TO 28									'values in the array of 17-28 are defined as burial space/burial space items, there can be status counted or excluded only.
			IF calc_array(i, 1) <> "" AND IsNumeric(calc_array(i, 1)) = TRUE AND calc_array(i, 2) = "excluded" THEN
				calc_array(i, 1) = calc_array(i, 1) * 1
				BS_BSI_total = BS_BSI_total + calc_array(i, 1)
			END IF
		NEXT															'the following two items need to be calculated as they can be labeled as BS/BSI items.
		IF calc_array(41, 1) <> "" AND IsNumeric(calc_array(41, 1)) = TRUE AND calc_array(41, 2) = "excluded" AND calc_array(41, 3) = "BS/BSI" THEN
			calc_array(41, 1) = calc_array(41, 1) * 1
			BS_BSI_total = BS_BSI_total + calc_array(41, 1)
		END IF
		IF calc_array(42, 1) <> "" AND IsNumeric(calc_array(42, 1)) = TRUE AND calc_array(42, 2) = "excluded" AND calc_array(42, 3) = "BS/BSI" THEN
			calc_array(42, 1) = calc_array(42, 1) * 1
			BS_BSI_total = BS_BSI_total + calc_array(42, 1)
		END IF

		'Totalling Unavailable Amount
		final_unavailable_total = 0 + remainder_unavailable_total						'we need to define that starting value of the unavailable total to include the remainder of any applied to bfe/unavailable
		FOR i = 0 to 42
			IF calc_array(i, 1) <> "" AND IsNumeric(calc_array(i, 1)) = TRUE AND calc_array(i, 2) = "unavailable" THEN
				calc_array(i, 1) = calc_array(i, 1) * 1
				final_unavailable_total = final_unavailable_total + calc_array(i, 1)
			END IF
		NEXT

		'Totalling Counted Amount
		final_counted_total = 0 + remainder_counted_total								'we need to define that starting value of the unavailable total to include the remainder of any applied to bfe/counted
		FOR i = 0 to 42
			IF calc_array(i, 1) <> "" AND IsNumeric(calc_array(i, 1)) = TRUE AND calc_array(i, 2) = "counted" THEN
				calc_array(i, 1) = calc_array(i, 1) * 1
				final_counted_total = final_counted_total + calc_array(i, 1)
			END IF
		NEXT

		'here we are defining a variable to be used in a special message box that will allow users to review before case noting and go back to fix things if needed.
		break_down_display = ""
		FOR i = 0 TO 42										'this will grab every value for assets that are not blank and bring them over to the double check message box.
			IF calc_array(i, 1) <> "" THEN
				break_down_display = break_down_display & vbCr & calc_array(i, 1) & ", " & calc_array(i, 0) & ",    " & calc_array(i, 3) & ",    " & calc_array(i, 2)
			END IF
		NEXT

		double_check_display = ""					'blank message box variable so that it won't carry information from previous runs if user chooses to alter data.
		double_check_display = Msgbox ("The script has finished calculating your Burial Assets. Please review the following information for accuracy." & vbCr &_
			"If this is accurate, press YES to continue." & vbCr &_
			"If this needs modification, press NO to retry." & vbCr &_
			"If you wish to cancel the script, press CANCEL." & vbCr &_
			break_down_display & vbCr & vbCr &_
			"Applied to BFE: " & BFE_total & vbCr &_
			"BS/BSI Excluded: " & BS_BSI_total & vbCr &_
			"Unavailable: " & final_unavailable_total & vbCr &_
			"Counted: " & final_counted_total, vbYesNoCancel + vbSystemModal + vbInformation, "PLEASE REVIEW")
		IF double_check_display = vbCancel THEN stopscript
	LOOP UNTIL double_check_display = vbYes					'we loop until the user decided they are happy with the results and then move on.
END FUNCTION

'SECTION 03: The script----------------------------------------------------------------------------------------------------
EMConnect "" 		'connecting to MAXIS
Call MAXIS_case_number_finder(MAXIS_case_number)	'grabbing the case number

insurance_policy_number = "none"			'establishing value of the variable

'calling the initial dialog
DO
	err_msg = "" 					'established the perimeter that err_msg = ""
	Dialog opening_dialog_01		'calls the initial dialog
	cancel_confirmation				'if cancel is pressed, this function gives the user the option to proceed or back out of the cancel request
	IF type_of_designated_account <> "None" AND isnumeric(counted_value_designated) = FALSE THEN err_msg = err_msg & vbNewLine & _
	"Designated Account Counted Value is not a number. Do not include letters or special characters."
	IF insurance_policy_number <> "none" AND isnumeric(insurance_counted_value) = FALSE THEN err_msg = err_msg & vbNewLine & _
	"Insurance Counted Value is not a number. Do not include letters or special characters."
	If programs = "Select one..." then err_msg = err_msg & vbNewLine & "* Select the program that you are evaluating this asset for."
	IF hh_member = "" then err_msg = err_msg & vbNewLine & "* Enter a HH member."

	If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
	If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
LOOP until ButtonPressed = open_dialog_next_button AND err_msg = ""

Do
	Do
		err_msg = ""
		Dialog burial_assets_dialog_01
		cancel_confirmation
		If type_of_burial_agreement = "Select One..." Then err_msg = err_msg & vbNewLine & "You must select a type of burial agreement. Select none if n/a."
		If type_of_burial_agreement <> "None" THEN
			If purchase_date = "" or IsDate(purchase_date) = FALSE then err_msg = err_msg & vbNewLine & " You must enter the purchase date."
			If issuer_name = "" then err_msg = err_msg & vbNewLine & "You must enter the issuer name."
			If policy_number = "" then err_msg = err_msg & vbNewLine & "You must enter the policy number."
			If face_value = "" or IsNumeric(face_value) = FALSE then err_msg = err_msg & vbNewLine & "You must enter the policy's face value."
		END IF
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = "" AND ButtonPressed = next_to_02_button
	DO														'if the type of burial agreement is NONE the script will skip the detail asset breakdown.
		IF type_of_burial_agreement <> "None" THEN CALL build_dynamic_burial_dialog(calc_array, BFE_total, BS_BSI_total, final_counted_total, final_unavailable_total)  'here we call the burial item dialogs, these are built dynamically with safe guarding built into function.
	LOOP until err_msg = ""
	DO
		actions_taken = InputBox("Actions Taken: ", "Actions taken")
	LOOP until actions_taken <> ""
LOOP until err_msg = ""

Call check_for_MAXIS(False) 'checking for an active MAXIS session

'SECTION 04: Finalizing totals-------------------------------------------------------------------------------------------------------
If counted_value_designated <> "" then final_counted_total = final_counted_total + cint(counted_value_designated)
If insurance_counted_value <> "" then final_counted_total = final_counted_total + cint(insurance_counted_value)

'SECTION 05: The CASE NOTE----------------------------------------------------------------------------------------------------
DIM MAXIS_service_row			'variables used for checking the headers to see if they need to be written on the top of a continuing case note.
DIM MAXIS_col

'first section of case note is dependant on what types of designated accounts were chosen.
start_a_blank_CASE_NOTE
CALL write_variable_in_case_note( "**BURIAL ASSETS -- Memb " & hh_member & " for " & programs & "**")
IF type_of_designated_account <> "None" then
	call write_variable_in_case_note("---Designated Account----")
	call write_bullet_and_variable_in_case_note("Type of designated account", type_of_designated_account)
	call write_bullet_and_variable_in_case_note("Account Identified", account_identifier)
	call write_bullet_and_variable_in_case_note("Reasons funds could not be separated", why_not_separated)
	call write_bullet_and_variable_in_case_note("Date account created", account_create_date)
	call write_bullet_and_variable_in_case_note("Counted Value", counted_value_designated)
	call write_bullet_and_variable_in_case_note("Info on BFE", BFE_information_designated)
END IF
IF insurance_policy_number <> "none" THEN
	call write_variable_in_case_note("---Non-Term Life Insurance----")
	call write_bullet_and_variable_in_case_note("Policy Number", insurance_policy_number)
	call write_bullet_and_variable_in_case_note("Insurance Company", insurance_company)
	call write_bullet_and_variable_in_case_note("Date policy created", insurance_create_date)
	call write_bullet_and_variable_in_case_note("CSV/FV designated to BFE", insurance_csv)
	call write_bullet_and_variable_in_case_note("Counted Value", insurance_counted_value)
	call write_bullet_and_variable_in_case_note("Info on BFE", insurance_BFE_steps_info)
END IF
IF type_of_burial_agreement <> "None" THEN						'if the type of burial agreement is NONE the script will skip the detail asset breakdown.
	If applied_BFE_check = 1 AND BFE_total = "" then CALL write_variable_in_case_note("* Applied $1500 of burial services to BFE.")
	CALL write_variable_in_case_note("* Type: " & type_of_burial_agreement & ". Purchase date: " & purchase_date & ".")
	CALL write_variable_in_case_note("* Issuer: " & issuer_name & ". Policy #: " & policy_number & ".")
	CALL write_bullet_and_variable_in_case_note("Face value", face_value)
	CALL write_bullet_and_variable_in_case_note("Funeral home", funeral_home)
	IF Primary_benficiary_check = 1 THEN Call write_variable_in_case_note ("* Primary beneficiary is: Any funeral provider whose interest may appear                irrevocably")
	IF Contingent_benficiary_check = 1 THEN Call write_variable_in_case_note ("* Contingent Beneficiary is: The estate of the insured")
	IF policy_CSV_check = 1 THEN Call write_variable_in_case_note ("* Policy's CSV is irrevocably designated to the funeral provider")

	'the following section will write the various sections of assets and then for each items in the arrays (in certain spots) it will write the value, item, and status of the asset

	CALL write_variable_in_case_note("--------------SERVICE--------------------AMOUNT----------STATUS------------")

	FOR i = 0 to 16
		case_note_page_four
		extra_spaces = ""
		If calc_array(i, 1) <> "" then
			new_service_heading
			call write_three_columns_in_case_note(3, (calc_array(i, 0) & ":"), 44, "$" & calc_array(i, 1), 54, calc_array(i, 2))
		End if
	NEXT

	FOR i = 41 to 42								'these two spots on the array have special handling to write only if they are services.
		case_note_page_four
		extra_spaces = ""
		If calc_array(i, 1) <> "" AND calc_array(i, 3) = "Services" THEN
			new_service_heading
			call write_three_columns_in_case_note(3, (calc_array(i, 0) & ":"), 44, "$" & calc_array(i, 1), 54, calc_array(i, 2))
		End if
	NEXT

	CALL write_variable_in_case_note("--------BURIAL SPACE/ITEMS---------------AMOUNT----------STATUS------------")

	FOR i = 17 to 28
		case_note_page_four
		extra_spaces = ""
		If calc_array(i, 1) <> "" then
			new_BS_BSI_heading
			call write_three_columns_in_case_note(3, (calc_array(i, 0) & ":"), 44, "$" & calc_array(i, 1), 54, calc_array(i, 2))
		End if
	NEXT

	FOR i = 41 to 42								'these two spots on the array have special handling to write only if they are BS/BSI.
		case_note_page_four
		extra_spaces = ""
		If calc_array(i, 1) <> "" AND calc_array(i, 3) = "BS/BSI" THEN
			new_BS_BSI_heading
			call write_three_columns_in_case_note(3, (calc_array(i, 0) & ":"), 44, "$" & calc_array(i, 1), 54, calc_array(i, 2))
		End if
	NEXT

	CALL write_variable_in_case_note("--------CASH ADVANCE ITEMS---------------AMOUNT----------STATUS------------")

	FOR i = 29 to 40
		case_note_page_four
		extra_spaces = ""
		If calc_array(i, 1) <> "" then
			new_CAI_heading
			call write_three_columns_in_case_note(3, (calc_array(i, 0) & ":"), 44, "$" & calc_array(i, 1), 54, calc_array(i, 2))
		End if
	NEXT

	FOR i = 41 to 42							'these two spots on the array have special handling to write only if they are CAI.
		case_note_page_four
		extra_spaces = ""
		If calc_array(i, 1) <> "" AND calc_array(i, 3) = "CAI" THEN
			new_CAI_heading
			call write_three_columns_in_case_note(3, (calc_array(i, 0) & ":"), 44, "$" & calc_array(i, 1), 54, calc_array(i, 2))
		End if
	NEXT

	CALL write_variable_in_case_note( "---------------------------------------------------------------------------")
	IF BFE_total <> "" THEN CALL write_variable_in_case_note( "* Total services/items applied to BFE: $" & BFE_total)
	CALL write_variable_in_case_note( "* Total BS/BSI excluded amount: $" & BS_BSI_total)
	CALL write_variable_in_case_note( "* Total unavailable CAI: $" & final_unavailable_total)
END IF


CALL write_variable_in_case_note( "---------------------------------------------------------------------------")
CALL write_variable_in_case_note( "* Total counted amount: $" & final_counted_total)
CALL write_variable_in_case_note( "* Actions taken: " & actions_taken)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

script_end_procedure("")
