'Required for statistical purposes===============================================================================
name_of_script = "BULK - ADD SSRT PANEL.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 60                      'manual run time in seconds
STATS_denomination = "C"       				'C is for each CASE
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
call changelog_update("01/08/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function HCRE_panel_bypass() 
	'handling for cases that do not have a completed HCRE panel
	PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
	Do
		EMReadscreen HCRE_panel_check, 4, 2, 50
		If HCRE_panel_check = "HCRE" then
			PF10	'exists edit mode in cases where HCRE isn't complete for a member
			PF3
		END IF
	Loop until HCRE_panel_check <> "HCRE"
End Function

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\GRH\GRH active cases 01-18.xlsx"

'dialog and dialog DO...Loop	
Do
	Do
			'The dialog is defined in the loop as it can change as buttons are pressed 
			BeginDialog info_dialog, 0, 0, 266, 105, "Add SSRT panel"
  				ButtonGroup ButtonPressed
    			PushButton 200, 40, 50, 15, "Browse...", select_a_file_button
    			OkButton 145, 85, 50, 15
    			CancelButton 200, 85, 50, 15
  				GroupBox 10, 5, 250, 75, "Using the Add SSRT panel script:"
  				Text 20, 20, 235, 20, "This script should be used when a list of cases on GRH need to have a SSRT panel added."
  				Text 15, 60, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  				EditBox 15, 40, 180, 15, file_selection_path
			EndDialog

			err_msg = ""
			Dialog info_dialog
			cancel_confirmation
			If ButtonPressed = select_a_file_button then
				If file_selection_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
					objExcel.Quit 'Closing the Excel file that was opened on the first push'
					objExcel = "" 	'Blanks out the previous file path'
				End If
				call file_selection_system_dialog(file_selection_path, ".xlsx") 'allows the user to select the file'
			End If
			If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data."
			If err_msg <> "" Then MsgBox err_msg
		Loop until err_msg = ""
		If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
		If err_msg <> "" Then MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
 Loop until are_we_passworded_out = false					'loops until user passwords back in

'NOW THE SCRIPT IS CHECKING STAT/FACI FOR EACH CASE.----------------------------------------------------------------------------------------------------
excel_row = 2 'Resetting the case row to investigate.

Do
	MAXIS_case_number= objExcel.cells(excel_row, 2).Value	're-establishing the case number to use for the case
    If trim(MAXIS_case_number) = "" then exit do
	
	'This Do...loop gets back to SELF
	back_to_self
	call navigate_to_MAXIS_screen("STAT", "PROG")
    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
	If PRIV_check = "PRIV" then
		case_status = "PRIV case, cannot access/update."
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the case number
		transmit
    Else 
		EMReadScreen grh_status, 4, 9, 74
		If grh_status <> "ACTV" then 
			case_status = "GRH case status is " & grh_status
			grh_active = false 
		Else 
			grh_active = true
		End if
	End if 
    
	If grh_active = true then 
		Call navigate_to_MAXIS_screen("STAT", "DISA")
		Call write_value_and_transmit ("01", 20, 76)	'For member 01 - All GRH cases should be for member 01. 
		EMReadScreen waiver_type, 1, 14, 59
		If waiver_type <> "_" then case_status = "Client is active on a waiver. Should not be Rate 2."
		
	 	Call HCRE_panel_bypass
	
		Call navigate_to_MAXIS_screen("STAT", "FACI")
		EMReadScreen member_number, 2, 4, 33
		If member_number <> "01" then 
			EmWriteScreen "01", 20, 76						'For member 01 - All GRH cases should be for member 01. 
			Call write_value_and_transmit ("01", 20, 79)	'1st version of FACI 
		End if 
	    'LOOKS FOR MULTIPLE STAT/FACI PANELS, GOES TO THE MOST RECENT ONE
	    EMReadScreen FACI_total_check, 1, 2, 78
		If FACI_total_check = "0" then 
			current_faci = False 
			case_status = "Case does not have a FACI panel."
		Else 
	        row = 14
	        Do 
	        	EMReadScreen date_out, 10, row, 71
		    	'msgbox "date out: " & date_out 
	        	If date_out = "__ __ ____" then 
		    	  	EMReadScreen grh_rate, 1, row, 34
		    	  	If grh_rate = "2" then 
		    	  		'msgbox grh_rate
		    	  		current_faci = TRUE
						EMReadScreen date_in, 10, row, 47
						EMReadScreen vendor_number, 8, 5, 43
	        	  		exit do
		    	  	ELSE
						current_faci = False 
						row = row + 1
					End if 
				Else 
	        		row = row + 1
					'msgbox row
					current_faci = False	
				End if 	
	        	If row = 19 then 
	        		transmit
	        		row = 14
	        	End if 
				EMReadScreen last_panel, 5, 24, 2
	        Loop until last_panel = "ENTER"	'This means that there are no other faci panels
		End if 
	    
	    If current_faci = False then case_status = "Unable to find FACI panel or current FACI is not Rate 2."
		'msgbox current_faci

	    'GETS FACI NAME AND PUTS IT IN SPREADSHEET, IF CLIENT IS IN FACI.
	    If current_faci = True then
	    	Call navigate_to_MAXIS_screen ("STAT", "SSRT")
			Call write_value_and_transmit ("01", 20, 76)	'For member 01 - All GRH cases should be for member 01. 
			EMReadScreen SSRT_total_check, 1, 2, 78
			
			If SSRT_total_check = "0" then
			 	Call write_value_and_transmit("NN", 20, 79)
				EMWriteScreen vendor_number, 5, 43		'Enters vendor number 
				transmit 'to make the vendor name and NPI appear
				EMReadScreen NPI_number, 10, 7, 43
				row = 10 
				Do 
					EMReadScreen open_row, 10, row, 47
					If open_row = "__ __ ____" then 
						Call create_MAXIS_friendly_date_with_YYYY(date_in, 0, row, 47)
						If date_out <> "__ __ ____" then Call create_MAXIS_friendly_date_with_YYYY(date_out, 0, row, 71)
						Added_SSRT = True 
						case_status = "Created SSRT panel."
						
						exit do 
					Else 
						row = row + 1
						Added_SSRT = False 
					End if 
				Loop until row = 15
				'msgbox "going to save the SSRT panel"
				PF3 'to save
			else 
				Do 
					EMReadScreen SSRT_current_panel, 1, 2, 73
					EMReadScreen vendor_check, 8, 5, 43
					If vendor_check = vendor_number then 
						case_status = "Case already has SSRT panel."
						EMReadScreen NPI_number, 10, 7, 43
						exit do 
					Else 
						transmit
						'msgbox "needs new SSRT panel."
						case_status = "SSRT does not exist. Please review case."
					End if 
					EMReadScreen last_panel, 5, 24, 2
		        Loop until last_panel = "ENTER"	'This means that there are no other faci panels
				
			End if
	    End if
 	End if 
	
	ObjExcel.Cells(excel_row, 12).Value = trim(NPI_number)
 	ObjExcel.Cells(excel_row, 13).Value = case_status

	excel_row = excel_row + 1 'setting up the script to check the next row.
LOOP UNTIL objExcel.Cells(excel_row, 2).Value = ""	'Loops until there are no more cases in the Excel list

'formatting the cells
FOR i = 1 to 13
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your list has been created.")