Enter fi'Script Created by Casey Love from Ramsey County 

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - CASE TRANSFER.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'Required for statistical purposes==========================================================================================
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 20                	'manual run time in seconds
STATS_denomination = "I"       			'I is for each Item
'END OF stats block=========================================================================================================

'DIALOGS----------------------------------------------------------------------
BeginDialog select_parameters_data_into_excel, 0, 0, 376, 390, "Select Parameters for Cases to Transfer"
  EditBox 75, 20, 130, 15, worker_number
  CheckBox 5, 60, 240, 10, "Check here to have the script query all the cases listed on REPT/ACTV.", query_all_check
  CheckBox 5, 110, 170, 10, "Exclude all cases with any Pending Program", exclude_pending_check
  CheckBox 5, 125, 120, 10, "Exclude all cases with IEVS DAILs", exclude_ievs_check
  CheckBox 140, 125, 120, 10, "Exclude all cases with PARIS DAILs", exclude_paris_check
  CheckBox 5, 140, 40, 10, "SNAP", SNAP_check
  CheckBox 90, 140, 90, 10, "Exclude all SNAP cases", exclude_snap_check
  CheckBox 190, 140, 75, 10, "SNAP ONLY cases", SNAP_Only_check
  CheckBox 15, 165, 60, 10, "ABAWD cases", SNAP_ABAWD_check
  CheckBox 90, 165, 90, 10, "Uncle Harry SNAP", SNAP_UH_check
  CheckBox 5, 185, 25, 10, "GA", ga_check
  CheckBox 40, 185, 30, 10, "MSA", msa_check
  CheckBox 90, 185, 100, 10, "Exclude all GA/MSA cases", exclude_ga_msa_check
  CheckBox 5, 200, 25, 10, "RCA", rca_check
  CheckBox 90, 200, 90, 10, "Exclude all RCA cases", exclude_RCA_check
  CheckBox 5, 215, 30, 10, "MFIP", mfip_check
  CheckBox 40, 215, 30, 10, "DWP", DWP_check
  CheckBox 90, 215, 95, 10, "Exclude all MFIP/DWP", exclude_mfip_dwp_check
  CheckBox 190, 215, 70, 10, "MFIP ONLY cases", MFIP_Only_check
  CheckBox 15, 245, 90, 10, "MFIP cases with at least", MFIP_tanf_check
  EditBox 105, 240, 20, 15, tanf_months
  CheckBox 15, 260, 85, 10, "Child Only MFIP cases", child_only_mfip_check
  CheckBox 105, 260, 90, 10, "Only monthly reporters", mont_rept_check
  CheckBox 5, 280, 50, 10, "Child Care", ccap_check
  CheckBox 90, 280, 95, 10, "Exclude Child Care cases", exclude_ccap_check
  CheckBox 5, 295, 40, 10, "GRH", GRH_check
  CheckBox 90, 295, 75, 10, "Exclude GRH cases", exclude_grh_check
  CheckBox 190, 295, 75, 10, "GRH ONLY cases", GRH_Only_check
  CheckBox 5, 310, 65, 10, "EA/EGA Pending", EA_check
  CheckBox 90, 310, 95, 10, "Exclude EA/EGA Pending", exclude_ea_check
  CheckBox 5, 325, 40, 10, "HC", HC_check
  CheckBox 90, 325, 75, 10, "Exclude HC cases", exclude_HC_check
  CheckBox 190, 325, 75, 10, "HC ONLY cases", HC_Only_check
  CheckBox 15, 345, 120, 10, "Medicare Savings Program Active", HC_msp_check
  CheckBox 15, 360, 40, 10, "Adult MA", adult_hc_check
  CheckBox 90, 360, 45, 10, "Family MA", family_hc_check
  CheckBox 15, 375, 40, 10, "LTC HC", ltc_HC_check
  CheckBox 90, 375, 50, 10, "Waiver HC", waiver_HC_check
  ButtonGroup ButtonPressed
    OkButton 270, 370, 50, 15
    CancelButton 325, 370, 50, 15
  GroupBox 10, 150, 190, 30, "SNAP Details"
  GroupBox 10, 230, 190, 45, "MFIP Details"
  GroupBox 10, 335, 190, 55, "HC Details"
  Text 215, 10, 155, 40, "Select the criteria you want the script to sort by. The script will then generate an Excel Spreadsheet of all the cases in the indicated worker number(s) that meet your selected criteria."
  Text 260, 55, 100, 35, "Another Pop Up will allow you select your transfer options before actually transferring cases."
  Text 130, 245, 65, 10, "TANF Months used."
  GroupBox 275, 105, 95, 255, "Hints"
  Text 280, 120, 85, 25, "Use 'Tab' to move between check boxes without your mouse."
  Text 280, 150, 85, 25, "Use the Spacebar to check and uncheck boxes without your mouse"
  Text 5, 25, 65, 10, "Worker(s) to check:"
  Text 5, 75, 235, 10, " This will not give a transfer option but will look for all possible criteria."
  Text 5, 40, 210, 20, "Enter last 3 digits of your workers' x1 numbers (ex: x100###), separated by a comma."
  Text 5, 95, 235, 10, "Check all that apply - What type of cases do you want to transfer?"
  Text 65, 5, 100, 10, "***Case Parameters to Pull***"
EndDialog

'THE SCRIPT-------------------------------------------------------------------------

'Determining specific county for multicounty agencies...
CALL worker_county_code_determination(worker_county_code, two_digit_county_code)

'Connects to BlueZone
EMConnect ""

'Shows dialogDialog pull_rept_data_into_Excel_dialog
Do	
	Do 
		Dialog select_parameters_data_into_excel
		cancel_confirmation
		err_msg = ""
		IF worker_number = "" then err_msg = err_msg & vbCr & "You must Select an X-Number to pull cases from."
		IF query_all_check = unchecked AND snap_check = unchecked AND mfip_check = unchecked AND DWP_check = unchecked AND EA_check = unchecked AND HC_check = unchecked AND ga_check = unchecked AND msa_check = unchecked AND GRH_check = unchecked Then err_msg = err_msg & _ 
		  vbCr & "You must select a type of program for the cases you want to transfer, please pick one - SNAP, MFIP, DWP, EA, HC, GA, MSA, or GRH"
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	Loop until query_all_check = checked OR snap_check = checked OR mfip_check = checked OR DWP_check = checked OR EA_check = checked OR HC_check = checked OR ga_check = checked OR msa_check = checked OR GRH_check = checked AND worker_number <> ""
	err_msg = ""
	If SNAP_check = unchecked then 
		IF SNAP_Only_check = checked OR SNAP_ABAWD_check = checked OR SNAP_UH_check = checked then err_msg = err_msg & vbCr & "If you select SNAP details, you must filter FOR SNAP cases - Check the SNAP box"
	End If 
	IF mfip_check = unchecked then 
		IF MFIP_Only_check = checked OR MFIP_tanf_check = checked OR child_only_mfip_check = checked OR mont_rept_check = checked then err_msg = err_msg & vbCr & " If you select MFIP details, you must filter FOR MFIP cases - check the MFIP box"
	End If 
	If MFIP_tanf_check = checked AND tanf_months = "" then err_msg = err_msg & vbCr & "If you want to filter for a certain number of TANF months, you must indicate how many months you want"
	IF HC_check = unchecked then 
		If HC_msp_check = checked OR adult_hc_check = checked OR family_hc_check = checked OR ltc_HC_check = checked OR waiver_HC_check = checked then err_msg = err_msg & vbCr & "If you select HC details, you must filter FOR HC cases - check the HC Box"
	End If 
	IF snap_check = checked AND exclude_snap_check = checked then err_msg = err_msg & vbCr & "You cannot filter for SNAP and Exclude SNAP cases - please pick one"
	IF mfip_check = checked AND exclude_mfip_dwp_check = checked then err_msg = err_msg & vbCr & "You cannot filter for MFIP and Exclude MFIP cases - please pick one"
	IF ccap_check = checked AND exclude_ccap_check = checked then err_msg = err_msg & vbCr & "You cannot filter for Child Care and Exclude Child Care - please pick one"
	IF EA_check = checked AND exclude_ea_check = checked then err_msg = err_msg & vbCr & "You cannot filter for EA/EGA and Exclude EA/EGA cases - please pick one"
	IF HC_check = checked AND exclude_HC_check = checked then err_msg = err_msg & vbCr & "You cannot filter for HC and Exclude HC cases - please pick one"
	If exclude_ga_msa_check = checked then 
		IF ga_check = checked OR msa_check = checked then err_msg = err_msg & vbCr & "You cannot filter for GA and/or MSA and Exclude GA/MSA cases - please pick one"
	End If 
	If GRH_check = checked AND exclude_grh_check = checked then err_msg = err_msg & vbCr & "You cannot filter for GRH and Exclude GRH cases - please pick one"
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
Loop until err_msg = ""

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_MAXIS(True)

'In order to make the code a little cleaner, this sets all the option check boxes to checked when the Query All option exists.
IF query_all_check = checked THEN 
	SNAP_ABAWD_check = checked 
	SNAP_UH_check = checked 
	MFIP_tanf_check = checked 
	child_only_mfip_check = checked
	mont_rept_check = checked 
	HC_msp_check = checked 
	adult_hc_check = checked 
	family_hc_check = checked 
	ltc_HC_check = checked 
	waiver_HC_check = checked 
	exclude_ievs_check = checked 
	exclude_paris_check = checked
	ccap_check = checked 
End IF

MsgBox "The script will now create an Excel Spreadsheet to display case information" & _ 
   vbCr & "Please do not alter this spreadsheet until after the script has completed." & _ 
   vbCR & "Altering the spreadsheet will not change how the script runs and will only cause the data listed on it to be incorrect."
   
le contents here
