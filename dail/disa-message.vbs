'Required for statistical purposes===============================================================================
name_of_script = "DAIL - DISA MESSAGE.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 64          'manual run time in seconds
STATS_denomination = "C"       'C is for Case
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
call changelog_update("11/01/2026", "Updated functionality - displays original DAIL PEPR/DISA message as well as the info from STAT/DISA. Allows HSR to then select whether to TIKL, create CASE/NOTE, or other. Significant logic added for error checking.", "Sam Begley-May, Hennepin County")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note. Updated background navigation coding.", "Ilse Ferris")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'====FUNCTIONS AND CONSTANTS BLOCK



' will need reordered based on usage order

Dim row, col, Original_DAIL, MEMB_name_to_save, MembNo_value, count
Dim DISA_bgn_date, DISA_end_date
Dim Cert_period_bgn, Cert_period_end
Dim Eld_wavr_bgn, Eld_wavr_end, GRH_bgn, GRH_end
Dim cash_disa_status, cash_disa_ver, fs_disa_status, fs_disa_ver, hc_disa_status, hc_disa_ver
Dim MEMB_No, last_first_name, person_info
Dim form_to_send, verification_process, form_needed, short_verif_info
Dim GLOBAL_CASH_DISA_STATUS, GLOBAL_CASH_DISA_VER, GLOBAL_FS_DISA_STATUS, GLOBAL_FS_DISA_VER, GLOBAL_HC_DISA_STATUS, GLOBAL_HC_DISA_VER
Dim origArray, newArray
Dim orig_cash_status, orig_cash_ver, orig_fs_status, orig_fs_ver, orig_hc_status, orig_hc_ver






' assigns the message taken from dail-scrubber.vbs to the Original_DAIL variable. This will allow functionality for both "DISA is ending" AND "DISA has ended"
Original_DAIL = DISA_check


' Building "GLOBAL" arrays of all the potential options from dropdown list boxes
GLOBAL_CASH_DISA_STATUS = Array (_
    "01 - RSDI Only Disability", _
    "02 - RSDI Only Blindness", _
    "03 - SSI, SSI/RSDI Disability", _
    "04 - SSI, SSI/RSDI Blindness", _
    "06 - SMRT/SSA Pend", _
    "08 - SMRT Certified Blindness", _
    "09 - Ill/Incapacity", _
    "10 - SMRT Certified Disability", _
    "BLANK" _
)

GLOBAL_CASH_DISA_VER = Array( _
    "1 - DHS161/Dr Statement", _
    "2 - SMRT Certified", _
    "3 - Certified for RSDI or SSI", _
    "6 - Other Document", _
    "7 - Professional Stmt of Need", _
    "N - No Verif Provided", _
    "BLANK" _
)

GLOBAL_FS_DISA_STATUS = Array( _
    "01 - RSDI Only Disability", _
    "02 - RSDI Only Blindness", _
    "03 - SSI, SSI/RSDI Disability", _
    "04 - SSI, SSI/RSDI Blindness", _
    "06 - SMRT/SSA Pend", _
    "08 - SMRT Certified Blindness", _
    "09 - Ill/Incapacity", _
    "10 - SMRT Certified Disability", _
    "11 - VA Determined Pd 100% Disa", _
    "12 - VA (Other Accept Disa)", _
    "13 - Cert RR Ret Disa & On Medi", _
    "14 - Othr Govt Perm Disa Ret Bnft", _
    "15 - Disability from MINE List", _
    "16 - Unable Prepare Purch Own Meal", _
    "BLANK" _
)

GLOBAL_FS_DISA_VER = Array( _
    "1 - DHS161/Dr Statement", _
    "2 - SMRT Certified", _
    "3 - Certified for RSDI or SSI", _
    "4 - Receipt of HC For Disa/Blind", _
    "5 - Wrk Judgement", _
    "6 - Other Document", _
    "7 - Out of State Ver Pending", _
    "N - No Verif Provided", _
    "BLANK" _
)

GLOBAL_HC_DISA_STATUS = Array( _
    "01 - RSDI Only Disability", _
    "02 - RSDI Only Blindness", _
    "03 - SSI, SSI/RSDI Disability", _
    "04 - SSI, SSI/RSDI Blindness", _
    "06 - SMRT/SSA Pend", _
    "08 - Certified Blindness", _
    "10 - Certified Disabled", _
    "11 - Spec Cat - Disa Child", _
    "20 - TEFRA - Disabled", _
    "21 - TEFRA - Blind", _
    "22 - MA-EPD", _
    "23 - MA/Waiver", _
    "24 - SSA/SMRT Appeal Pending", _
    "26 - SSA/SMRT Disa Deny", _
    "BLANK" _
)

GLOBAL_HC_DISA_VER = Array( _
    "1 - DHS161/Dr Statement", _
    "2 - SMRT Certified", _
    "3 - Certified for RSDI or SSI", _
    "6 - Other Document", _
    "7 - Case Manager Determination", _
    "8 - LTC Consult Services", _
    "N - No Verif Provided", _
    "BLANK" _
)

GLOBAL_CASH_VERIF_PROCESSES = Array( _
        "MOF send through ECF", _
        "SMI checked and verified", _
        "HC Active - SMRT submitted through ISDS/ECF", _
        "Other documentation (please note)", _
        "Other client contact (please note)"_
)

GLOBAL_FS_VERIF_PROCESSES = Array( _
        "MOF send through ECF", _
        "HC Active - SMRT submitted through ISDS/ECF", _
        "SMI checked and verified", _
        "Other documentation (please note)", _
        "Other client contact (please note)"_
)

GLOBAL_HC_VERIF_PROCESSES = Array( _
        "MOF send through ECF", _
        "SMRT submitted through ISDS/ECF", _
        "SMI checked and verified", _
        "Other documentation (please note)", _
        "Other client contact (please note)"_
)

' This returns to the DAIL/DAIL screen for deletion. Will need to be put in the FuncLib eventually as all other scripts called from the DAIL scrubber should have this functionality
' leave the parameters (MEMB_name_to_save, Original_DAIL) as written and they will use the variables already in use for CASE/NOTEing and TIKLing
Function search_for_correct_DAIL_and_delete(MEMB_name_to_save, Original_DAIL)
    'Check on the right screen
    EMReadScreen WORKERS_check, 7, 2, 26
    ' The (DAIL) is located in the same place on both the DAIL/"" submenu and the DAIL/DAIL page. The word "WORKERS" is in a different location on both pages. This will keep the DAIL script from getting stuck on the submenu
    If WORKERS_check <> "WORKERS" then script_end_procedure("You are not in DAIL for the final deletion. The script will now end")


    'Setting the row/col markers to start on 5,4 coordinates to avoid the non-DAIL information triggering accidental matches
    row = 5
    col = 4
    EMSearch MEMB_name_to_save, row, col
    'If MEMB_name_to_save cannot be found, loop through DAIL/DAIL with PF8. If the last_page_check triggers, end the script; otherwise, look for the name again and keep looping until the script finds it.
    If row = 0 and col = 0 then
        Do
            PF8
            EMReadScreen last_page_check, 21, 24, 2
                If last_page_check = "THIS IS THE LAST PAGE" then script_end_procedure("The original DAIL could not be found for deletion. The script will now end.")
            EMSearch MEMB_name_to_save, row, col
        Loop until row <> 0 and col <> 0
    End if

    'On the page the last_first_name was found, look for the DAIL message. 
    EMSearch Original_DAIL, row, col 
    'Failsafe for if the person is on one page and their DAIL is on the next. Their name carries over to the next page - so the script will check for the name and then the message and keep looping as long as the name appears
      If row = 0 and col = 0 then
        Do
          PF8
          EMSearch MEMB_name_to_save, row, col
            If row = 0 and col = 0 then script_end_procedure("The original DAIL could not be found associated with that client. The script will now end.")
          EMSearch Original_DAIL, row, col
        Loop until row <> 0 and col <> 0
      End If
    
'Final deletion logic - if the row and col parameters are not 0, that means that the Original_DAIL has been found, which means "D" can delete the DAIL using the row coordinate. Otherwise, end the script
    If row <> 0 and col <> 0 then
      EMWriteScreen "D", row, 3
      transmit
      MsgBox "The original DAIL was found and deleted"
    Else
        script_end_procedure("The original DAIL could not be found for deletion. The script will now end.")
    End If
End Function


'Build the dropdown list for the dialog box. Using chr(9) with JOIN was the only way I found to split the string correctly. 
Function BuildDropDownList(currentValue, allOptions)
    Dim tempArray()
    Dim i

    ' Start array with the current value
    ReDim tempArray(0)
    tempArray(0) = currentValue

    ' Add the remaining options, avoiding duplicates
    For i = 0 To UBound(allOptions)
        If allOptions(i) <> currentValue Then
            ReDim Preserve tempArray(UBound(tempArray) + 1)
            tempArray(UBound(tempArray)) = allOptions(i)
        End If
    Next

    ' Join into a DropListBox‑compatible string
    BuildDropDownList = Join(tempArray, chr(9))
End Function

'Rebuild the dropdown lists while also creating an array of the changes - this allows the changes to appear on the top of the dropdown box and trigger any error checking logic
Function BuildChangedList(origValues, newValues)
    Dim changed(), changedCount, i
    ReDim changed(0)
    changedCount = 0

    For i = 0 To UBound(origValues)
        If origValues(i) <> newValues(i) Then
            If changedCount > 0 Then ReDim Preserve changed(changedCount)
            changed(changedCount) = i
            changedCount = changedCount + 1
        End If
    Next

    If changedCount = 0 Then changed(0) = -1
    BuildChangedList = changed
End Function

'Function to read the disa panel before displaying information in dialog box. Because this calls directly from the DAIL, the correct person's panel should be displayed already. The screen check is only there in case of errors or misnavigation from user input
Function Read_DISA()
    'double-check that user hasn't accidentally navigated away from the DISA panel
    EMReadScreen DISA_chk, 4, 2, 45
    If DISA_chk <> "DISA" then script_end_procedure("Not on a DISA message, the script will now end.") 


    'Read the member number; trim for spaces; and add the matched name variable to create the person_info variable that's used for CASE/NOTEs and dialog boxes throughout
    EMReadScreen MEMB_No, 2, 4, 33
    MEMB_No = trim(MEMB_No)
    person_info = MEMB_No & " " & MEMB_name_to_save


    'Read and format the disability begin date
    EMReadScreen DISA_bgn_date_month, 2, 6, 47
    EMReadScreen DISA_bgn_date_day, 2, 6, 50
    EMReadScreen DISA_bgn_date_year, 4, 6, 53
    If DISA_bgn_date_month = "__" OR DISA_bgn_date_day = "__" OR DISA_bgn_date_year = "____" Then
        DISA_bgn_date = "BLANK"
    Else
        DISA_bgn_date = DISA_bgn_date_month & "/" & DISA_bgn_date_day & "/" & DISA_bgn_date_year
    End If

    'Read and format the disability end date
    EMReadScreen DISA_end_date_month, 2, 6, 69
    EMReadScreen DISA_end_date_day, 2, 6, 72
    EMReadScreen DISA_end_date_year, 4, 6, 75
    If DISA_end_date_month = "__" OR DISA_end_date_day = "__" OR DISA_end_date_year = "____" Then
        DISA_end_date = "BLANK"
    Else
        DISA_end_date = DISA_end_date_month & "/" & DISA_end_date_day & "/" & DISA_end_date_year
    End If

    ' Read and format the certification period beginning date
    EMReadScreen Cert_period_bgn_month, 2, 7, 47
    EMReadScreen Cert_period_bgn_day, 2, 7, 50
    EMReadScreen Cert_period_bgn_year, 4, 7, 53
    If Cert_period_bgn_month = "__" OR Cert_period_bgn_day = "__" OR Cert_period_bgn_year = "____" Then
        Cert_period_bgn = "BLANK"
    Else
        Cert_period_bgn = Cert_period_bgn_month & "/" & Cert_period_bgn_day & "/" & Cert_period_bgn_year
    End If

    ' Read and format the certification period end date
    EMReadScreen Cert_period_end_month, 2, 7, 69
    EMReadScreen Cert_period_end_day, 2, 7, 72
    EMReadScreen Cert_period_end_year, 4, 7, 75
    If Cert_period_end_month = "__" Or Cert_period_end_day = "__" Or Cert_period_end_year = "____" Then
        Cert_period_end = "BLANK"
    Else
        Cert_period_end = Cert_period_end_month & "/" & Cert_period_end_day & "/" & Cert_period_end_year
    End If

    ' Read and format the elder waiver begin date
    EMReadScreen Eld_wavr_bgn_month, 2, 8, 47
    EMReadScreen Eld_wavr_bgn_day, 2, 8, 50
    EMReadScreen Eld_wavr_bgn_year, 4, 8, 53
    If Eld_wavr_bgn_month = "__" OR Eld_wavr_bgn_day = "__" OR Eld_wavr_bgn_year = "____" Then
        Eld_wavr_bgn = "BLANK"
    Else
        Eld_wavr_bgn = Eld_wavr_bgn_month & "/" & Eld_wavr_bgn_day & "/" & Eld_wavr_bgn_year
    End If

    ' Read and format the elder waiver end month
    EMReadScreen Eld_wavr_end_month, 2, 8, 69
    EMReadScreen Eld_wavr_end_day, 2, 8, 72
    EMReadScreen Eld_wavr_end_year, 4, 8, 75
    If Eld_wavr_end_month = "__" Or Eld_wavr_end_day = "__" Or Eld_wavr_end_year = "____" Then
        Eld_wavr_end = "BLANK"
    Else
        Eld_wavr_end = Eld_wavr_end_month & "/" & Eld_wavr_end_day & "/" & Eld_wavr_end_year
    End If


    ' Read and format the GRH related end date (MSA/County)
    EMReadScreen GRH_bgn_month, 2, 9, 47
    EMReadScreen GRH_bgn_day, 2, 9, 50
    EMReadScreen GRH_bgn_year, 4, 9, 53
    If GRH_bgn_month = "__" OR GRH_bgn_day = "__" Or GRH_bgn_year = "____" Then
        GRH_bgn = "BLANK"
    Else
        GRH_bgn = GRH_bgn_month & "/" & GRH_bgn_day & "/" & GRH_bgn_year
    End IF

    ' Read and format the GRH related end date (MSA/County)
    EMReadScreen GRH_end_month, 2, 9, 69
    EMReadScreen GRH_end_day, 2, 9, 72
    EMReadScreen GRH_end_year, 4, 9, 75
    If GRH_end_month = "__" OR GRH_end_day = "__" OR GRH_end_year = "____" Then
        GRH_end = "BLANK"
    Else
        GRH_end = GRH_end_month & "/" & GRH_end_day & "/" & GRH_end_year
    End if
    
' Disability statuses and verifications
    EMReadScreen cash_disa_status, 2, 11, 59
        If cash_disa_status = "01" then cash_disa_status = "01 - RSDI Only Disability"
        If cash_disa_status = "02" then cash_disa_status = "02 - RSDI Only Blindness"
        If cash_disa_status = "03" then cash_disa_status = "03 - SSI, SSI/RSDI Disability"
        If cash_disa_status = "04" then cash_disa_status = "04 - SSI, SSI/RSDI Blindness"
        If cash_disa_status = "06" then cash_disa_status = "06 - SMRT/SSA Pend"
        If cash_disa_status = "08" then cash_disa_status = "08 - SMRT Certified Blindness"
        If cash_disa_status = "09" then cash_disa_status = "09 - Ill/Incapacity"
        If cash_disa_status = "10" then cash_disa_status = "10 - SMRT Certified Disability"
        If cash_disa_status = "__" then cash_disa_status = "BLANK"

    EMReadScreen cash_disa_ver, 1, 11, 69
        If cash_disa_ver = "1" then cash_disa_ver = "1 - DHS161/Dr Statement" 
        If cash_disa_ver = "2" then cash_disa_ver = "2 - SMRT Certified"  
        If cash_disa_ver = "3" then cash_disa_ver = "3 - Certified for RSDI or SSI"
        If cash_disa_ver = "6" then cash_disa_ver = "6 - Other Document"
        If cash_disa_ver = "7" then cash_disa_ver = "7 - Professional Stmt of Need"
        If cash_disa_ver = "N" then cash_disa_ver = "N - No Verif Provided"
        If cash_disa_ver = "_" then cash_disa_ver = "BLANK"

    EMReadScreen fs_disa_status, 2, 12, 59
        If fs_disa_status = "01" then fs_disa_status = "01 - RSDI Only Disability"
        If fs_disa_status = "02" then fs_disa_status = "02 - RSDI Only Blindness"
        If fs_disa_status = "03" then fs_disa_status = "03 - SSI, SSI/RSDI Disability"
        If fs_disa_status = "04" then fs_disa_status = "04 - SSI, SSI/RSDI Blindness"
        If fs_disa_status = "06" then fs_disa_status = "06 - SMRT/SSA Pend"
        If fs_disa_status = "08" then fs_disa_status = "08 - SMRT Certified Blindness"
        If fs_disa_status = "09" then fs_disa_status = "09 - Ill/Incapacity"
        If fs_disa_status = "10" then fs_disa_status = "10 - SMRT Certified Disability"
        If fs_disa_status = "11" then fs_disa_status = "11 - VA Determined Pd 100% Disa"
        If fs_disa_status = "12" then fs_disa_status = "12 - VA (Other Accept Disa)"
        If fs_disa_status = "13" then fs_disa_status = "13 - Cert RR Ret Disa & On Medi"
        If fs_disa_status = "14" then fs_disa_status = "14 - Othr Govt Perm Disa Ret Bnft"
        If fs_disa_status = "15" then fs_disa_status = "15 - Disability from MINE List"
        If fs_disa_status = "16" then fs_disa_status = "16 - Unable Prepare Purch Own Meal"
        if fs_disa_status = "__" then fs_disa_status = "BLANK"

    EMReadScreen fs_disa_ver, 1, 12, 69
        If fs_disa_ver = "1" then fs_disa_ver = "1 - DHS161/Dr Statement"
        If fs_disa_ver = "2" then fs_disa_ver = "2 - SMRT Certified"
        If fs_disa_ver = "3" then fs_disa_ver = "3 - Certified for RSDI or SSI"
        If fs_disa_ver = "4" then fs_disa_ver = "4 - Receipt of HC For Disa/Blind"
        If fs_disa_ver = "5" then fs_disa_ver = "5 - Wrk Judgement"
        If fs_disa_ver = "6" then fs_disa_ver = "6 - Other Document"
        If fs_disa_ver = "7" then fs_disa_ver = "7 - Out of State Ver Pending"
        If fs_disa_ver = "N" then fs_disa_ver = "N - No Verif Provided"
        if fs_disa_ver = "_" then fs_disa_ver = "BLANK"


    EMReadScreen hc_disa_status, 2, 13, 59
        If hc_disa_status = "01" then hc_disa_status = "01 - RSDI Only Disability"
        If hc_disa_status = "02" then hc_disa_status = "02 - RSDI Only Blindness"
        If hc_disa_status = "03" then hc_disa_status = "03 - SSI, SSI/RSDI Disability"
        If hc_disa_status = "04" then hc_disa_status = "04 - SSI, SSI/RSDI Blindness"
        If hc_disa_status = "06" then hc_disa_status = "06 - SMRT/SSA Pend"
        If hc_disa_status = "08" then hc_disa_status = "08 - Certified Blindness"
        If hc_disa_status = "10" then hc_disa_status = "10 - Certified Disabled"
        If hc_disa_status = "11" then hc_disa_status = "11 - Spec Cat - Disa Child"
        If hc_disa_status = "20" then hc_disa_status = "20 - TEFRA - Disabled"
        If hc_disa_status = "21" then hc_disa_status = "21 - TEFRA - Blind"
        If hc_disa_status = "22" then hc_disa_status = "22 - MA-EPD"
        If hc_disa_status = "23" then hc_disa_status = "23 - MA/Waiver"
        If hc_disa_status = "24" then hc_disa_status = "24 - SSA/SMRT Appeal Pending"
        If hc_disa_status = "26" then hc_disa_status = "26 - SSA/SMRT Disa Deny"
        If hc_disa_status = "__" then hc_disa_status = "BLANK"

    EMReadScreen hc_disa_ver, 1, 13, 69
        If hc_disa_ver = "1" then hc_disa_ver = "1 - DHS161/Dr Statement"
        If hc_disa_ver = "2" then hc_disa_ver = "2 - SMRT Certified"
        If hc_disa_ver = "3" then hc_disa_ver = "3 - Certified for RSDI or SSI"
        If hc_disa_ver = "6" then hc_disa_ver = "6 - Other Document"
        If hc_disa_ver = "7" then hc_disa_ver = "7 - Case Manager Determination"
        If hc_disa_ver = "8" then hc_disa_ver = "8 - LTC Consult Services"
        If hc_disa_ver = "N" then hc_disa_ver = "N - No Verif Provided"
        If hc_disa_ver = "_" then hc_disa_ver = "BLANK"


' Set the "original" variables to the variables used for displaying selections from the dropdown List. This allows for the creation of the changed array later; and enables correct displays on dropdowns.    
orig_cash_status = cash_disa_status
orig_cash_ver    = cash_disa_ver
orig_fs_status   = fs_disa_status
orig_fs_ver      = fs_disa_ver
orig_hc_status   = hc_disa_status
orig_hc_ver      = hc_disa_ver  

'create arrays of all statuses for dropdown boxes in dialogs later
Dim results
    results = Array( _
        cash_disa_status, _
        cash_disa_ver, _
        fs_disa_status, _
        fs_disa_ver, _
        hc_disa_status, _
        hc_disa_ver, _
        GLOBAL_CASH_DISA_STATUS, _
        GLOBAL_CASH_DISA_VER, _
        GLOBAL_FS_DISA_STATUS, _
        GLOBAL_FS_DISA_VER, _
        GLOBAL_HC_DISA_STATUS, _
        GLOBAL_HC_DISA_VER _
    )


' this could be cleaner....
    If CASH_disa_option_status <> "BLANK" and CASH_disa_ver_type <> "BLANK" then
        CASH_active = true
    Else
        Cash_active = false
    End If

    If FS_disa_option_status <> "BLANK" and FS_disa_ver_type <> "BLANK" then
        FS_active = true
    Else
        FS_active = false
    End If

    If HC_disa_option_status <> "BLANK" and HC_disa_ver_type <> "BLANK" then
        HC_active = true
    Else
        HC_active = false
    End If


Read_DISA = results

End Function



Function Update_DISA_Panel_Loop()

    Dim changedIndexes, idx, i

    origArray = Array(orig_cash_status, orig_fs_status, orig_hc_status, orig_cash_ver, orig_fs_ver, orig_hc_ver)
    newArray = Array(CASH_disa_option_status, FS_disa_option_status, HC_disa_option_status, CASH_disa_ver_type, FS_disa_ver_type, HC_disa_ver_type)

    Do
    EMReadScreen DISA_chk, 4, 2, 45
    If DISA_chk <> "DISA" Then script_end_procedure("Not on a DISA message, the script will now end.")
    call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = false


    changedIndexes = BuildChangedList(origArray, newArray)

If UBound(changedIndexes) = -1 Then
    MsgBox "No changes were made..."
    Exit Function
End If


    PF9

    For i = 0 To UBound(changedIndexes)
        idx = changedIndexes(i)

        Select Case idx
            Case 0
                If newArray(idx) = "BLANK" Then
                    EMWriteScreen "  ", 11, 59
                Else
                    EMWriteScreen Left(newArray(idx), 2), 11, 59
                End If
            Case 1
                If newArray(idx) = "BLANK" then
                    EMWriteScreen "  ", 12, 59
                Else
                    EMWriteScreen Left(newArray(idx), 2), 12, 59
                End If
            Case 2
                If newArray(idx) = "BLANK" then
                    EMWriteScreen "  ", 13, 59
                Else
                    EMWriteScreen Left(newArray(idx), 2), 13, 59
                End If
            Case 3
                If newArray(idx) = "BLANK" then
                    EMWriteScreen " ", 11, 69
                Else
                    EMWriteScreen Left(newArray(idx), 1), 11, 69
                End If
            Case 4
                If newArray(idx) = "BLANK" then
                    EMWriteScreen " ", 12, 69
                Else
                    EMWriteScreen Left(newArray(idx), 1), 12, 69
                End If
            Case 5
                If newArray(idx) = "BLANK" then
                    EMWriteScreen " ", 13, 69
                Else
                    EMWriteScreen Left(newArray(idx), 1), 13, 69
                End If
        End Select
    Next
    
' Need additional handling in here for all blanks apeparing


    transmit

    orig_cash_status = newArray(0)
    orig_fs_status = newArray(1)
    orig_hc_status = newArray(2)
    orig_cash_ver = newArray(3)
    orig_fs_ver = newArray(4)
    orig_hc_ver = newArray(5)



End Function


' Function that validates the correct 
Function DISA_Fields_Are_Valid(CASH_disa_option_status, CASH_disa_ver_type, FS_disa_option_status, FS_disa_ver_type, HC_disa_option_status, HC_disa_ver_type)
    
    DISA_Fields_Are_Valid = False

    If CASH_disa_option_status = "BLANK" And CASH_disa_ver_type = "BLANK" And _
       FS_disa_option_status = "BLANK" And FS_disa_ver_type = "BLANK" And _
       HC_disa_option_status = "BLANK" And HC_disa_ver_type = "BLANK" Then
        MsgBox "All disability status and verification fields are blank. Please correct the information before proceeding or end the script and process manually."
        Exit Function
    End If


    'CASH status and verification checking; also assigns variables based on correct selections if function does not exit. Currently no handling for SSA/SMRT Pending, as that shouldn't trigger the DAIL message
    If (CASH_disa_option_status = "BLANK") <> (CASH_disa_ver_type = "BLANK") Then
        MsgBox "CASH Disability status and verification must either both be filled or both be blank."
        Exit Function
    End If

    If CASH_disa_option_status = "01 - RSDI Only Disability" or CASH_disa_option_status = "02 - RSDI Only Blindness" or CASH_disa_option_status = "03 - SSI, SSI/RSDI Disability" or CASH_disa_option_status = "04 - SSI, SSI/RSDI Blindness" then
        ' Note to Sam - check what type of things trigger 6 because for not, just being left out
        If CASH_disa_ver_type <> "3 - Certified for RSDI or SSI" and CASH_disa_ver_type <> "N - No Verif Provided" then
            MsgBox "Disability statuses related to RSDI and SSI must be verified through SMI and have a verification type of ""3 - Certified for RSDI or SSI"". Please correct this error"
            Exit Function
        End If
        Form_needed = false
        short_verif_info = "SMI"
        verification_process = "Use SMI to verify that client has RSDI and/or SSI award letter."
    End If

    If CASH_disa_option_status = "08 - SMRT Certified Blindness" or CASH_disa_option_status = "10 - SMRT Certified Disability" then
        If CASH_disa_ver_type <> "2 - SMRT Certified" and CASH_disa_ver_type <> "N - No Verif Provided" then
            MsgBox "Disability statuses related to SMRT must have a verification type of ""2 - SMRT Certified"". Please correct this error."
            Exit Function
        End If
        ' MsgBox "Please note that most CASH programs should not be referred to SMRT for a disability determination. It is assumed you have selected this option due to the client having a usable SMRT determination from a different program."
        Form_needed = true
        form_to_send = "SMRT Referral"
        short_verif_info = "ISDS/ECF"
        verification_process = "Submit a referral through ISDS. If you do not have access to ISDS, please use ECF to create a referral for ISDS submission."
    End If

    If CASH_disa_option_status = "09 - Ill/Incapacity" then
        If CASH_disa_ver_type <> "1 - DHS161/Dr Statement" and CASH_disa_ver_type <> "6 - Other Document" and CASH_disa_ver_type <> "7 - Professional Stmt of Need" and CASH_disa_ver_type <> "N - No Verif Provided" then
            MsgBox "The selected verification type does not match the disability status selected for the client's CASH program. Please correct this error."
            Exit Function
        End If
        Form_needed = true
        form_to_send = "DHS 2114, Medical Opinion Form"
        short_verif_info = "ECF"
        verification_process = "Use ECF to contact client and request an updated Medical Opinion Form"
    End If 



    'FS status and verification checking. Currently no handling for SSA/SMRT Pending, as that shouldn't trigger the DAIL message
    If (FS_disa_option_status = "BLANK") <> (FS_disa_ver_type = "BLANK") Then
        MsgBox "SNAP/FS Disability status and verification must either both be filled or both be blank."
        Exit Function
    End If

    If FS_disa_option_status = "01 - RSDI Only Disability" or FS_disa_option_status = "02 - RSDI Only Blindness" or FS_disa_option_status = "03 - SSI, SSI/RSDI Disability" or FS_disa_option_status = "04 - SSI, SSI/RSDI Blindness" then
        If FS_disa_ver_type <> "3 - Certified for RSDI or SSI" and FS_disa_ver_type <> "N - No Verif Provided" then
            MsgBox "Disability statuses related to RSDI and SSI must be verified through SMI and have a verification type of ""3 - Certified for RSDI or SSI"". Please correct this error"
            Exit Function
        End If
        Form_needed = false
        short_verif_info = "SMI"
        verification_process = "Use SMI to verify that client has RSDI and/or SSI award letter."
    End If

    If FS_disa_option_status = "08 - SMRT Certified Blindness" or FS_disa_option_status = "10 - SMRT Certified Disability" then
        If FS_disa_ver_type <> "2 - SMRT Certified" and FS_disa_ver_type <> "N - No Verif Provided" then
            MsgBox "Disability statuses related to SMRT must have a verification type of ""2 - SMRT Certified"". Please correct this error."
            Exit Function
        End If
        ' MsgBox "Please note that SNAP cases should not be referred to SMRT for a disability determination. It is assumed you have selected this option due to the client having a usable SMRT determination from a different program."
        Form_needed = true
        form_to_send = "SMRT Referral"
        short_verif_info = "ISDS/ECF"
        verification_process = "Submit a referral through ISDS. If you do not have access to ISDS, please use ECF to create a referral for ISDS submission."
    End If

    If FS_disa_option_status = "11 - VA Determined Pd 100% Disa" or FS_disa_option_status = "12 - VA (Other Accept Disa)" or FS_disa_option_status = "13 - Cert RR Ret Disa & On Medi" or FS_disa_option_status = "14 - Othr Govt Perm Disa Ret Bnft" or FS_disa_option_status = "15 - Disability from MINE List" or FS_disa_option_status = "16 - Unable Prepare Purch Own Meals" then
        If FS_disa_ver_type <> "1 - DHS161/Dr Statement" and FS_disa_ver_type <> "N - No Verif Provided" and FS_disa_ver_type <> "4 - Receipt of HC For Disa/Blind" and FS_disa_ver_type <> "5 - Wrk Judgement" and FS_disa_ver_type <> "6 - Other Document" and FS_disa_ver_type <> "7 - Out of State Ver Pending" then
            MsgBox "The selected verification type does not match the disability status selected for the client's FS/SNAP program. Please correct this error."
            Exit Function
        End If
        Form_needed = true
        form_to_send = "DHS 2114, Medical Opinion Form (or other relevant form based on disability status)"
        short_verif_info = "ECF"
        verification_process = "Use ECF to contact client and request an updated Medical Opinion Form, or other verification documentation based on disability status"
    End If

    If FS_disa_option_status = "09 - Ill/Incapacity" then
        If FS_disa_ver_type <> "1 - DHS161/Dr Statement" and FS_disa_ver_type <> "N - No Verif Provided" and FS_disa_ver_type <> "4 - Receipt of HC For Disa/Blind" and FS_disa_ver_type <> "5 - Wrk Judgement" and FS_disa_ver_type <> "6 - Other Document" and FS_disa_ver_type <> "7 - Out of State Ver Pending" then
            MsgBox "The selected verification type does not match the disability status selected for the client's FS/SNAP program. Please correct this error."
            Exit Function
        End If
        Form_needed = true
        form_to_send = "DHS 2114, Medical Opinion Form"
        short_verif_info = "ECF"
        verification_process = "Use ECF to contact client and request an updated Medical Opinion Form"
    End If 

    'HC verification and status checking. No checking for SSA/SMRT Pending
    If (HC_disa_option_status = "BLANK") <> (HC_disa_ver_type = "BLANK") Then
        MsgBox "HC Disability status and verification must either both be filled or both be blank."
        Exit Function
    End If

    If HC_disa_option_status = "01 - RSDI Only Disability" or HC_disa_option_status = "02 - RSDI Only Blindness" or HC_disa_option_status = "03 - SSI, SSI/RSDI Disability" or HC_disa_option_status = "04 - SSI, SSI/RSDI Blindness" then
        If HC_disa_ver_type <> "3 - Certified for RSDI or SSI" and HC_disa_ver_type <> "N - No Verif Provided" then
            MsgBox "Disability statuses related to RSDI and SSI must be verified through SMI and have a verification type of ""3 - Certified for RSDI or SSI"". Please correct this error"
            Exit Function
        End If
        Form_needed = false
        short_verif_info = "SMI"
        verification_process = "Use SMI to verify that client has RSDI and/or SSI award letter."
    End If

    If HC_disa_option_status = "08 - Certified Blindness" or HC_disa_option_status = "10 - Certified Disabled" or HC_disa_option_status = "11 - Spec Cat - Disa Child" or HC_disa_option_status = "20 - TEFRA - Disabled" or HC_disa_option_status = "21 - TEFRA - Blind" or HC_disa_option_status = "MA-EPD" then
        If HC_disa_ver_type <> "2 - SMRT Certified" and HC_disa_ver_type <> "N - No Verif Provided" then
            MsgBox "This script currently only supports the ""2 - SMRT Certified"" verification type. Please process this case manually if a different verification type is required."
            Exit Function
        End If
        Form_needed = true
        form_to_send = "SMRT Referral"
        short_verif_info = "ISDS/ECF"
        verification_process = "Submit a referral through ISDS. If you do not have access to ISDS, please use ECF to create a referral for ISDS submission."
    End If


    If CASH_disa_option_status = "06 - SMRT/SSA Pend" or FS_disa_option_status = "06 - SMRT/SSA Pend" or HC_disa_option_status = "MA/Waiver" or HC_disa_option_status = "SSA/SMRT Appeal Pending" or HC_disa_option_status = "SSA/SMRT Disa Deny" then
        MsgBox "You have selected a disability status that the DAIL scrubber currently does not support. Please process this case manually."
        Exit Function
    End If


    If CASH_disa_option_status <> "BLANK" and CASH_disa_ver_type <> "BLANK" then
        CASH_active = true
    Else
        Cash_active = false
    End If

    If FS_disa_option_status <> "BLANK" and FS_disa_ver_type <> "BLANK" then
        FS_active = true
    Else
        FS_active = false
    End If

    If HC_disa_option_status <> "BLANK" and HC_disa_ver_type <> "BLANK" then
        HC_active = true
    Else
        HC_active = false
    End If

DISA_Fields_Are_Valid = True

End Function

'Function-izing the logic to create an array of the HH MEMBs and loop through each of their panels, matching the MEMB_name_to_save from our DAIL selection to the name on the panel
' MEMB_name_to_save = keep this written as the parameter
Function MEMB_Name_DAIL_matching(MEMB_name_to_save)

    Dim MEMB_loop_array()
    ReDim MEMB_loop_array(0)

    count = 0

    row = 5

    Do
        EMReadScreen MEMB_number_from_comparison, 2, row, 3

        MembNo_value = Trim(MEMB_number_from_comparison)

        If MembNo_value = "" then
            Exit Do
        End If

        If IsNumeric(MembNo_value) then
            ReDim Preserve MEMB_loop_array(count)
            MEMB_loop_array(count) = MembNo_value
            count = count + 1
        End If

        Row = row + 1
    Loop

    For i = 0 To UBound(MEMB_loop_array)
        mna = MEMB_loop_array(i)

        Do
            EMWriteScreen mna, 20, 76
            transmit
            EMReadScreen panel_exist_check, 14, 24, 13
                If panel_exist_check = "DOES NOT EXIST" then
                    Exit Do
                End If
                
                EMReadScreen name_comparison_on_panel, 30, 4, 37
                name_comparison_on_panel = trim(name_comparison_on_panel)
                    
                name_comparison_on_panel = replace(name_comparison_on_panel, ", ", ",")
                    If name_comparison_on_panel = MEMB_name_to_save then
                        exit function
                    Else
                        Exit Do
                    End If 
                
        Loop

    Next

    call script_end_procedure("no panel matching the Member Name: " & MEMB_name_to_save & " could be found. The script will now end.")

End Function


Function write_DISA_DAIL_CASE_NOTE(Original_DAIL)


    'Creating logic for listing out active programs and potentially different disability statuses in the same casenote
    If CASH_active = true and FS_active = true and HC_active = true then
        current_programs_active = "CASH, FS, HC"
        If CASH_disa_option_status = FS_disa_option_status and CASH_disa_option_status = HC_disa_option_status then
            disability_type = CASH_disa_option_status
        Elseif CASH_disa_option_status = FS_disa_option_status and CASH_disa_option_status <> HC_disa_option_status then
            disability_type = CASH_disa_option_status & " , " & HC_disa_option_status
        Elseif CASH_disa_option_status <> FS_disa_option_status and CASH_disa_option_status = HC_disa_option_status then
            disability_type = CASH_disa_option_status & " , " & FS_disa_option_status
        End If

    Elseif CASH_active = true and FS_active = true and HC_active = false then
        current_programs_active = "CASH, FS"
        If CASH_disa_option_status = FS_disa_option_status then
            disability_type = CASH_disa_option_status
        Elseif CASH_disa_option_status <> FS_disa_option_status then
            disability_type = CASH_disa_option_status & " , " & FS_disa_option_status
        End if 
    
    Elseif CASH_active = false and FS_active = true and HC_active = true then
        current_programs_active = "FS, HC"
        If FS_disa_option_status = HC_disa_option_status then
            disability_type = FS_disa_option_status
        Elseif FS_disa_option_status <> HC_disa_option_status then
            disability_type = FS_disa_option_status & " , " & HC_disa_option_status
        End if

    Elseif CASH_active = true and FS_active = false and HC_active = true then
        current_programs_active = "CASH, HC"
        If CASH_disa_option_status = HC_disa_option_status then
            disability_type = CASH_disa_option_status
        Elseif CASH_disa_option_status <> HC_disa_option_status then
            disability_type = CASH_disa_option_status & " , " & HC_disa_option_status
        End If

    Elseif CASH_active = true and FS_active = false and HC_active = false then
        current_programs_active = "CASH"
        disability_type = CASH_disa_option_status

    Elseif CASH_active = false and FS_active = true and HC_active = false then
        current_programs_active = "FS"
        disability_type = FS_disa_option_status

    Elseif CASH_active = false and FS_active = false and HC_active = true then
        current_programs_active = "HC"
        disability_type = HC_disa_option_status

    End If



    If Original_DAIL = "DISABILITY IS ENDING IN 60 DAYS - REVIEW DISABILITY STATUS" then
        '


        ' if box is checked, do TIKL
        call navigate_to_MAXIS_screen("CASE", "NOTE")
        PF9
        
        call write_variable_in_CASE_NOTE(Original_DAIL)
        call write_variable_in_CASE_NOTE("---------")
        call write_variable_in_CASE_NOTE(trim(person_info) & " - needs new verification for their disability status") 
        call write_variable_in_CASE_NOTE("Client's current disability status(es) are " & disability_type & " for " & current_programs_active)
            If Still_need_TIKL = Checked then 
            call write_variable_in_CASE_NOTE("Verification from client requested through and a TIKL was created. The case worker noted the following process was used: " & verification) 
            Else 
            call write_variable_in_CASE_NOTE("Verification was reveiwed from client through " & short_verif_info)
            End if 
        call write_variable_in_CASE_NOTE(TIKL_note_text)
        call write_variable_in_CASE_NOTE("---")
        call write_variable_in_CASE_NOTE(worker_signature)
    
    Elseif Original_DAIL = "DISA HAS ENDED - REVIEW DISA STATUS OR REDETERMINE ELIG" then
        call navigate_to_MAXIS_screen("CASE", "NOTE")
        PF9

        call write_variable_in_CASE_NOTE(Original_DAIL)
        call write_variable_in_CASE_NOTE("---------")
        call write_variable_in_CASE_NOTE(trim(person_info) & "'s disability status has ended as of " & DISA_end_date) 
        call write_variable_in_CASE_NOTE("Client's disability status(es) were " & disability_type & ". Updated documentation has not received.")
        call write_variable_in_CASE_NOTE("The case worker noted the following information regarding eligibility change(s):")
        If CASH_active = true then
            call write_variable_in_CASE_NOTE("CASH eligibility: " & CASHedit)
            call write_variable_in_CASE_NOTE("=======")
        End if 
        If FS_active = true then
            call write_variable_in_CASE_NOTE("FS eligibility: " & FSedit)
            call write_variable_in_CASE_NOTE("=======")
        End If 
        If HC_active = true then
            call write_variable_in_CASE_NOTE("HC eligibility: " & HCedit)
            call write_variable_in_CASE_NOTE("=======")
        End if

        Call write_variable_in_CASE_NOTE("---")
        call write_variable_in_CASE_NOTE(worker_signature) 

    End If

End Function







Function Find_MEMB_and_message_for_DAIL_string_match(MEMB_name_to_save, Original_DAIL)

'Pulling DAIL-screen specifics for easier deletion later
row = 1
col = 1
EMSearch "--(", row, col
    If row = 5 then
        EMReadScreen other_MEMB_name, 40, row, col+3
        other_MEMB_name = Left(other_MEMB_name, InStr(other_MEMB_name, ")")-1)
    Else
        EMReadScreen MEMB_name, 40, 5, 5
        MEMB_name = Left(MEMB_name, InStr(MEMB_name, "--")-1)
    End If
    
'saving MEMB name for comparison and usage later
If MEMB_name = "" and other_MEMB_name <> "" then
    MEMB_name_to_save = other_MEMB_name
Elseif MEMB_name <> "" and other_MEMB_name = "" then
    MEMB_name_to_save = MEMB_name
End If

row = 6
col = 1

'Double check DAIL message is still correct
EMSearch Original_DAIL, row, col
If row = 0 or col = 0 then
script_end_procedure("The script will now end because the DAIL went missing after the scrubber was started.")
End If

If col <> 0 and row = 6 then
    'Opens the Edit Summary (SUMM) screen 
    EMWriteScreen "S", 6, 3
    transmit
Else
    script_end_procedure("The script will now end because the DAIL went missing after the scrubber was started")    
End If


End Function














'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER------------------As such, it does NOT include protections to be ran independently--------------------

'Connects to MAXIS
EMConnect ""
Dialog1 = ""

'pulls and stores the MEMB name (since no number is listed); finds and saves the string of the original DAIL message
Call Find_MEMB_and_message_for_DAIL_string_match(MEMB_name_to_save, Original_DAIL)

' Navigates to the STAT/DISA screen. 
EMWriteScreen "DISA", 20, 71
transmit

'Finds the correct panel based on looping through an array of all MEMB No.s on a case and comparing the name on the panel (if it exists) to the name found on the DAIL
call MEMB_Name_DAIL_matching(MEMB_name_to_save)

'Reading all of the data and creating arrays for dropdown lists in later dialogs
Call MAXIS_case_number_finder(MAXIS_case_number)
disaDATA = Read_DISA()



'=========might need to be it's own function for cleanup?
'Build CASH Status dropdown list from disaDATA
currentCASHstatus = cash_disa_status
CASH_allStatusOptions = GLOBAL_CASH_DISA_STATUS
CASH_status_optionsString = BuildDropDownList(currentCASHstatus, CASH_allStatusOptions)

'Build CASH Verifications dropdown list from disa DATA
currentCASHver = cash_disa_ver
CASH_allVerOptions = GLOBAL_CASH_DISA_VER
CASH_ver_optionsString = BuildDropDownList(currentCASHver, CASH_allVerOptions)

'Build FS Status dropdown List from disaDATA
currentFSstatus = fs_disa_status
FS_allStatusOptions = GLOBAL_FS_DISA_STATUS
FS_status_optionsString = BuildDropDownList(currentFSstatus, FS_allStatusOptions)

'Build FS Verifications dropdown List from disaDATA
currentFSver = fs_disa_ver
FS_allVerOptions = GLOBAL_FS_DISA_VER
FS_ver_optionsString = BuildDropDownList(currentFSver, FS_allVerOptions)

'Build HC Status dropdown List from disaDATA
currentHCstatus = hc_disa_status
HC_allStatusOptions = GLOBAL_HC_DISA_STATUS
HC_status_optionsString = BuildDropDownList(currentHCstatus, HC_allStatusOptions)

'Build HC Verifications dropdown list from disaDATA
currentHCver = hc_disa_ver
HC_allVerOptions = GLOBAL_HC_DISA_VER
HC_ver_optionsString = BuildDropDownList(currentHCver, HC_allVerOptions)

currentCASHverificationprocess = verification_process
CASH_allverificationprocesses = GLOBAL_CASH_VERIF_PROCESSES
CASH_combobox_string = BuildDropDownList(currentCASHverificationprocess, CASH_allverificationprocesses)

currentFSverificationprocess = verification_process
FS_allverificationprocesses = GLOBAL_FS_VERIF_PROCESSES
FS_combobox_string = BuildDropDownList(currentFSverificationprocess, FS_allverificationprocesses)

currentHCverificationprocess = verification_process
HC_allverificationprocesses = GLOBAL_HC_VERIF_PROCESSES
HC_combobox_string = BuildDropDownList(currentHCverificationprocess, HC_allverificationprocesses)

'Clearing Dialog1
Dialog1 = ""
' Initial Dialog Box framwork that displays the STAT/DISA information
BeginDialog Dialog1, 0, 0, 690, 275, "STAT/DISA Information"
    ButtonGroup ButtonPressed
        OkButton 570, 250, 50, 15
        CancelButton 630, 250, 50, 15
    GroupBox 10, 35, 320, 185, "DISA Dates"
        Text 20, 55, 295, 15, "Disability Begin: " & DISA_bgn_date
        Text 20, 75, 295, 15, "Disability End: " & DISA_end_date
        Text 20, 95, 295, 15, "Certification Begin: " & Cert_period_bgn
        Text 20, 115, 295, 15, "Certification End: " & Cert_period_end
        Text 20, 135, 295, 15, "Elder Waiver Begin: " & Eld_wavr_bgn
        Text 20, 155, 295, 15, "Elder Waiver End: " & Eld_wavr_end
        Text 20, 175, 295, 15, "GRH Begin: " & GRH_bgn
        Text 20, 195, 295, 15, "GRH End: " & GRH_end
    GroupBox 345, 65, 330, 165, "DISA Status and Verification"
        Text 355, 85, 310, 15, "Cash Status: " & cash_disa_status
        Text 355, 105, 310, 15, "Cash Verification: " & cash_disa_ver
        Text 355, 125, 310, 15, "FS Status: " & fs_disa_status
        Text 355, 145, 310, 15, "FS Verification: " & fs_disa_ver
        Text 355, 165, 310, 15, "HC Status: " & hc_disa_status
        Text 355, 185, 310, 15, "HC Verification: " & hc_disa_ver
        Text 355, 10, 135, 15, "CASE NUMBER: " & MAXIS_case_number
        Text 15, 225, 300, 15, "Select Action You Would Like to Take"
        Text 355, 30, 320, 15, "MEMBER: " & person_info
        Text 15, 10, 320, 15, "DAIL MSG: " & original_DAIL
    DropListBox 15, 240, 180, 20, "TIKL and CASE/NOTE Updated Verification Request"+chr(9)+"Disability Ended - CASE/NOTE details"+chr(9)+"CASE/NOTE only"+chr(9)+"Delete DAIL only", DISA_option_list
EndDialog

'Initialize [STAT/DISA Information] Dialog Box
Do
    Dialog Dialog1
    cancel_confirmation
        If DISA_option_list = "TIKL and CASE/NOTE Updated Verification Request" And ButtonPressed = OK then
            'clear the dialog for rebuilding
            Dialog1 = ""
            'Dialog for right after "TIKL and CASE/NOTE Updated Verification Request"
                BeginDialog Dialog1, 0, 0, 290, 105, "Verification Source Check"
                    ButtonGroup ButtonPressed
                        PushButton 170, 40, 100, 25, "YES - Proceed to TIKL", TIKL_button
                        PushButton 170, 70, 100, 25, "NO - Update DISA Panel", Update_Before_TIKL
                    Text 5, 5, 150, 20, "The current verification source(s) is listed as"
                    Text 5, 30, 155, 15, "CASH Verif: " & cash_disa_ver
                    Text 5, 50, 155, 15, "FS Verif: " & fs_disa_ver
                    Text 5, 70, 155, 15, "HC Verif: " & hc_disa_ver
                    Text 175, 10, 95, 20, "Is this correct?"
                EndDialog
                    Do
                    Dialog Dialog1
                    cancel_confirmation
                        If ButtonPressed = Update_Before_TIKL then 
                            'clear the dialog for rebuilding
                            Dialog1 = ""
                            'First dialog for updating the DISA panel if requested from Dialog_TIKL_checking
                                BeginDialog Dialog1, 0, 0, 351, 115, "Update DISA Panel Options"
                                ButtonGroup ButtonPressed
                                    OkButton 290, 65, 50, 15
                                    CancelButton 290, 85, 50, 15
                                DropListBox 60, 20, 115, 20, CASH_status_optionsString, CASH_disa_option_status
                                DropListBox 60, 50, 115, 20, FS_status_optionsString, FS_disa_option_status
                                DropListBox 60, 80, 115, 20, HC_status_optionsString, HC_disa_option_status
                                DropListBox 195, 20, 80, 25, CASH_ver_optionsString, CASH_disa_ver_type
                                DropListBox 195, 50, 80, 25, FS_ver_optionsString, FS_disa_ver_type
                                DropListBox 195, 80, 80, 25, HC_ver_optionsString, HC_disa_ver_type
                                Text 5, 20, 50, 15, "CASH - "
                                Text 5, 50, 50, 15, "FS - "
                                Text 5, 80, 50, 15, "HC - "
                                Text 65, 5, 105, 15, "Disability Status"
                                Text 195, 5, 105, 15, "Verification Type"
                                EndDialog
                            
                            Do
                                'initialize the dialog with the [Update DISA Panel Options v1] - 
                                Dialog dialog1
                                cancel_confirmation
                                    If ButtonPressed = OK then 
                                        If DISA_Fields_Are_Valid(CASH_disa_option_status, CASH_disa_ver_type, FS_disa_option_status, FS_disa_ver_type, HC_disa_option_status, HC_disa_ver_type) then 
                                            call Update_DISA_Panel_Loop()
                                        End If
                                    Dialog1 = ""
                                    'establish new dialog
                                    'Second/Final dialog for updating the DISA panel - it shows the new selection and allows for the user to either press "correct" and proceed to the TIKL and CASE/NOTE; or loop through making more changes
                                        BeginDialog Dialog1, 0, 0, 421, 115, "Update DISA Panel Options"
                                            ButtonGroup ButtonPressed
                                                CancelButton 320, 80, 50, 15
                                                PushButton 300, 30, 95, 15, "Correct - Proceed to TIKL", TIKL_Correct
                                                PushButton 290, 55, 115, 15, "More Changes Made - Update", Update_Again
                                            DropListBox 60, 20, 115, 20, CASH_status_optionsString, CASH_disa_option_status
                                            DropListBox 60, 50, 115, 20, FS_status_optionsString, FS_disa_option_status
                                            DropListBox 60, 80, 115, 20, HC_status_optionsString, HC_disa_option_status
                                            DropListBox 195, 20, 80, 25, CASH_ver_optionsString, CASH_disa_ver_type
                                            DropListBox 195, 50, 80, 25, FS_ver_optionsString, FS_disa_ver_type
                                            DropListBox 195, 80, 80, 25, HC_ver_optionsString, HC_disa_ver_type
                                            Text 5, 20, 50, 15, "CASH"
                                            Text 5, 50, 50, 15, "FS"
                                            Text 5, 80, 50, 15, "HC"
                                            Text 65, 5, 105, 15, "Disability Status"
                                            Text 195, 5, 105, 15, "Verification Type"
                                        EndDialog
                                    Do 
                                        'Initialize the dialog with the [update DISA panel Options v2]
                                        Dialog Dialog1
                                        cancel_confirmation
                                            If ButtonPressed = Update_Again Then
                                                If DISA_Fields_Are_Valid(CASH_disa_option_status, CASH_disa_ver_type, FS_disa_option_status, FS_disa_ver_type, HC_disa_option_status, HC_disa_ver_type) then
                                                    call Update_DISA_Panel_Loop()
                                                End If
                                            End If
                                            If ButtonPressed = TIKL_Correct then
                                                DAIL_DISA_TIKL_goahead = false 
                                                If DISA_Fields_Are_Valid(CASH_disa_option_status, CASH_disa_ver_type, FS_disa_option_status, FS_disa_ver_type, HC_disa_option_status, HC_disa_ver_type) then
                                                    Call Update_DISA_Panel_Loop()
                                                    DAIL_DISA_TIKL_goahead = true
                                                End If
                                            End If    
                                        call check_for_password(are_we_passworded_out)
                                    Loop until ButtonPressed = TIKL_Correct and are_we_passworded_out = false and DAIL_DISA_TIKL_goahead = true
                                        Dialog1 = ""
                                        Still_need_TIKL = checked
                                        BeginDialog Dialog1, 0, 0, 326, 170, "Final Verification Submission Step"
                                            ButtonGroup ButtonPressed
                                                OkButton 230, 145, 40, 15
                                                CancelButton 275, 145, 40, 15
                                                PushButton 5, 145, 65, 15, "One Policy", verify_cash_disa_policy
                                                PushButton 75, 145, 65, 15, "Two Policy", verify_snap_disa_policy
                                                PushButton 145, 145, 65, 15, "Three Policy", verify_hc_disa_policy
                                            GroupBox 5, 30, 95, 40, "CASH"
                                            GroupBox 110, 30, 95, 40, "FS"
                                            GroupBox 215, 30, 95, 40, "HC"
                                            ComboBox 15, 45, 75, 15, CASH_combobox_string, CASH_Final_Dropbox
                                            ComboBox 120, 45, 75, 15, FS_combobox_string, FS_Final_Dropbox
                                            ComboBox 225, 45, 75, 15, HC_combobox_string, HC_Final_Dropbox
                                            EditBox 10, 100, 300, 15, verif_notes_from_user
                                            CheckBox 15, 125, 250, 10, "Uncheck if TIKL is no longer needed", Still_need_TIKL
                                            Text 10, 75, 295, 20, "Enter any notes related to any actions you took or changes you made during the verification request process."
                                            Text 10, 5, 300, 20, "Verify and/or change the method you used to request updated disability verification."
                                        EndDialog
                                            Do
                                            Dialog dialog1
                                            

                                            call check_for_password(are_we_passworded_out)
                                            Loop until are_we_passworded_out = false

                                        If DAIL_DISA_TIKL_goahead = true then
                                            If Still_need_TIKL = checked then
                                                call create_TIKL(form_to_send & " was submitted 30 days ago. Please review case, and send another request if applicable.", 30, date, False, TIKL_note_text)
                                            End If
                                            call write_DISA_DAIL_CASE_NOTE(Original_DAIL)
                                            'Now for DAIL deletion
                                            call navigate_to_MAXIS_screen("DAIL", "DAIL") 
                                            EMWriteScreen MAXIS_case_number, 20, 38
                                            Transmit 
                                            call search_for_correct_DAIL_and_delete(MEMB_name_to_save, Original_DAIL)
                                        End If
                                    End If
                                Loop Until ButtonPressed <> Update_Before_TIKL
                            End If
                        If ButtonPressed = TIKL_button then
                            Dialog1 = ""
                            Still_need_TIKL = checked
                            BeginDialog Dialog1, 0, 0, 326, 170, "Final Verification Submission Step"
                                ButtonGroup ButtonPressed
                                    OkButton 230, 145, 40, 15
                                    CancelButton 275, 145, 40, 15
                                    PushButton 5, 145, 65, 15, "One Policy", verify_cash_disa_policy
                                    PushButton 75, 145, 65, 15, "Two Policy", verify_snap_disa_policy
                                    PushButton 145, 145, 65, 15, "Three Policy", verify_hc_disa_policy
                                GroupBox 5, 30, 95, 40, "CASH"
                                GroupBox 110, 30, 95, 40, "FS"
                                GroupBox 215, 30, 95, 40, "HC"
                                ComboBox 15, 45, 75, 15, CASH_combobox_string, CASH_Final_Dropbox
                                ComboBox 120, 45, 75, 15, "", FS_Final_Dropbox
                                ComboBox 225, 45, 75, 15, "", HC_Final_Dropbox
                                EditBox 10, 100, 300, 15, Edit1
                                CheckBox 15, 125, 250, 10, "Uncheck if TIKL is no longer needed", Still_need_TIKL
                                Text 10, 75, 295, 20, "Enter any notes related to any actions you took or changes you made during the verification request process."
                                Text 10, 5, 300, 20, "Verify and/or change the method you used to request updated disability verification."
                            EndDialog
                                Do
                                Dialog dialog1
                                

                                call check_for_password(are_we_passworded_out)
                                Loop until are_we_passworded_out = false

                                If Still_need_TIKL = checked then
                                    call create_TIKL(form_to_send & " was submitted using" & short_verif_info & "30 days ago. Please review case, and send another request if applicable.", 30, date, False, TIKL_note_text)
                                End If
                            call write_DISA_DAIL_CASE_NOTE(Original_DAIL)
                            'Now for DAIL deletion
                            call navigate_to_MAXIS_screen("DAIL", "DAIL") 
                            EMWriteScreen MAXIS_case_number, 20, 38
                            Transmit 
                            call search_for_correct_DAIL_and_delete(MEMB_name_to_save, Original_DAIL)    
                        End If
                    Call check_for_password(are_we_passworded_out)
                Loop until are_we_passworded_out = false 
        End If    

        If DISA_option_list = "Disability Ended - CASE/NOTE" then
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 316, 375, "Dialog"
                ButtonGroup ButtonPressed
                    OkButton 195, 345, 50, 15
                    CancelButton 255, 345, 50, 15
                GroupBox 10, 5, 285, 50, "Current End Dates"
                Text 20, 25, 100, 20, "DISA end date:  & DISA_end_date"
                Text 165, 25, 100, 20, "Cert end date:  & Cert_period_end"
                Text 15, 60, 295, 10, "If end date(s) are incorrect, cancel the script and process case manually"
            GroupBox 10, 75, 285, 70, "CASH Programs"
                Text 20, 90, 120, 30, "CASH status:"
                Text 150, 85, 135, 35, "***Please note: a disability ending can impact eligibility status for CASH programs. Please note below what eligibility changes, if any, occurred."
                EditBox 15, 125, 275, 15, CASHedit
            GroupBox 10, 155, 285, 100, "FS/SNAP Programs"
                Text 20, 170, 120, 30, "FS/SNAP status:"
                Text 145, 165, 135, 55, "***Please note: a disability ending can impact eligibility status for FS/SNAP programs. Please run the ABAWD exemption finder and note below what eligibility changes, if any, occurred."
                EditBox 15, 230, 275, 15, FSedit
                ButtonGroup ButtonPressed
                PushButton 20, 205, 110, 15, "Run ABAWD Exemption Finder", ABAWD_script_trigger
            GroupBox 10, 260, 285, 75, "HC Programs"
                EditBox 15, 310, 275, 15, HCedit
                Text 20, 275, 120, 30, "HC status:"
                Text 150, 270, 135, 35, "***Please note: a disability ending can impact eligibility status for CASH programs. Please note below what eligibility changes, if any, occurred."
            EndDialog
            Do
            Dialog Dialog1
            cancel_confirmation
                If ButtonPressed = ABAWD_script_trigger then
                    call ABAWD_FSET_exemption_finder()
                End if

            call check_for_password(are_we_passworded_out)
            Loop until are_we_passworded_out = false and ButtonPressed <> cancel
            call write_DISA_DAIL_CASE_NOTE(Original_DAIL)
            'Now for DAIL deletion
            call navigate_to_MAXIS_screen("DAIL", "DAIL") 
            EMWriteScreen MAXIS_case_number, 20, 38
            Transmit 
            call search_for_correct_DAIL_and_delete(MEMB_name_to_save, Original_DAIL)    
        End If
        If DISA_option_list = "CASE/NOTE only" And ButtonPressed = OK Then
            call write_DISA_DAIL_CASE_NOTE(Original_DAIL)
        End If
        'If "Delete DAIL only" is the selection from the drop-down menu of the main dialog, it will trigger a confirmation dialog box reminding the user that this is only deleting the DAIL and taking no action
        If DISA_option_list = "Delete DAIL only" And ButtonPressed = OK Then
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 191, 105, "Deletion Confirmation"
                Text 5, 5, 175, 60, "You have selected that you would like to delete the DAIL message for this case. This will only delete the DAIL message and will not take any action nor create any CASE/NOTE. Do you wish to proceed?"
                ButtonGroup ButtonPressed
                    PushButton 10, 70, 75, 25, "Yes - Delete", Delete_yes_only
                    PushButton 100, 70, 75, 25, "No - End Script", End_script_no_delete_only
            EndDialog
            Do
                Dialog Dialog1
                If ButtonPressed = Delete_yes_only then    
                    call back_to_SELF()
                    call navigate_to_MAXIS_screen("DAIL", "DAIL") 
                    EMWriteScreen MAXIS_case_number, 20, 38
                    Transmit 
                    call search_for_correct_DAIL_and_delete(MEMB_name_to_save, Original_DAIL)
                ElseIf ButtonPressed = End_script_no_delete_only then
                    script_end_procedure("The script will now end.")
                End If
            Loop until ButtonPressed <> ""    
        End If

        call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false 

script_end_procedure("script done. This is testing only")


'----------------------------------------------------------------------------------------------------Closing Project Documentation 
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------10-29-26------------------Discussed with Dave during early drafts
'--Tab orders reviewed & confirmed----------------------------------------------n/a
'--Mandatory fields all present & Reviewed--------------------------------------
'--All variables in dialog match mandatory fields-------------------------------
'Review dialog names for content and content fit in dialog----------------------
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------
'--Include script category and name somewhere on first dialog-------------------
'--Create a button to reference instructions------------------------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------
'--CASE:NOTE Header doesn't look funky------------------------------------------
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------n/a-----------------------Due to deletion functionality add, CASE/NOTE not left in edit mode
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used --
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------
'--MAXIS_background_check reviewed (if applicable)------------------------------
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------
'--script_end_procedures (w/ or w/o error messaging)----------------------------
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------SAM, check for all the "tee hees" you were testing with! 
'--Remove testing code/unnecessary code-----------------------------------------
'--Review/update SharePoint instructions----------------------------------------
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------

