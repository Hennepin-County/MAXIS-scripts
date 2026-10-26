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
call changelog_update("10/20/26", "Updated functionality - displays original DAIL PEPR/DISA message as well as the info from STAT/DISA. Allows HSR to then select whether to TIKL, create CASE/NOTE, or other.", "Sam Begley-May, Hennepin County")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note. Updated background navigation coding.", "Ilse Ferris")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'-------new functions



Original_DAIL = "DISABILITY IS ENDING IN 60 DAYS - REVIEW DISABILITY STATUS"

Dim row, col, Original_DAIL
Dim DISA_bgn_date, DISA_end_date
Dim Cert_period_bgn, Cert_period_end
Dim Eld_wavr_bgn, Eld_wavr_end, GRH_bgn, GRH_end
Dim cash_disa_status, cash_disa_ver
Dim fs_disa_status, fs_disa_ver, hc_disa_status, hc_disa_ver
Dim MEMB_No, last_first_name, person_info
Dim form_to_send
Dim GLOBAL_CASH_DISA_STATUS, GLOBAL_CASH_DISA_VER, GLOBAL_FS_DISA_STATUS, GLOBAL_FS_DISA_VER, GLOBAL_HC_DISA_STATUS, GLOBAL_HC_DISA_VER
Dim origArray, newArray
Dim orig_cash_status, orig_cash_ver, orig_fs_status, orig_fs_ver, orig_hc_status, orig_hc_ver


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

Function search_for_correct_DAIL_and_delete(last_first_name, Original_DAIL)
    'Check on the right screen
    EMReadScreen DAIL_check, 7, 2, 26
    ' The (DAIL) is located on both the DAIL/"" page and the DAIL/DAIL page. The word "WORKERS" is in a different location on both pages. This will keep the DAIL script from getting stuck on the submenu
    If DAIL_check <> "WORKERS" then script_end_procedure("You are not in DAIL for the final deletion. The script will now end")

    row = 5
    col = 4
    
    EMSearch last_first_name, row, col
    
    If row = 0 and col = 0 then
        Do
            PF8
            EMReadScreen last_page_check, 21, 24, 2
                If last_page_check = "THIS IS THE LAST PAGE" then script_end_procedure("The original DAIL could not be found for deletion. The script will now end.")
                Exit Do
            EMSearch last_first_name, row, col
        Loop until row <> 0 and col <> 0
    End if

    EMSearch Original_DAIL, row, col 
    'Failsafe for if the person is on one page and their DAIL is on the next 
      If row = 0 and col = 0 then
        Do
          PF8
          EMSearch last_first_name, row, col
            If row = 0 and col = 0 then script_end_procedure("The original DAIL could not be found associated with that client. The script will now end.")
            Exit Do
          EMSearch Original_DAIL, row, col
        Loop until row <> 0 and col <> 0
      End If
    
      EMWriteScreen "D", row, 3
      transmit
      MsgBox "The original DAIL was found and deleted"

End Function

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


Function Read_DISA()
    'double-check that user hasn't accidentally navigated away from the DISA panel
    EMReadScreen DISA_chk, 4, 2, 45
    If DISA_chk <> "DISA" then script_end_procedure("Not on a DISA message, the script will now end.") ' Note to Sam - can we make a more specific DAIL end procedure, where it does everything a script end procedure does but also goes back to the DAIL scrubber?

    EMReadScreen MEMB_No, 2, 4, 33
    trim(MEMB_No)
    EMReadScreen last_first_name, 33, 4, 37
    trim(last_first_name)
    person_info = MEMB_No & " " & last_first_name

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
        If cash_disa_ver = "1" then 
            cash_disa_ver = "1 - DHS161/Dr Statement" 
            form_to_send = "Medical Opinion Form"
        End If
        If cash_disa_ver = "2" then 
            cash_disa_ver = "2 - SMRT Certified"  
            form_to_send = "SMRT"
        End If
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
        If hc_disa_status = "10" then hc_disa_status = "10 - Certified Disabiled"
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

    
orig_cash_status = cash_disa_status
orig_cash_ver    = cash_disa_ver
orig_fs_status   = fs_disa_status
orig_fs_ver      = fs_disa_ver
orig_hc_status   = hc_disa_status
orig_hc_ver      = hc_disa_ver  

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

    Read_DISA = results


End Function



Function Update_DISA_Panel_Loop()
    Dim changedIndexes, idx, i

    origArray = Array(orig_cash_status, orig_fs_status, orig_hc_status, orig_cash_ver, orig_fs_ver, orig_hc_ver)
    newArray = Array(CASH_disa_option_status, FS_disa_option_status, HC_disa_option_status, CASH_disa_ver_type, FS_disa_ver_type, HC_disa_ver_type)

    EMReadScreen DISA_chk, 4, 2, 45
    If DISA_chk <> "DISA" Then script_end_procedure("Not on a DISA message, the script will now end.")

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

Function DISA_Fields_Are_Valid(CASH_disa_option_status, CASH_disa_ver_type, FS_disa_option_status, FS_disa_ver_type, HC_disa_option_status, HC_disa_ver_type)
    
    DISA_Fields_Are_Valid = False

    If CASH_disa_option_status = "BLANK" And CASH_disa_ver_type = "BLANK" And _
       FS_disa_option_status = "BLANK" And FS_disa_ver_type = "BLANK" And _
       HC_disa_option_status = "BLANK" And HC_disa_ver_type = "BLANK" Then
        MsgBox "All disability status and verification fields are blank. Please correct the information before proceeding."
        Exit Function
    End If

    If (CASH_disa_option_status = "BLANK") <> (CASH_disa_ver_type = "BLANK") Then
        MsgBox "CASH Disability status and verification must either both be filled or both be blank."
        Exit Function
    End If

    If (FS_disa_option_status = "BLANK") <> (FS_disa_ver_type = "BLANK") Then
        MsgBox "SNAP/FS Disability status and verification must either both be filled or both be blank."
        Exit Function
    End If

    If (HC_disa_option_status = "BLANK") <> (HC_disa_ver_type = "BLANK") Then
        MsgBox "HC Disability status and verification must either both be filled or both be blank."
        Exit Function
    End If

    DISA_Fields_Are_Valid = True

End Function


'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER------------------As such, it does NOT include protections to be ran independently--------------------

'Connects to MAXIS
EMConnect ""
Dialog1 = ""

'Opens the Edit Summary (SUMM) screen - note to Sam, this may eventually need adjusted based on how the scrubber gets updated. 
'EMWriteScreen "S", 6, 3
'transmit

' Navigates to the STAT/DISA screen. The script will automatically go to the MEMB Ref # that is associated with this DAIL
'EMWriteScreen "DISA", 20, 71
'transmit
EMReadScreen MEMB_ref_num, 2, 4, 33

' =======Reading all of the data
'Pulls the case number
Call MAXIS_case_number_finder(MAXIS_case_number)
disaDATA = Read_DISA()

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




'===========



' Sam - this will need error handling eventually 
Dialog1 = ""
' Initial Dialog Box that displays the STAT/DISA information
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
  DropListBox 15, 240, 180, 20, "TIKL and CASE/NOTE updated verif request"+chr(9)+"CASE/NOTE only"+chr(9)+"Delete DAIL only", DISA_option_list
EndDialog

    Do
        Dialog Dialog1
            If DISA_option_list = "TIKL and CASE/NOTE updated verif request" And ButtonPressed = OK then
                'clear the dialog for rebuilding
                Dialog1 = ""
                'Dialog for right after "TIKL and CASE/NOTE updated verif request"
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
                            
                            Dialog dialog1
                            cancel_confirmation
                                If ButtonPressed = OK and DISA_Fields_Are_Valid(CASH_disa_option_status, CASH_disa_ver_type, FS_disa_option_status, FS_disa_ver_type, HC_disa_option_status, HC_disa_ver_type) = true then call Update_DISA_Panel_Loop()
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
                                    Dialog Dialog1
                                        If ButtonPressed = Update_Again and DISA_Fields_Are_Valid(CASH_disa_option_status, CASH_disa_ver_type, FS_disa_option_status, FS_disa_ver_type, HC_disa_option_status, HC_disa_ver_type) = true then 
                                            call Update_DISA_Panel_Loop()
                                        End If
                                        If ButtonPressed = Cancel then
                                            call script_end_procedure("You have cancelled the script. It will now end.")
                                        End If
                                        call check_for_password(are_we_passworded_out)
                                Loop until ButtonPressed = TIKL_Correct and are_we_passworded_out = false and DISA_Fields_Are_Valid(CASH_disa_option_status, CASH_disa_ver_type, FS_disa_option_status, FS_disa_ver_type, HC_disa_option_status, HC_disa_ver_type) = true           
                        Loop Until ButtonPressed <> Update_Before_TIKL
                        If ButtonPressed = cancel then script_end_procedure("You have chosen to cancel the script, the script will now end.")
                        If ButtonPressed = TIKL_Correct and DISA_Fields_Are_Valid(CASH_disa_option_status, CASH_disa_ver_type, FS_disa_option_status, FS_disa_ver_type, HC_disa_option_status, HC_disa_ver_type) then
                            call create_TIKL(form_to_send & " was submitted 30 days ago. Please review case, and send another request if applicable.", 30, date, False, TIKL_note_text)
                            call navigate_to_MAXIS_screen("CASE", "NOTE")
                            PF9
                            call write_variable_in_CASE_NOTE("DISABILITY IS ENDING IN 60 DAYS - REVIEW DISABILITY STATUS")
                            call write_variable_in_CASE_NOTE("---------")
                            call write_variable_in_CASE_NOTE(person_info & " needs new verification for their disability status")
                            call write_variable_in_CASE_NOTE("Client was sent " & form_to_send & " and a TIKL was created")
                            call write_variable_in_CASE_NOTE(TIKL_note_text)
                            Call write_variable_in_CASE_NOTE("---")
                            call write_variable_in_CASE_NOTE(worker_signature)
                            'Now for DAIL deletion
                            call navigate_to_MAXIS_screen("DAIL", "DAIL") 
                            EMWriteScreen MAXIS_case_number, 20, 38
                            Transmit 
                            call search_for_correct_DAIL_and_delete(last_first_name, Original_DAIL)

                        End If
                    End If
                    If ButtonPressed = TIKL_button then
                        call create_TIKL(form_to_send & " was submitted 30 days ago. Please review case, and send another request if applicable.", 30, date, False, TIKL_note_text)
                        call navigate_to_MAXIS_screen("CASE", "NOTE")
                        PF9
                        call write_variable_in_CASE_NOTE("DISABILITY IS ENDING IN 60 DAYS - REVIEW DISABILITY STATUS")
                        call write_variable_in_CASE_NOTE("---------")
                        call write_variable_in_CASE_NOTE(person_info & " needs new verification for their disability status")
                        call write_variable_in_CASE_NOTE("Client was sent " & form_to_send & " and a TIKL was created")
                        call write_variable_in_CASE_NOTE(TIKL_note_text)
                        Call write_variable_in_CASE_NOTE("---")
                        call write_variable_in_CASE_NOTE(worker_signature)
                        'Now for DAIL deletion
                        call navigate_to_MAXIS_screen("DAIL", "DAIL") 
                        EMWriteScreen MAXIS_case_number, 20, 38
                        Transmit 
                        call search_for_correct_DAIL_and_delete(last_first_name, Original_DAIL)    
                    End If
                    Call check_for_password(are_we_passworded_out)
                Loop until are_we_passworded_out = false 

            End If    
            If DISA_option_list = "CASE/NOTE only" And ButtonPressed = OK Then
                        call navigate_to_MAXIS_screen("CASE", "NOTE")
                        PF9
                        call write_variable_in_CASE_NOTE("DISABILITY IS ENDING IN 60 DAYS - REVIEW DISABILITY STATUS")
                        call write_variable_in_CASE_NOTE("---------")
                        call write_variable_in_CASE_NOTE(person_info & " needs new verification for their disability status")
                        call write_variable_in_CASE_NOTE("Client was sent " & form_to_send & " and a TIKL was created")
                        call write_variable_in_CASE_NOTE(TIKL_note_text)
                        Call write_variable_in_CASE_NOTE("---")
                        call write_variable_in_CASE_NOTE(worker_signature)  
            End If
            If DISA_option_list = "Delete DAIL only" And ButtonPressed = OK Then
                call navigate_to_MAXIS_screen("DAIL", "DAIL") 
                EMWriteScreen MAXIS_case_number, 20, 38
                Transmit 
                call search_for_correct_DAIL_and_delete(last_first_name, Original_DAIL)
            End If

        
        call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = false 




script_end_procedure("script done. This is testing only")
