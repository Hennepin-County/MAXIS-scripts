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

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

Function create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust)
    'All 10-day cutoff dates are provided in POLI/TEMP TE19.132
    '--- This function creates and saves a TIKL message in DAIL/WRIT. 
    '~~~~~ TIKL_text: Text that the TIKL message will say. 
    '~~~~~ num_of_days: how many days the TIKL should be set for. Must be a numeric or a numeric value. 
    '~~~~~ date_to_start: this determines which date to start counting the number of days to TIKL out from. Ex: date to use today's date, or Application_date to use within the CAF. 
    '~~~~~ ten_day_adjust: True or False. True to adjust the TIKL date to the 1st day of the next month if after 10 day cutoff, False to NOT adjust to 10 day cutoff.
    '===== Keywords: MAXIS, TIKL
    
    TIKL_date = DateAdd("D", num_of_days, date_to_start)    'Creates the TIKL date based on the number of days and date to start chosen by the user    
    If cdate(TIKL_date) < date then
        msgbox "Unable to create TIKL, the TIKL date is a past date. Please manually track this case and action."   'fail-safe in case the TIKL date created is in the past. DAIL/WRIN does not allow past dates.
    Else 
        If ten_day_adjust = True then 
            TIKL_mo = right("0" & DatePart("m",    TIKL_date), 2) 'Creating new month and year variables to determine which ten day cut off date to use  
            TIKL_yr = right(      DatePart("yyyy", TIKL_date), 2) 
            
            IF TIKL_mo = "01" AND TIKL_yr = "20" THEN
                ten_day_cutoff = #01/21/2020#
            ELSEIF TIKL_mo = "02" AND TIKL_yr = "20" THEN
                ten_day_cutoff = #02/19/2020#
            ELSEIF TIKL_mo = "03" AND TIKL_yr = "20" THEN
                ten_day_cutoff = #03/19/2020#
            ELSEIF TIKL_mo = "04" AND TIKL_yr = "20" THEN
                ten_day_cutoff = #04/20/2020#
            ELSEIF TIKL_mo = "05" AND TIKL_yr = "20" THEN
                ten_day_cutoff = #05/21/2020#
            ELSEIF TIKL_mo = "06" AND TIKL_yr = "20" THEN
                ten_day_cutoff = #06/18/2020#
            ELSEIF TIKL_mo = "07" AND TIKL_yr = "20" THEN
                ten_day_cutoff = #07/21/2020#
            ELSEIF TIKL_mo = "08" AND TIKL_yr = "20" THEN
                ten_day_cutoff = #08/20/2020#
            ELSEIF TIKL_mo = "09" AND TIKL_yr = "20" THEN
                ten_day_cutoff = #09/18/2020#
            ELSEIF TIKL_mo = "10" AND TIKL_yr = "20" THEN
                ten_day_cutoff = #10/21/2020#
            ELSEIF TIKL_mo = "11" AND TIKL_yr = "20" THEN
                ten_day_cutoff = #11/19/2020#
            ELSEIF TIKL_mo = "12" AND TIKL_yr = "20" THEN
                ten_day_cutoff = #12/21/2020#    
            ELSEIF TIKL_mo = "12" AND TIKL_yr = "19" THEN
                ten_day_cutoff = #12/19/2019#
            Else 
                missing_date = True 'in case TIKL time spans exceed 10 day cut off calendar. 
            End if
            
            If missing_date = True then 
                TIKL_date = TIKL_date 'defaults to the date set by the user 
            Else 
                'Determining the TIKL date based on if past 10 day cut off or not. 
                If cdate(TIKL_date) > cdate(ten_day_cutoff) then
                    'Date of the 1st of the next month where negative action can be taken is determined & becomes the TIKL_date
                    new_TIKL_mo = right(DatePart("m",    DateAdd("m", 1, TIKL_date)), 2)
                    new_TIKL_yr = right(DatePart("yyyy", DateAdd("m", 1, TIKL_date)), 2)
                    TIKL_date = new_TIKL_mo & "/01/" & new_TIKL_yr
                End if
            End if 
        End if 
        'Creating the TIKL message 
        Call navigate_to_MAXIS_screen("DAIL", "WRIT")           
        call create_MAXIS_friendly_date(TIKL_date, 0, 5, 18)    '0 is the date as all the adjustments are already determined. 
        Call write_variable_in_TIKL(TIKL_text)
        PF3 'to save & exit 
    End if
End Function

MAXIS_case_number = "286750"
Call create_TIKL("testing the TIKLs", 10, date, true)
stopscript