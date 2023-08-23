'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - ELIGIBILITY SUMMARY.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 500                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'END FUNCTIONS LIBRARY BLOCK================================================================================================
other_county_redirect = True
run_locally = False
henn_county_git_hub_repo = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/"
call run_from_GitHub(henn_county_git_hub_repo & "notes/eligibility-summary.vbs")
