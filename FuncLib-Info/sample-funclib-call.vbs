'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY ===============================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN                    'Shouldn't load FuncLib if it already loaded once
    IF run_locally = FALSE or run_locally = "" THEN    'If the scripts are set to run locally, it skips this and uses an FSO
        IF use_master_branch = TRUE THEN               'Defined in global variables, defaults to RELEASE if not defined
            FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
        Else
            FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
        End if
        SET req = CreateObject("Msxml2.XMLHttp.6.0")                'Creates an object to get a FuncLib_URL
        req.open "GET", FuncLib_URL, FALSE                          'Attempts to open the FuncLib_URL
        req.send                                                    'Sends request
        IF req.Status = 200 THEN                                    '200 means great success
            Set fso = CreateObject("Scripting.FileSystemObject")    'Creates an FSO
            Execute req.responseText                                'Executes the script code
        ELSE                                                        'Error message
            critical_error_msgbox = MsgBox ("Something has gone wrong. Could not connect to FuncLib code stored on GitHub." &_
                                            vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL &_
                                            vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection." &_
                                            vbNewLine & vbNewLine &_
                                            "Consult a scripts administrator with any questions.", _
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
'END FUNCTIONS LIBRARY BLOCK ====================================================================================================