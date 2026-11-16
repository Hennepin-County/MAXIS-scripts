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

' --- Function to display any array ---
Sub DisplayArray(arr, title)
    Dim i

    ' Validate input: Check if it's an array
    If IsArray(arr) = False Then
        WScript.Echo "Error: Provided variable is not an array."
        Exit Sub
    End If

    ' Optional title
    If Len(title) > 0 Then
        WScript.Echo title
    End If

    ' Handle empty arrays
    If UBound(arr) < 0 Then
        WScript.Echo "(Array is empty)"
        Exit Sub
    End If

    ' Loop through array and display each element
    For i = LBound(arr) To UBound(arr)
        WScript.Echo "Index " & i & ": " & arr(i)
    Next
End Sub

' --- Example usage ---

' Fixed-size array
Dim fruits
fruits = Array("Apple", "Banana", "Cherry")
Call DisplayArray(fruits, "Fruits Array:")

' Dynamic array
Dim colors()
ReDim colors(2)
colors(0) = "Red"
colors(1) = "Green"
colors(2) = "Blue"
Call DisplayArray(colors, "Colors Array:")

' Invalid input test
Dim notArray
notArray = "Hello"
Call DisplayArray(notArray, "Test with non-array:")






' Procedure to display a one-dimensional array
Sub DisplayArray1D(arr)
    Dim i
    WScript.Echo "One-Dimensional Array:"
    For i = LBound(arr) To UBound(arr)
        WScript.Echo "Index " & i & ": " & arr(i)
    Next
End Sub

' Procedure to display a two-dimensional array
Sub DisplayArray2D(arr)
    Dim row, col
    WScript.Echo "Two-Dimensional Array:"
    For row = LBound(arr, 1) To UBound(arr, 1)
        For col = LBound(arr, 2) To UBound(arr, 2)
            WScript.Echo "Row " & row & ", Col " & col & ": " & arr(row, col)
        Next
    Next
End Sub