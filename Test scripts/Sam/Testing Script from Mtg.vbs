

Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName
user_ID_for_validation = ucase(windows_user_ID)


MsgBox "Being run by: " & user_ID_for_validation, vbInformation, "User Validation"