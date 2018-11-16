phone_number = InputBox("What is the phone number?")

phone_number = trim(phone_number)
If Instr(phone_number, "-") = 0 Then
    If len(phone_number) = 7 Then
        phone_number = left(phone_number, 3) & "-" & right(phone_number, 4)
    ElseIf len(phone_number) = 10 Then
        phone_number = left(phone_number, 3) & "-" & mid(phone_number, 4, 3) & "-" & right(phone_number, 4)
    End If
End If

MsgBox phone_number
