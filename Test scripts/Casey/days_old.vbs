EMConnect ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 141, 100, "Dialog"
  EditBox 40, 30, 90, 15, child_name
  EditBox 45, 50, 85, 15, child_b_day
  ButtonGroup ButtonPressed
    OkButton 80, 75, 50, 15
  Text 40, 10, 65, 10, "How old are you?"
  Text 10, 35, 25, 10, "Name:"
  Text 10, 55, 35, 10, "Birthday:"
EndDialog

Do
    child_name = ""
    child_b_day = ""

    dialog Dialog1

    days_old = DateDiff("d", child_b_day, date)
    weeks_old = DateDiff("w", child_b_day, date)
    months_old = DateDiff("m", child_b_day, date)
    years_old = DateDiff("yyyy", child_b_day, date)

    review_how_old_msg = MsgBox("*** " & child_name & " ***" & vbCr & vbCr & "Is " & days_old & " days old." & vbCr & "Is " & weeks_old & " weeks old." & vbCr & "Is " & months_old & " months old." & vbCr & "Is " & years_old & " years old.", vbOKCancel)

Loop until review_how_old_msg = vbCancel
