HOLIDAYS_ARRAY = Array(#9/3/18#, #11/12/18#, #11/22/18#, #11/23/18#, #12/24/18#, #12/25/18#, #1/1/19#, #1/21/19#, #2/18/19#, #5/27/19#, #7/4/19#)

function is_date_holiday_or_weekend(date_to_review, boolean_variable)
    non_working_day = FALSE
    day_of_week = WeekdayName(WeekDay(date_to_review))
    If day_of_week = "Saturday" OR day_of_week = "Sunday" Then non_working_day = TRUE
    For each holiday in HOLIDAYS_ARRAY
        If holiday = date_to_review Then non_working_day = TRUE
    Next
    boolean_variable = non_working_day
end function

function change_date_to_soonest_working_day(date_to_change)
    Do
        is_holiday = FALSE
        For each holiday in HOLIDAYS_ARRAY
            If holiday = date_to_change Then
                is_holiday = TRUE
                date_to_change = DateAdd("d", -1, date_to_change)
            End If
        Next
        If WeekdayName(WeekDay(date_to_change)) = "Saturday" Then date_to_change = DateAdd("d", -1, date_to_change)
        If WeekdayName(WeekDay(date_to_change)) = "Sunday" Then date_to_change = DateAdd("d", -2, date_to_change)
    Loop until is_holiday = FALSE
end function

the_day = #7/25/18#

CALL is_date_holiday_or_weekend(the_day, yes_or_no)
Call change_date_to_soonest_working_day(the_day)

MsgBox "The day is a non working day - " & yes_or_no & vbNewLine & "New date - " & the_day & " - " & WeekdayName(WeekDay(the_day))
