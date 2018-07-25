'sort dates script

dates_array = array(#12/23/2013#, #5/3/2010#, #6/28/2018#, #2/17/2015#)

dim ordered_dates ()
redim ordered_dates(0)

days =  0
'for each date_in_order in ordered_dates
do
    redim preserve ordered_dates(days)
    prev_date = ""
    for each thing in dates_array
        check_this_date = TRUE
        'all_dates = join(ordered_dates, "~")
        ' MsgBox all_dates
        ' thing_string = thing & ""
        ' MsgBox "in string is " & instr(thing_string, all_dates)
        ' MsgBox thing_string
        ' if all_dates = "" OR instr(thing_string, all_dates) = 0 Then
        For each known_date in ordered_dates
            if known_date = thing Then check_this_date = FALSE
            MsgBox "known dates is " & known_date & vbNewLine & "thing is " & thing & vbNewLine & "match - " & check_this_date
        next
        if check_this_date = TRUE Then
            if prev_date = "" Then
                prev_date = thing
            Else
                if DateDiff("d", prev_date, thing) <0 then
                    prev_date = thing
                end if
            end if
        end if
    next
    'MsgBox prev_date
    ordered_dates(days) = prev_date
    'all_dates = join(ordered_dates, "~")
    days = days + 1
'next
loop until days > UBOUND(dates_array)


for each thing in ordered_dates
    MsgBox thing
next
