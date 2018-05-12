'Function create_outlook_task()
    Set objOutlook = CreateObject("Outlook.Application")
    Set objTask = objOutlook.CreateItem(olTaskItem)
    
    objTask.Subject = "Test"
    objTask.Body = "testing task creation."
    objTask.DueDate = #01/31/2017 10:00 AM#
    objTask.ReminderSet = True
    objTask.ReminderTime = #01/31/2017 09:00 AM#
    objTask.Save
'End Function