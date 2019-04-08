EMConnect ""

BeginDialog Dialog1, 0, 0, 326, 215, "Dialog"
  ButtonGroup ButtonPressed
    OkButton 210, 195, 50, 15
    CancelButton 265, 195, 50, 15
    PushButton 15, 25, 90, 75, "BUTTON 1", button_one_btn
    PushButton 120, 25, 90, 75, "BUTTON 2", button_two_btn
    PushButton 225, 25, 90, 75, "BUTTON 3", button_three_btn
    PushButton 15, 110, 90, 75, "BUTTON 4", button_four_btn
    PushButton 120, 110, 90, 75, "BUTTON 5", button_five_btn
    PushButton 225, 110, 90, 75, "BUTTON 6", button_six_btn
  Text 130, 10, 55, 10, "PICK A BUTTON"
EndDialog

Do
    dialog Dialog1
    If buttonpressed = 0 Then stopscript

    'flower'
    If ButtonPressed = button_one_btn Then the_web_page = "<img src='https://images.unsplash.com/photo-1468444326310-637139573e31?ixlib=rb-1.2.1&ixid=eyJhcHBfaWQiOjEyMDd9&auto=format&fit=crop&w=1000&q=80' height = 400 width = 500>"
    'mushroom'
    If ButtonPressed = button_two_btn Then the_web_page = "<img src='http://blogs.discovermagazine.com/crux/files/2018/10/amanita-muscaria.jpg' height = 400 width = 600>"
    'tree'
    If ButtonPressed = button_three_btn Then the_web_page = "<img src='https://www.naturehills.com/media/catalog/product/cache/74c1057f7991b4edb2bc7bdaa94de933/o/a/oak-tree-full-425x425.jpg' height = 500 width = 400>"
    'duck'
    If ButtonPressed = button_four_btn Then the_web_page = "<img src='https://www.purelypoultry.com/images/pekin-ducklings_01.jpg' height = 500 width = 400>"
    'squirrel'
    If ButtonPressed = button_five_btn Then the_web_page = "<img src='https://upload.wikimedia.org/wikipedia/commons/1/1c/Squirrel_posing.jpg' height = 400 width = 500>"
    'owl'
    If ButtonPressed = button_six_btn Then the_web_page = "<img src='http://wildlife.ohiodnr.gov/portals/wildlife/Species%20and%20Habitats/Species%20Guide%20Index/Images/greathornedowl.jpg' height = 600 width = 400>"

    Set objExplorer = CreateObject("InternetExplorer.Application")

    With objExplorer
        .Navigate "about:blank"
        .ToolBar = 0
        .StatusBar = 0
        .Left = 100
        .Top = 100
        .Width = 1300
        .Height = 1300
        .Document.Title = "Important image!"
    End With

    Do While objExplorer.Busy
        WScript.Sleep 200
    Loop

    objExplorer.Document.Body.InnerHTML = the_web_page
    objExplorer.Document.focus()
    objExplorer.Visible = TRUE

    ready_to_leave_message = MsgBox("Look at this Picture!" & vbCr & vbCr & "Are you done?", vbQuestion + vbYesNo, "Like this one?")
    objExplorer.Quit
Loop until ready_to_leave_message = vbYes
