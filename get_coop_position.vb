Sub getJD()
'
' getJD Macro
'
' Keyboard Shortcut: Ctrl+g
'
    Range("G4").Select
    
    Dim bot As New WebDriver
    bot.Start "chrome", "http://google.com"
    
    ' login and navigate to the 1st page
    Call startUp(bot)
    
    Dim pg As Integer
    'loop the page
    For pg = 1 To 29
    
        Call processThePage(bot, pg)
    
    Next pg
    
    'close the browser
    bot.Quit
    
End Sub
Sub startUp(bot As WebDriver)
    
    bot.Get "https://waterlooworks.uwaterloo.ca/home.htm"
    'click student
    bot.FindElementByXPath("/html/body/div[3]/div/div/div[5]/div/div/a[1]").Click
    
    'need to login or not
    If bot.FindElementByXPath("//*[@id='username']").IsDisplayed() Then
    
        bot.FindElementByXPath("//*[@id='username']").SendKeys ("UWATERLOOUSERNAME")
        bot.FindElementByXPath("//*[@id='password']").SendKeys ("UWATERLOOPASSWORD")
        bot.FindElementByXPath("//*[@id='cas-submit']/input").Click
       
    End If
    
    'click hire waterloo coop'
    bot.FindElementByXPath("//*[@id='closeNav']/div/ul/div/ul/li[2]/a").Click
    
    Application.Wait (Now + TimeValue("0:00:5"))
    'click my program'
    bot.FindElementByXPath("//*[@id='quickSearchCountsContainer']/table/tbody/tr[1]/td[2]/a").Click
                            
    'click not interested list
    'bot.FindElementByXPath("//*[@id='mainContentDiv']/div[2]/div/div/div/div[2]/div[3]/div[2]/div[3]/div[2]/div/div/a[2]").Click

End Sub

Sub processThePage(bot As WebDriver, pg As Integer)

    'click to go to the page
    s1 = "#postingsTablePlaceholder > div:nth-child(4) > div > ul > li:nth-child("
    s2 = ") > a"
    slt = s1 + CStr(pg + 2) + s2  'the selector
    
    s1 = "//*[@id='postingsTablePlaceholder']/div[1]/div/ul/li["
    s2 = "]/a"
    slt = s1 + CStr(pg + 2) + s2
    
    If pg <> 1 Then
        'bot.FindElementByCss(slt).Click
'        Application.Wait (Now + TimeValue("0:00:2"))
        bot.ExecuteScript ("window.scrollTo(0,0);") ' document.body.scrollHeight);")
        bot.Mouse.MoveTo bot.FindElementByXPath(slt)
        bot.FindElementByXPath(slt).Click
    End If
    
    Dim i As Integer
    'loop the item
    For i = 1 To 100
        Call handleTheItem(bot, pg, i)
    Next i
End Sub

Sub handleTheItem(bot As WebDriver, pg As Integer, i As Integer)
    'copy the info from list
    jid = bot.FindElementByXPath("//*[@id='postingsTable']/tbody/tr[" + CStr(i) + "]/td[3]").Text
    
    ts = "//*[@id='posting" + jid + "']/td[4]/span/a"
    tmp = bot.FindElementByXPath("//*[@id='posting" + jid + "']/td[4]").Attribute("innerHTML")

    If InStr(tmp, "<strong") > 0 Then
        ts = "//*[@id='posting" + jid + "']/td[4]/strong/span/a"
    End If
    
    If InStr(tmp, "span[2]") > 0 Then
        ts = "//*[@id='posting" + jid + "']/td[4]/span[2]/a"
    End If
    
    
    Title = bot.FindElementByXPath(ts).Text
    org = bot.FindElementByXPath("//*[@id='posting" + jid + "']/td[5]/span").Text
    division = bot.FindElementByXPath("//*[@id='posting" + jid + "']/td[6]/span").Text
    opening = bot.FindElementByXPath("//*[@id='posting" + jid + "']/td[7]").Text
    city = bot.FindElementByXPath("//*[@id='posting" + jid + "']/td[9]").Text
    Level = bot.FindElementByXPath("//*[@id='posting" + jid + "']/td[10]").Text
    apps = bot.FindElementByXPath("//*[@id='posting" + jid + "']/td[11]").Text
    deadline = bot.FindElementByXPath("//*[@id='posting" + jid + "']/td[12]").Text
    
    'click the item
    'If i Mod 15 = 0 Then
    '    bot.Mouse.MoveTo bot.FindElementByXPath(slt)
    'End If
    
    bot.FindElementByXPath(ts).Click
    bot.SwitchToNextWindow
    
    'go to the new tab for that item
    'workTerm = bot.FindElementByXPath("//*[@id='postingDiv']/div[1]/div[2]/table/tbody/tr[1]/td[2]").Text
    'Address = bot.FindElementByXPath("//*[@id='postingDiv']/div[1]/div[2]/table/tbody/tr[8]/td[2]").Text
    'province = bot.FindElementByXPath("//*[@id='postingDiv']/div[1]/div[2]/table/tbody/tr[9]/td[2]").Text
    'zip = bot.FindElementByXPath("//*[@id='postingDiv']/div[1]/div[2]/table/tbody/tr[10]/td[2]").Text
    'country = bot.FindElementByXPath("//*[@id='postingDiv']/div[1]/div[2]/table/tbody/tr[12]/td[2]").Text
    'term = bot.FindElementByXPath("//*[@id='postingDiv']/div[1]/div[2]/table/tbody/tr[13]/td[2]").Text
    'jobSummary = bot.FindElementByXPath("//*[@id='postingDiv']/div[1]/div[2]/table/tbody/tr[14]/td[2]").Text
    'responsibility = bot.FindElementByXPath("//*[@id='postingDiv']/div[1]/div[2]/table/tbody/tr[15]/td[2]").Text
    'skill = bot.FindElementByXPath("//*[@id='postingDiv']/div[1]/div[2]/table/tbody/tr[16]/td[2]").Text
    'transHouse = bot.FindElementByXPath("//*[@id='postingDiv']/div[1]/div[2]/table/tbody/tr[17]/td[2]").Text
    'compensation = bot.FindElementByXPath("//*[@id='postingDiv']/div[1]/div[2]/table/tbody/tr[18]/td[2]").Text
    'targetDegree = bot.FindElementByXPath("//*[@id='postingDiv']/div[1]/div[2]/table/tbody/tr[19]/td[2]").Text
    
    'get all the info
    
        'get # of rows
        tblHTML = bot.FindElementByXPath("//*[@id='postingDiv']/div[1]/div[2]/table").Attribute("innerHTML")
        tgs = Split(tblHTML, "<tr>")
        rn = UBound(tgs) - LBound(tgs)
        
        Dim info(12, 2) As String
        
        Call GetInfo(bot, rn, info)
        
    'print to pdf
    'bot.FindElementByXPath("//*[@id='mainContentDiv']/div[2]/div/a[3]").Click
    
    ActiveWorkbook.Sheets("post").Select
    
    'copy the content
    theRow = (pg - 1) * 100 + i + 1
    Cells(theRow, 1).Value2 = jid
    Cells(theRow, 2).Value2 = Title
    Cells(theRow, 3).Value2 = org
    Cells(theRow, 4).Value2 = division
    Cells(theRow, 5).Value2 = opening
    Cells(theRow, 6).Value2 = city
    Cells(theRow, 7).Value2 = Level
    Cells(theRow, 8).Value2 = apps
    Cells(theRow, 9).Value2 = info(0, 1) 'workTerm
    Cells(theRow, 10).Value2 = info(5, 1) 'term
    Cells(theRow, 11).Value2 = info(6, 1) 'jobSummary
    Cells(theRow, 12).Value2 = info(7, 1) 'responsibility
    Cells(theRow, 13).Value2 = info(8, 1) 'skill
    Cells(theRow, 14).Value2 = info(9, 1) 'transHouse
    Cells(theRow, 15).Value2 = info(10, 1) 'compensation
    Cells(theRow, 16).Value2 = info(1, 1) 'Address
    Cells(theRow, 17).Value2 = info(2, 1) 'province
    Cells(theRow, 18).Value2 = info(3, 1) 'zip
    Cells(theRow, 19).Value2 = info(4, 1) 'country
    Cells(theRow, 20).Value2 = deadline
    Cells(theRow, 21).Value2 = info(11, 1) 'targetDegree
     
    'close this tab
     bot.SwitchToPreviousWindow
     bot.SwitchToPreviousWindow.Close
     bot.SwitchToPreviousWindow

    'go to the previous page tab
End Sub

Sub GetInfo(bot, rn, info)
    info(0, 0) = "Work Term:"
    info(1, 0) = "Job - Address Line One:"
    info(2, 0) = "Job - Province / State:"
    info(3, 0) = "Job - Postal Code / Zip Code (X#X #X#):"
    info(4, 0) = "Job - Country:"
    info(5, 0) = "Work Term Duration:"
    info(6, 0) = "Job Summary:"
    info(7, 0) = "Job Responsibilities:"
    info(8, 0) = "Required Skills:"
    info(9, 0) = "Transportation and Housing:"
    info(10, 0) = "Compensation and Benefits Information:"
    info(11, 0) = "Targeted Degrees and Disciplines:"
    
    
    i = 1
    Do While i < rn
        keyPath = "//*[@id='postingDiv']/div[1]/div[2]/table/tbody/tr[" + CStr(i) + "]/td[1]"
        valPath = "//*[@id='postingDiv']/div[1]/div[2]/table/tbody/tr[" + CStr(i) + "]/td[2]"
        Key = Trim(bot.FindElementByXPath(keyPath).Text)
        Vl = Trim(bot.FindElementByXPath(valPath).Text)
        
        j = 0
        Do While j >= 0 And j < 12
            If Key = info(j, 0) Then
                info(j, 1) = Vl
                j = -1
            Else
                j = j + 1
            End If
        Loop
    
        'check extra tr
        xtr = bot.FindElementByXPath(valPath).Attribute("innerHTML")
        xtrarray = Split(xtr, "<tr")
        n = UBound(xtrarray) - LBound(xtrarray)
        
        i = i + 1
        rn = rn - n

Loop
    
End Sub