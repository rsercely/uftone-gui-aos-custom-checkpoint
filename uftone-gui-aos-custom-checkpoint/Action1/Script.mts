Const minAcceptableScore  = 9.0
Browser("Advantage Shopping").Navigate "http://advantageonlineshopping.com"
Browser("Advantage Shopping").Page("Advantage Shopping").Link("HeadphonesCategoryTxt").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").Link("HeadphonesCategoryTxt")_;_script infofile_;_ZIP::ssf1.xml_;_
if Browser("Advantage Shopping").Page("Advantage Shopping").Image("fetchImage?image_id=2200").Exist then @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").Image("fetchImage?image id2200")_;_script infofile_;_ZIP::ssf2.xml_;_
	wait 1
	Browser("Advantage Shopping").Page("Advantage Shopping").Image("fetchImage?image_id=2200").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").Image("fetchImage?image id2200")_;_script infofile_;_ZIP::ssf2.xml_;_
end if
score = Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Product Score").GetROProperty ("innerText")
score = cDbl(score)
If score >= minAcceptableScore  Then
	reporter.ReportEvent micPass, "check rating", "Rating exceeded expectations. It was " & score & " and needed to exceed " & minAcceptableScore
else
	reporter.ReportEvent micFail , "check rating", "Rating is poor. It was " & score & " and needed to exceed " & minAcceptableScore
End If
'Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Product Score").Check CheckPoint("9.3") @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("9.3")_;_script infofile_;_ZIP::ssf3.xml_;_
Browser("Advantage Shopping").Page("Advantage Shopping").Link("HOME").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").Link("HOME")_;_script infofile_;_ZIP::ssf4.xml_;_
