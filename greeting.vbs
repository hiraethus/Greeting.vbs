Dim greeting, speech, name, objNetwork

'get user's name
Set objNetwork = CreateObject("WScript.Network")
name=objNetwork.UserName

Set speech=CreateObject("sapi.spvoice")

If Hour(Time) < 12 Then
	greeting="Good Morning, " & name
Else 
	If Hour(Time) >= 12 And Hour(Time) < 18 Then
		greeting = "Good Afternoon, " & name
	Else
		greeting = "Good Evening " & name
	End If
End If
speech.Speak greeting

day_of_week=WeekdayName(Weekday(Date))
day_in_month=Day(Date)
month_name=MonthName(Month(Date))
the_year=Year(Date)
greeting2="Today is " & day_of_week & ", the " & day_in_month & " of " & month_name & " " & the_year
speech.Speak greeting2

'download verse of the day page
'modified from http://markalexanderbain.suite101.com/how-to-use-vbscript-to-download-a-web-page-a89661
dim xmlhttp : set xmlhttp = createobject("msxml2.xmlhttp.3.0")
xmlhttp.open "get", "http://www.biblegateway.com/votd/get/?format=html&version=ESV", false
xmlhttp.send

'Prepare a regular expression object
Set myRegExp = New RegExp
dim quote_text
myRegExp.IgnoreCase = True
'myRegExp.Global = True
myRegExp.Pattern = "&ldquo;(.*)&rdquo;"
quote_text = myRegExp.Replace(xmlhttp.responseText, "$1")
WScript.Echo quote_text
speech.Speak quote_text