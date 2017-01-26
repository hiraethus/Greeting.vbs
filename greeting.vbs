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

Set xDoc = CreateObject("Microsoft.XMLDOM")
xDoc.async = False

If xDoc.Load("http://www.biblegateway.com/votd/get/?format=xml&version=ESV") Then
	Dim sXPath: sXPath = "/votd/content"
	Set quotation = xDoc.selectSingleNode (sXPath)
	speech.Speak Trim(quotation.Text)

	Dim sXPath2 : sXPath2 = "/votd/reference"
	Set reference = xDoc.selectSingleNode(sXPath2)
	speech.Speak reference.Text
End If

Set xDoc = Nothing
