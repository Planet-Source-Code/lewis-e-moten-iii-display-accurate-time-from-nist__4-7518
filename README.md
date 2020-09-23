<div align="center">

## Display Accurate Time from NIST


</div>

### Description

Display accurate time to your visitors! This script requests the UTC time from a server that is syncronized with an atomic clock. It then adjusts the time according to timezone and if daylight savings time is in effect.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis E\. Moten III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md)
**Level**          |Beginner
**User Rating**    |5.0 (35 globes from 7 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[System Services/ Functions](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/system-services-functions__4-23.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-e-moten-iii-display-accurate-time-from-nist__4-7518/archive/master.zip)





### Source Code

```
<%
' Author: Lewis Moten
' Email: lewis@moten.com
' URL: http://www.lewismoten.com
' Not a requirement, but I do suggest that you link to, visit
' or promote my website by referring it to others in newsgroups,
' email, co-workers, etc. If this code appears on a website
' that allows interaction, please submit votes, comments, reviews,
' etc.
'	List of time servers:
'	http://boulder.nist.gov/timefreq/service/time-servers.html
' Server used for this demonstration
' time.nist.gov
' Format returned by server - Daytime Protocol (RFC-867)
' JJJJJ YR-MO-DA HH:MM:SS TT L H msADV UST(NIST) OTM
' Example:
' 52399 02-05-05 02:11:02 50 0 0 49.4 UTC(NIST) *
' JJJJJ
'	Modified Julian Date (MJD). Last 5 digits of julian
'	date (count of days since January 1, 4713 B.C.). Add
'	2.4 Million to get actual Julian Date.
' YR-MO-DA
'	Year, Month, Day
' HH:MM:SS
'	Hours, Minutes, Seconds (in Coordinated Universal Time - UTC)
' TT
'	00 - Standard Time
'	50 - Daylight Savings Time
'	1 to 49 - Days within current month until daylight
'		savings time adjustement approaches
' L
'	Leap seconds keep UTC time adjusted with earths
'	rotation (every 1 to 1 1/2 years)
'	0 - leap second will not occur this month.
'	1 - 61 seconds will appear in last minute of month
'	2 - 59 seconds will appear in last minute of month
' H
'	0 - server is healthy
'	1 - time may be in error by up to 5 seconds
'	2 - time is known to be wrong by more then 5 seconds
'	4 - Hardware/Software failure - amount of time error is unknown.
' msADV
'	milliseconds NIST has advanced time to compensate for
'	network delays.
' UTC(NIST)
'	Signifies UTC time comes from National Institute of
'	Standards and Technology.
' OTM
'	On-Time marker. Signifies that the arrival time of the
'	time recieved from server should be accurate.
' ################################################################
' Begin Code
' ################################################################
' Server to query time from
Const TimeServer = "http://time.nist.gov:13"
' Define your timezone here
Const TimeZoneOffset = -5
' Set to true or false if you observe daylight savings time
Const DST = True
' Use XML HTTP object to request web page content
'Set Spider = Server.CreateObject ("MSXML2.XMLHTTP.3.0")
'Set Spider = Server.CreateObject ("MSXML2.ServerXMLHTTP")
Set Spider = Server.CreateObject ("Microsoft.XMLHTTP")
Spider.Open "GET", TimeServer, False, "", ""
Spider.Send
NIST = Spider.ResponseText
Set Spider = Nothing
' Parse UTC date
UTC = Mid(NIST, 11, 2) & "/" & Mid(NIST, 14, 2) & "/" & Mid(NIST, 8, 2) & " " & Mid(NIST, 16, 9)
' Is daylight savings in effect?
IsDaylightSavings = CInt(Mid(NIST, 26, 2)) = 50 Or (Month(UTC) > 6 AND CInt(Mid(NIST, 26, 2)) > 0)
' Create the local time
LocalTime = DateAdd("h", TimeZoneOffset, UTC)
' Modify for daylight savings
If DST Then
	If IsDaylightSavings Then LocalTime = DateAdd("h", 1, LocalTime)
End If
' Write out the results
Response.Write "Time Server URL: " & TimeServer & "<BR>"
Response.Write "Server Responded with: " & NIST & "<BR>"
Response.Write "Parsed UTC Date: " & UTC & "<BR>"
Response.Write "Your timezone offset: " & TimeZoneOffset & "<BR>"
Response.Write "You observe daylight savings time: " & DST & "<BR>"
Response.Write "Your local time is: " & LocalTime & "<BR>"
%>
```

