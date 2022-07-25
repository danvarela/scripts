'This Script will ping a list of hosts and present the results in an auto launched Web Page
' and a .CSV file for Excel.  The output file names Ping-Results.htm and Ping-Results.csv in 
' folder that the script is running in.
' Users do need to copy the files that they need since they may be overwriten
'
' This script was written by Dan Long (danlong27@gmail.com)
'
dim fso, today, oshell
cr=VBCrLf
q=chr(34)  'Quote charcter, used in text strings
c=chr(44)  'Comma charcter, used in text strings
'WebFile="c:\temp\Ping-Results.htm"  'This is the HTML results file
'DatFile="c:\temp\Ping-Results.csv"  'This is the CSV results file
WebFile="Ping-Results.htm"  'This is the HTML results file
DatFile="Ping-Results.csv"  'This is the CSV results file
PingList=""

'This section checks for Arguments on the command line.
Set objArgs = WScript.Arguments

'This section checks to see if there are any arguments.
ArgCount=objargs.count
if ArgCount=0 then  'checks for no argument and sets the variable to no
	PingList=Inputbox("What is the full path to the file with the host list","Pinger")
	If PingList="" then
		Wscript.echo "You need to have a file with a list of hosts to Ping" & VBCrLf & "You can use the format Cscript Pinger??.vbs hostlist.txt" & VBCrLf & "   (the ?? refers to the version of Pinger)" & VBCrLf & "Or you can drag hostlist.txt and drop it on Pinger??.vbs"
		Wscript.quit
	End if
    'x=msgbox("Use"&cr&cr&"PINGER.VBS hostlist.txt"&cr&cr& _
    '"hostlist.txt is a text file with a server name on each line followed by a CR LF"& _
    'cr&"If there are spaces in the name the name must be in quotes"&cr&cr& _
    '"The Results are in C:\Temp\Ping-Results.htm and Ping-Results.csv", _
    '0,"Pinger")
    'Wscript.quit
end if
'wscript.echo argcount
'wscript.quit

'If ArgCount>1 then
'  x=msgbox("Use"&cr&cr&"PINGER.VBS hostlist.txt"&cr&cr& _
'    "hostlist.txt is a text file with a server name on each line followed by a CR LF"& _
'    cr&"If there are spaces in the name the name must be in quotes"&cr&cr& _
'    "The Results are in C:\Temp\Ping-Results.htm and Ping-Results.csv", _
'    0,"Pinger")
'    Wscript.quit
'end if

If Pinglist="" then
	PingList=ObjArgs(0)
End if

'wscript.echo pinglist

today1=date
today=cstr(today1)
Now1=time
rnow=cstr(now1)
cr=VBCrLf

'wscript.echo today
set fso=createobject("Scripting.FileSystemObject")
'PingList=inputbox("Path to the list of servers to Ping?"&cr&cr& _
'  "This is a text file that has a server name or IP address on seperate lines"&cr&cr& _
'  "The Results are in C:\Temp\Ping-Results.htm","Pinger","C:\")

If PingList="" then Wscript.quit
set file1=fso.opentextfile(PingList,1)
set wshshell=wscript.createobject("wscript.shell")
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
'msgbox(scriptdir)

On Error Resume Next
err.clear
set file2=fso.createtextfile(scriptdir & "\" & WebFile,true)
'msgbox "Web File" & vbcrlf & "Error number " & err.number
if err.number <> 0 then
	Msgbox "the Web file is Locked and can't be recreated.  Ending Script"
	Wscript.quit
end if
err.clear
Set file3=fso.createtextfile(scriptdir & "\" & DatFile,true)
'msgbox "csv file" & vbcrlf & "Error number " & err.number
if err.number <> 0 then
	msgbox "The .CSV file is locked and can't be recreated.  Ending Script"
	Wscript.quit
end if
On error goto 0

file2.write "<HTML>" & VBCrLf
file2.write "<TITLE>Whats Up Now</Title>"  & VBCrLf
file2.write "<HEAD>" & VBCrLf
file2.write "<B>The List of servers you wanted to check as of " & today & " at " & rnow &"</B>" & "<P>"
file2.write "The results can be found in " & scriptdir & "\" &  WebFile & " and " & scriptdir & "\" &  DatFile & "<P>"
'file2.write "<BR>"  & VBCrLf
file2.write "<style>"  & VBCrLf
file2.write "table, th, td {"  & VBCrLf
file2.write "    border: 1px solid black;"  & VBCrLf
file2.write "    border-collapse: collapse;"  & VBCrLf
file2.write "}" & VBCrLf
file2.write "th {"  & VBCrLf
file2.write "    Text-align: Left;" & VBCrLf
file2.write "    padding:1px 15px 1px 2px;" & VBCrLf
file2.write "}"  & VBCrLf
file2.write "</style>"  & VBCrLf
file2.write "</HEAD>" & VBCrLf
'file2.write "<BR>" & VBCrLf
file2.write "<BR>" & VBCrLf
file2.write "<BODY>" & VBCrLf
file2.write "<table style=" & q & "auto" & q & ">" & VBCrLf
file2.write "<tr>"  & VBCrLf
'file2.write "   <th width="& q & "60%"& q & "> Host Name</th>"  & VBCrLf
'file2.write "   <th width="& q & "30%"& q & "> IP Address</th>"  & VBCrLf
'file2.write "   <th width="& q & "10%"& q & "> Status</th>"  & VBCrLf
file2.write "   <th>Pinged Name</th>"  & VBCrLf
'file2.write "   <th>Replied Name</th>" & VBCrLf
file2.write "   <th>IP Address</th>"  & VBCrLf
file2.write "   <th>Response Time</th>" & VBCrLf
file2.write "   <th> Time To Live</th>" & VBCrLf
file2.write "   <th>Status</th>"  & VBCrLf

file2.write "</tr>"  & VBCrLf

file3.write q&pinglist&q&c&q&" was used "&q&c&q&" to generate this file on " & today& " at " & rnow &q&cr
file3.write q&"Pinged Name"&q&c&q&"Status"&q&c&q&"IP Address"&q&c&q&"Response time"&q&c&q&"Time to Live"&q&cr

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

do until file1.atendofstream = true
	mychar=file1.readline
  if mychar = "" then
    file2.write "<BR>" & VBCrLf
    mychar=file1.readline
  end if

'mychar="127.0.0.1"

	Set colPings = objWMIService.ExecQuery _
	 ("Select * From Win32_PingStatus where Address = '" & MyChar & "'")
	If Err.number = 0 Then
	  Err.Clear
	  For Each objPing in colPings
		If Err = 0 Then
		  Err.Clear
		  If objPing.StatusCode = 0 Then
			Stat="! Up"
			IP = objPing.ProtocolAddress
			RespName = objPing.ProtocolAddressResolved
			'msgbox(respname)
			'msgbox "Bytes Sent: " & objPing.BufferSize
			RespTime = objPing.ResponseTime & " ms"
			TTL = objPing.ResponseTimeToLive & " seconds"
		  Else
			Stat = MyChar & " did not respond to ping."
			RepTime = objPing.StatusCode
		  End If
		Else
		  Err.Clear
		  WScript.Echo "Unable to call Win32_PingStatus on " & strComputer & "."      
		End If
	  Next
	Else
	  Err.Clear
	  WScript.Echo "Unable to call Win32_PingStatus on " & strComputer & "."
	End If






'	set oexec=wshshell.exec("ping -n 1 -w 300 " & Mychar)
'
'	Stat=""
'	IP=""
'	do until oexec.stdout.atendofstream
'   MyChar1=oexec.stdout.readline()
'   if instr(1,MyChar1,"[",1) > 0 then
'      x=instr(1,MyChar1,"[",1)
'      y=instr(1,MyChar1,"]",1)
'      ip=mid(mychar1,x+1,y-x-1)
'    end if
'		if instr(1,MyChar1,"Reply from",1) > 0 then stat="up"
''      y=instr(1,mychar1,":",1)
''      ip=mid(mychar1,11,y-11)
'    end if

		'output=output & oexec.stdout.readline()
	'loop
	if stat="! Up" then
		file2.write "<tr>"  & VBCrLf
		file2.write "   <th>" & Mychar & "</th>"  & VBCrLf
		'file2.write "   <th>" & RespName & "</th>" & VBCrLf
		file2.write "   <th>" & ip & "</th>"  & VBCrLf
		file2.write "   <th>" & RespTime & "</th>" & VBCrLf
		file2.write "   <th>" & TTL & "</th>" & VBCrLf
		file2.write "   <th>! Up</th>"  & VBCrLf
		file2.write "</tr>"  & VBCrLf
		'file2.write mychar & " at "& ip & " is <B>up</B> " & "<P>"
		file3.write q & mychar &q& c &q& "! Up" &q& c &q& ip &q& c &q& RespTime &q& c &q& TTL &q& cr
	else
		
		file2.write "<tr>"  & VBCrLf
		file2.write "   <th>" & Mychar & "</th>"  & VBCrLf
		'file2.write "   <th>" & RespName & "</th>" & VBCrLf
		file2.write "   <th>" & ip & "</th>"  & VBCrLf
		file2.write "   <th>" & RespTime & "</th>" & VBCrLf
		file2.write "   <th>" & TTL & "</th>" & VBCrLf
		file2.write "   <th>" & stat & "</th>"  & VBCrLf
		file2.write "</tr>"  & VBCrLf
		'file2.write mychar & " at " & ip & " is <B>down</B>" & "<P>"
		file3.write q& mychar &q&c&q& stat &q&c&q& ip &q& c &q& RespTime &q& c &q& TTL &q&cr
	end if
	Mychar=""
	RespName=""
	IP=""
	RespTime=""
	TTL=""
	Stat=""
	
loop


file2.write "</table>" & VBCrLf
file2.write "</BODY>" & VBCrLf
file2.write "</HTML>" & VBCrLf

set oshell=wscript.createobject("wscript.shell")
'oshell.run "iexplore " & scriptdir & "\" & WebFile
oshell.run "msedge.exe " & chr(34) & scriptdir & "\" & WebFile & chr(34)

' Ping command status codes.

' Success (0)
' Buffer Too Small (11001)
' Destination Net Unreachable (11002)
' Destination Host Unreachable (11003)
' Destination Protocol Unreachable (11004)
' Destination Port Unreachable (11005)
' No Resources (11006)
' Bad Option (11007)
' Hardware Error (11008)
' Packet Too Big (11009)
' Request Timed Out (11010)
' Bad Request (11011)
' Bad Route (11012)
' TimeToLive Expired Transit (11013)
' TimeToLive Expired Reassembly (11014)
' Parameter Problem (11015)
' Source Quench (11016)
' Option Too Big (11017)
' Bad Destination (11018)
' Negotiating IPSEC (11032)
' General Failure (11050)
