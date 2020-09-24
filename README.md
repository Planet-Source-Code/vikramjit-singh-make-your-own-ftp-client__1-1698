<div align="center">

## Make Your Own FTP Client


</div>

### Description

Make a simple FTP Client that allows you to read and write to a remote computer
 
### More Info
 
The IP address and local file name , path.

Set Project Refrences to MSINET.OCX before you run the code

Or you could set Project components and check on Microsoft Internet transfer control...then drag the MSINET control onto the form.In that case comment the line

'Dim Inet1 As New InetCtlsObjects.Inet

'

----

NOTE

----

' This code runs fine on a local intranet... for ALL versions of VB.

' This code has also been tested by me to work on the INTERNET for VB5

' (SP3). if you have VB5 PLEASE upgrade to SP3...to resolve known

' bugs in INET. The code will then run like a breeze. VB 5 SP3 is FREE

' at http://www.microsoft.com/msdownload/vstudio/sp97/vb.asp


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Vikramjit Singh](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vikramjit-singh.md)
**Level**          |Unknown
**User Rating**    |5.9 (616 globes from 105 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vikramjit-singh-make-your-own-ftp-client__1-1698/archive/master.zip)





### Source Code

```
' Dim Inet1 As New InetCtlsObjects.Inet
Dim FTPHostname As String
Dim Response As String
Public Sub writefile(pathname As String, filename As String, IPaddress As String)
'note ..your ip addres specified should be that of an anonymous FTP Server.
'otherwise use ftp://ftp.microsoft.com kind of syntax
 FTPLogin
 FTPHostname = IPaddress
 Inet1.Execute FTPHostname, "PUT " & pathname & filename & " /" & filename
 Do While Inet1.StillExecuting
 DoEvents
 Loop
 Exit Sub
End Sub
Public Sub getfile(pathname As String, filename As String, IPaddress As String)
'note ..your ip addres specified should be that of an anonymous FTP Server.
'otherwise use ftp://ftp.microsoft.com kind of syntax
 FTPLogin
 FTPHostname = IPaddress
 Inet1.Execute FTPHostname, "GET " & filename & " " & pathname & filename
 Do While Inet1.StillExecuting
 DoEvents
 Loop
 Exit Sub
End Sub
Private Sub FTPLogin()
With Inet1
.Password = "Pass"
.UserName = "Anonymous"
End With
End Sub
```

