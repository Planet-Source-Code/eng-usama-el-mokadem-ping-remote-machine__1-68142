<div align="center">

## Ping remote machine

<img src="PIC2007315917229399.JPG">
</div>

### Description

Ping remote machine (TCP/IP)
 
### More Info
 
IP Address or Hostname

Ping remote machine let you know if this machine work or not

Statues. Fail or Success


<span>             |<span>
---                |---
**Submitted On**   |2007-03-15 09:20:02
**By**             |[Eng\. Usama El\-Mokadem](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/eng-usama-el-mokadem.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Jokes/ Humor](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/jokes-humor__1-40.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Ping\_remot2053913152007\.zip](https://github.com/Planet-Source-Code/eng-usama-el-mokadem-ping-remote-machine__1-68142/archive/master.zip)

### API Declarations

```
Const ICMP_SUCCESS As Long = 0
Const WS_VERSION_REQD As Long = &amp;H101
Const INADDR_NONE = &amp;HFFFF
Type WSADATA
  wVersion     As Integer
  wHighVersion   As Integer
  szDescription(0 To 256) As Byte
  szSystemStatus(0 To 128) As Byte
  iMaxSockets   As Long
  iMaxUDPDG    As Long
  lpVendorInfo   As Long
End Type
Type IP_OPTION_INFORMATION
  Ttl       As Byte
  Tos       As Byte
  Flags      As Byte
  OptionsSize   As Byte
  OptionsData   As Long
End Type
Type ICMP_ECHO_REPLY
  address     As Long
  Status      As Long
  RoundTripTime  As Long
  DataSize     As Long
  Reserved     As Integer
  ptrData     As Long
  Options     As IP_OPTION_INFORMATION
  Data       As String * 250
End Type
Type HOSTENT
  hName      As Long
  hAliases    As Long
  hAddrType    As Integer
  hLength     As Integer
  hAddrList    As Long
End Type
Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal cp As String) As Long
Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Declare Function IcmpSendEcho Lib "icmp.dll" _
  (ByVal IcmpHandle As Long, _
  ByVal DestinationAddress As Long, _
  ByVal RequestData As String, _
  ByVal RequestSize As Long, _
  ByVal RequestOptions As Long, _
  ReplyBuffer As ICMP_ECHO_REPLY, _
  ByVal ReplySize As Long, _
  ByVal Timeout As Long) As Long
Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
```





