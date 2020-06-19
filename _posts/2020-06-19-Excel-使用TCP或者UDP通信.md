---
layout:     post
title:      在VBA中使用TCP或者UDP进行通信
subtitle:   Excel 工具，在VBA中使用TCP或者UDP向别的进程发送数据。
date:       2020-06-19
author:     zl
header-img: img/post_excel_wise.webp
catalog: true
tags:
    - VBA
    - excel
    - TCP
    - UDP
---
> VBA中通信总结
#  VBA中使用tcp/udp总结

​	数次的使用winsock的失败，甚至让我对vba失去了信心。我想要VBA中使用tcp或者udp通信方式将excel中的数据传输到其他的进程，为了达到这个目的，我尝试过直接使用winsock，也尝试使用OSWINSCK，vbRichClient。前者也是在winsock上进行的封装，我也不知道后者是怎么做的，但是其参考资料特别少，而且往往不能使用。

​	还好VBA有着很好的COM互操作库，对照着 [微软win32app](http://msdn.microsoft.com/en-us/library/windows/desktop/ms740673(v=vs.85).aspx) 总算将C++版本的TCP翻译为VBA版本。主要使用ws2_32.dll函数。

首先是UDP版本：

```vb
Const IP = "127.0.0.1"
Const Port = 6666

Const INVALID_SOCKET = -1
Const WSADESCRIPTION_LEN = 256
Private Const SOCKET_ERROR As Long = -1 'const #define SOCKET_ERROR            (-1)

Enum AF
  AF_UNSPEC = 0
  AF_INET = 2
  AF_IPX = 6
  AF_APPLETALK = 16
  AF_NETBIOS = 17
  AF_INET6 = 23
  AF_IRDA = 26
  AF_BTH = 32
End Enum

Enum sock_type
   SOCK_STREAM = 1
   SOCK_DGRAM = 2
   SOCK_RAW = 3
   SOCK_RDM = 4
   SOCK_SEQPACKET = 5
End Enum

Enum Protocol
   IPPROTO_ICMP = 1
   IPPROTO_IGMP = 2
   BTHPROTO_RFCOMM = 3
   IPPROTO_TCP = 6
   IPPROTO_UDP = 17
   IPPROTO_ICMPV6 = 58
   IPPROTO_RM = 113
End Enum

'Type sockaddr
'   sa_family As Integer
'   sa_data(0 To 13) As Byte
'End Type

Private Type sockaddr_in
  sin_family As Integer
  sin_port(0 To 1) As Byte
  sin_addr(0 To 3) As Byte
  sin_zero(0 To 7) As Byte
End Type

Private Type socket
   pointer As Long
End Type

Private Type LPWSADATA_Type
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To WSADESCRIPTION_LEN) As Byte
   szSystemStatus(0 To WSADESCRIPTION_LEN) As Byte
   iMaxSockets As Integer
   iMaxUdpDg As Integer
   lpVendorInfo As Long
End Type

Private Declare PtrSafe Function WSAGetLastError Lib "Ws2_32.dll" () As Integer
Private Declare PtrSafe Function WSAStartup Lib "Ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As LPWSADATA_Type) As Long
Private Declare PtrSafe Function sendto Lib "Ws2_32.dll" (ByVal socket As Long, ByVal buf As LongPtr, ByVal length As Long, ByVal flags As Long, ByRef toaddr As sockaddr_in, tolen As Long) As Long
Private Declare PtrSafe Function f_socket Lib "Ws2_32.dll" Alias "socket" (ByVal AF As Long, ByVal stype As Long, ByVal Protocol As Long) As Long
Private Declare PtrSafe Function closesocket Lib "Ws2_32.dll" (ByVal socket As Long) As Long
Private Declare PtrSafe Sub WSACleanup Lib "Ws2_32.dll" ()
Private Declare PtrSafe Function VarPtrArray Lib "VBE7" Alias "VarPtr" (Var() As Any) As LongPtr


Sub SendPacket(Message As String, IP As String, Port As Integer)

   Dim ConnectSocket As socket
   Dim wsaData As LPWSADATA_Type
   Dim iResult As Integer: iResult = 0
   Dim send_sock As sock_type: send_sock = INVALID_SOCKET
   Dim iFamily As AF: iFamily = AF_INET
   Dim iType As Integer: iType = SOCK_DGRAM

   
   Dim iProtocol As Integer: iProtocol = 0

   Dim SendBuf(0 To 1023) As Byte
   Dim BufLen As Integer: BufLen = 1024
   Dim RecvAddr As sockaddr_in: RecvAddr.sin_family = AF_INET ' : RecvAddr.sin_port = Port
   Dim SplitArray As Variant: SplitArray = Split(IP, ".")

   RecvAddr.sin_addr(0) = 127
   RecvAddr.sin_addr(1) = 0
   RecvAddr.sin_addr(2) = 0
   RecvAddr.sin_addr(3) = 1

   Dim abytPortAsBytes() As Byte
   abytPortAsBytes() = PortLongToBytes(Port)
   
RecvAddr.sin_port(1) = abytPortAsBytes(1)
RecvAddr.sin_port(0) = abytPortAsBytes(0)
   
   For buf = 1 To Len(Message)
      SendBuf(buf - 1) = Asc(Mid(Message, buf, 1))
   Next buf
   SendBuf(buf + 1) = 0

   iResult = WSAStartup(&H202, wsaData)
   If iResult <> 0 Then
      MsgBox ("WSAStartup failed: " & iResult)
      Exit Sub
   End If

   send_sock = f_socket(iFamily, iType, iProtocol)
   If send_sock = INVALID_SOCKET Then
      
      Errno = WSAGetLastError()
      MsgBox ("sock error: " & Errno)
      Exit Sub
   End If
  ' iResult = sendto(send_sock, buffer(0), Size, 0, RecvAddr, Len(RecvAddr))
   
   iResult = sendto(send_sock, VarPtr(SendBuf(0)), Len(Message), 0, RecvAddr, Len(RecvAddr))
   MsgBox ("sendto: " & iResult)
   Debug.Print (RecvAddr.sin_addr(0))
   
 ' BufLen, 0, RecvAddr, Len(RecvAddr))
   If iResult = -1 Then
      MsgBox ("sendto failed with error: " & IP)
      closesocket (send_sock)
      Call WSACleanup
      Exit Sub
   End If

   iResult = closesocket(send_sock)
   If iResult <> 0 Then
      MsgBox ("closesocket failed with error : " & WSAGetLastError())
      Call WSACleanup
   End If
End Sub

Private Function PortLongToBytes(ByVal lPort As Integer) As Byte()
    ReDim abytReturn(0 To 1) As Byte
    abytReturn(0) = lPort \ 256
    abytReturn(1) = lPort Mod 256
    PortLongToBytes = abytReturn()
End Function

Sub send()vb
    Call SendPacket("test", IP, Port)
End Sub

```

然后是TCP版本

```vb
Option Explicit
 
Private Const INVALID_SOCKET = -1
Private Const WSADESCRIPTION_LEN = 256
Private Const SOCKET_ERROR As Long = -1 'const #define SOCKET_ERROR            (-1)
 
Private Enum AF
    AF_UNSPEC = 0
    AF_INET = 2
    AF_IPX = 6
    AF_APPLETALK = 16
    AF_NETBIOS = 17
    AF_INET6 = 23
    AF_IRDA = 26
    AF_BTH = 32
End Enum
 
Private Enum sock_type
    SOCK_STREAM = 1
    SOCK_DGRAM = 2
    SOCK_RAW = 3
    SOCK_RDM = 4
    SOCK_SEQPACKET = 5
End Enum
 
Private Enum Protocol
    IPPROTO_ICMP = 1
    IPPROTO_IGMP = 2
    BTHPROTO_RFCOMM = 3
    IPPROTO_TCP = 6
    IPPROTO_UDP = 17
    IPPROTO_ICMPV6 = 58
    IPPROTO_RM = 113
End Enum
 
'Private Type sockaddr
'    sa_family As Integer
'    sa_data(0 To 13) As Byte
'End Type
 
Private Type sockaddr_in
    sin_family As Integer
    sin_port(0 To 1) As Byte
    sin_addr(0 To 3) As Byte
    sin_zero(0 To 7) As Byte
End Type
 
 
'typedef UINT_PTR        SOCKET;
Private Type udtSOCKET
    pointer As Long
End Type
 
 
 
' typedef struct WSAData {
'  WORD           wVersion;
'  WORD           wHighVersion;
'  char           szDescription[WSADESCRIPTION_LEN+1];
'  char           szSystemStatus[WSASYS_STATUS_LEN+1];
'  unsigned short iMaxSockets;
'  unsigned short iMaxUdpDg;
'  char FAR       *lpVendorInfo;
'} WSADATA, *LPWSADATA;
 
Private Type udtWSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADESCRIPTION_LEN) As Byte
    szSystemStatus(0 To WSADESCRIPTION_LEN) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
 
'int errorno = WSAGetLastError()
Private Declare PtrSafe Function WSAGetLastError Lib "Ws2_32" () As Integer
 
'   int WSAStartup(
'  __in   WORD wVersionRequested,
'  __out  LPWSADATA lpWSAData
');
Private Declare PtrSafe Function WSAStartup Lib "Ws2_32" (ByVal wVersionRequested As Integer, ByRef lpWSAData As udtWSADATA) As Long 'winsockErrorCodes2
 
 
'    SOCKET WSAAPI socket(
'  __in  int af,
'  __in  int type,
'  __in  int protocol
');
 
Private Declare PtrSafe Function ws2_socket Lib "Ws2_32" Alias "socket" (ByVal AF As Long, ByVal stype As Long, ByVal Protocol As Long) As Long
 
Private Declare PtrSafe Function ws2_closesocket Lib "Ws2_32" Alias "closesocket" (ByVal socket As Long) As Long
 
'int recv(
'  SOCKET s,
'  char   *buf,
'  int    len,
'  int    flags
');
Private Declare PtrSafe Function ws2_recv Lib "Ws2_32" Alias "recv" (ByVal socket As Long, ByVal buf As LongPtr, ByVal length As Long, ByVal flags As Long) As Long
 
'int WSAAPI connect(
'  SOCKET         s,
'  const sockaddr *name,
'  int            namelen
');
 
Private Declare PtrSafe Function ws2_connect Lib "Ws2_32" Alias "connect" (ByVal S As LongPtr, ByRef name As sockaddr_in, ByVal namelen As Long) As Long
 
'int WSAAPI send(
'  SOCKET     s,
'  const char *buf,
'  int        len,
'  int        flags
');
Private Declare PtrSafe Function ws2_send Lib "Ws2_32" Alias "send" (ByVal S As LongPtr, ByVal buf As LongPtr, ByVal BufLen As Long, ByVal flags As Long) As Long
 
 
Private Declare PtrSafe Function ws2_shutdown Lib "Ws2_32" Alias "shutdown" (ByVal S As Long, ByVal how As Long) As Long
 
Private Declare PtrSafe Sub WSACleanup Lib "Ws2_32" ()
 
Private Enum eShutdownConstants
    SD_RECEIVE = 0  '#define SD_RECEIVE      0x00
    SD_SEND = 1     '#define SD_SEND         0x01
    SD_BOTH = 2     '#define SD_BOTH         0x02
End Enum
 
Sub TestPortLongToBytes()
    'redis is on port number 6379
    Dim abytPortAsBytes() As Byte
    abytPortAsBytes() = PortLongToBytes(6379)
    Debug.Assert abytPortAsBytes(0) = 24
    Debug.Assert abytPortAsBytes(1) = 235
End Sub
 
Private Function PortLongToBytes(ByVal lPort As Integer) As Byte()
    ReDim abytReturn(0 To 1) As Byte
    abytReturn(0) = lPort \ 256
    abytReturn(1) = lPort Mod 256
    PortLongToBytes = abytReturn()
End Function
 
Sub TestWS2SendAndReceive()
 
    Dim sResponse As String
    Const clJavaPort As Long = 6666
    If WS2SendAndReceive(clJavaPort, "1+3" & vbCrLf, sResponse) Then
        MsgBox (sResponse)
    End If
End Sub
 
 
Public Function WS2SendAndReceive(ByVal lPort As Long, ByVal sText As String, ByRef psResponse As String) As Boolean
    'https://docs.microsoft.com/en-gb/windows/desktop/api/winsock/nf-winsock-recv
    If Right$(sText, 2) <> vbCrLf Then Err.Raise vbObjectError, , "Best suffix your sends with a new line (vbCrLf)"
    psResponse = ""
    '//----------------------
    '// Declare PtrSafe and initialize variables.
    Dim iResult As Integer: iResult = 0
    Dim wsaData As udtWSADATA
 
    Dim ConnectSocket As Long
 
    Dim clientService As sockaddr_in
 
    Dim SendBuf() As Byte
    SendBuf = StrConv(sText, vbFromUnicode)
 
    Const recvbuflen As Long = 512
    Dim recvbuf(0 To recvbuflen - 1) As Byte
 
    '//----------------------
    '// Initialize Winsock
    Dim eResult As Long 'winsockErrorCodes2
    eResult = WSAStartup(&H202, wsaData)
    If eResult <> 0 Then
        MsgBox "WSAStartup failed with error: " & eResult
        WS2SendAndReceive = False
        GoTo SingleExit
    End If
 
 
    '//----------------------
    '// Create a SOCKET for connecting to server
    ConnectSocket = ws2_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If ConnectSocket = INVALID_SOCKET Then
        Dim eLastError As Long 'winsockErrorCodes2
        eLastError = WSAGetLastError()
        MsgBox "socket failed with error: " & eLastError
        Call WSACleanup
        WS2SendAndReceive = False
        GoTo SingleExit
    End If
 
 
    '//----------------------
    '// The sockaddr_in structure specifies the address family,
    '// IP address, and port of the server to be connected to.
    clientService.sin_family = AF_INET
 
    clientService.sin_addr(0) = 127
    clientService.sin_addr(1) = 0
    clientService.sin_addr(2) = 0
    clientService.sin_addr(3) = 1
 
    Dim abytPortAsBytes() As Byte
    abytPortAsBytes() = PortLongToBytes(lPort)
 
    clientService.sin_port(1) = 235  '* 6379
    clientService.sin_port(0) = 24
 
    clientService.sin_port(1) = abytPortAsBytes(1)
    clientService.sin_port(0) = abytPortAsBytes(0)
 
    '//----------------------
    '// Connect to server.
 
    iResult = ws2_connect(ConnectSocket, clientService, LenB(clientService))
    If (iResult = SOCKET_ERROR) Then
 
        eLastError = WSAGetLastError()
 
        MsgBox "connect failed with 1 error: " & eLastError
        Call ws2_closesocket(ConnectSocket)
        Call WSACleanup
        WS2SendAndReceive = False
        GoTo SingleExit
    End If
 
    '//----------------------
    '// Send an initial buffer
    Dim sendbuflen As Long
    sendbuflen = UBound(SendBuf) - LBound(SendBuf) + 1
    iResult = ws2_send(ConnectSocket, VarPtr(SendBuf(0)), sendbuflen, 0)
    
    ' MsgBox iResult
    
    If (iResult = SOCKET_ERROR) Then
        eLastError = WSAGetLastError()
        MsgBox "send failed with error: " & eLastError
        Call ws2_closesocket(ConnectSocket)
        Call WSACleanup
        WS2SendAndReceive = False
        GoTo SingleExit
    End If
 
    'MsgBox "Bytes Sent: ", iResult
 
    '// shutdown the connection since no more data will be sent
    iResult = ws2_shutdown(ConnectSocket, SD_SEND)
    If (iResult = SOCKET_ERROR) Then
 
        eLastError = WSAGetLastError()
        MsgBox "shutdown failed with error: " & eLastError
        Call ws2_closesocket(ConnectSocket)
        Call WSACleanup
        WS2SendAndReceive = False
        GoTo SingleExit
    End If
 
    ' receive only one MsgBox (TODO handle when buffer is not large enough)
 
    iResult = ws2_recv(ConnectSocket, VarPtr(recvbuf(0)), recvbuflen, 0)
    If (iResult > 0) Then
        'MsgBox "Bytes received: ", iResult
    ElseIf (iResult = 0) Then
        MsgBox "Connection closed"
        WS2SendAndReceive = False
        Call ws2_closesocket(ConnectSocket)
        Call WSACleanup
        GoTo SingleExit
    Else
        eLastError = WSAGetLastError()
        MsgBox "recv failed with error: " & eLastError
    End If
 
    psResponse = Left$(StrConv(recvbuf, vbUnicode), iResult)
 
    'MsgBox psResponse
 
    '// close the socket
    iResult = ws2_closesocket(ConnectSocket)
    If (iResult = SOCKET_ERROR) Then
 
        eLastError = WSAGetLastError()
        MsgBox "close failed with error: " & eLastError
        
        Call WSACleanup
        WS2SendAndReceive = False
        GoTo SingleExit
    End If
 
    Call WSACleanup
    WS2SendAndReceive = True
 
SingleExit:
    Exit Function
ErrHand:
 
End Function

```

#### 参考文献

1. http://exceldevelopmentplatform.blogspot.com/2019/01/vba-sockets-ruby-java-interop-nirvana.html
2. https://stackoverflow.com/questions/48917030/sendto-sends-garbage-vba