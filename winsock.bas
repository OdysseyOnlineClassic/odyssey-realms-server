Attribute VB_Name = "modWinsock"
'Visual Basic 4.0 and above winsock declares and functions modules
'requires a msghook for async methods and functions. This declares
'file was originally obtained from the alt.winsock.programming news
'group, alot has been added and modified since that time. However I
'would like to credit the people on that newsgroup for the information
'they contributed, and now I pass it back...
'
'NOTES:
' I haven't been able to get the WSAAsyncGetXbyY functions to work properly
'under windows95(tm) aside from that ALL functions "SHOULD" work just fine.
'any questions about this file may be posted to alt.winsock.programming
'   Topaz..
'
Option Explicit

'windows declares here
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal cb&)
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

'WINSOCK DEFINES START HERE

Global Const FD_SETSIZE = 64

Type fd_set
    fd_count As Integer
    fd_array(FD_SETSIZE) As Integer
End Type

'same
Type timeval
    tv_sec As Long
    tv_usec As Long
End Type

Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Const hostent_size = 16

Type servent
    s_name As Long
    s_aliases As Long
    s_port As Integer
    s_proto As Long
End Type
Const servent_size = 14

Type protoent
    p_name As Long
    p_aliases As Long
    p_proto As Integer
End Type
Const protoent_size = 10

Global Const IPPROTO_TCP = 6
Global Const IPPROTO_UDP = 17

Global Const INADDR_NONE = &HFFFFFFFF
Global Const INADDR_ANY = &H0

Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
Public Const sockaddr_size = 16
Dim saZero As sockaddr

Global Const WSA_DESCRIPTIONLEN = 256
Global Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1

Global Const WSA_SYS_STATUS_LEN = 128
Global Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Global Const INVALID_SOCKET = -1
Global Const SOCKET_ERROR = -1

Global Const SOCK_STREAM = 1
Global Const SOCK_DGRAM = 2

Global Const MAXGETHOSTSTRUCT = 1024

Global Const AF_INET = 2
Global Const PF_INET = 2

Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type

Global Const SOL_SOCKET = &HFFFF&
Global Const SO_DEBUG = &H1&         ' 0x0001 turn on debugging info recording
Global Const SO_ACCEPTCONN = &H2&    ' 0x0002 socket has had listen()
Global Const SO_REUSEADDR = &H4&     ' 0x0004 allow local address reuse
Global Const SO_KEEPALIVE = &H8&     ' 0x0008 keep connections alive
Global Const SO_DONTROUTE = &H10&    ' 0x0010 just use interface addresses
Global Const SO_BROADCAST = &H20&    ' 0x0020 permit sending of broadcast messages
Global Const SO_USELOOPBACK = &H40&  ' 0x0040 bypass hardware when possible
Global Const SO_LINGER = &H80&       ' 0x0080 linger on close if data present
Global Const SO_OOBINLINE = &H100&   ' 0x0100 leave received OOB data in line
Global Const SO_DONTLINGER = Not SO_LINGER
'' Additional options
Global Const SO_SNDBUF = &H1001&   ' 0x1001 send buffer size
Global Const SO_RCVBUF = &H1002&   ' 0z1002 receive buffer size
Global Const SO_SNDLOWAT = &H1003&    ' 0x1003 send low-water mark
Global Const SO_RCVLOWAT = &H1004&    ' 0x1004 receive low-water mark
Global Const SO_SNDTIMEO = &H1005&    ' 0x1005 send timeout
Global Const SO_RCVTIMEO = &H1006&    ' 0x1006 receive timeout
Global Const SO_ERROR = &H1007&    ' 0x1007 get error status and clear
Global Const SO_TYPE = &H1008&     ' 0x1008 get socket type
'' TCP options
Global Const TCP_NODELAY = &H1    ' 0x0001

Global Const FD_READ = &H1&
Global Const FD_WRITE = &H2&
Global Const FD_OOB = &H4&
Global Const FD_ACCEPT = &H8&
Global Const FD_CONNECT = &H10&
Global Const FD_CLOSE = &H20&

'SOCKET FUNCTIONS
Declare Function accept Lib "ws2_32.dll" (ByVal S As Long, addr As sockaddr, addrlen As Long) As Long
Declare Function bind Lib "ws2_32.dll" (ByVal S As Long, addr As sockaddr, ByVal namelen As Long) As Long
Declare Function closesocket Lib "ws2_32.dll" (ByVal S As Long) As Long
Declare Function connect Lib "ws2_32.dll" (ByVal S As Long, addr As sockaddr, ByVal namelen As Long) As Long
Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal S As Long, ByVal cmd As Long, argp As Long) As Long
Declare Function getpeername Lib "ws2_32.dll" (ByVal S As Long, sname As sockaddr, namelen As Long) As Long
Declare Function getsockname Lib "ws2_32.dll" (ByVal S As Long, sname As sockaddr, namelen As Long) As Long
Declare Function getsockopt Lib "ws2_32.dll" (ByVal S As Long, ByVal Level As Long, ByVal optname As Long, ByVal optval As String, optlen As Long) As Long
Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Long) As Integer
Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Declare Function listen Lib "ws2_32.dll" (ByVal S As Long, ByVal backlog As Long) As Long
Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Long) As Integer
Declare Function recv Lib "ws2_32.dll" (ByVal S As Long, Buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Declare Function recvfrom Lib "ws2_32.dll" (ByVal S As Long, ByVal Buf As String, ByVal buflen As Long, ByVal flags As Long, from As sockaddr, fromlen As Long) As Long
Declare Function ws_select Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, timeout As timeval) As Long
Declare Function send Lib "ws2_32.dll" (ByVal S As Long, Buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Declare Function sendto Lib "ws2_32.dll" (ByVal S As Long, Buf As Any, ByVal buflen As Long, ByVal flags As Long, to_addr As sockaddr, ByVal tolen As Long) As Long
Declare Function setsockopt Lib "ws2_32.dll" (ByVal S As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Declare Function ShutDown Lib "ws2_32.dll" Alias "shutdown" (ByVal S As Long, ByVal how As Long) As Long
Declare Function Socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
'DATABASE FUNCTIONS
Declare Function gethostbyaddr Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Declare Function gethostname Lib "ws2_32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
Declare Function getservbyport Lib "ws2_32.dll" (ByVal Port As Long, ByVal proto As String) As Long
Declare Function getservbyname Lib "ws2_32.dll" (ByVal serv_name As String, ByVal proto As String) As Long
Declare Function getprotobynumber Lib "ws2_32.dll" (ByVal proto As Long) As Long
Declare Function getprotobyname Lib "ws2_32.dll" (ByVal proto_name As String) As Long
'WINDOWS EXTENSIONS
Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Declare Sub WSASetLastError Lib "ws2_32.dll" (ByVal iError As Long)
Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
Declare Function WSAIsBlocking Lib "ws2_32.dll" () As Long
Declare Function WSAUnhookBlockingHook Lib "ws2_32.dll" () As Long
Declare Function WSASetBlockingHook Lib "ws2_32.dll" (ByVal lpBlockFunc As Long) As Long
Declare Function WSACancelBlockingCall Lib "ws2_32.dll" () As Long
'WSAASYNCGETXBYY FUNCTIONS DON'T WORK RELIABLY UNDER 32BIT WINDOWS
Declare Function WSAAsyncGetServByName Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal serv_name As String, ByVal proto As String, ByVal Buf As String, ByVal buflen As Long) As Long
Declare Function WSAAsyncGetServByPort Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Port As Long, ByVal proto As String, ByVal Buf As String, ByVal buflen As Long) As Long
Declare Function WSAAsyncGetProtoByName Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal proto_name As String, ByVal Buf As String, ByVal buflen As Long) As Long
Declare Function WSAAsyncGetProtoByNumber Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal number As Long, ByVal Buf As String, ByVal buflen As Long) As Long
Declare Function WSAAsyncGetHostByName Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal host_name As String, ByVal Buf As String, ByVal buflen As Long) As Long
Declare Function WSAAsyncGetHostByAddr Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal addr As String, ByVal addr_len As Long, ByVal addr_type As Long, ByVal Buf As String, ByVal buflen As Long) As Long
Declare Function WSACancelAsyncRequest Lib "ws2_32.dll" (ByVal hAsyncTaskHandle As Long) As Long
Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal S As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Declare Function WSARecvEx Lib "ws2_32.dll" (ByVal S As Long, Buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long

'SOME STUFF I ADDED
Global Const INVALID_PORT = -1  'added by me
Global Const INVALID_PROTO = -1    'added by me

Global MySocket%
Global SockReadBuffer$

Global Const WSA_NoName = "Unknown"

Global WSAStartedUp%    'Flag to keep track of whether winsock WSAStartup wascalled

Global WSARecvActive%   'Flag to indicate an async recv is in progress
Global WSASendActive%   'Flag to indicate an async send is in progress
Public Function SendData(ByVal S&, St As String) As Long
    Dim TheMsg() As Byte
    TheMsg = ""
    TheMsg = StrConv(St, vbFromUnicode)
    If UBound(TheMsg) > -1 Then
        SendData = send(S, TheMsg(0), UBound(TheMsg) + 1, 0)
    End If
End Function
Function Receive(S As Long) As String
    Dim Buf() As Byte, L As Long, St As String
    ReDim Buf(256) As Byte
    L = recv(S, Buf(0), 255, 0)
    If L > 0 Then
        ReDim Preserve Buf(L - 1)
        St = StrConv(Buf, vbUnicode)
        Receive = St
    End If
End Function
'this function should work on 16 and 32 bit systems
Function WSAGetAsyncBufLen(ByVal lParam As Long) As Long
'used only by the async lookups and wparam will
'be the task handle for these events
    On Error Resume Next
    '#define WSAGETASYNCBUFLEN(lParam)           LOWORD(lParam)
    WSAGetAsyncBufLen = lParam And &HFFFF&
    If Err Then
        WSAGetAsyncBufLen = 0
    End If
End Function

'this function should work on 16 and 32 bit systems
Function WSAGetSelectEvent(ByVal lParam As Long) As Long
    On Error Resume Next
    '#define WSAGETSELECTEVENT(lParam)            LOWORD(lParam)
    WSAGetSelectEvent = (lParam And &HFFFF&)
    If Err Then
        WSAGetSelectEvent = 0
    End If
End Function



'this function should work on 16 and 32 bit systems
Function WSAGetAsyncError(ByVal lParam As Long) As Long
    On Error Resume Next
    'WSAGETASYNCERROR(lParam) HIWORD(lParam)
    WSAGetAsyncError = (lParam \ &H10000 And &HFFFF&)
    If Err Then
        WSAGetAsyncError = 0
    End If
End Function

'this function DOES work on 16 and 32 bit systems
Function AddrToIP(ByVal AddrOrIP$) As String
    On Error Resume Next
    AddrToIP$ = getascip(GetHostByNameAlias(AddrOrIP$))
    If Err Then AddrToIP$ = "255.255.255.255"
End Function

Function ConnectSock(ByVal host$, ByVal Port&, retIpPort$, ByVal HWndToMsg&, Msg As Long, ByVal Async%) As Long
    Dim S&, SelectOps&, dummy&
    Dim sockin As sockaddr

    SockReadBuffer$ = ""
    sockin = saZero
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    If sockin.sin_port = INVALID_PORT Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If

    sockin.sin_addr = GetHostByNameAlias(host$)
    If sockin.sin_addr = INADDR_NONE Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    retIpPort$ = getascip$(sockin.sin_addr) & ":" & ntohs(sockin.sin_port)

    S = Socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
    If S < 0 Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    If SetSockLinger(S, 1, 0) = SOCKET_ERROR Then
        If S > 0 Then
            dummy = closesocket(S)
        End If
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    If Not Async Then
        If connect(S, sockin, sockaddr_size) <> 0 Then
            If S > 0 Then
                dummy = closesocket(S)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
        SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
        If WSAAsyncSelect(S, HWndToMsg, ByVal Msg, ByVal SelectOps) Then
            If S > 0 Then
                dummy = closesocket(S)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
    Else
        SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
        If WSAAsyncSelect(S, HWndToMsg, ByVal Msg, ByVal SelectOps) Then
            If S > 0 Then
                dummy = closesocket(S)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
        If connect(S, sockin, sockaddr_size) <> -1 Then
            If S > 0 Then
                dummy = closesocket(S)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
    End If
    If setsockopt(S, SOL_SOCKET, SO_RCVBUF, 8192&, 4) = SOCKET_ERROR Then
        If S > 0 Then
            dummy = closesocket(S)
        End If
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    'If setsockopt(S, SOL_SOCKET, SO_SNDBUF, 8192&, 4) = SOCKET_ERROR Then
    '    If S > 0 Then
    '        dummy = closesocket(S)
    '    End If
    '    ConnectSock = INVALID_SOCKET
    '    Exit Function
    'End If
    'If setsockopt(S, IPPROTO_TCP, TCP_NODELAY, 1&, 1) = SOCKET_ERROR Then
    '    If S > 0 Then
    '        dummy = closesocket(S)
    '    End If
    '    ConnectSock = INVALID_SOCKET
    '    Exit Function
    'End If
    ConnectSock = S
End Function
Function SetSockLinger(ByVal SockNum&, ByVal OnOff%, ByVal LingerTime%) As Long
    Dim ret&
    Dim LingBuf$, Linger As LingerType

    LingBuf = Space(4)
    Linger.l_onoff = OnOff
    Linger.l_linger = LingerTime
    MemCopy ByVal LingBuf, Linger, 4
    ret = setsockopt(SockNum, SOL_SOCKET, SO_LINGER, ByVal LingBuf, 4)
    'Debug.Print "Error Setting Linger info: "; ret

    LingBuf = Space(4)
    Linger.l_onoff = 0   'reset this so we can get an honest look
    Linger.l_linger = 0  'reset this so we can get an honest look

    ret = getsockopt(SockNum, SOL_SOCKET, SO_LINGER, LingBuf, 4)
    MemCopy Linger, ByVal LingBuf, 4
    'Debug.Print "Error Getting Linger info: "; ret
    'Debug.Print "Linger is on if nonzero: "; Linger.l_onoff
    'Debug.Print "Linger time if linger is on: "; Linger.l_linger
    SetSockLinger = ret
End Function
Sub EndWinsock()
    Dim ret&
    If WSAIsBlocking() Then
        ret = WSACancelBlockingCall()
    End If
    ret = WSACleanup()
    WSAStartedUp = False
End Sub

Function getascip(ByVal inn As Long) As String
    On Error Resume Next
    Dim lpStr&
    Dim nStr&
    Dim retString$

    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        getascip = "255.255.255.255"
        Exit Function
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    MemCopy ByVal retString, ByVal lpStr, nStr
    retString = Left(retString, nStr)
    getascip = retString
    If Err Then getascip = "255.255.255.255"
End Function
'this function DOES work on 32bit and 16 bit systems
Function GetHostByAddress(ByVal addr As Long) As String
    On Error Resume Next
    Dim phe&
    Dim heDestHost As HostEnt
    Dim hostname$
    phe = gethostbyaddr(addr, 4, PF_INET)
    'Debug.Print phe
    If phe <> 0 Then
        MemCopy heDestHost, ByVal phe, hostent_size
        'Debug.Print heDestHost.h_name
        'Debug.Print heDestHost.h_aliases
        'Debug.Print heDestHost.h_addrtype
        'Debug.Print heDestHost.h_length
        'Debug.Print heDestHost.h_addr_list

        hostname = String(256, 0)
        MemCopy ByVal hostname, ByVal heDestHost.h_name, 256
        GetHostByAddress = Left(hostname, InStr(hostname, vbNullChar) - 1)
    Else
        GetHostByAddress = WSA_NoName
    End If
    If Err Then GetHostByAddress = WSA_NoName
End Function
'this function DOES work on 16 and 32 bit systems
Function GetHostByNameAlias(ByVal hostname$) As Long
    On Error Resume Next
    'Return IP address as a long, in network byte order

    Dim phe&    ' pointer to host information entry
    Dim heDestHost As HostEnt    'hostent structure
    Dim addrList&
    Dim retIP&
    'first check to see if what we have been passed is a valid IP
    retIP = inet_addr(hostname)
    If retIP = INADDR_NONE Then
        'it wasn't an IP, so do a DNS lookup
        phe = gethostbyname(hostname)
        If phe <> 0 Then
            'Pointer is non-null, so copy in hostent structure
            MemCopy heDestHost, ByVal phe, hostent_size
            'Now get first pointer in address list
            MemCopy addrList, ByVal heDestHost.h_addr_list, 4
            MemCopy retIP, ByVal addrList, heDestHost.h_length
        Else
            'its not a valid address
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
    If Err Then GetHostByNameAlias = INADDR_NONE
End Function

'this function DOES work on 16 and 32 bit systems
Function GetLocalHostName() As String
    Dim dummy&
    Dim LocalName$
    Dim S$
    On Error Resume Next
    LocalName = String(256, 0)
    LocalName = WSA_NoName
    dummy = 1
    S = String(256, 0)
    dummy = gethostname(S, 256)
    If dummy = 0 Then
        S = Left(S, InStr(S, vbNullChar) - 1)
        If Len(S) > 0 Then
            LocalName = S
        End If
    End If
    GetLocalHostName = LocalName
    If Err Then GetLocalHostName = WSA_NoName
End Function

Function GetPeerAddress(ByVal S&) As String
    Dim addrlen&
    Dim ret&

    On Error Resume Next
    Dim sa As sockaddr
    addrlen = sockaddr_size
    ret = getpeername(S, sa, addrlen)
    If ret = 0 Then
        GetPeerAddress = SockaddressToString(sa)
    Else
        GetPeerAddress = ""
    End If
    If Err Then GetPeerAddress = ""
End Function
'this function should work on 16 and 32 bit systems

Function GetPortFromString(ByVal PortStr$) As Long
'sometimes users provide ports outside the range of a VB
'integer, so this function returns an integer for a string
'just to keep an error from happening, it converts the
'number to a negative if needed
    On Error Resume Next
    If Val(PortStr$) > 32767 Then
        GetPortFromString = CInt(Val(PortStr$) - &H10000)
    Else
        GetPortFromString = Val(PortStr$)
    End If
    If Err Then GetPortFromString = 0
End Function
Function GetProtocolByName(ByVal protocol$) As Long
    Dim tmpShort&

    On Error Resume Next
    Dim ppe&
    Dim peDestProt As protoent
    ppe = getprotobyname(protocol)
    If ppe = 0 Then
        tmpShort = Val(protocol)
        If tmpShort <> 0 Or protocol = "0" Or protocol = "" Then
            GetProtocolByName = htons(tmpShort)
        Else
            GetProtocolByName = INVALID_PROTO
        End If
    Else
        MemCopy peDestProt, ByVal ppe, protoent_size
        GetProtocolByName = peDestProt.p_proto
    End If
    If Err Then GetProtocolByName = INVALID_PROTO
End Function

Function GetServiceByName(ByVal service$, ByVal protocol$) As Long
    Dim serv&

    On Error Resume Next
    Dim pse&
    Dim seDestServ As servent
    pse = getservbyname(service, protocol)
    If pse <> 0 Then
        MemCopy seDestServ, ByVal pse, servent_size
        GetServiceByName = seDestServ.s_port
    Else
        serv = Val(service)
        If serv <> 0 Then
            GetServiceByName = htons(serv)
        Else
            GetServiceByName = INVALID_PORT
        End If
    End If
    If Err Then GetServiceByName = INVALID_PORT
End Function
Function GetSockAddress(ByVal S&) As String
    Dim addrlen&
    Dim ret&
    On Error Resume Next
    Dim sa As sockaddr
    Dim szRet$
    szRet = String(32, 0)
    addrlen = sockaddr_size
    ret = getsockname(S, sa, addrlen)
    If ret = 0 Then
        GetSockAddress = SockaddressToString(sa)
    Else
        GetSockAddress = ""
    End If
    If Err Then GetSockAddress = ""
End Function
'this function should work on 16 and 32 bit systems
Function GetWSAErrorString(ByVal errnum&) As String
    On Error Resume Next
    Select Case errnum
    Case 10004: GetWSAErrorString = "Interrupted system call."
    Case 10009: GetWSAErrorString = "Bad file number."
    Case 10013: GetWSAErrorString = "Permission Denied."
    Case 10014: GetWSAErrorString = "Bad Address."
    Case 10022: GetWSAErrorString = "Invalid Argument."
    Case 10024: GetWSAErrorString = "Too many open files."
    Case 10035: GetWSAErrorString = "Operation would block."
    Case 10036: GetWSAErrorString = "Operation now in progress."
    Case 10037: GetWSAErrorString = "Operation already in progress."
    Case 10038: GetWSAErrorString = "Socket operation on nonsocket."
    Case 10039: GetWSAErrorString = "Destination address required."
    Case 10040: GetWSAErrorString = "Message too long."
    Case 10041: GetWSAErrorString = "Protocol wrong type for socket."
    Case 10042: GetWSAErrorString = "Protocol not available."
    Case 10043: GetWSAErrorString = "Protocol not supported."
    Case 10044: GetWSAErrorString = "Socket type not supported."
    Case 10045: GetWSAErrorString = "Operation not supported on socket."
    Case 10046: GetWSAErrorString = "Protocol family not supported."
    Case 10047: GetWSAErrorString = "Address family not supported by protocol family."
    Case 10048: GetWSAErrorString = "Address already in use."
    Case 10049: GetWSAErrorString = "Can't assign requested address."
    Case 10050: GetWSAErrorString = "Network is down."
    Case 10051: GetWSAErrorString = "Network is unreachable."
    Case 10052: GetWSAErrorString = "Network dropped connection."
    Case 10053: GetWSAErrorString = "Software caused connection abort."
    Case 10054: GetWSAErrorString = "Connection reset by peer."
    Case 10055: GetWSAErrorString = "No buffer space available."
    Case 10056: GetWSAErrorString = "Socket is already connected."
    Case 10057: GetWSAErrorString = "Socket is not connected."
    Case 10058: GetWSAErrorString = "Can't send after socket shutdown."
    Case 10059: GetWSAErrorString = "Too many references: can't splice."
    Case 10060: GetWSAErrorString = "Connection timed out."
    Case 10061: GetWSAErrorString = "Connection refused."
    Case 10062: GetWSAErrorString = "Too many Levels of symbolic links."
    Case 10063: GetWSAErrorString = "File name too long."
    Case 10064: GetWSAErrorString = "Host is down."
    Case 10065: GetWSAErrorString = "No route to host."
    Case 10066: GetWSAErrorString = "Directory not empty."
    Case 10067: GetWSAErrorString = "Too many processes."
    Case 10068: GetWSAErrorString = "Too many users."
    Case 10069: GetWSAErrorString = "Disk quota exceeded."
    Case 10070: GetWSAErrorString = "Stale NFS file handle."
    Case 10071: GetWSAErrorString = "Too many Levels of remote in path."
    Case 10091: GetWSAErrorString = "Network subsystem is unusable."
    Case 10092: GetWSAErrorString = "Winsock DLL cannot support this application."
    Case 10093: GetWSAErrorString = "Winsock not initialized."
    Case 10101: GetWSAErrorString = "Disconnect."
    Case 11001: GetWSAErrorString = "Host not found."
    Case 11002: GetWSAErrorString = "Nonauthoritative host not found."
    Case 11003: GetWSAErrorString = "Nonrecoverable error."
    Case 11004: GetWSAErrorString = "Valid name, no data record of requested type."
    Case Else:
    End Select
End Function

'this function DOES work on 16 and 32 bit systems
Function IpToAddr(ByVal AddrOrIP$) As String
    On Error Resume Next
    IpToAddr = GetHostByAddress(GetHostByNameAlias(AddrOrIP$))
    If Err Then IpToAddr = WSA_NoName
End Function

Function IrcGetAscIp(ByVal IPL$) As String
'this function is IRC specific, it expects a long ip stored in Network byte order, in a string
'the kind that would be parsed out of a DCC command string
    On Error GoTo IrcGetAscIPError:
    Dim lpStr&
    Dim nStr&
    Dim retString$
    Dim inn&
    If Val(IPL) > 2147483647 Then
        inn = Val(IPL) - 4294967296#
    Else
        inn = Val(IPL)
    End If
    inn = ntohl(inn)
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        IrcGetAscIp = "0.0.0.0"
        Exit Function
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    MemCopy ByVal retString, ByVal lpStr, nStr
    retString = Left(retString, nStr)
    IrcGetAscIp = retString
    Exit Function
IrcGetAscIPError:
    IrcGetAscIp = "0.0.0.0"
    Exit Function
    Resume
End Function
'this function DOES work on 16 and 32 bit systems
Function IrcGetLongIp(ByVal AscIp$) As String
'this function converts an ascii ip string into a long ip in network byte order
'and stick it in a string suitable for use in a DCC command.
    On Error GoTo IrcGetLongIpError:
    Dim inn&
    inn = inet_addr(AscIp)
    inn = htonl(inn)
    If inn < 0 Then
        IrcGetLongIp = CVar(inn + 4294967296#)
        Exit Function
    Else
        IrcGetLongIp = CVar(inn)
        Exit Function
    End If
    Exit Function
IrcGetLongIpError:
    IrcGetLongIp = "0"
    Exit Function
    Resume
End Function

Function ListenForConnect(ByVal Port&, ByVal HWndToMsg&, Msg As Long) As Long
    Dim S&, dummy&
    Dim SelectOps&
    Dim sockin As sockaddr

    sockin = saZero     'zero out the structure
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    If sockin.sin_port = INVALID_PORT Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    sockin.sin_addr = htonl(INADDR_ANY)
    If sockin.sin_addr = INADDR_NONE Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    S = Socket(PF_INET, SOCK_STREAM, 0)
    If S < 0 Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    If bind(S, sockin, sockaddr_size) Then
        If S > 0 Then
            dummy = closesocket(S)
        End If
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    SelectOps = FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
    If WSAAsyncSelect(S, HWndToMsg, ByVal Msg, ByVal SelectOps) Then
        If S > 0 Then
            dummy = closesocket(S)
        End If
        ListenForConnect = SOCKET_ERROR
        Exit Function
    End If

    If listen(S, 1) Then
        If S > 0 Then
            dummy = closesocket(S)
        End If
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    ListenForConnect = S
End Function
'this function should work on 16 and 32 bit systems
Function SockaddressToString(sa As sockaddr) As String
    On Error Resume Next
    SockaddressToString = getascip(sa.sin_addr)    '& ":" & ntohs(sa.sin_port)
    If Err Then SockaddressToString = ""
End Function

Function StartWinsock(desc$) As Long
    Dim ret&
    Dim WinsockVers&

    Dim wsadStartupData As WSADataType
    WinsockVers = &H101   'Vers 1.1

    If WSAStartedUp = False Then
        ret = 1
        ret = WSAStartup(WinsockVers, wsadStartupData)
        If ret = 0 Then
            WSAStartedUp = True
            'Debug.Print "wVersion="; VBntoaVers(wsadStartupData.wVersion), "wHighVersion="; VBntoaVers(wsadStartupData.wHighVersion)
            'Debug.Print "szDescription="; wsadStartupData.szDescription
            'Debug.Print "szSystemStatus="; wsadStartupData.szSystemStatus
            'Debug.Print "iMaxSockets="; wsadStartupData.iMaxSockets, "iMaxUdpDg="; wsadStartupData.iMaxUdpDg
            desc = wsadStartupData.szDescription
        Else
            WSAStartedUp = False
        End If
    End If
    StartWinsock = WSAStartedUp
End Function
Function VBntoaVers$(ByVal vers&)
    On Error Resume Next
    Dim szVers$
    szVers = String(5, 0)
    szVers = (vers And &HFF) & "." & ((vers And &HFF00) / 256)
    VBntoaVers = szVers
End Function
'this function should work on 16 and 32 bit systems
'this function uses MemCopy to transfer data from
'2 integers into 2 strings, then combines the strings
'and copys that data into a long, ineffect MAKELONG()
Function WSAMakeSelectReply(ByVal TheEvent%, ByVal TheError%) As Long
    Dim EventStr$, ErrorStr$, BothStr$, TheLong&
    EventStr = Space(2)
    ErrorStr = Space(2)
    BothStr = Space(4)
    MemCopy ByVal EventStr, TheEvent, 2
    MemCopy ByVal ErrorStr, TheError, 2
    BothStr = EventStr & ErrorStr
    If Len(BothStr) = 4 Then
        MemCopy TheLong, ByVal BothStr, 4
    End If
    WSAMakeSelectReply = TheLong
End Function


