Attribute VB_Name = "modMain"
Option Explicit

Public REN As String, UserName As String, NM As String, Step As Integer, RoomName As String
Public buffer As String, Connected As Boolean, Password As String, packet As String, mp As Long
Public InRoom As Boolean, USRNM As String, ChrCode As String
Public GroupName() As String, lstFriendID() As String
Public IsfrmMainLoaded As Boolean
Public lstOnLineFriend As String, lstStatus As String
Public GroupError As Boolean
Public lstNowChatting As String
Public LocalPort As Long
Public frmNewPager() As New frmPager
Public lstSize As String

Public IsLoggedIn As Boolean
Public Const MyToolTip As Integer = 64

Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * MyToolTip
End Type
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
   
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" _
   (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public SysTrayIcon As NOTIFYICONDATA

Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64

Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(1 To LF_FACESIZE) As Byte
End Type


Type NEWTEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
        ntmFlags As Long
        ntmSizeEM As Long
        ntmCellHeight As Long
        ntmAveWidth As Long
End Type

' ntmFlags field flags
Public Const NTM_REGULAR = &H40&
Public Const NTM_BOLD = &H20&
Public Const NTM_ITALIC = &H1&

'  tmPitchAndFamily flags
Public Const TMPF_FIXED_PITCH = &H1
Public Const TMPF_VECTOR = &H2
Public Const TMPF_DEVICE = &H8
Public Const TMPF_TRUETYPE = &H4

Public Const ELF_VERSION = 0
Public Const ELF_CULTURE_LATIN = 0

'  EnumFonts Masks
Public Const RASTER_FONTTYPE = &H1
Public Const DEVICE_FONTTYPE = &H2
Public Const TRUETYPE_FONTTYPE = &H4
Public Const DEFAULT_CHARSET = 1


Public Declare Function EnumFontFamilies Lib "gdi32" Alias _
    "EnumFontFamiliesA" _
    (ByVal hdc As Long, ByVal lpszFamily As String, _
    ByVal lpEnumFontFamProc As Long, lParam As Any) As Long

Public Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" _
    (ByVal hdc As Long, lpLogFont As LOGFONT, _
    ByVal lpEnumFontProc As Long, _
    ByVal lParam As Long, ByVal dw As Long) _
As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
    ByVal hdc As Long) As Long
    
Type POINTAPI

    x As Integer
    Y As Integer

End Type
Type ConvertPOINTAPI

    xy As Long

End Type

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Public Const WM_SYSCOMMAND = &H112
Public Const WM_PAINT = &HF

Public Const MOUSE_MOVE = &HF012
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_GRAYTEXT = 17


Public Const DSTINVERT = &H550009   ' (DWORD) dest = (NOT dest)


Public Const GWW_HWNDPARENT = (-8)
Public ToolbarLoaded As Integer


Public Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _
    ByVal FontType As Long, lParam As ComboBox) As Long
    
    Dim FaceName As String
    Dim FullName As String
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    lParam.AddItem Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
    EnumFontFamProc = 1
'    EnumFontFamExProc = 1
End Function


Public Sub FillComboWithFonts(CBO As ComboBox)
    Dim hdc As Long
    CBO.Clear
    hdc = GetDC(CBO.hwnd)

    EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamProc, CBO
        
'    EnumFontFamiliesEx hDC, lpNLF, AddressOf EnumFontFamExProc, CBO.hWnd, 0
    ReleaseDC CBO.hwnd, hdc
End Sub

Public Function CalcSize(PckLen As Integer) As String
    Dim FstNum As String
    FstNum = 0
    Do While PckLen > 255
    FstNum = FstNum + 1
    PckLen = PckLen - 256
    Loop
    CalcSize = Chr$(FstNum) & Chr$(PckLen)
End Function

Public Function Chat_ChatLogin() As String
    REN = "109À€" & UserName & "À€1À€" & UserName & "À€6À€abcdeÀ€"
    Chat_ChatLogin = "YMSG" & Chr(9) & String(4, 0) & Chr(Len(REN)) & Chr(0) & Chr(&H96) & String(4, 0) & NM & REN
End Function

Public Function Chat_RoomLogin() As String
    REN = "1À€" & UserName & "À€104À€" & RoomName & "À€"
    Chat_RoomLogin = "YMSG" & Chr(9) & String(4, 0) & Chr(Len(REN)) & Chr(0) & Chr(&H98) & String(4, 0) & NM & REN
End Function

Public Function Chat_ChatSend(rmtext As String) As String
    REN = "1À€" & UserName & "À€104À€" & RoomName & "À€117À€" & rmtext & "À€124À€1À€"
    Chat_ChatSend = "YMSG" & Chr(9) & String(4, 0) & Chr(Len(REN)) & Chr(0) & Chr(&HA8) & String(4, 0) & NM & REN
End Function

Public Sub PM_Send(txtToWhom As String, PMData As String, sckSend As Winsock)
    Dim SendData As String
    REN = "5À€" & txtToWhom & "À€4À€" & UserName & "À€8À€None" & "À€14À€" & PMData & "À€97À€1À€"
    SendData = "YMSG" & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(Len(REN)) & Chr(0) & Chr(6) & Chr(0) & Chr(0) & Chr(0) & Chr(1) & NM & REN
    sckSend.SendData SendData
End Sub

Public Sub BootHim(txtToWhom As String, sck As Winsock)
    Dim Data2Send
    REN = "4À€" & UserName & "À€2À€booÀ€5À€" & txtToWhom & "À€5À€" & txtToWhom & "§À€5À€" & txtToWhom & "§§À€5À€" & txtToWhom & "§§§À€13À€4À€49À€PEERTOPEERÀ€14À€2À€16À€0À€À€"
    Data2Send = "YMSG" & Chr(10) & Chr(0) & Chr(11) & Chr(0) & Chr(0) & Chr(Len(REN)) & Chr(0) & Chr(77) & Chr(255) & Chr(11) & Chr(11) & Chr(11) & NM & REN
    sck.SendData Data2Send
End Sub

Public Sub Add_Me(txtToWhom As String, SendData As String, Optional GroupName As String = "Friends", Optional sckAdd As Winsock)
    Dim Data2Send As String
    REN = "1À€" & UserName & "À€7À€" & txtToWhom & "À€14À€" & SendData & "À€65À€" & GroupName & "À€"
    Data2Send = "YMSG" & Chr(10) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(Len(REN)) & Chr(0) & Chr(131) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & NM & REN
    sckAdd.SendData Data2Send
End Sub

Public Function CamInvite(txtToWhom As String) As String
    REN = "49À€WEBCAMINVITEÀ€14À€ À€13À€0À€1À€" & UserName & "À€5À€" & txtToWhom & "À€1002À€2À€"
    CamInvite = "YMSG" & Chr(9) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(Len(REN)) & Chr(0) & "K" & Chr(0) & Chr(0) & Chr(0) & Chr(22) & NM & REN
End Function

Public Function VoiceInvite(WhoTo As String, from As String, sck As Winsock) As String
    REN = "1À€" & from & "À€5À€" & WhoTo & "À€57À€" & from & "-21723À€13À€1À€"
    VoiceInvite = "YMSG" & Chr(9) & String(4, 0) & Chr(Len(REN)) & Chr(0) & "J" & String(4, 0) & NM & REN
    sck.SendData VoiceInvite
End Function

Sub SendFiles(WhoTo As String, WhoFrom As String, FileName As String, sckSendFile As Winsock)
    On Error Resume Next
    Dim B, a, c As String
    B = "5À€" & WhoTo & "À€4À€" & WhoFrom & "À€49À€FILEXFERÀ€1À€" & WhoFrom & "À€13À€1À€27À€" & FileName & "À€28À€720896À€20À€"
    Dim mp As Integer
    mp = Len(B)
    a = "YMSG" & Chr(9) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(mp) & Chr(0) & "M" & Chr(0) & Chr(0) & Chr(0) & Chr(&H16) & Chr(&H6B) & Chr(&HD3) & Chr(&H30) & Chr(&H30)
    c = a & B
    sckSendFile.SendData c
End Sub

Sub ChangeStatus(StatusID As Integer, sckStatus As Winsock, Optional CustomMessage As String, Optional IsBusy As Boolean = True)
    Dim a As String
    REN = "10À€" & CStr(StatusID)
    If StatusID = 99 Then
        REN = REN & "À€19À€" & CustomMessage & "À€47À€"
    Else
        REN = REN & "À€47À€"
    End If
    If IsBusy Then
        REN = REN & "1À€"
    Else
        REN = REN & "0À€"
    End If
    a = "YMSG" & Chr(10) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(Len(REN)) & Chr(0) & Chr(3) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & NM & REN
    sckStatus.SendData a
End Sub

Public Sub Print2Rich(PersonName As String, txtData As String, rtb As RichTextBox)
    
End Sub
