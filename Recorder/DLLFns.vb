Imports System.Drawing.Imaging
Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure KeyboardInput
    Public wVk As UShort
    Public wScan As UShort
    Public dwFlags As UInteger
    Public time As UInteger
    Public dwExtraInfo As IntPtr
End Structure
<StructLayout(LayoutKind.Sequential)>
Public Structure MouseInput
    Public dx As Integer
    Public dy As Integer
    Public mouseData As UInteger
    Public dwFlags As UInteger
    Public time As UInteger
    Public dwExtraInfo As IntPtr
End Structure
<StructLayout(LayoutKind.Sequential)>
Public Structure HardwareInput
    Public uMsg As UInteger
    Public wParamL As UShort
    Public wParamH As UShort
End Structure
<StructLayout(LayoutKind.Sequential)>
Public Structure RECT
    Public Left As Integer
    Public Top As Integer
    Public Right As Integer
    Public Bottom As Integer
End Structure
<StructLayout(LayoutKind.Explicit)>
Public Structure InputUnion
    <FieldOffset(0)>
    Public mi As MouseInput
    <FieldOffset(0)>
    Public ki As KeyboardInput
    <FieldOffset(0)>
    Public hi As HardwareInput
End Structure
<StructLayout(LayoutKind.Sequential)>
Public Structure SPOINT
    Public X As Integer
    Public Y As Integer
End Structure
Public Structure Input
    Public type As Integer
    Public u As InputUnion
End Structure
<Flags>
Public Enum SpecialFolders
    DOWNLOAD = -1
    DOCUMENT = 1
    MYPICTURE = 2
    MYMUSIC = 3
    MYVIDEO = 4
    USER = 5
    DESKTOP = 6
    WINDOWS = 7
    PROGRAMFILES = 8
    PROGRAMFILES86 = 9
    TEMP = 10
    APP = 11
End Enum
<Flags>
Public Enum InputType
    Mouse = 0
    Keyboard = 1
    Hardware = 2
End Enum
<Flags>
Public Enum KeyEventF
    KeyDown = &H0
    ExtendedKey = &H1
    KeyUp = &H2
    Unicode = &H4
    Scancode = &H8
End Enum
<Flags>
Public Enum MouseEventF
    Absolute = &H8000
    HWheel = &H1000
    Move = &H1
    MoveNoCoalesce = &H2000
    LeftDown = &H2
    LeftUp = &H4
    RightDown = &H8
    RightUp = &H10
    MiddleDown = &H20
    MiddleUp = &H40
    VirtualDesk = &H4000
    Wheel = &H800
    XDown = &H80
    XUp = &H100
End Enum
<Flags>
Public Enum VKey
    NA = -1
    BAK = &H8
    TAB = &H9
    CLEAR = &HC
    Enter = &HD
    SHIFT = &H10
    LSHIFT = &HA0
    RSHIFT = &HA1
    CONTROL = &H11
    LCONTROL = &HA2
    RCONTROL = &HA3
    ALT = &H12
    LALT = &HA4
    RALT = &HA5
    PAUSE = &H13
    CAPSLOCK = &H14
    ESCAPE = &H1B
    SPACE = &H20
    PAGEUP = &H21
    PAGEDOWN = &H22
    BEND = &H23
    HOME = &H24
    LEFT = &H25
    UP = &H26
    RIGHT = &H27
    DOWN = &H28
    BSelect = &H29
    PRINT = &H2A
    PRINTSCREEN = &H2C
    INSERT = &H2D
    DELETE = &H2E
    HELP = &H2F
    F1 = &H70
    F2 = &H71
    F3 = &H72
    F4 = &H73
    F5 = &H74
    F6 = &H75
    F7 = &H76
    F8 = &H77
    F9 = &H78
    F10 = &H79
    F11 = &H7A
    F12 = &H7B
    F13 = &H7C
    F14 = &H7D
    F15 = &H7E
    F16 = &H7F
    SLEEP = &H5F
    NUMLOCK = &H90
    SCROLL = &H91
    WIN = &H5B
    LWIN = &H5B
    RWIN = &H5C
    APPS = &H5D 'Like "WIN" act like left mouse click
    VOLMUTE = &HAD
    VOLUP = &HAF
    VOLDOWN = &HAE
    PLAY = &HFA
    NUMPAD0 = &H60
    NUMPAD1 = &H61
    NUMPAD2 = &H62
    NUMPAD3 = &H63
    NUMPAD4 = &H64
    NUMPAD5 = &H65
    NUMPAD6 = &H66
    NUMPAD7 = &H67
    NUMPAD8 = &H68
    NUMPAD9 = &H69
    NUMPADMULTIPLY = &H6A
    NUMPADADD = &H6B
    NUMPADSUBTRACT = &H6D
    NUMPADDIVIDE = &H6F
    NUMPADSEPARATOR = &H6C
    B0 = &H30
    B1 = &H31
    B2 = &H32
    B3 = &H33
    B4 = &H34
    B5 = &H35
    B6 = &H36
    B7 = &H37
    B8 = &H38
    B9 = &H39
    A = &H41
    B = &H42
    C = &H43
    D = &H44
    E = &H45
    F = &H46
    G = &H47
    H = &H48
    I = &H49
    J = &H4A
    K = &H4B
    L = &H4C
    M = &H4D
    N = &H4E
    O = &H4F
    P = &H50
    Q = &H51
    R = &H52
    S = &H53
    T = &H54
    U = &H55
    V = &H56
    W = &H57
    X = &H58
    Y = &H59
    Z = &H5A
    Zen = &HC0 '"~" key
    MINUS = &HBD '"-" Key not in numpad
    PLUS = &HBB '"+" Key not in numpad
    OPENPRAKETS = &HDB '"[" key
    CLOSEPRAKETS = &HDD '"]" key
    BACKSLASH = &HDC '"\" key
    SCOMMA = &HBA '";" key
    QUOTE = &HDE '"'" key
    COMMA = &HBC '"," Key
    PERIOD = &HBE '"." Key
    SLASH = &HBF '"/" key
End Enum
Public Class InpPoint
    Public m_X As Integer
    Public m_Y As Integer

    Public Sub New(ByVal X As Integer, ByVal Y As Integer)
        m_X = X
        m_Y = Y
    End Sub
    Public Sub New()
        m_X = 0
        m_Y = 0
    End Sub
End Class

Public Class DLLFns
    Private Const WM_GETTEXT As Integer = &HD 'used in SendMessage DLL function to get window title
    Private Const WM_SETTEXT As Integer = &HC 'used in SendMessage DLL function to set window text

#Region "DLL Import"
    <DllImport("user32.dll")>
    Private Shared Function GetAsyncKeyState(ByVal vKey As Integer) As Short
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SendInput(ByVal nInputs As UInteger, ByVal pInputs As Input(), ByVal cbSize As Integer) As UInteger
    End Function
    <DllImport("user32.dll")>
    Private Shared Function GetMessageExtraInfo() As IntPtr
    End Function
    <DllImport("user32.dll")>
    Private Shared Function GetCursorPos(<Out> ByRef lpPoint As SPOINT) As Boolean
    End Function
    <DllImport("User32.dll")>
    Private Shared Function SetCursorPos(ByVal x As Integer, ByVal y As Integer) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowRect(ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetTopWindow(ByVal hWnd As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function GetWindowText(ByVal hWnd As IntPtr, ByVal lpString As System.Text.StringBuilder, ByVal cch As Integer) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetKeyboardLayout(ByVal idThread As UInteger) As IntPtr
    End Function
    <DllImport("shell32.dll")>
    Private Shared Function SHGetKnownFolderPath(<MarshalAs(UnmanagedType.LPStruct)> ByVal rfid As Guid, ByVal dwFlags As UInt32, ByVal hToken As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> ByRef pszPath As System.Text.StringBuilder) As Int32
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function ShowWindow(ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function SetWindowPos(ByVal hWnd As IntPtr, ByVal hWndInsertAfternCmdShow As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal CX As Integer, ByVal CY As Integer, ByVal XuFlags As UInteger) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>'(hWnd, IntPtr.Zero, "Intermediate D3D Window", Nothing)
    Private Shared Function FindWindowEx(ByVal parentHandle As IntPtr, ByVal childAfter As IntPtr, ByVal lclassName As String, ByVal windowTitle As String) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>'e.g:FindWindow(Nothing, "*Untitled - Notepad")
    Private Shared Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>'(hWnd, WM_GETTEXT, 200, Hndl)
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>'SendMessage(Hwnd, WM_SETTEXT, 0, str as string)
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, <MarshalAs(UnmanagedType.LPWStr)> ByVal lParam As String) As IntPtr
    End Function

#End Region
#Region "Private Functions"
    Private Shared Function GetFolderPath(ByVal Optional FolderID As SpecialFolders = SpecialFolders.DOWNLOAD) As String
        Dim RetVal As String = ""

        Select Case FolderID
            Case SpecialFolders.DOWNLOAD
                Dim FolderDownloads As New Guid("374DE290-123F-4565-9164-39C4925E467B")
                Dim sb As New System.Text.StringBuilder(128)

                SHGetKnownFolderPath(FolderDownloads, 0, IntPtr.Zero, sb)
                RetVal = sb.ToString
            Case SpecialFolders.DESKTOP
                RetVal = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            Case SpecialFolders.DOCUMENT
                RetVal = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Case SpecialFolders.MYMUSIC
                RetVal = Environment.GetFolderPath(Environment.SpecialFolder.MyMusic)
            Case SpecialFolders.MYPICTURE
                RetVal = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
            Case SpecialFolders.MYVIDEO
                RetVal = Environment.GetFolderPath(Environment.SpecialFolder.MyVideos)
            Case SpecialFolders.PROGRAMFILES
                RetVal = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
            Case SpecialFolders.PROGRAMFILES86
                RetVal = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
            Case SpecialFolders.TEMP
                RetVal = Path.GetTempPath()
            Case SpecialFolders.USER
                RetVal = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
            Case SpecialFolders.WINDOWS
                RetVal = Environment.GetFolderPath(Environment.SpecialFolder.Windows)
            Case SpecialFolders.APP
                RetVal = Application.StartupPath
        End Select

        Return RetVal
    End Function
    Private Shared Function ReplaceCDate(ByVal InpStr As String, ByVal Format As String) As String
        Dim dt As Date = Date.Now
        Dim RetVal As String = InpStr

        Do While RetVal.IndexOf("{CDATE") > -1
            Dim FoundAt As Integer = RetVal.IndexOf("{CDATE")
            Dim StartStr As String = RetVal.Substring(0, FoundAt)
            Dim EndStr As String = RetVal.Substring(FoundAt + Len("{CDATE"))
            Dim EndAt As Integer = EndStr.IndexOf("}")
            Dim myDate As Date = dt.Date

            If Not EndAt = -1 Then
                Dim ExtDatesPeriod As String = EndStr.Substring(0, EndAt).ToLower
                Dim ExtDates As String = ""

                EndStr = EndStr.Substring(EndAt + Len("}"))
                If ExtDatesPeriod.Length > 0 Then
                    Try
                        ExtDates = ExtDatesPeriod.Substring(0, ExtDatesPeriod.Length - 1)
                        ExtDatesPeriod = ExtDatesPeriod.Substring(ExtDatesPeriod.Length - 1)
                        If ExtDatesPeriod = "y" Then ExtDatesPeriod = "yyyy"
                        If ExtDatesPeriod = "w" Then ExtDatesPeriod = "ww"
                        myDate = DateAdd(ExtDatesPeriod, Integer.Parse(ExtDates), dt)
                    Catch ex As Exception

                    End Try

                End If
            End If

            RetVal = myDate.ToString(Format, CultureInfo.CreateSpecificCulture("en-US"))
            RetVal = StartStr & RetVal & EndStr
        Loop

        Return RetVal
    End Function
    Private Shared Sub SendKeyboardInput(ByVal kbInputs As KeyboardInput())
        Dim inputs As Input() = New Input(kbInputs.Length - 1) {}

        For i As Integer = 0 To kbInputs.Length - 1
            inputs(i) = New Input With {
                    .type = CInt(InputType.Keyboard),
                    .u = New InputUnion With {
                        .ki = kbInputs(i)
                    }
                }
        Next

        SendInput(CUInt(inputs.Length), inputs, Marshal.SizeOf(GetType(Input)))
    End Sub
    Private Shared Sub SendMouseInput(ByVal mInputs As MouseInput())
        Dim inputs As Input() = New Input(mInputs.Length - 1) {}

        For i As Integer = 0 To mInputs.Length - 1
            inputs(i) = New Input With {
                    .type = CInt(InputType.Mouse),
                    .u = New InputUnion With {
                        .mi = mInputs(i)
                    }
                }
        Next

        SendInput(CUInt(inputs.Length), inputs, Marshal.SizeOf(GetType(Input)))
    End Sub
    'get list of Keyboardinput instant for list of VKeys
    Private Shared Function getKeyBoardInputs(ByVal Keys() As VKey) As KeyboardInput()
        Dim Index As Integer
        Dim KIs As New List(Of KeyboardInput)
        Dim tmpKI As KeyboardInput

        If IsNothing(Keys) Then Return Nothing

        For Index = 0 To Keys.Count - 1
            tmpKI = New KeyboardInput
            tmpKI.wVk = Keys(Index) : tmpKI.dwFlags = CUInt((KeyEventF.ExtendedKey Or KeyEventF.KeyDown))
            KIs.Add(tmpKI)
        Next
        For Index = 0 To Keys.Count - 1
            tmpKI = New KeyboardInput
            tmpKI.wVk = Keys(Index) : tmpKI.dwFlags = CUInt((KeyEventF.ExtendedKey Or KeyEventF.KeyUp))
            KIs.Add(tmpKI)
        Next
        Return KIs.ToArray
    End Function
    Private Shared Function GetVKey(ByVal Key As String, ByVal Optional CaseSensitive As Boolean = True) As VKey()
        Dim Index As Integer
        Dim DataFrom() As String = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "`", "~", "-", "_", "=", "+", "[", "{", "]", "}", "\", "|", ";", ":", "'", """", ",", "<", ".", ">", "/", "?", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", " ", "  ", "{ALT}", "{CTRL}", "{SHIFT}", "{WIN}", "{WIN2}", "{ESC}", "{PAGEUP}", "{PAGEDOWN}", "{F1}", "{F2}", "{F3}", "{F4}", "{F5}", "{F6}", "{F7}", "{F8}", "{F9}", "{F10}", "{F11}", "{F12}", "{TAB}", "{CAPS}", "{NUM}", "{INS}", "{DEL}", "{HOM}", "{END}", "{ENTER}", "{BAK}", "{LEFT}", "{RIGHT}", "{UP}", "{DOWN}"}
        Dim DataTo(,) As VKey = {{VKey.NA, VKey.A}, {VKey.NA, VKey.B}, {VKey.NA, VKey.C}, {VKey.NA, VKey.D}, {VKey.NA, VKey.E}, {VKey.NA, VKey.F}, {VKey.NA, VKey.G}, {VKey.NA, VKey.H}, {VKey.NA, VKey.I}, {VKey.NA, VKey.J}, {VKey.NA, VKey.K}, {VKey.NA, VKey.L}, {VKey.NA, VKey.M}, {VKey.NA, VKey.N}, {VKey.NA, VKey.O}, {VKey.NA, VKey.P}, {VKey.NA, VKey.Q}, {VKey.NA, VKey.R}, {VKey.NA, VKey.S}, {VKey.NA, VKey.T}, {VKey.NA, VKey.U}, {VKey.NA, VKey.V}, {VKey.NA, VKey.W}, {VKey.NA, VKey.X}, {VKey.NA, VKey.Y}, {VKey.NA, VKey.Z}, {VKey.NA, VKey.B0}, {VKey.NA, VKey.B1}, {VKey.NA, VKey.B2}, {VKey.NA, VKey.B3}, {VKey.NA, VKey.B4}, {VKey.NA, VKey.B5}, {VKey.NA, VKey.B6}, {VKey.NA, VKey.B7}, {VKey.NA, VKey.B8}, {VKey.NA, VKey.B9}, {VKey.NA, VKey.Zen}, {VKey.SHIFT, VKey.Zen}, {VKey.NA, VKey.MINUS}, {VKey.SHIFT, VKey.MINUS}, {VKey.NA, VKey.PLUS}, {VKey.SHIFT, VKey.PLUS}, {VKey.NA, VKey.OPENPRAKETS}, {VKey.SHIFT, VKey.OPENPRAKETS}, {VKey.NA, VKey.CLOSEPRAKETS}, {VKey.SHIFT, VKey.CLOSEPRAKETS}, {VKey.NA, VKey.BACKSLASH}, {VKey.SHIFT, VKey.BACKSLASH}, {VKey.NA, VKey.SCOMMA}, {VKey.SHIFT, VKey.SCOMMA}, {VKey.NA, VKey.QUOTE}, {VKey.SHIFT, VKey.QUOTE}, {VKey.NA, VKey.COMMA}, {VKey.SHIFT, VKey.COMMA}, {VKey.NA, VKey.PERIOD}, {VKey.SHIFT, VKey.PERIOD}, {VKey.NA, VKey.SLASH}, {VKey.SHIFT, VKey.SLASH}, {VKey.SHIFT, VKey.B1}, {VKey.SHIFT, VKey.B2}, {VKey.SHIFT, VKey.B3}, {VKey.SHIFT, VKey.B4}, {VKey.SHIFT, VKey.B5}, {VKey.SHIFT, VKey.B6}, {VKey.SHIFT, VKey.B7}, {VKey.SHIFT, VKey.B8}, {VKey.SHIFT, VKey.B9}, {VKey.SHIFT, VKey.B0}, {VKey.NA, VKey.SPACE}, {VKey.NA, VKey.TAB}, {VKey.NA, VKey.ALT}, {VKey.NA, VKey.CONTROL}, {VKey.NA, VKey.SHIFT}, {VKey.NA, VKey.WIN}, {VKey.NA, VKey.APPS}, {VKey.NA, VKey.ESCAPE}, {VKey.NA, VKey.PAGEUP}, {VKey.NA, VKey.PAGEDOWN}, {VKey.NA, VKey.F1}, {VKey.NA, VKey.F2}, {VKey.NA, VKey.F3}, {VKey.NA, VKey.F4}, {VKey.NA, VKey.F5}, {VKey.NA, VKey.F6}, {VKey.NA, VKey.F7}, {VKey.NA, VKey.F8}, {VKey.NA, VKey.F9}, {VKey.NA, VKey.F10}, {VKey.NA, VKey.F11}, {VKey.NA, VKey.F12}, {VKey.NA, VKey.TAB}, {VKey.NA, VKey.CAPSLOCK}, {VKey.NA, VKey.NUMLOCK}, {VKey.NA, VKey.INSERT}, {VKey.NA, VKey.DELETE}, {VKey.NA, VKey.HOME}, {VKey.NA, VKey.BEND}, {VKey.NA, VKey.Enter}, {VKey.NA, VKey.BAK}, {VKey.NA, VKey.LEFT}, {VKey.NA, VKey.RIGHT}, {VKey.NA, VKey.UP}, {VKey.NA, VKey.DOWN}}
        Dim UKey As String = Key.ToUpper

        For Index = 0 To DataFrom.Count - 1
            If UKey = DataFrom(Index) Then
                If DataTo(Index, 0) = VKey.NA Then
                    If CaseSensitive And Index < 25 Then 'in case of upper key
                        If UKey = Key Then 'if Capital Needed then add Shift if capslock not active 
                            If Control.IsKeyLocked(Keys.CapsLock) Then Return {DataTo(Index, 1)} Else Return {VKey.SHIFT, DataTo(Index, 1)}
                        Else
                            If Control.IsKeyLocked(Keys.CapsLock) Then Return {VKey.SHIFT, DataTo(Index, 1)} Else Return {DataTo(Index, 1)}
                        End If
                    Else
                        Return {DataTo(Index, 1)}
                    End If
                Else
                    Return {DataTo(Index, 0), DataTo(Index, 1)}
                End If
            End If
        Next

        Return Nothing
    End Function
    Private Shared Function GetCurrentLang() As String
        Dim ProcIdXL As Integer = 0
        Dim hWnd As IntPtr = GetForegroundWindow()
        Dim ProcID As Integer = GetWindowThreadProcessId(hWnd, ProcIdXL)
        Dim culturePtr As IntPtr = GetKeyboardLayout(ProcID)
        Dim Index As Integer

        For Index = 0 To InputLanguage.InstalledInputLanguages.Count - 1
            Dim myLang As InputLanguage = InputLanguage.InstalledInputLanguages(Index)
            If myLang.Handle = culturePtr Then
                Return myLang.Culture.TwoLetterISOLanguageName.ToLower
            End If
        Next

        Return ""
        'Dim xproc As Process
        'xproc = Process.GetProcessById(ProcIdXL)
        'xproc.Threads(Index).CurrentUICulture = New CultureInfo("en-US")
        'System.Threading.Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")
    End Function
    Public Shared Sub Sleep(ByVal Optional Waitms As Integer = 100) 'shared, don't check m_Imergancy
        Dim Index As Double

        If Waitms = 0 Then Return
        For Index = 0 To Int(Waitms / 10)
            Application.DoEvents()
            Threading.Thread.Sleep(10)
        Next
    End Sub
#End Region
#Region "Helper Function"
    Public Shared Function ReplaceConst(ByVal InpStr As String, ByVal Format As String) As String
        Dim dt As Date = Date.Now
        Dim RetVal As String = InpStr

        RetVal = RetVal.Replace("{DOWNLOAD}", GetFolderPath())
        RetVal = RetVal.Replace("{DOCUMENT}", GetFolderPath(SpecialFolders.DOCUMENT))
        RetVal = RetVal.Replace("{MYPICTURE}", GetFolderPath(SpecialFolders.MYPICTURE))
        RetVal = RetVal.Replace("{MYMUSIC}", GetFolderPath(SpecialFolders.MYMUSIC))
        RetVal = RetVal.Replace("{MYVIDEO}", GetFolderPath(SpecialFolders.MYVIDEO))
        RetVal = RetVal.Replace("{USER}", GetFolderPath(SpecialFolders.USER))
        RetVal = RetVal.Replace("{DESKTOP}", GetFolderPath(SpecialFolders.DESKTOP))
        RetVal = RetVal.Replace("{WINDOWS}", GetFolderPath(SpecialFolders.WINDOWS))
        RetVal = RetVal.Replace("{PROGRAMFILES}", GetFolderPath(SpecialFolders.PROGRAMFILES))
        RetVal = RetVal.Replace("{PROGRAMFILES86}", GetFolderPath(SpecialFolders.PROGRAMFILES86))
        RetVal = RetVal.Replace("{TEMP}", GetFolderPath(SpecialFolders.TEMP))
        RetVal = RetVal.Replace("{APP}", GetFolderPath(SpecialFolders.APP))
        RetVal = ReplaceCDate(RetVal, Format)

        Return RetVal
    End Function
    Public Shared Function GetActiveWinTitle() As String()
        Dim ProcIdXL As Integer = 0
        Dim hWnd As IntPtr = GetForegroundWindow()
        Dim ProcID As Integer = GetWindowThreadProcessId(hWnd, ProcIdXL)
        Dim title As New System.Text.StringBuilder(256)
        GetWindowText(hWnd, title, title.Capacity)
        hWnd = GetTopWindow(hWnd)
        Return {hWnd.ToString, title.ToString} 'Proc.MainModule.FileName
    End Function
    Public Shared Function GetActiveWinPos() As InpPoint
        Dim hWnd As IntPtr = GetForegroundWindow()
        Dim rct As RECT

        If GetWindowRect(hWnd, rct) Then
            Return New InpPoint(rct.Left, rct.Top)
        Else
            Return Nothing
        End If
    End Function
    Public Shared Function GetCursorPosition() As SPOINT
        Dim point As SPOINT = Nothing
        GetCursorPos(point)
        Return point
    End Function
    Public Shared Sub SetCursorPosition(ByVal x As Integer, ByVal y As Integer)
        SetCursorPos(x, y)
    End Sub
    Public Shared Sub ClickKey(ByVal scanCode As UShort)
        Dim inputs = New KeyboardInput() {New KeyboardInput With {
                .wScan = scanCode,
                .dwFlags = CUInt((KeyEventF.KeyDown Or KeyEventF.Scancode)),
                .dwExtraInfo = GetMessageExtraInfo()
            }, New KeyboardInput With {
                .wScan = scanCode,
                .dwFlags = CUInt((KeyEventF.KeyUp Or KeyEventF.Scancode)),
                .dwExtraInfo = GetMessageExtraInfo()
            }}
        SendKeyboardInput(inputs)
    End Sub
    'Take screenshot starting from imgX,imgY with size=(imgwidth,imgHeight)
    Public Shared Function TakeScreenShot(ByVal Optional imgX As Integer = 0, ByVal Optional imgY As Integer = 0, ByVal Optional imgWidth As Integer = -1, ByVal Optional imgHeight As Integer = -1) As Bitmap
        Dim screenSize As Size
        Dim screenGrab As Bitmap
        Dim g As Graphics
        Dim iWidth As Integer = My.Computer.Screen.Bounds.Width
        Dim iHeight As Integer = My.Computer.Screen.Bounds.Height

        If imgWidth > 0 And imgWidth < iWidth Then iWidth = imgWidth
        If imgHeight > 0 And imgHeight < iHeight Then iHeight = imgHeight

        screenSize = New Size(iWidth, iHeight)
        screenGrab = New Bitmap(iWidth, iHeight, PixelFormat.Format24bppRgb)
        g = Graphics.FromImage(screenGrab)
        g.CopyFromScreen(New Point(imgX, imgY), New Point(0, 0), screenSize)
        Return screenGrab
    End Function
    'Get URL result
    Public Shared Function GetURL(ByVal myURL As String) As String()
        Dim RetVal() As String
        Dim webRequest As WebRequest
        Dim webresponse As WebResponse
        Dim inStream As StreamReader

        Try
            webRequest = WebRequest.Create(myURL)
            webresponse = webRequest.GetResponse()

            inStream = New StreamReader(webresponse.GetResponseStream())
            RetVal = inStream.ReadToEnd().Replace(vbLf, "").Split(vbNewLine)

            Return RetVal
        Catch ex As Exception

        End Try
    End Function
    'Search Image in Image
    Public Shared Function searchImage(ByRef Src As Bitmap, ByRef Search As Bitmap, ByVal Optional margin As Integer = 15) As InpPoint

        If IsNothing(Src) Or IsNothing(Search) Then Return Nothing
        If Src.Width < Search.Width OrElse Src.Height < Search.Height Then Return Nothing

        '-- Prepare optimizations
        Dim CMargin As Integer = 3 * margin
        Dim sr As New Rectangle(0, 0, Src.Width, Src.Height)
        Dim br As New Rectangle(0, 0, Search.Width, Search.Height)
        Dim srcLock As BitmapData = Src.LockBits(sr, Imaging.ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb)
        Dim bmpLock As BitmapData = Search.LockBits(br, Imaging.ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb)
        Dim sStride As Integer = srcLock.Stride
        Dim bStride As Integer = bmpLock.Stride
        Dim srcSz As Integer = sStride * Src.Height
        Dim bmpSz As Integer = bStride * Search.Height
        Dim srcBuff(srcSz) As Byte
        Dim bmpBuff(bmpSz) As Byte
        Dim x, y, x2, y2, sx, sy, bx, by, sw, sh, bw, bh As Integer
        Dim dr, dg, db As Integer
        Dim p As InpPoint = Nothing


        Marshal.Copy(srcLock.Scan0, srcBuff, 0, srcSz)
        Marshal.Copy(bmpLock.Scan0, bmpBuff, 0, bmpSz)
        ' we don't need to lock the image anymore as we have a local copy
        Search.UnlockBits(bmpLock)
        Src.UnlockBits(srcLock)

        bw = Search.Width
        bh = Search.Height
        sw = Src.Width - bw      ' limit scan to only what we need. the extra corner
        sh = Src.Height - bh     ' point we need is taken care of in the loop itself.
        bx = 0 : by = 0

        '-- Scan source for bitmap
        For y = 0 To sh
            sy = y * sStride
            For x = 0 To sw
                sx = sy + x * 3
                '-- Find start point/pixel
                dr = srcBuff(sx + 2) : dr -= bmpBuff(2) : dr = Math.Abs(dr)
                dg = srcBuff(sx + 1) : dg -= bmpBuff(1) : dg = Math.Abs(dg)
                db = srcBuff(sx + 0) : db -= bmpBuff(0) : db = Math.Abs(db)

                If (dr <= margin AndAlso dg <= margin AndAlso db <= margin) OrElse (dr + dg + db) < CMargin Then '-- We have a pixel match, check the region
                    p = New InpPoint(x, y)
                    '-- We have a pixel match, check the region
                    For y2 = 0 To bh - 1
                        by = y2 * bStride
                        For x2 = 0 To bw - 1
                            bx = by + x2 * 3
                            sy = (y + y2) * sStride
                            sx = sy + (x + x2) * 3

                            dr = srcBuff(sx + 2) : dr -= bmpBuff(bx + 2) : dr = Math.Abs(dr)
                            dg = srcBuff(sx + 1) : dg -= bmpBuff(bx + 1) : dg = Math.Abs(dg)
                            db = srcBuff(sx + 0) : db -= bmpBuff(bx + 0) : db = Math.Abs(db)
                            If (dr > margin OrElse dg > margin OrElse db > margin) AndAlso (dr + dg + db) >= CMargin Then '-- Not matching, continue checking                                '
                                p = Nothing
                                sy = y * sStride
                                Exit For
                            End If

                        Next
                        If IsNothing(p) Then Exit For
                    Next
                End If
                If Not IsNothing(p) Then Exit For
            Next
            If Not IsNothing(p) Then Exit For
        Next
        bmpBuff = Nothing
        srcBuff = Nothing

        Return p
    End Function
    'recomended rather than compareimage2 as faster than compareImage2 (4ms:6ms) also could add tolerance
    Public Shared Function CompareImages(ByRef srcImg As Bitmap, ByRef trgImg As Bitmap) As Boolean
        Dim srcRect As New Rectangle(0, 0, srcImg.Width, srcImg.Height)
        Dim trgRect As New Rectangle(0, 0, trgImg.Width, trgImg.Height)
        Dim srcData As System.Drawing.Imaging.BitmapData = srcImg.LockBits(srcRect, Imaging.ImageLockMode.ReadWrite, srcImg.PixelFormat)
        Dim trgData As System.Drawing.Imaging.BitmapData = trgImg.LockBits(trgRect, Imaging.ImageLockMode.ReadWrite, trgImg.PixelFormat)
        Dim srcLength As Integer = (srcData.Height * srcData.Stride) - 1
        Dim trgLength As Integer = (trgData.Height * trgData.Stride) - 1
        Dim srcBytes(srcLength) As Byte
        Dim trgBytes(trgLength) As Byte

        System.Runtime.InteropServices.Marshal.Copy(srcData.Scan0, srcBytes, 0, srcBytes.Length)
        System.Runtime.InteropServices.Marshal.Copy(trgData.Scan0, trgBytes, 0, trgBytes.Length)

        srcImg.UnlockBits(srcData)
        trgImg.UnlockBits(trgData)

        If srcBytes.Length <> trgBytes.Length Then Return False
        For i As Integer = 0 To srcBytes.Length - 1
            If Not srcBytes(i) = trgBytes(i) Then Return False
        Next

        Return True
    End Function
    Public Shared Function CompareImages2(ByRef srcImg As Bitmap, ByRef trgImg As Bitmap) As Boolean
        Dim ms1 As New MemoryStream
        Dim ms2 As New MemoryStream
        srcImg.Save(ms1, ImageFormat.Bmp)
        srcImg.Save(ms1, ImageFormat.Bmp)

        For i = 0 To CInt(ms1.Length) - 1
            If ms1.ReadByte() <> ms2.ReadByte Then Return False
        Next

        Return True
    End Function
#End Region

#Region "Public Function"
    'Check if any key from VKey structure pressed
    Public Shared Function isKeyPressed(ByVal KeyCode As VKey) As Boolean
        Return GetAsyncKeyState(KeyCode)
    End Function
    'Change current language in the taskbar
    Public Shared Sub DO_SetLang(ByVal NeedLang As String)
        Dim TotalLang As Integer = InputLanguage.InstalledInputLanguages.Count

        For Index = 0 To TotalLang - 1
            Dim CurrLang As String = GetCurrentLang()
            If Not CurrLang = NeedLang.ToLower Then
                DO_PressKeys({VKey.ALT, VKey.SHIFT})
                Sleep()
            Else
                Exit For
            End If
        Next
    End Sub
    'Move mouse to absolute position
    Public Shared Sub DO_MouseMoveAbs(ByVal X As Integer, ByVal Y As Integer, ByVal Optional HitClick As Boolean = False)
        Dim fx As Double = 65536 / (Screen.PrimaryScreen.Bounds.Width - 1)
        Dim fy As Double = 65536 / (Screen.PrimaryScreen.Bounds.Height - 1)
        Dim MI As New MouseInput()

        MI.dx = X * fx
        MI.dy = Y * fy
        MI.dwFlags = CUInt(MouseEventF.Move Or MouseEventF.Absolute) 'Absolute
        SendMouseInput({MI})
        If HitClick Then DO_MouseClick()
    End Sub
    'Move mouse to Relative position
    Public Shared Sub DO_MouseMoveRel(ByVal X As Integer, ByVal Y As Integer, ByVal Optional HitClick As Boolean = False)
        Dim fx As Double = 65536 / (Screen.PrimaryScreen.Bounds.Width - 1)
        Dim fy As Double = 65536 / (Screen.PrimaryScreen.Bounds.Height - 1)
        Dim MI As New MouseInput()
        Dim point As SPOINT = GetCursorPosition()

        MI.dx = (point.X + X) * fx
        MI.dy = (point.Y + Y) * fy
        MI.dwFlags = CUInt(MouseEventF.Move Or MouseEventF.Absolute) 'Move relative
        SendMouseInput({MI})
        If HitClick Then DO_MouseClick()
    End Sub
    'Fire Mouse Click
    Public Shared Sub DO_MouseClick()
        Dim MI As New MouseInput

        MI.dwFlags = CUInt(MouseEventF.LeftUp Or MouseEventF.LeftDown) 'Click
        SendMouseInput({MI})
    End Sub
    'Fire Mouse left Down
    Public Shared Sub DO_MouseDrag(ByVal X As Integer, ByVal Y As Integer)
        Dim fx As Double = 65536 / (Screen.PrimaryScreen.Bounds.Width - 1)
        Dim fy As Double = 65536 / (Screen.PrimaryScreen.Bounds.Height - 1)
        Dim MI As New MouseInput()

        MI.dwFlags = CUInt(MouseEventF.LeftDown) : SendMouseInput({MI})
        MI.dx = X * fx : MI.dy = Y * fy
        MI.dwFlags = CUInt(MouseEventF.Move Or MouseEventF.Absolute) 'Absolute
        SendMouseInput({MI})
        MI.dwFlags = CUInt(MouseEventF.LeftUp) : SendMouseInput({MI})
    End Sub
    'Fire right Mouse Click
    Public Shared Sub DO_MouseRClick()
        Dim MI As New MouseInput

        MI.dwFlags = CUInt(MouseEventF.RightUp Or MouseEventF.RightDown) 'Click
        SendMouseInput({MI})
    End Sub
    'Press special Key sequance like {WIN}+R or {ALT}+{SHIFT}+{ESC}
    Public Shared Sub DO_PressKeys(ByVal Keys() As VKey, ByVal Optional Waitms As Integer = 100)
        Dim KIs() As KeyboardInput = getKeyBoardInputs(Keys)
        If Not IsNothing(KIs) Then SendKeyboardInput(KIs)
        Sleep(Waitms)
    End Sub
    Public Shared Sub DO_PressKeys(ByVal Keys As String)
        Dim Values() As String = Keys.Split("+")
        Dim Index1, Index2 As Integer
        Dim KIs As New List(Of VKey)

        For Index1 = 0 To Values.Count - 1
            Dim myKey() As VKey = GetVKey(Values(Index1), False)
            If Not IsNothing(myKey) Then
                For Index2 = 0 To myKey.Count - 1
                    If Not myKey(Index2) = VKey.NA Then KIs.Add(myKey(Index2))
                Next
            End If
        Next
        DO_PressKeys(KIs.ToArray)
    End Sub
    'Press Keys and then enter
    Public Shared Sub DO_Type(ByVal InpStr As String, ByVal Optional DateFormat As String = "dd/MM/yyy hh:mm:ss", ByVal Optional HitEnter As Boolean = False)
        Dim Index As Integer

        InpStr = ReplaceConst(InpStr, DateFormat)
        For Index = 0 To InpStr.Length - 1
            Dim MyChar As String = InpStr.Substring(Index, 1)
            If MyChar = "{" Then 'catch if special key passed
                Dim Found As Integer = InpStr.IndexOf("}", Index)
                If Not Found = -1 Then
                    Dim Mystr As String = InpStr.Substring(Index, Found - Index + 1)
                    Dim MytmpKeys() As VKey = GetVKey(Mystr)
                    If Not IsNothing(MytmpKeys) Then
                        MyChar = Mystr
                        Index += Mystr.Length - 1
                    End If
                End If
            End If
            Dim MyKeys() As VKey = GetVKey(MyChar)
            If Not IsNothing(MyKeys) Then DO_PressKeys(MyKeys, 0)
        Next
        If HitEnter Then DO_PressKeys({VKey.Enter})
    End Sub
#End Region

End Class
