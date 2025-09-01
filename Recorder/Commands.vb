Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Net
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices

Public Class Commands
    Private LoopFlag As Boolean = False
    Private LoopCmds As New List(Of String)
    Private LoopCount As Integer = 0
    Private RunFlag As Boolean = True 'used in CHECKCLIP Command

    Public m_OldPnt() As Integer = {-1, -1}
    Public m_Lang As String = ""
    Public m_DateFormat As String = "dd-MM-yyyy"
    Public m_ImerganceBreak As Boolean = False
    Public m_MaxLines As Integer
    Public m_CurLine As Integer
    Public m_Status As Integer '0:Still Running,-1:IDLE,>2 Error in Line,<1 Emergency Break at line

#Region "Private Function"
    'Check if imergency stop fired or ESC key pressed
    Private Function checkStop() As Boolean
        Application.DoEvents()
        If DLLFns.isKeyPressed(27) Then m_ImerganceBreak = True 'Esc = End running script
        If m_ImerganceBreak Then Return True
        Return False
    End Function
    'used in "Sleep", Hult for 10ms and check emergency break flag
    Private Function DoSleep(ByVal Optional Waitms As Integer = 10) As Boolean
        For Index = 0 To Int(Waitms / 10)
            If checkStop() Then Return False
            Threading.Thread.Sleep(10)
        Next

        If checkStop() Then Return False 'in case Waitms=0
        Return True
    End Function
    'Search desktop for Image within Waitms (or -1 for forever)
    Private Function SearchScreenCaptureImage(ByRef SrchImag As Bitmap, ByVal Optional Waitms As Double = -1) As InpPoint
        Dim myStopwatch As Stopwatch = Stopwatch.StartNew()

        If Waitms = -1 Then Waitms = 10000 : myStopwatch.Stop()

        Do While Waitms - myStopwatch.ElapsedMilliseconds > 0
            Dim ImgBitmap As Bitmap = DLLFns.TakeScreenShot()
            Dim myPnt As InpPoint = DLLFns.searchImage(ImgBitmap, SrchImag)

            If Not IsNothing(myPnt) Then
                myStopwatch = Nothing
                Return myPnt
            End If
            If Not DoSleep() Then Exit Do
        Loop

        myStopwatch = Nothing
        Return Nothing
    End Function
    'Check if desktop screen capture image changed starting from point (imgX,imgY) sized (srchImg.width,srchImg.height) within waitms
    Private Function checkImageChanged(ByVal imgX As Integer, ByVal imgY As Integer, ByRef srcImag As Bitmap, ByVal Optional Waitms As Double = -1) As Boolean
        Dim myStopwatch As Stopwatch = Stopwatch.StartNew()

        If Waitms = -1 Then Waitms = 10000 : myStopwatch.Stop()

        Do While Waitms - myStopwatch.ElapsedMilliseconds > 0
            Dim ImgBitmap As Bitmap = DLLFns.TakeScreenShot(imgX, imgY, srcImag.Width, srcImag.Height)
            If Not DLLFns.CompareImages(srcImag, ImgBitmap) Then
                myStopwatch = Nothing
                ImgBitmap.Dispose()
                Return True
            Else
                ImgBitmap.Dispose()
            End If
            If Not DoSleep() Then Exit Do
        Loop

        myStopwatch = Nothing
        Return False
    End Function
    'Check active window title contain text within Waitms (or -1 for forever)
    Private Function CheckWinTitle(ByVal Title As String, ByVal Optional Waitms As Double = -1, ByVal Optional MustChange As Boolean = False) As InpPoint
        Dim myStopwatch As Stopwatch = Stopwatch.StartNew()
        Dim myPnt As InpPoint = Nothing
        Dim OldHWND As String = ""

        If Waitms = -1 Then Waitms = 10000 : myStopwatch.Stop()
        Do While IsNothing(myPnt) Or Waitms - myStopwatch.ElapsedMilliseconds > 0
            Dim WinTitle() As String = DLLFns.GetActiveWinTitle()
            If WinTitle(1).ToLower.Contains(Title.ToLower) Then
                If OldHWND = "" And MustChange Then OldHWND = WinTitle(0) & WinTitle(1)
                If Not OldHWND = WinTitle(0) & WinTitle(1) Then
                    myPnt = DLLFns.GetActiveWinPos()
                    If Not IsNothing(myPnt) Then myStopwatch = Nothing : Return myPnt
                End If
            End If

            If Not DoSleep() Then Return Nothing
        Loop

        myStopwatch = Nothing
        Return Nothing
    End Function
#End Region
#Region "Command Functions"
    Private Sub Cmd_Move(ByVal X As Double, ByVal Y As Double, ByVal Optional ClickFlag As Boolean = False)
        Dim point As SPOINT = DLLFns.GetCursorPosition()
        If X < 0 Then X = point.X
        If Y < 0 Then Y = point.Y
        DLLFns.DO_MouseMoveAbs(X, Y, ClickFlag)
        m_OldPnt = {X, Y}
    End Sub
    Private Sub Cmd_MoveRel(ByVal X As Double, ByVal Y As Double, ByVal Optional ClickFlag As Boolean = False)
        Dim point As SPOINT = DLLFns.GetCursorPosition()
        If m_OldPnt(0) = -1 And m_OldPnt(1) = -1 Then m_OldPnt = {point.X, point.Y}
        m_OldPnt(0) += X
        m_OldPnt(1) += Y
        DLLFns.DO_MouseMoveAbs(m_OldPnt(0), m_OldPnt(1))
    End Sub
    Private Sub Cmd_Drag(ByVal X As Double, ByVal Y As Double)
        Dim point As SPOINT = DLLFns.GetCursorPosition()
        If X < 0 Then X = point.X
        If Y < 0 Then Y = point.Y
        DLLFns.DO_MouseDrag(X, Y)
        m_OldPnt = {X, Y}
    End Sub
    Private Sub Cmd_DragRel(ByVal X As Double, ByVal Y As Double)
        Dim point As SPOINT = DLLFns.GetCursorPosition()
        If m_OldPnt(0) = -1 And m_OldPnt(1) = -1 Then m_OldPnt = {point.X, point.Y}
        m_OldPnt(0) += X
        m_OldPnt(1) += Y
        DLLFns.DO_MouseDrag(m_OldPnt(0), m_OldPnt(1))
    End Sub
    Private Sub Cmd_CheckClip(ByVal DataStr As String)
        Dim ClipBoard As String = My.Computer.Clipboard.GetText()
        If Not Regex.Match(ClipBoard, DataStr, RegexOptions.IgnoreCase).Success Then RunFlag = False
    End Sub
    Private Sub Cmd_CheckWin(ByVal DataStr As String, ByVal Optional Waitms As Double = -1, ByVal Optional MustChange As Boolean = False) 'Waitms=-1 wait forever
        Dim mypnt As InpPoint = CheckWinTitle(DataStr, Waitms, MustChange)

        If IsNothing(mypnt) Then
            RunFlag = False 'in case not found
        Else
            DLLFns.DO_MouseMoveAbs(mypnt.m_X, mypnt.m_Y)
            m_OldPnt = {mypnt.m_X, mypnt.m_Y}
        End If

    End Sub
    Private Sub Cmd_CheckImg(ByVal DataStr As String, ByVal Optional Waitms As Double = -1) 'Waitms=-1 wait forever
        Dim ImgFile As String = DLLFns.ReplaceConst(DataStr, m_DateFormat)
        Dim SrchImag As Bitmap
        Dim myPnt As InpPoint

        If Not File.Exists(ImgFile) Then RunFlag = False : Return
        SrchImag = New Bitmap(ImgFile)
        myPnt = SearchScreenCaptureImage(SrchImag, Waitms)

        If IsNothing(myPnt) Then
            RunFlag = False 'in case not found
        Else
            DLLFns.DO_MouseMoveAbs(myPnt.m_X, myPnt.m_Y)
            m_OldPnt = {myPnt.m_X, myPnt.m_Y}
        End If

        SrchImag.Dispose() : SrchImag = Nothing
    End Sub
    Private Sub Cmd_CheckImgChange(ByVal IWidth As Integer, ByVal IHeight As Integer, ByVal Optional Waitms As Double = -1) 'Waitms=-1 wait forever
        Dim point As SPOINT = DLLFns.GetCursorPosition()
        Dim SrchImag As Bitmap = DLLFns.TakeScreenShot(point.X, point.Y, IWidth, IHeight)


        If Not checkImageChanged(point.X, point.Y, SrchImag, Waitms) Then
            RunFlag = False 'in case not changed
        End If

        SrchImag.Dispose() : SrchImag = Nothing
    End Sub
    Private Function Cmd_LoadURL(ByVal myURL As String) As Integer
        Dim Index As Integer
        Dim Cmds() As String = DLLFns.GetURL(myURL)

        If Cmds Is Nothing Then Return -1
        For Index = 0 To Cmds.Count - 1
            If DoCmd(Cmds(Index)) = -1 Then Return -1
        Next

        Return 0
    End Function
    Private Function Cmd_Loop() As Integer
        Dim RetVal, Index1, Index2 As Integer

        'm_MaxLines += (LoopCount - 2) * (LoopCmds.Count - 2)
        For Index1 = 0 To LoopCount - 2
            m_CurLine -= LoopCmds.Count - 2
            For Index2 = 0 To LoopCmds.Count - 2 'last cmd="LOOPEND"
                RetVal = DoCmd(LoopCmds(Index2))
                If RetVal < 0 Then Return RetVal
                m_CurLine += 1 ' Index2 + (Index1 * (LoopCmds.Count - 2))
            Next
        Next

        Return 0
    End Function

    'Check help.txt for more details about command structure, return 0:Move to next command,-1:Error,-2:force break,>0: Jump to line
    Private Function DoCmd(ByVal CMD As String) As Integer
        Dim myCMD As New Command(CMD)
        Dim RetVal As Integer = -1

        If checkStop() Then Return -2
        If LoopFlag Then LoopCmds.Add(CMD) 'in case of in Loop then add the command (even if it's comements)
        If myCMD.m_Command = "" Then Return 0
        If Not myCMD.m_Command = "ENDCHECK" And Not RunFlag Then Return 0 'incase RunFlag=true the don't do commands other than "Endcheck"

        If Not m_Lang = "" Then DLLFns.DO_SetLang(m_Lang) 'each time run command be sure language is correct as it might been changed by user
        Select Case myCMD.m_Command
            Case "RUNURL"
                RetVal = Cmd_LoadURL(myCMD.m_Text)
            Case "LOOP"
                LoopCmds.Clear() : LoopFlag = True : LoopCount = myCMD.m_Number : RetVal = 0
            Case "LOOPEND"
                LoopFlag = False : RetVal = Cmd_Loop()
            Case "CHECKCLIP" 'Check clipboard
                Cmd_CheckClip(myCMD.m_Text) : RetVal = 0
            Case "CHECKWIN" 'Check Current Window Title
                Cmd_CheckWin(myCMD.m_Text, myCMD.m_Waitms, False) : myCMD.m_Waitms = 0 : RetVal = 0
            Case "CHECKWINCHANGED" 'Check Current Window Title changed
                Cmd_CheckWin(myCMD.m_Text, myCMD.m_Waitms, True) : myCMD.m_Waitms = 0 : RetVal = 0
            Case "CHECKIMG"
                Cmd_CheckImg(myCMD.m_Text, myCMD.m_Waitms) : myCMD.m_Waitms = 0 : RetVal = 0
            Case "CHECKIMGCHANGE"
                Cmd_CheckImgChange(myCMD.m_X, myCMD.m_Y, myCMD.m_Waitms) : myCMD.m_Waitms = 0 : RetVal = 0
            Case "ENDCHECK"
                RunFlag = True : RetVal = 0
            Case "JUMP" 'JUmp to Non Zero Based line (1,2,3,...)
                RetVal = myCMD.m_Number
                If RetVal < 0 Then RetVal = 0
            Case "MOVE"
                Cmd_Move(myCMD.m_X, myCMD.m_Y) : RetVal = 0
            Case "MOVE&CLICK"
                Cmd_Move(myCMD.m_X, myCMD.m_Y, True) : RetVal = 0
            Case "MOVEREL"
                Cmd_MoveRel(myCMD.m_X, myCMD.m_Y) : RetVal = 0
            Case "MOVEREL&CLICK"
                Cmd_MoveRel(myCMD.m_X, myCMD.m_Y, True) : RetVal = 0
            Case "CLICK"
                DLLFns.DO_MouseClick() : RetVal = 0
            Case "RCLICK"
                DLLFns.DO_MouseRClick() : RetVal = 0
            Case "DRAG"
                Cmd_Drag(myCMD.m_X, myCMD.m_Y) : RetVal = 0
            Case "DRAGREL"
                Cmd_DragRel(myCMD.m_X, myCMD.m_Y) : RetVal = 0
            Case "KEY_PRESS"
                DLLFns.DO_PressKeys(myCMD.m_Text) : RetVal = 0
            Case "TYPE"
                DLLFns.DO_Type(myCMD.m_Text, m_DateFormat, True) : RetVal = 0
            Case "TYPENE"
                DLLFns.DO_Type(myCMD.m_Text, m_DateFormat) : RetVal = 0
            Case "SET"
                Dim UpCMD As String = myCMD.m_Text.ToUpper
                If UpCMD = "CAPSON" And Not Control.IsKeyLocked(Keys.CapsLock) Then DLLFns.DO_PressKeys({VKey.CAPSLOCK})
                If UpCMD = "CAPSOFF" And Control.IsKeyLocked(Keys.CapsLock) Then DLLFns.DO_PressKeys({VKey.CAPSLOCK})
                If UpCMD = "NUMON" And Not Control.IsKeyLocked(Keys.NumLock) Then DLLFns.DO_PressKeys({VKey.NUMLOCK})
                If UpCMD = "NUMOFF" And Control.IsKeyLocked(Keys.NumLock) Then DLLFns.DO_PressKeys({VKey.NUMLOCK})
                If UpCMD.StartsWith("LANG") Then m_Lang = myCMD.m_Text.Substring(4).ToLower : DLLFns.DO_SetLang(m_Lang)
                If UpCMD.StartsWith("DATE") Then m_DateFormat = myCMD.m_Text.Substring(4).Replace(" ", "")
                RetVal = 0
        End Select

        If RetVal = 0 Then
            If Not DoSleep(myCMD.m_Waitms) Then RetVal = -2
        End If
        Return RetVal
    End Function
#End Region

#Region "Public Function"
    'read file and do command
    Public Sub DoFile(ByVal FilePath As String)
        Dim reader As IO.StreamReader = New IO.StreamReader(FilePath)
        Dim Lines() As String = reader.ReadToEnd().Replace(vbLf, "").Split(vbNewLine)
        Dim Index, Result As Integer

        reader.Close()
        m_MaxLines = Lines.Count
        m_ImerganceBreak = False
        RunFlag = True
        For Index = 0 To Lines.Count - 1
            m_CurLine = Index : m_Status = 0 'still running
            Result = DoCmd(Lines(Index)) 'normal return -1
            If Result = -1 Then m_CurLine = 0 : m_Status = Index + 1 : Return  'error then return line #
            If Result = -2 Then m_CurLine = 0 : m_Status = -(Index + 2) : Return  'Emergency Stop at -line
            If Result > 0 Then 'in case Jump
                Index = Result - 1
                If Index > Lines.Count - 1 Then m_CurLine = 0 : m_Status = -1 : Return  'IDLE
            End If

        Next
        m_CurLine = 0 : m_Status = -1 'IDLE
    End Sub
#End Region

End Class
