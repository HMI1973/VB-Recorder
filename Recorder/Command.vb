Public Class Command
    Public m_Command As String
    Public m_X As Integer
    Public m_Y As Integer
    Public m_Text As String 'in case of "GETURL" or "CHECKCLIP" or "CHECKWIN" or "CHECKWINCHANGE" or "CHECKIMG"
    Public m_Number As Integer 'in case "LOOP" or "JUMP"
    Public m_Waitms As Double

    Public Sub Init(ByVal CMD As String, ByVal X As Integer, ByVal Y As Integer, ByVal myTxt As String, ByVal myNum As Integer, ByVal myWaitms As Double)
        m_Command = CMD.ToUpper.Replace(" ", "")
        m_X = X
        m_Y = Y
        m_Text = myTxt
        m_Number = myNum
        m_Waitms = myWaitms

        If m_Command.StartsWith("//") Then m_Command = "" ' Comment line
    End Sub
    Public Sub Init(ByVal FullCMD As String)
        Dim myCMD() As String = FullCMD.Split("|")

        Init(myCMD(0), 0, 0, "", 0, 0)
        If myCMD.Count = 3 Then m_Waitms = Double.Parse(myCMD(2)) 'read Wait mSec
        If myCMD.Count > 1 Then
            Dim mPos() As String = myCMD(1).Split(",")
            If mPos.Count = 1 Then
                m_Text = mPos(0)
                If IsNumeric(mPos(0)) Then m_Number = CInt(mPos(0))
            End If
            If mPos.Count = 2 Then
                m_X = CInt(mPos(0))
                m_Y = CInt(mPos(1))
            End If
        End If
    End Sub


    Public Sub New(ByVal CMD As String)
        Init(CMD)
    End Sub
End Class
