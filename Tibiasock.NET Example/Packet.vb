Imports TibiasockNET
Imports System.Runtime.InteropServices

Module Packet
    Delegate Function Int() As Integer

    Public Tibia_PID As Long
    Public Tibia_Base As Long = &H400000
    Public Tibia_hWnd As Integer
    Public Tibia_Handle As IntPtr

    Public GameState As Int = Function() (&H7B9EA8 - &H400000) + Tibia_Base ' 8 = online, 0 = offline


    Public Enum TextMessageColor As Byte

        YellowConsole = 1 '// 2, 3
        PurpleConsole = 4 '// 21
        '5 and 6 crashes
        'YellowConsole = 9 '// 11
        NPCTeal = 10
        '11 crashes
        RedConsoleOnly = 12
        '13 nothing
        'RedConsoleOnly = 14
        '15 crashes
        WhiteServerLogAndBottomGameWindow = 16
        RedServerLogAndCenterGameWindow = 17
        WhiteServerLogAndCenterGameWindow = 18
        WhiteBottomGameWindow = 19
        GreenCenterGameWindow = 20
        '21 crashes
        '22 crashes
        '23 crashes

    End Enum


    Public Enum enumTalkTypes As Byte
        Talk = 1
        Whisper = 2
        Yell = 3
        PrivateMessage = 5
        Channel = 7
        NPC = 11

    End Enum

    Public Enum enumChatChannels As Byte
        GuildChat = 0
        WorldChat = 3
        EnglishChat = 4
        TradeChat = 5
        RookTradeChat = 6
        HelpChat = 7
        OwnPrivateChannel = 14
        PrivateChannel1 = 17
    End Enum
    Public Enum Directions
        North = 0
        East = 1
        South = 2
        West = 3
        NorthEast = 4
        SouthEast = 5
        SouthWest = 6
        NorthWest = 7
        Center = 8

        FaceNorth = 0
        FaceEast = 1
        FaceSouth = 2
        FaceWest = 3
    End Enum


    Public Sub RecvMsg(ByVal Message As String, ByVal Color As TextMessageColor)
        If IsOnline() = False Then Exit Sub
        If Message.Length > 256 Then Exit Sub

        Dim packet As New NetworkMessage
        packet.AddByte(&HB4)
        packet.AddByte(Color)
        packet.AddString(Message)

        SendPacketToClient(Tibia_PID, packet.RawData)
    End Sub

    Public Sub RecvMsgUsingNETDLL(ByVal Message As String, ByVal Color As TextMessageColor)
        If IsOnline() = False Then Exit Sub
        If Message.Length > 256 Then Exit Sub

        Dim packet As New NetworkMessage
        packet.AddByte(&HB4)
        packet.AddByte(Color)
        packet.AddString(Message)

        TibiasockNET.SendPacketToClient(Tibia_PID, packet.RawData)

    End Sub


    Public Sub SendMessage(ByRef Message As String, Optional ByVal TalkType As enumTalkTypes = enumTalkTypes.Talk)
        If IsOnline() = False Then Exit Sub
        If Message.Length > 256 Then Exit Sub


        Dim packet As New NetworkMessage
        packet.AddByte(&H96)
        packet.AddByte(TalkType)
        packet.AddString(Message)

        SendPacketToServer(Tibia_PID, packet.RawData)

    End Sub


    Public Sub SendMsgChannel(ByVal Message As String, ByVal Channel As enumChatChannels)
        If IsOnline() = False Then Exit Sub

        Dim packet As New NetworkMessage
        packet.AddByte(&H96)
        packet.AddByte(enumTalkTypes.Channel)
        packet.AddUInt16(Channel)
        packet.AddString(Message)

        SendPacketToServer(Tibia_PID, packet.RawData)
    End Sub

    Public Sub SendPM(ByRef CharName As String, ByRef Message As String)
        If IsOnline() = False Then Exit Sub
        If Message.Length > 256 Then Exit Sub
        If CharName.Length > 40 Then Exit Sub


        Dim packet As New NetworkMessage
        packet.AddByte(&H96)
        packet.AddByte(enumTalkTypes.PrivateMessage)
        packet.AddString(CharName)
        packet.AddString(Message)

        SendPacketToServer(Tibia_PID, packet.RawData)

    End Sub

    Public Sub Dance()
        LookTo(Directions.FaceWest)
        Threading.Thread.Sleep(300)
        LookTo(Directions.FaceNorth)
        Threading.Thread.Sleep(300)
        LookTo(Directions.FaceEast)
        Threading.Thread.Sleep(300)
        LookTo(Directions.FaceSouth)
    End Sub

    Public Sub LookTo(ByVal sDirection As Directions)
        If Not IsOnline() Then Exit Sub
        Dim packet As New NetworkMessage
        packet.AddByte(&H6F + sDirection)
        SendPacketToServer(Tibia_PID, packet.RawData)
    End Sub
    Public Function IsOnline() As Boolean
        Dim Buffer As Integer
        Dim res As Boolean = False
        If Tibia_PID = 0 Then res = False
        Try
            Buffer = ReadInt(Tibia_Handle, GameState.Invoke)
            Dim lastError As Integer = Marshal.GetLastWin32Error()
            If Buffer = 8 Then res = True
        Catch

        End Try
        Return res
    End Function
End Module
