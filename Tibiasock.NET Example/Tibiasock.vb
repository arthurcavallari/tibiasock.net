
Imports System.Runtime.InteropServices



Module Tibiasock

    'Private Const OUTGOINGDATASTREAM As Integer = &H7B8F28 ' 9.54
    'Private Const OUTGOINGDATALEN As Integer = &H9EDA78
    'Private Const SENDOUTGOINGPACKET As Integer = &H512F20
    'Private Const INCOMINGDATASTREAM As Integer = &H9EDA64
    'Private Const PARSERFUNC As Integer = &H462D50


    Private Const SENDOUTGOINGPACKET As Integer = &H514900 '961
    Private Const OUTGOINGDATASTREAM As Integer = &H7B6F50 '961
    Private Const OUTGOINGDATALEN As Integer = &H9D1FD8 '961

    Private Const INCOMINGDATASTREAM As Integer = &H9D1FC4 '961

    Private Const PARSERFUNC As Integer = &H463330 '961

    Private Const INFINITE As UInteger = &HFFFFFFFFUI

#Region "Helper Methods"
    Private Function OpenAndSuspendThread(ByVal threadID As ULong) As IntPtr
        Dim proc As Process = Process.GetProcessById(threadID)
        Dim pOpenThread As IntPtr
        Dim mtime As Double = 0.0
        Dim tmptime As Double = 0.0
        Dim pid As Integer

        For Each pT As ProcessThread In proc.Threads

            tmptime = Math.Max(mtime, pT.UserProcessorTime.TotalSeconds)
            If tmptime > mtime Then
                pid = pT.Id
                mtime = tmptime
            End If

        Next
        pOpenThread = OpenThread((ThreadAccess.GET_CONTEXT Or ThreadAccess.SUSPEND_RESUME Or ThreadAccess.SET_CONTEXT), False, CUInt(pid))
        SuspendThread(pOpenThread)

        Return pOpenThread

    End Function

    Private Sub ResumeAndCloseThread(ByVal thread As IntPtr)
        ResumeThread(thread)
        CloseHandle(thread)
    End Sub

    Private Sub ExecuteRemoteCode(ByVal process As IntPtr, ByVal codeAddress As IntPtr, ByVal arg As UInteger)

        Dim WorkThread As IntPtr = CreateRemoteThread(process, 0, 0, codeAddress, arg, 0, 0)
        WaitForSingleObject(WorkThread, INFINITE)
        CloseHandle(WorkThread)
    End Sub

    Private Function Rebase(ByVal address As ULong, ByVal base As ULong) As ULong

        Return CULng((address - &H400000) + base)

    End Function

    Private Function CreateOutgoingBuffer(ByVal dataBuffer() As Byte, ByVal length As Integer) As Byte()

        Dim actualBuffer(1024) As Byte
        Dim size As Integer = Marshal.SizeOf(dataBuffer(0)) * dataBuffer.Length

        Dim pnt As IntPtr = Marshal.AllocHGlobal(size)

        Marshal.Copy(dataBuffer, 0, pnt, length - 8)
        Marshal.Copy(pnt, actualBuffer, 8, length - 8)
        Marshal.FreeHGlobal(pnt)

        Return actualBuffer

    End Function

    Private Sub WriteIncomingBuffer(ByVal process As IntPtr, ByVal recvStream As Integer, ByVal data() As Byte, ByVal length As Integer, ByVal position As Integer)

        Dim DataPointer As Integer
        WriteInt(process, recvStream + 4, length)
        WriteInt(process, recvStream + 8, position)

        DataPointer = ReadInt(process, recvStream)
        WriteBytes(process, DataPointer, data, length)
    End Sub
    Private Function CreateRemoteBuffer(ByVal process As IntPtr, ByVal dataBuffer() As Byte, ByVal length As Integer) As IntPtr

        Dim RemoteBufferPointer As IntPtr = VirtualAllocEx(process, 0, length, AllocationType.Commit, MemoryProtection.ExecuteReadWrite)
        WriteBytes(process, RemoteBufferPointer, dataBuffer, length)
        Return RemoteBufferPointer
    End Function
#End Region


    Public Sub SendPacketToServerEx(ByVal process_id As IntPtr, ByVal dataBuffer() As Byte, ByVal SendStreamData As Integer, ByVal SendStreamLength As Integer, ByVal SendPacketCall As Integer)

        Dim MainThread As IntPtr = OpenAndSuspendThread(process_id)

        Dim OldLength As Integer
        Dim OldData(1024) As Byte

        Dim length As Integer = dataBuffer.Length

        Dim process As IntPtr = OpenProcess(PROCESS_ALL_ACCESS, False, process_id)

        OldLength = ReadInt(process, SendStreamLength)
        OldData = ReadBytes(process, SendStreamData, OldLength)

        length += 8

        Dim actualBuffer() As Byte = CreateOutgoingBuffer(dataBuffer, length)
        WriteInt(process, SendStreamLength, length)
        WriteBytes(process, SendStreamData, actualBuffer, length)

        ExecuteRemoteCode(process, SendPacketCall, 1)

        WriteInt(process, SendStreamLength, OldLength)
        WriteBytes(process, SendStreamData, OldData, OldLength)

        ResumeAndCloseThread(MainThread)
    End Sub
    Public Sub SendPacketToServer(ByVal process_id As IntPtr, ByVal dataBuffer() As Byte)

        Dim ImageBase As Integer = Process.GetProcessById(process_id).MainModule.BaseAddress
        Dim SendStreamData As Integer = Rebase(OUTGOINGDATASTREAM, ImageBase)
        Dim SendStreamLength As Integer = Rebase(OUTGOINGDATALEN, ImageBase)
        Dim SendPacketCall As Integer = Rebase(SENDOUTGOINGPACKET, ImageBase)
        SendPacketToServerEx(process_id, dataBuffer, SendStreamData, SendStreamLength, SendPacketCall)
    End Sub




    Public Sub SendPacketToClientEx(ByVal process_id As IntPtr, ByVal dataBuffer() As Byte, ByVal RecvStream As Integer, ByVal ParserCall As Integer)

        Dim length As Integer = dataBuffer.Length

        Dim process As IntPtr = OpenProcess(PROCESS_ALL_ACCESS, False, process_id)

        Dim MainThread As IntPtr = OpenAndSuspendThread(process_id)
        Dim DataPointer As Integer
        Dim OldLength As Integer, OldPosition As Integer
        Dim OldDataBuffer(4096) As Byte
        OldLength = ReadInt(process, RecvStream + 4)
        OldPosition = ReadInt(process, RecvStream + 8)
        DataPointer = ReadInt(process, RecvStream)
        OldDataBuffer = ReadBytes(process, DataPointer, OldLength)

        WriteIncomingBuffer(process, RecvStream, dataBuffer, length, 0)

        Dim CodeCave() As Byte = {&HB8, &H0, &H0, &H0, &H0, &HFF, &HD0, &HC3} ' MOV EAX, <DWORD> | CALL EAX | RETN

        Dim pnt As IntPtr = Marshal.AllocHGlobal(4)
        Dim b() As Byte = BitConverter.GetBytes(ParserCall)
        Marshal.Copy(b, 0, pnt, 4)
        Marshal.Copy(pnt, CodeCave, 1, 4)
        Marshal.FreeHGlobal(pnt)


        Dim CodeCavePointer As IntPtr = CreateRemoteBuffer(process, CodeCave, 10)

        ExecuteRemoteCode(process, CodeCavePointer, 0)

        VirtualFreeEx(process, CodeCavePointer, 10, FreeType.RELEASE)
        WriteIncomingBuffer(process, RecvStream, OldDataBuffer, OldLength, OldPosition)

        ResumeAndCloseThread(MainThread)
    End Sub
    Public Sub SendPacketToClient(ByVal process_id As IntPtr, ByVal dataBuffer() As Byte)

        Dim ImageBase As Integer = Process.GetProcessById(process_id).MainModule.BaseAddress
        Dim RecvStream As Integer = Rebase(INCOMINGDATASTREAM, ImageBase)
        Dim ParserCall As Integer = Rebase(PARSERFUNC, ImageBase)
        SendPacketToClientEx(process_id, dataBuffer, RecvStream, ParserCall)
    End Sub

End Module
