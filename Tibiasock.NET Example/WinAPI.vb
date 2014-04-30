Imports System.Runtime.InteropServices

Public Module WinAPI
    <DllImport("kernel32.dll")> _
    Public Function OpenThread(ByVal dwDesiredAccess As ThreadAccess, ByVal bInheritHandle As Boolean, ByVal dwThreadId As UInteger) As IntPtr
    End Function
    <DllImport("kernel32.dll")> _
    Public Function SuspendThread(ByVal hThread As IntPtr) As UInteger
    End Function

    <DllImport("kernel32.dll")> _
    Public Function ResumeThread(ByVal hThread As IntPtr) As UInt32
    End Function

    <DllImport("Kernel32.dll")> _
    Public Function WaitForSingleObject(ByVal hHandle As UInteger, ByVal dwMilliseconds As UInteger) As UInteger
    End Function

    <DllImport("kernel32.dll")> _
    Public Function CreateRemoteThread(ByVal hProcess As IntPtr, ByVal lpThreadAttributes As IntPtr, ByVal dwStackSize As UInteger, ByVal lpStartAddress As IntPtr, ByVal lpParameter As IntPtr, ByVal dwCreationFlags As UInteger, _
 ByVal lpThreadId As IntPtr) As IntPtr
    End Function

    <DllImport("kernel32.dll", SetLastError:=True, ExactSpelling:=True)> _
    Public Function VirtualAllocEx(ByVal hProcess As IntPtr, ByVal lpAddress As IntPtr, _
     ByVal dwSize As UInteger, ByVal flAllocationType As UInteger, _
     ByVal flProtect As UInteger) As IntPtr
    End Function

    <DllImport("kernel32.dll")> _
    Public Function VirtualFreeEx(ByVal hProcess As IntPtr, _
                      ByVal lpAddress As IntPtr, _
                      ByVal dwSize As Integer, _
                      ByVal dwFreeType As FreeType) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)> _
    Private Function ReadProcessMemory(ByVal hProcess As IntPtr, ByVal lpBaseAddress As IntPtr, <Out()> ByVal lpBuffer As Byte(), ByVal dwSize As Integer, ByRef lpNumberOfBytesRead As Integer) As Boolean
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)> _
    Private Function WriteProcessMemory(ByVal hProcess As IntPtr, ByVal lpBaseAddress As IntPtr, <Out()> ByVal lpBuffer As Byte(), ByVal nSize As System.UInt32, <Out()> ByRef lpNumberOfBytesWritten As IntPtr) As Boolean
    End Function
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Public Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Integer, ByRef lpdwProcessId As Integer) As Integer

    Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
    Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As IntPtr) As IntPtr

    Public Const PROCESS_ALL_ACCESS = &H1F0FFF


    <Flags()> _
    Public Enum FreeType As UInteger
        DECOMMIT = &H4000
        RELEASE = &H8000
    End Enum
    <Flags()> _
    Public Enum ThreadAccess As Integer
        TERMINATE = (&H1)
        SUSPEND_RESUME = (&H2)
        GET_CONTEXT = (&H8)
        SET_CONTEXT = (&H10)
        SET_INFORMATION = (&H20)
        QUERY_INFORMATION = (&H40)
        SET_THREAD_TOKEN = (&H80)
        IMPERSONATE = (&H100)
        DIRECT_IMPERSONATION = (&H200)
    End Enum

    <Flags()> _
    Public Enum AllocationType
        Commit = &H1000
        Reserve = &H2000
        Decommit = &H4000
        Release = &H8000
        Reset = &H80000
        Physical = &H400000
        TopDown = &H100000
        WriteWatch = &H200000
        LargePages = &H20000000
    End Enum

    <Flags()> _
    Public Enum MemoryProtection
        Execute = &H10
        ExecuteRead = &H20
        ExecuteReadWrite = &H40
        ExecuteWriteCopy = &H80
        NoAccess = &H1
        [ReadOnly] = &H2
        ReadWrite = &H4
        WriteCopy = &H8
        GuardModifierflag = &H100
        NoCacheModifierflag = &H200
        WriteCombineModifierflag = &H400
    End Enum

#Region "Memory"
    Public Function ReadBytes(ByVal ProcessHandle As IntPtr, ByVal address As Long, ByVal bytesToRead As UInteger) As Byte()
        Dim ptrBytesRead As IntPtr
        Dim buffer As Byte() = New Byte(bytesToRead - 1) {}
        ReadProcessMemory(ProcessHandle, New IntPtr(address), buffer, bytesToRead, ptrBytesRead)
        Return buffer
    End Function
    Public Function ReadByte(ByVal ProcessHandle As IntPtr, ByVal address As Long) As Byte
        Return ReadBytes(ProcessHandle, address, 1)(0)
    End Function
    Public Function ReadInt(ByVal ProcessHandle As IntPtr, ByVal address As Long) As Integer
        Return BitConverter.ToInt32(ReadBytes(ProcessHandle, address, 4), 0)
    End Function
    Public Function ReadLong(ByVal ProcessHandle As IntPtr, ByVal address As Long) As Long
        Return BitConverter.ToInt64(ReadBytes(ProcessHandle, address, 4), 0)
    End Function
    Public Function ReadUInt(ByVal ProcessHandle As IntPtr, ByVal address As Long) As UInteger
        Return BitConverter.ToUInt32(ReadBytes(ProcessHandle, address, 4), 0)
    End Function
    Public Function ReadULong(ByVal ProcessHandle As IntPtr, ByVal address As Long) As ULong
        Return BitConverter.ToUInt64(ReadBytes(ProcessHandle, address, 4), 0)
    End Function

    Public Function ReadString(ByVal handle As IntPtr, ByVal address As Long) As String
        Return ReadString(handle, address, 0)
    End Function

    Public Function ReadString(ByVal handle As IntPtr, ByVal address As Long, ByVal length As UInteger) As String
        If length > 0 Then
            Dim buffer As Byte()
            buffer = ReadBytes(handle, address, length)
            Return System.Text.ASCIIEncoding.[Default].GetString(buffer).Split(New [Char]())(0)
        Else
            Dim s As String = ""
            Dim temp As Byte = ReadByte(handle, System.Math.Max(System.Threading.Interlocked.Increment(address), address - 1))
            While temp <> 0
                s += ChrW(temp)
                temp = ReadByte(handle, System.Math.Max(System.Threading.Interlocked.Increment(address), address - 1))
            End While
            Return s
        End If
    End Function
    Public Function WriteBytes(ByVal ProcessHandle As IntPtr, ByVal address As Long, ByVal bytes As Byte(), ByVal length As UInteger) As Boolean
        Dim bytesWritten As IntPtr
        Dim result As Integer = WriteProcessMemory(ProcessHandle, New IntPtr(address), bytes, length, bytesWritten)
        Return result <> 0
    End Function

    Public Function WriteBytes(ByVal ProcessHandle As IntPtr, ByVal address As Long, ByVal bytes As Byte()) As Boolean
        Dim bytesWritten As IntPtr
        Dim result As Integer = WriteProcessMemory(ProcessHandle, New IntPtr(address), bytes, bytes.Length, bytesWritten)
        Return result <> 0
    End Function
    Public Function WriteByte(ByVal ProcessHandle As IntPtr, ByVal address As Long, ByVal value As Byte) As Boolean
        Return WriteBytes(ProcessHandle, address, New Byte() {value})
    End Function

    Public Function WriteInt(ByVal ProcessHandle As IntPtr, ByVal address As Long, ByVal value As Integer) As Boolean
        Dim bytes As Byte() = BitConverter.GetBytes(value)
        Return WriteBytes(ProcessHandle, address, bytes)
    End Function

    Public Function WriteLong(ByVal ProcessHandle As IntPtr, ByVal address As Long, ByVal value As Long) As Boolean
        Dim bytes As Byte() = BitConverter.GetBytes(value)
        Return WriteBytes(ProcessHandle, address, bytes)
    End Function

    Public Function WriteUInt(ByVal ProcessHandle As IntPtr, ByVal address As Long, ByVal value As UInteger) As Boolean
        Dim bytes As Byte() = BitConverter.GetBytes(value)
        Return WriteBytes(ProcessHandle, address, bytes)
    End Function

    Public Function WriteULong(ByVal ProcessHandle As IntPtr, ByVal address As Long, ByVal value As ULong) As Boolean
        Dim bytes As Byte() = BitConverter.GetBytes(value)
        Return WriteBytes(ProcessHandle, address, bytes)
    End Function

    Public Function WriteString(ByVal ProcessHandle As IntPtr, ByVal address As Long, ByVal value As String) As Boolean
        Dim encoding As New System.Text.ASCIIEncoding()
        Dim bytes As Byte() = encoding.GetBytes(value + Chr(0))
        Return WriteBytes(ProcessHandle, address, bytes)
    End Function
#End Region
End Module
