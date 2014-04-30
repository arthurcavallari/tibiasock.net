<ComClass(TibiasockCOM.ClassId, TibiasockCOM.InterfaceId, TibiasockCOM.EventsId)> _
Public Class TibiasockCOM


#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "3d628888-71b8-43b1-9fc0-d06103744c6b"
    Public Const InterfaceId As String = "ae074ee4-2e0c-460f-8745-ad28c11d7811"
    Public Const EventsId As String = "9e3c5fb5-6cb0-4ddc-9ea8-caf12ffc1f67"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub
    Public Sub SendPacketToServerEx(ByRef process_id As IntPtr, ByRef dataBuffer() As Byte, ByRef SendStreamData As Integer, ByRef SendStreamLength As Integer, ByRef SendPacketCall As Integer)
        Tibiasock.SendPacketToServerEx(process_id, dataBuffer, SendStreamData, SendStreamLength, SendPacketCall)
    End Sub
    Public Sub SendPacketToServer(ByRef process_id As IntPtr, ByRef dataBuffer() As Byte)
        Tibiasock.SendPacketToServer(process_id, dataBuffer)

    End Sub

    Public Sub SendPacketToClientEx(ByRef process_id As IntPtr, ByRef dataBuffer() As Byte, ByRef RecvStream As Integer, ByRef ParserCall As Integer)
        Tibiasock.SendPacketToClientEx(process_id, dataBuffer, RecvStream, ParserCall)
    End Sub
    Public Sub SendPacketToClient(ByRef process_id As IntPtr, ByRef dataBuffer() As Byte)

        Tibiasock.SendPacketToClient(process_id, dataBuffer)
    End Sub
End Class


