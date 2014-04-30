Imports System.Text

Public Class NetworkMessage
    Private buffer As Byte()
    Private m_position As Integer, m_length As Integer, bufferSize As Integer = 16394

#Region "Contructors"
    Public Sub New()
        buffer = New Byte(bufferSize - 1) {}
        m_position = 2
    End Sub
#End Region

#Region "Properties"
    Public Property Length() As Integer
        Get
            Return m_length
        End Get
        Set(ByVal value As Integer)
            m_length = value
        End Set
    End Property

    Public Property Position() As Integer
        Get
            Return m_position
        End Get
        Set(ByVal value As Integer)
            m_position = value
        End Set
    End Property

    Public ReadOnly Property Data() As Byte()
        Get
            Dim t As Byte() = New Byte(m_length - 1) {}
            Array.Copy(buffer, t, m_length)
            Return t
        End Get
    End Property
    Public ReadOnly Property RawData() As Byte()
        Get
            Dim t As Byte() = New Byte(m_length - 3) {}
            Array.Copy(buffer, 2, t, 0, m_length - 2)
            Return t
        End Get
    End Property
#End Region
    Public Function GetPacketHeaderSize() As Integer
        Return 2
    End Function

    Public Sub PrepareToSend()
        Dim lenght As UInt16 = Data.Length - 2
        Position = 0
        AddUInt16(lenght)
    End Sub
    Public Sub AddByte(ByVal value As Byte)
        If 1 + m_length > bufferSize Then
            Throw New Exception("NetworkMessage buffer is full.")
        End If
        AddBytes(New Byte() {value})
    End Sub

    Public Sub AddBytes(ByVal value As Byte())
        If value.Length + m_length > bufferSize Then
            Throw New Exception("NetworkMessage buffer is full.")
        End If
        Array.Copy(value, 0, buffer, m_position, value.Length)
        m_position += value.Length
        If m_position > m_length Then
            m_length = m_position
        End If
    End Sub

    Public Sub AddString(ByVal value As String)
        AddUInt16(CUShort(value.Length))
        AddBytes(System.Text.ASCIIEncoding.[Default].GetBytes(value))
    End Sub

    Public Sub AddUInt16(ByVal value As UShort)
        AddBytes(BitConverter.GetBytes(value))
    End Sub

    Public Sub AddUInt32(ByVal value As UInteger)
        AddBytes(BitConverter.GetBytes(value))
    End Sub
End Class