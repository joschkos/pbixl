Imports System.Runtime.InteropServices
Imports System.Net.NetworkInformation
Imports System.Net

Public Module TCP_Get
    Declare Auto Function GetExtendedTcpTable Lib "iphlpapi.dll" (ByVal pTCPTable As IntPtr, ByRef OutLen As Integer, ByVal Sort As Boolean, ByVal IpVersion As Integer, ByVal dwClass As Integer, ByVal Reserved As Integer) As Integer
    Const TCP_TABLE_OWNER_PID_ALL As Integer = 5
    <StructLayout(LayoutKind.Sequential)>
    Public Structure MIB_TCPTABLE_OWNER_PID
        Public NumberOfEntries As Integer
        Public Table As IntPtr
    End Structure
    <StructLayout(LayoutKind.Sequential)>
    Public Structure MIB_TCPROW_OWNER_PID
        Public state As Integer
        Public localAddress As UInteger
        Public localPort1 As Byte
        Public localPort2 As Byte
        Public localPort3 As Byte
        Public localPort4 As Byte
        Public RemoteAddress As UInteger
        Public remotePort As Integer
        Public PID As Integer
    End Structure

    Structure TcpConnection
        Public State As TcpState
        Public localAddress As String
        Public localPort As Integer
        Public localPort1 As Byte
        Public localPort2 As Byte
        Public localPort3 As Byte
        Public localPort4 As Byte
        Public RemoteAddress As String
        Public remotePort As Integer
        Public Proc As String
        Public PID As Integer
    End Structure

    Function TCPConns() As MIB_TCPROW_OWNER_PID()

        TCPConns = Nothing
        Dim cb As Integer
        GetExtendedTcpTable(Nothing, cb, False, 2, TCP_TABLE_OWNER_PID_ALL, 0)
        Dim tcptable As IntPtr = Marshal.AllocHGlobal(cb)
        If GetExtendedTcpTable(tcptable, cb, False, 2, TCP_TABLE_OWNER_PID_ALL, 0) = 0 Then
            Dim tab As MIB_TCPTABLE_OWNER_PID = Marshal.PtrToStructure(tcptable, GetType(MIB_TCPTABLE_OWNER_PID))
            Dim Mibs(tab.NumberOfEntries - 1) As MIB_TCPROW_OWNER_PID
            Dim row As IntPtr
            For i As Integer = 0 To tab.NumberOfEntries - 1
                row = New IntPtr(tcptable.ToInt64 + Marshal.SizeOf(tab.NumberOfEntries) + Marshal.SizeOf(GetType(MIB_TCPROW_OWNER_PID)) * i)
                Mibs(i) = Marshal.PtrToStructure(row, GetType(MIB_TCPROW_OWNER_PID))
            Next
            TCPConns = Mibs
        End If
        Marshal.FreeHGlobal(tcptable)
    End Function


    Function MIB_ROW_To_TCP(ByVal row As MIB_TCPROW_OWNER_PID) As TcpConnection
        Dim tcp As New TcpConnection With {
            .State = DirectCast(row.state, TcpState)
            }
        Dim ipad As New IPAddress(row.localAddress)
        tcp.localAddress = ipad.ToString
        Dim localPort As Integer = ((row.localPort1 + 8) _
            + (row.localPort2 _
            + ((row.localPort3 + 24) _
            + (row.localPort4 + 16))))
        tcp.localPort = BitConverter.ToUInt16(New Byte() {row.localPort2, row.localPort1}, 0)
        ipad = New IPAddress(row.RemoteAddress)
        tcp.RemoteAddress = ipad.ToString
        tcp.remotePort = row.remotePort / 256 + (row.remotePort Mod 256) * 256
        Dim p As Process = Process.GetProcessById(row.PID)
        tcp.Proc = p.ProcessName & " (" & row.PID.ToString & ")"
        tcp.PID = row.PID
        p.Dispose()
        Return tcp
    End Function
End Module