Imports ExcelDna.Integration
Imports ExcelDna.Integration.Rtd
Imports System.Threading
Imports System.Runtime.InteropServices

Public Module MyFunctions




    Public intInterval As Integer

    <ExcelFunction(Description:="pbixl Timer", Category:="pbixl", IsMacroType:=False, IsVolatile:=False, IsHidden:=False, Name:="pbixl.Timer")>
    Public Function GetTimer(Interval As Integer)

        Try
            'Return XlCall.RTD("SqlQuery.RtdTest.MyRtdServer", Nothing, Interval.ToString)
            Return XlCall.RTD("pbixl.RtdTest.MyRtdServer", Nothing, Interval.ToString)
        Catch ex As Exception
            Return ex.Message
        End Try


    End Function




End Module


'''''''''''''''''''''''
' The RTD server using the ExcelRtdServer base class.
Namespace RtdTest
    <ComVisible(True)>
    Public Class MyRtdServer
        Inherits ExcelRtdServer
        Dim myTimer As Timer

        Dim topics As New List(Of Topic)

        Protected Overrides Function ServerStart() As Boolean
            myTimer = New Timer(AddressOf Tick, Nothing, 1000, 1000)
            Return True
        End Function

        Protected Overrides Sub ServerTerminate()
            myTimer = Nothing
        End Sub



        Dim lstTopic As New List(Of xtopic)

        Protected Overrides Function ConnectData(topic As Topic, topicInfo As IList(Of String), ByRef newValues As Boolean) As Object

            Try
                topics.Add(topic)
                Dim x As New xtopic
                x.topic = topic
                x.InterVal = CInt(topicInfo(0))
                x.Report = Now
                x.NextRun = Now.AddMilliseconds(-100)
                lstTopic.Add(x)
                Return GetTime(x.InterVal)
            Catch ex As Exception
                Return ex.Message
            End Try


        End Function

        Protected Overrides Sub DisconnectData(topic As Topic)

            'Remove Topics
            For i As Integer = Me.lstTopic.Count - 1 To 0 Step -1
                If Me.lstTopic.Item(i).topic Is topic Then
                    Me.lstTopic.RemoveAt(i)
                End If
            Next i

            Me.topics.Remove(topic)
            If topics.Count = 0 Then

            End If
        End Sub

        Private Sub Tick(ByVal state As Object)

            For i As Integer = Me.topics.Count - 1 To 0 Step -1
                Dim InterVal As Integer = -1
                For Each xt In lstTopic
                    If xt.topic Is Me.topics.Item(i) Then
                        intInterval = xt.InterVal
                        Exit For
                    End If
                Next xt
                Try
                    Me.topics.Item(i).UpdateValue(GetTime(intInterval))
                Catch ex As Exception
                End Try
            Next i

        End Sub

        Private Function GetTime(interVal As String) As String

            For Each xt In Me.lstTopic
                If xt.InterVal = CInt(interVal) Then
                    If Now > xt.NextRun Then
                        xt.Report = Now
                        xt.NextRun = xt.Report.AddSeconds(CInt(interVal))
                        Return xt.Report.ToString("HH:mm:ss")
                    Else
                        Return xt.Report.ToString("HH:mm:ss")
                    End If
                End If

            Next xt

            Return DateTime.Now.ToString("HH:mm:ss")
        End Function


        Private Class xtopic
            Public Property topic As Topic
            Public Property InterVal As Integer
            Public Property NextRun As DateTime
            Public Property Report As DateTime
        End Class



    End Class





End Namespace