Imports System.ServiceProcess
Imports System.Timers
Imports log4net

Partial Public Class Service1
    Inherits ServiceBase

    Shared ReadOnly Logger As ILog = LogManager.GetLogger("Service1")
    Private timerTwoSeconds As Timer = New Timer(20000)
    Private accessPrintTask As AccessPrintTask = New AccessPrintTask()

    Friend Sub OnDebug()
        OnStart(Nothing)
    End Sub

    Public Sub New()
        InitializeComponent()
    End Sub

    Protected Overrides Sub OnStart(ByVal args As String())
        Logger.Info("OnStart..")
        AddHandler timerTwoSeconds.Elapsed, New ElapsedEventHandler(AddressOf TimerTwoSeconds_Elapsed)
        timerTwoSeconds.Enabled = True
    End Sub

    Private Sub TimerTwoSeconds_Elapsed(ByVal sender As Object, ByVal e As ElapsedEventArgs)
        Logger.Info("Timestamp..")

        Try
            Dim task1 As Task = Task.Factory.StartNew(Sub() accessPrintTask.Start())
        Catch ex As Exception
            Logger.[Error](ex)
        End Try
    End Sub

    Protected Overrides Sub OnStop()
        Logger.Info("OnStop..")
        timerTwoSeconds.[Stop]()
    End Sub
End Class
