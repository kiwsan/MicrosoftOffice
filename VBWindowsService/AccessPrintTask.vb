Imports System.IO
Imports log4net
Imports Microsoft.Office.Interop.Access

Public Class AccessPrintTask

    Shared ReadOnly Logger As ILog = LogManager.GetLogger("AccessPrintTask")
    Private Shared _locker As Object = New Object()

    Public Sub New()

    End Sub

    Public Sub Start()
        Try

            If KillProcess("MSACCESS") Then

                SyncLock _locker
                    Logger.Info("Thread safe..")
                    Dim fileName As String = "D:\\MSAccessDatabase\\MSAccessDatabase.accdb"

                    If File.Exists(fileName) Then
                        Logger.Info("Open..")

                        Dim microsoftAccess As New Application
                        microsoftAccess = CreateObject("Access.Application")
                        microsoftAccess.DoCmd.Maximize()
                        microsoftAccess.DoCmd.Minimize()

                        'microsoftAccess.Application.Visible = True
                        microsoftAccess.OpenCurrentDatabase(fileName)
                        Dim myName = microsoftAccess.Run("GetName")
                        Logger.Info(String.Format("My Name: {0}", myName))

                        If microsoftAccess IsNot Nothing Then
                            microsoftAccess.Quit()
                            Runtime.InteropServices.Marshal.ReleaseComObject(microsoftAccess)
                            microsoftAccess = Nothing
                            Logger.Info("Quit..")
                        End If
                    End If
                End SyncLock
            End If

        Catch ex As Exception
            Logger.Error(ex)
            Logger.Info("KillProcess..")
            KillProcess("MSACCESS")
        End Try
    End Sub

    Private Shared Function KillProcess(name As String) As Boolean

        For Each item As Process In Process.GetProcesses().Where(Function(x) x.ProcessName.Contains(name))
            If Process.GetCurrentProcess().Id = item.Id Then
                Continue For
            End If
            If item.ProcessName.Contains(name) Then
                item.Kill()
                Return True
            End If
        Next

        Return True
    End Function

End Class
