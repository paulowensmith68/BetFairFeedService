Imports System.ServiceProcess
Imports System.Text
Public Class BetFairFeedService
    Inherits System.ServiceProcess.ServiceBase

    Private worker As New Worker()

    Protected Overrides Sub OnStart(ByVal args() As String)

        Dim wt As System.Threading.Thread
        Dim ts As System.Threading.ThreadStart
        gobjEvent.WriteToEventLog("Service Start Banner:    ************************************************")
        gobjEvent.WriteToEventLog("Service Start Banner:    *       BefFairFeedService STARTING            *")
        gobjEvent.WriteToEventLog("Service Start Banner:    ************************************************")
        gobjEvent.WriteToEventLog("Windows Service OnStart method starting service.")

        ts = AddressOf worker.DoWork
        wt = New System.Threading.Thread(ts)

        wt.Start()

    End Sub

    Protected Overrides Sub OnStop()
        worker.StopWork()
    End Sub

End Class

Public Class Worker

    Private m_thMain As System.Threading.Thread
    Private m_booMustStop As Boolean = False
    Private m_rndGen As New Random(Now.Millisecond)
    Private Shared stdOutput As StringBuilder = Nothing
    Private Shared stdNumOutputLines As Integer = 0
    Private Shared errOutput As StringBuilder = Nothing
    Private Shared errNumOutputLines As Integer = 0
    Private intFileNumber As Integer = FreeFile()
    Public Sub StopWork()

        m_booMustStop = True

        gobjEvent.WriteToEventLog("Service Stopping Banner: ************************************************")
        gobjEvent.WriteToEventLog("Service Stopping Banner: *       BefFairFeedService STOPPED             *")
        gobjEvent.WriteToEventLog("Service Stopping Banner: ************************************************")

        If Not m_thMain Is Nothing Then

            If Not m_thMain.Join(100) Then

                m_thMain.Abort()

            End If

        End If

    End Sub
    Public Sub DoWork()

        '----------------------------------------------------------'
        'Purpose:   Worker thread.
        '----------------------------------------------------------'

        m_thMain = System.Threading.Thread.CurrentThread

        Dim i As Integer = m_rndGen.Next
        Dim blnReturnStatus As Boolean
        Dim intMins As Integer
        Dim intAdapterCycleEveryMillisecs As Integer = My.Settings.ProcessCycleEverySecs * 1000

        m_thMain.Name = "Thread" & i.ToString
        gobjEvent.WriteToEventLog("Windows worker thread : " + m_thMain.Name + " created.")

        ' Write log entries for configuration settings
        gobjEvent.WriteToEventLog("WorkerThread : Cycle every (secs) : " + My.Settings.ProcessCycleEverySecs.ToString)
        gobjEvent.WriteToEventLog("WorkerThread : Cycle every (millisecs) : " + intAdapterCycleEveryMillisecs.ToString)


        While Not m_booMustStop

            ' Call start process and set status
            blnReturnStatus = StartProcess()

            ' Check status and issue warning
            If blnReturnStatus = False Then
                gobjEvent.WriteToEventLog("WorkerThread : Process returned failed status, service will continue", EventLogEntryType.Warning)
            End If

            '-------------------------------------------------
            '-  Issue heartbeat message every service cycle  -
            '-------------------------------------------------
            If intMins = 0 Then
                gobjEvent.WriteToEventLog("Windows worker thread : Heartbeat.......")
            End If

            '-------------------------------------------------
            '-  Now sleep, you beauty.                       -
            '-------------------------------------------------
            System.Threading.Thread.Sleep(intAdapterCycleEveryMillisecs)

        End While

    End Sub
    Function StartProcess() As Boolean

        ' Define static variables shared by class methods.
        Dim intElapsedTimeMillisecs As Integer = 0


        Try

            Dim marketCountriesUkOnly As HashSet(Of String)
            marketCountriesUkOnly = New HashSet(Of String)({"GB"})
            Dim marketCountriesEurope As HashSet(Of String)
            marketCountriesEurope = New HashSet(Of String)({"FR", "DE", "IT", "ES"})

            ' Refresh database from BetFair API interface


            ' Football - eventTypeId = 1
            ' MarketTypeCodes applicable : "MATCH_ODDS", "HALF_TIME", "HALF_TIME_FULL_TIME", "OVER_UNDER_25", "CORRECT_SCORE", "BOTH_TEAMS_TO_SCORE"
            '
            gobjEvent.WriteToEventLog("StartProcess:    *--------------------------------------------------")
            gobjEvent.WriteToEventLog("StartProcess:    *-----  BefFairFeedService - Updating Football ----")
            gobjEvent.WriteToEventLog("StartProcess:    *--------------------------------------------------")
            Dim BetFairDatabase1 As New BetFairDatabaseClass()
            BetFairDatabase1.PollBetFairEvents(1, "MATCH_ODDS", 50, marketCountriesUkOnly)
            BetFairDatabase1 = Nothing

            ' Match new odds
            Dim BetFairDbOddsMatch1 As New BetFairDatabaseClass()
            BetFairDbOddsMatch1.MatchWithBookmakers(1, "MATCH_ODDS")
            BetFairDbOddsMatch1 = Nothing

            ' Europe leagues
            Dim BetFairDatabase2 As New BetFairDatabaseClass()
            BetFairDatabase2.PollBetFairEvents(1, "MATCH_ODDS", 20, marketCountriesEurope)
            BetFairDatabase2 = Nothing

            Dim BetFairDatabase3 As New BetFairDatabaseClass()
            BetFairDatabase3.PollBetFairEvents(1, "HALF_TIME", 50, marketCountriesUkOnly)
            BetFairDatabase3 = Nothing

            Dim BetFairDatabase4 As New BetFairDatabaseClass()
            BetFairDatabase4.PollBetFairEvents(1, "HALF_TIME_FULL_TIME", 50, marketCountriesUkOnly)
            BetFairDatabase4 = Nothing

            Dim BetFairDatabase5 As New BetFairDatabaseClass()
            BetFairDatabase5.PollBetFairEvents(1, "OVER_UNDER_25", 50, marketCountriesUkOnly)
            BetFairDatabase5 = Nothing

            Dim BetFairDatabase6 As New BetFairDatabaseClass()
            BetFairDatabase6.PollBetFairEvents(1, "CORRECT_SCORE", 50, marketCountriesUkOnly)
            BetFairDatabase6 = Nothing

            Dim BetFairDatabase7 As New BetFairDatabaseClass()
            BetFairDatabase7.PollBetFairEvents(1, "BOTH_TEAMS_TO_SCORE", 50, marketCountriesUkOnly)
            BetFairDatabase7 = Nothing

            ' ------------------------------
            ' Horse Racing - eventTypeId = 7
            ' ------------------------------
            ' MarketTypeCodes applicable : "WIN", "PLACE"
            '
            gobjEvent.WriteToEventLog("StartProcess:    *------------------------------------------------------")
            gobjEvent.WriteToEventLog("StartProcess:    *-----  BefFairFeedService - Updating Horse Racing ----")
            gobjEvent.WriteToEventLog("StartProcess:    *------------------------------------------------------")
            Dim BetFairHorseRacingDatabase1 As New BetFairDatabaseClass()
            BetFairHorseRacingDatabase1.PollBetFairEvents(7, "WIN", 50, marketCountriesUkOnly)
            BetFairHorseRacingDatabase1 = Nothing

            Dim BetFairHorseRacingDatabase2 As New BetFairDatabaseClass()
            BetFairHorseRacingDatabase2.PollBetFairEvents(7, "PLACE", 50, marketCountriesUkOnly)
            BetFairHorseRacingDatabase2 = Nothing

        Catch ex As Exception

            gobjEvent.WriteToEventLog("StartProcess : Process has been killed, general error : " & ex.Message, EventLogEntryType.Error)
            Return False

        End Try

        Return True

    End Function

End Class

