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
            marketCountriesEurope = New HashSet(Of String)({"FR", "DE", "IT", "ES", "PT", "NL", "GR", "RU", "UA", "TR", "BE", "CH", "CZ", "HR"})

            '' Set-up specific competitions
            'Dim competitionIds As HashSet(Of String)
            'competitionIds = New HashSet(Of String)({"78601", "4527196"})
            Dim competitionIds As ISet(Of String) = New HashSet(Of String)()

            ' Football is eventId 1
            Dim separators() As String = {",", " "}
            Dim eventArray() As String = (My.Settings.CompetitionEventIds.Split(separators, StringSplitOptions.RemoveEmptyEntries))
            For Each eventId In eventArray
                competitionIds.Add(eventId)
            Next


            ' Soccer - eventTypeId = 1
            ' MarketTypeCodes applicable : "MATCH_ODDS", "HALF_TIME", "HALF_TIME_FULL_TIME", "OVER_UNDER_25", "CORRECT_SCORE", "BOTH_TEAMS_TO_SCORE"
            '
            If My.Settings.StreamSportId = 1 Then

                gobjEvent.WriteToEventLog("StartProcess:    *--------------------------------------------------")
                gobjEvent.WriteToEventLog("StartProcess:    *-----  BefFairFeedService - Updating Football ----")
                gobjEvent.WriteToEventLog("StartProcess:    *--------------------------------------------------")


                If My.Settings.Process_MATCH_ODDS Then

                    Dim BetFairDatabase1 As New BetFairDatabaseClass()
                    BetFairDatabase1.PollBetFairEvents(1, "MATCH_ODDS", My.Settings.NumberOfUkEvents, marketCountriesUkOnly, True)
                    BetFairDatabase1 = Nothing

                    ' Europe leagues
                    Dim BetFairDatabase2 As New BetFairDatabaseClass()
                    BetFairDatabase2.PollBetFairEvents(1, "MATCH_ODDS", My.Settings.NumberOfEuropenEvents, marketCountriesEurope, False)
                    BetFairDatabase2 = Nothing

                    'Dim BetFairDatabase3 As New BetFairDatabaseClass()
                    'BetFairDatabase3.PollBetFairCompetition(1, "MATCH_ODDS", My.Settings.NumberOfUkEvents, competitionIds, False)
                    'BetFairDatabase3 = Nothing

                    ' Match new odds
                    Dim BetFairDbOddsMatch1 As New BetFairDatabaseClass()
                    BetFairDbOddsMatch1.MatchSoccerWithBookmakers(1, "MATCH_ODDS")
                    BetFairDbOddsMatch1 = Nothing

                End If


                If My.Settings.Process_HALF_TIME Then

                    Dim BetFairDatabase1 As New BetFairDatabaseClass()
                    BetFairDatabase1.PollBetFairEvents(1, "HALF_TIME", My.Settings.NumberOfUkEvents, marketCountriesUkOnly, True)
                    BetFairDatabase1 = Nothing

                    ' Europe leagues
                    Dim BetFairDatabase2 As New BetFairDatabaseClass()
                    BetFairDatabase2.PollBetFairEvents(1, "HALF_TIME", My.Settings.NumberOfEuropenEvents, marketCountriesEurope, False)
                    BetFairDatabase2 = Nothing

                    'Dim BetFairDatabase3 As New BetFairDatabaseClass()
                    'BetFairDatabase3.PollBetFairCompetition(1, "HALF_TIME", My.Settings.NumberOfUkEvents, competitionIds, False)
                    'BetFairDatabase3 = Nothing

                    ' Match new odds
                    Dim BetFairDbOddsMatch1 As New BetFairDatabaseClass()
                    BetFairDbOddsMatch1.MatchSoccerWithBookmakers(1, "HALF_TIME")
                    BetFairDbOddsMatch1 = Nothing
                End If


                If My.Settings.Process_HALF_TIME_FULL_TIME Then

                    Dim BetFairDatabase1 As New BetFairDatabaseClass()
                    BetFairDatabase1.PollBetFairEvents(1, "HALF_TIME_FULL_TIME", My.Settings.NumberOfUkEvents, marketCountriesUkOnly, True)
                    BetFairDatabase1 = Nothing

                    ' Europe leagues
                    Dim BetFairDatabase2 As New BetFairDatabaseClass()
                    BetFairDatabase2.PollBetFairEvents(1, "HALF_TIME_FULL_TIME", My.Settings.NumberOfEuropenEvents, marketCountriesEurope, False)
                    BetFairDatabase2 = Nothing

                    'Dim BetFairDatabase3 As New BetFairDatabaseClass()
                    'BetFairDatabase3.PollBetFairCompetition(1, "HALF_TIME_FULL_TIME", My.Settings.NumberOfUkEvents, competitionIds, False)
                    'BetFairDatabase3 = Nothing

                    ' Match new odds
                    Dim BetFairDbOddsMatch1 As New BetFairDatabaseClass()
                    BetFairDbOddsMatch1.MatchSoccerWithBookmakers(1, "HALF_TIME_FULL_TIME")
                    BetFairDbOddsMatch1 = Nothing

                End If


                If My.Settings.Process_OVER_UNDER Then

                    'Dim BetFairDatabase1 As New BetFairDatabaseClass()
                    'BetFairDatabase1.PollBetFairEvents(1, "OVER_UNDER_25", My.Settings.NumberOfUkEvents, marketCountriesUkOnly, True)
                    'BetFairDatabase1 = Nothing

                    '' Europe leagues
                    'Dim BetFairDatabase2 As New BetFairDatabaseClass()
                    'BetFairDatabase2.PollBetFairEvents(1, "OVER_UNDER_25", My.Settings.NumberOfEuropenEvents, marketCountriesEurope, False)
                    'BetFairDatabase2 = Nothing

                    Dim BetFairDatabase0 As New BetFairDatabaseClass()
                    BetFairDatabase0.PollBetFairCompetition(1, "OVER_UNDER_05", My.Settings.NumberOfUkEvents, competitionIds, False)
                    BetFairDatabase0 = Nothing

                    ' Match new odds
                    Dim BetFairDbOddsMatch0 As New BetFairDatabaseClass()
                    BetFairDbOddsMatch0.MatchSoccerWithBookmakers(1, "OVER_UNDER_05")
                    BetFairDbOddsMatch0 = Nothing



                    Dim BetFairDatabase1 As New BetFairDatabaseClass()
                    BetFairDatabase1.PollBetFairCompetition(1, "OVER_UNDER_15", My.Settings.NumberOfUkEvents, competitionIds, False)
                    BetFairDatabase1 = Nothing

                    ' Match new odds
                    Dim BetFairDbOddsMatch1 As New BetFairDatabaseClass()
                    BetFairDbOddsMatch1.MatchSoccerWithBookmakers(1, "OVER_UNDER_15")
                    BetFairDbOddsMatch1 = Nothing


                    Dim BetFairDatabase2 As New BetFairDatabaseClass()
                    BetFairDatabase2.PollBetFairCompetition(1, "OVER_UNDER_25", My.Settings.NumberOfUkEvents, competitionIds, False)
                    BetFairDatabase2 = Nothing

                    ' Match new odds
                    Dim BetFairDbOddsMatch2 As New BetFairDatabaseClass()
                    BetFairDbOddsMatch2.MatchSoccerWithBookmakers(1, "OVER_UNDER_25")
                    BetFairDbOddsMatch2 = Nothing



                    Dim BetFairDatabase3 As New BetFairDatabaseClass()
                    BetFairDatabase3.PollBetFairCompetition(1, "OVER_UNDER_35", My.Settings.NumberOfUkEvents, competitionIds, False)
                    BetFairDatabase3 = Nothing

                    ' Match new odds
                    Dim BetFairDbOddsMatch3 As New BetFairDatabaseClass()
                    BetFairDbOddsMatch3.MatchSoccerWithBookmakers(1, "OVER_UNDER_35")
                    BetFairDbOddsMatch3 = Nothing



                    Dim BetFairDatabase4 As New BetFairDatabaseClass()
                    BetFairDatabase4.PollBetFairCompetition(1, "OVER_UNDER_45", My.Settings.NumberOfUkEvents, competitionIds, False)
                    BetFairDatabase4 = Nothing

                    ' Match new odds
                    Dim BetFairDbOddsMatch4 As New BetFairDatabaseClass()
                    BetFairDbOddsMatch4.MatchSoccerWithBookmakers(1, "OVER_UNDER_45")
                    BetFairDbOddsMatch4 = Nothing



                    Dim BetFairDatabase5 As New BetFairDatabaseClass()
                    BetFairDatabase5.PollBetFairCompetition(1, "OVER_UNDER_55", My.Settings.NumberOfUkEvents, competitionIds, False)
                    BetFairDatabase5 = Nothing

                    ' Match new odds
                    Dim BetFairDbOddsMatch5 As New BetFairDatabaseClass()
                    BetFairDbOddsMatch5.MatchSoccerWithBookmakers(1, "OVER_UNDER_55")
                    BetFairDbOddsMatch5 = Nothing



                    Dim BetFairDatabase6 As New BetFairDatabaseClass()
                    BetFairDatabase6.PollBetFairCompetition(1, "OVER_UNDER_65", My.Settings.NumberOfUkEvents, competitionIds, False)
                    BetFairDatabase6 = Nothing

                    ' Match new odds
                    Dim BetFairDbOddsMatch6 As New BetFairDatabaseClass()
                    BetFairDbOddsMatch6.MatchSoccerWithBookmakers(1, "OVER_UNDER_65")
                    BetFairDbOddsMatch6 = Nothing



                    Dim BetFairDatabase7 As New BetFairDatabaseClass()
                    BetFairDatabase7.PollBetFairCompetition(1, "OVER_UNDER_75", My.Settings.NumberOfUkEvents, competitionIds, False)
                    BetFairDatabase7 = Nothing

                    ' Match new odds
                    Dim BetFairDbOddsMatch7 As New BetFairDatabaseClass()
                    BetFairDbOddsMatch7.MatchSoccerWithBookmakers(1, "OVER_UNDER_75")
                    BetFairDbOddsMatch7 = Nothing




                    Dim BetFairDatabase8 As New BetFairDatabaseClass()
                    BetFairDatabase8.PollBetFairCompetition(1, "OVER_UNDER_85", My.Settings.NumberOfUkEvents, competitionIds, False)
                    BetFairDatabase8 = Nothing

                    ' Match new odds
                    Dim BetFairDbOddsMatch8 As New BetFairDatabaseClass()
                    BetFairDbOddsMatch8.MatchSoccerWithBookmakers(1, "OVER_UNDER_85")
                    BetFairDbOddsMatch8 = Nothing



                End If


                If My.Settings.Prcoess_CORRECT_SCORE Then

                    Dim BetFairDatabase1 As New BetFairDatabaseClass()
                    BetFairDatabase1.PollBetFairEvents(1, "CORRECT_SCORE", My.Settings.NumberOfUkEvents, marketCountriesUkOnly, True)
                    BetFairDatabase1 = Nothing

                    ' Europe leagues
                    Dim BetFairDatabase2 As New BetFairDatabaseClass()
                    BetFairDatabase2.PollBetFairEvents(1, "CORRECT_SCORE", My.Settings.NumberOfEuropenEvents, marketCountriesEurope, False)
                    BetFairDatabase2 = Nothing

                    'Dim BetFairDatabase3 As New BetFairDatabaseClass()
                    'BetFairDatabase3.PollBetFairCompetition(1, "CORRECT_SCORE", My.Settings.NumberOfUkEvents, competitionIds, False)
                    'BetFairDatabase3 = Nothing

                    ' Match new odds
                    Dim BetFairDbOddsMatch1 As New BetFairDatabaseClass()
                    BetFairDbOddsMatch1.MatchSoccerWithBookmakers(1, "CORRECT_SCORE")
                    BetFairDbOddsMatch1 = Nothing

                End If

                If My.Settings.Prcoess_CORRECT_SCORE Then

                    'Dim BetFairDatabase7 As New BetFairDatabaseClass()
                    'BetFairDatabase7.PollBetFairEvents(1, "BOTH_TEAMS_TO_SCORE", My.Settings.NumberOfUkEvents, marketCountriesUkOnly, True)
                    'BetFairDatabase7 = Nothing

                End If

            End If

            ' ------------------------------
            ' Horse Racing - eventTypeId = 7
            ' ------------------------------
            ' MarketTypeCodes applicable : "WIN", "PLACE"
            '
            If My.Settings.StreamSportId = 7 Then

                gobjEvent.WriteToEventLog("StartProcess:    *------------------------------------------------------")
                gobjEvent.WriteToEventLog("StartProcess:    *-----  BefFairFeedService - Updating Horse Racing ----")
                gobjEvent.WriteToEventLog("StartProcess:    *------------------------------------------------------")
                Dim BetFairHorseRacingDatabase1 As New BetFairDatabaseClass()
                BetFairHorseRacingDatabase1.PollBetFairEvents(7, "WIN", My.Settings.NumberOfUkEvents, marketCountriesUkOnly, True)
                BetFairHorseRacingDatabase1 = Nothing

                Dim BetFairHorseRacingDatabase2 As New BetFairDatabaseClass()
                BetFairHorseRacingDatabase2.PollBetFairEvents(7, "PLACE", My.Settings.NumberOfUkEvents, marketCountriesUkOnly, True)
                BetFairHorseRacingDatabase2 = Nothing

            End If

        Catch ex As Exception

            gobjEvent.WriteToEventLog("StartProcess : Process has been killed, general error : " & ex.Message, EventLogEntryType.Error)
            Return False

        End Try

        Return True

    End Function

End Class

