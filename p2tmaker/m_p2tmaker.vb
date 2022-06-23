
Imports System.IO

Module m_p2tmaker

    Private log As New List(Of String)
    Private logFileName As String = String.Empty
    Private HeavyRainLimit As Double = 50
    Private maxRain As Double = Double.NaN
    Private twiceHeavyRain As Integer = 0
    Private reportHeavyRain As Boolean = False
    Private reportSnowMelt As Boolean = False
    Private reportIrrigation As Boolean = False
    Private p2tFileInfo As FileInfo

    Const hearderRowNo As Integer = 2

    Private keepLogFileName As Boolean = False

    <DebuggerStepThrough>
    Private Sub add2Log(entry As String)

        log.Add(entry)
        Console.WriteLine(log.Last)

        Try

            If recursive And ArgPathRecurisve <> String.Empty Then
                logFileName =
                        Path.Combine(
                            ArgPathRecurisve,
                            "p2t.log")
            ElseIf Not keepLogFileName Then

                If PRZMRunDir = String.Empty Then
                    logFileName =
                        Path.Combine(
                            Environment.CurrentDirectory,
                            "p2t.log")

                Else
                    logFileName =
                        Path.Combine(
                            PRZMRunDir,
                            "p2t.log")

                    keepLogFileName = True

                End If

            End If

            File.WriteAllLines(
                path:=logFileName,
                contents:=log.ToArray)

        Catch ex As Exception
            Console.WriteLine("IO Error writing log to" & vbCrLf &
                              logFileName & vbCrLf &
                              ex.Message)
        End Try


    End Sub


    Sub Main()

        'switch regional settings to English
        My.Application.ChangeCulture("en-US")

        'final output
        Dim out As New List(Of String)
        Dim P2TFileName As String = String.Empty
        Dim SWASHNameKonvention As Boolean
        Dim stdCMPNames As Boolean

        'get cmd line arguments
        getCMDArgs()

        add2Log("")
        add2Log("")
        add2Log(
            Join(
                SourceArray:=getApplnInfo(Leadingstring:=("   ")),
                Delimiter:=vbCrLf))

        ' and parse them

        getDescription()

        getWarmUp()

        getMonthlyAverage()

        If Not monthlyAverage Then
            getMRT()
        End If

        getSeasonOnly()

        getmaxPREC()

        getEXP()

        getRecursive()

        If Not recursive Then
            'cmd = zts:=...
            getSingleZTS()
        End If

        'or cmd = path:=...
        If ZTSfiles2go.Count = 0 Then
            getAllZTSinDir()
        End If

        For Each ZTSFileName As String In ZTSfiles2go

            Try
                PRZMRunDir = Path.GetDirectoryName(ZTSFileName)
            Catch ex As Exception
                add2Log(
                    entry:=("IO Error:=").PadLeft(logLen) & vbCrLf &
                            "Can't get run directory from ZTSFile " & vbCrLf &
                            ex.Message)
                End
            End Try

            'check for new przm run directory
            If OldPRZMRunDir = String.Empty OrElse
               OldPRZMRunDir <> PRZMRunDir Then

                OldPRZMRunDir = PRZMRunDir

                add2Log("")

                add2Log(
                    entry:=(" ").PadLeft(logLen) & " ****************************************************** ")

                add2Log(
                    entry:=("Actual path:=").PadLeft(logLen) & Path.GetDirectoryName(path:=ZTSFileName))

                add2Log(
                    entry:=(" ").PadLeft(logLen) & " ****************************************************** ")


                'get cmp names from master.fpj
                stdCMPNames = getCMPNames(baseName:=Path.GetFileNameWithoutExtension(ZTSFileName),
                            userCMPnames:=userCMPNames)

                'get SWASH numbers from przm.pzm
                SWASHNameKonvention = getSWASHNos()

                'get crop from filename
                Try

                    Crop = getPRZMCropFromFileName(ZTSFileName:=ZTSFileName)
                    add2Log(entry:=("Crop:=").PadLeft(logLen) & Crop)

                    'Scenario = getPRZMScenarioFromFilename(ZTSFileName:=ZTSFileName)
                    'add2Log(entry:=("Scenario:=").PadLeft(logLen) & Scenario)

                Catch ex As Exception
                    add2Log(
                           entry:=("Parsing Error:=").PadLeft(logLen) &
                           "Can't parse crop or scenario name from filename " & vbCrLf &
                           ZTSfiles2go.First & vbCrLf & ex.Message)
                End Try

            Else
                add2Log("")
            End If


            add2Log(
                entry:=(" ").PadLeft(logLen) & " ****************************************************** ")

            add2Log(
                entry:=("ZTS:=").PadLeft(logLen) & Path.GetFileName(ZTSFileName))

            Scenario = getPRZMScenarioFromFilename(ZTSFileName:=ZTSFileName)
            add2Log(entry:=("Scenario:=").PadLeft(logLen) & Scenario)

            'init
            applns.Clear()
            applnsSeason.Clear()
            p2tHeader.Clear()
            p2tDataParent.Clear()
            p2tDataMet01.Clear()
            p2tDataMet02.Clear()
            HeavyRain.Clear()
            twiceHeavyRain = 0
            maxRain = 0
            irrigation.Clear()
            seasonStart = New Date
            seasonEnd = New Date
            out.Clear()

            'get data
            If Not createP2T(
                ZTSFileName:=ZTSFileName) Then

                Console.Beep()
                add2Log(entry:="Major error creating p2t data")
                Continue For

            End If

            If irrigation.Count <> 0 Then
                add2Log(
                    entry:=("Irrigation:=").PadLeft(logLen) & "true")
            Else
                add2Log(
                    entry:=("Irrigation:=").PadLeft(logLen) & "false")
            End If

            If HeavyRain.Count <> 0 Then
                add2Log(
                   entry:=("Rain > " & HeavyRainLimit & "mm :=").PadLeft(logLen) & HeavyRain.Count & " times")

                add2Log(
                   entry:=("Rain > 2*" & HeavyRainLimit & "mm:=").PadLeft(logLen) & twiceHeavyRain & " times")

                add2Log(
                   entry:=("Max rainfall:=").PadLeft(logLen) & maxRain & " mm")
            End If

            If Not createHeader(
                ParMet:=eParMet.Par,
                ZTSFileName:=ZTSFileName) Then

                Console.Beep()
                add2Log(entry:="Major error creating p2t header")
                Continue For

            End If

            out.AddRange(p2tHeader)
            out.AddRange(p2tDataParent)

            Try

                If SWASHno <> -99 Then
                    P2TFileName = Path.Combine(
                                        PRZMRunDir,
                                        SWASHno.ToString("00000") & "-C1.p2t")
                Else
                    P2TFileName = Path.Combine(
                                    PRZMRunDir,
                                    Path.GetFileNameWithoutExtension(
                                            path:=ZTSFileName) & "-" & Parent & ".p2t")

                    P2TFileName = Replace(Expression:=P2TFileName,
                                             Find:="--",
                                             Replacement:="-",
                                             Compare:=CompareMethod.Text)

                End If

                File.WriteAllLines(
                        path:=P2TFileName,
                        contents:=out.ToArray)

                p2tFileInfo = New FileInfo(P2TFileName)

                With p2tFileInfo

                    If Not .Exists Then
                        add2Log(
                            entry:=(Parent & " p2t:=").PadLeft(logLen) & " doesn't exist")

                        Process.Start(fileName:=logFileName)
                        End

                    Else
                        add2Log(
                            entry:=(Parent & " p2t:=").PadLeft(logLen) & .Name)
                        add2Log(
                           entry:=(" ").PadLeft(logLen) & Math.Round(.Length / 1000000, 0) & "MB")
                    End If

                End With

            Catch ex As Exception

                add2Log(
                    entry:=("IO Error:=").PadLeft(logLen) & ex.Message)

                Process.Start(fileName:=logFileName)
                End

            End Try


            If Met01 <> String.Empty Then

                createHeader(
                    ParMet:=eParMet.Met01,
                    ZTSFileName:=ZTSFileName)

                out.Clear()
                out.AddRange(p2tHeader)
                out.AddRange(p2tDataMet01)

                Try

                    If SWASHno <> -99 Then
                        P2TFileName = Path.Combine(
                                        PRZMRunDir,
                                        SWASHno.ToString("00000") & "-C2.p2t")
                    Else
                        P2TFileName = Path.Combine(
                                            PRZMRunDir,
                                            Path.GetFileNameWithoutExtension(
                                                path:=ZTSFileName) & "-" & Met01 & ".p2t")

                        P2TFileName = Replace(Expression:=P2TFileName,
                                             Find:="--",
                                             Replacement:="-",
                                             Compare:=CompareMethod.Text)

                    End If


                    File.WriteAllLines(
                            path:=P2TFileName,
                            contents:=out.ToArray)

                    p2tFileInfo = New FileInfo(P2TFileName)

                    With p2tFileInfo

                        If Not .Exists Then
                            add2Log(
                            entry:=(Met01 & " p2t:=").PadLeft(logLen) & " doesn't exist")

                            Process.Start(fileName:=logFileName)
                            End

                        Else
                            add2Log(
                            entry:=(Met01 & " p2t:=").PadLeft(logLen) & .FullName)
                            add2Log(
                           entry:=(" ").PadLeft(logLen) & Math.Round(.Length / 1000000, 0) & "MB")
                        End If

                    End With

                Catch ex As Exception

                    add2Log(
                        entry:=("IO Error:=").PadLeft(logLen) & ex.Message)

                    Process.Start(fileName:=logFileName)
                    End

                End Try

            End If

            If Met02 <> String.Empty Then

                createHeader(
                    ParMet:=eParMet.Met02,
                    ZTSFileName:=ZTSFileName)

                out.Clear()
                out.AddRange(p2tHeader)
                out.AddRange(p2tDataMet02)

                Try

                    If SWASHno <> -99 Then
                        P2TFileName = Path.Combine(
                                        PRZMRunDir,
                                        SWASHno.ToString("00000") & "-C3.p2t")
                    Else
                        P2TFileName = Path.Combine(
                                            PRZMRunDir,
                                            Path.GetFileNameWithoutExtension(
                                                path:=ZTSFileName) & "-" & Met02 & ".p2t")

                        P2TFileName = Replace(Expression:=P2TFileName,
                                              Find:="--",
                                              Replacement:="-",
                                              Compare:=CompareMethod.Text)
                    End If

                    File.WriteAllLines(
                            path:=P2TFileName,
                            contents:=out.ToArray)

                    p2tFileInfo = New FileInfo(P2TFileName)

                    With p2tFileInfo

                        If Not .Exists Then

                            add2Log(
                            entry:=(Met02 & " p2t:=").PadLeft(logLen) & " doesn't exist")

                            Process.Start(fileName:=logFileName)
                            End

                        Else
                            add2Log(
                            entry:=(Met02 & " p2t:=").PadLeft(logLen) & .FullName)
                            add2Log(
                           entry:=(" ").PadLeft(logLen) & Math.Round(.Length / 1000000, 0) & "MB")
                        End If

                    End With

                Catch ex As Exception

                    add2Log(
                        entry:=("IO Error:=").PadLeft(logLen) & ex.Message)

                    Process.Start(fileName:=logFileName)
                    End

                End Try

            End If


            If reportHeavyRain Then
                add2Log("")
                add2Log(entry:=(" ").PadLeft(logLen) & "Heavy rainfall events (>=" & HeavyRainLimit & "mm)")
                add2Log(
                    Join(
                        SourceArray:=HeavyRain.ToArray,
                        Delimiter:=vbCrLf))
            End If

            If reportIrrigation Then
                add2Log("")
                add2Log("Irrigation dates and volumes")
                add2Log(
                    Join(
                        SourceArray:=irrigation.ToArray,
                        Delimiter:=vbCrLf))
            End If

        Next

    End Sub


#Region "    internal stuff"

    ' row where then data in the zts file starts
    Const ZTSDataStartRowNo As Integer = 3
    ' row where then header in the zts file starts
    Const ZTSHeaderRowNo As Integer = 2
    'format the log file
    Private logLen As Integer = 20
    Private stdPos As Integer = 30


    ''' <summary>
    ''' Std. member of the zts file
    ''' </summary>
    ''' <remarks></remarks>
    Private Enum eZTSHeader

        TRS = 0
        PRZ = 1

        EventYear = 2
        EventMonth
        EventDay

        RUNF
        ESLS
        PRCP
        INFL

        RFLX1
        EFLX1

        RFLX2
        EFLX2

        RFLX3
        EFLX3

    End Enum


    ''' <summary>
    ''' Position of the appln info in zts file
    ''' </summary>
    ''' <remarks></remarks>
    Private posTPAP As Integer = -1

    ''' <summary>
    ''' Position of the irrigation info in zts file
    ''' </summary>
    ''' <remarks></remarks>
    Private posIRRG As Integer = -1


    Private Function getApplnInfo(Optional Leadingstring As String = "*  ") As String()

        Dim out As New List(Of String)


        With out
            .Add(Leadingstring & My.Application.Info.ProductName &
                             "   v" & My.Application.Info.Version.ToString)


            .Add(Leadingstring & "Started on " & Now.ToLongDateString)
            .Add(Leadingstring & "        at " & Now.ToLongTimeString)

            .Add(Leadingstring & "on machine " & Environment.MachineName &
                                          " (" & Environment.ProcessorCount & " CPUs)")
            .Add(Leadingstring & "        OS " & Environment.OSVersion.VersionString)
            .Add(Leadingstring & "Culture    " & My.Application.Culture.ToString)
            .Add(Leadingstring)

        End With

        Return out.ToArray

    End Function

    Private Sub writeUsage(Optional Leadingstring As String = "*  ")

        Console.Clear()
        add2Log(entry:=Leadingstring)
        add2Log(entry:=Leadingstring)
        add2Log(entry:=
            Join(
                SourceArray:=getApplnInfo,
                Delimiter:=vbCrLf))
        add2Log(entry:=Leadingstring & "Usage:")
        add2Log(entry:=Leadingstring)
        add2Log(entry:=Leadingstring & "convert a single zts file")
        add2Log(entry:=Leadingstring & "zts:='full zts file path'")
        add2Log(entry:=Leadingstring & "OR")
        add2Log(entry:=Leadingstring & "convert all zts files in a project directory")
        add2Log(entry:=Leadingstring & "path:='full path to directory with zts files'")
        add2Log(entry:=Leadingstring & "-------------------------------------------------")
        add2Log(entry:=Leadingstring & "get ZTS files recursively")
        add2Log(entry:=Leadingstring & "recursive:=true/false")
        add2Log(entry:=Leadingstring & "-------------------------------------------------")
        add2Log(entry:=Leadingstring & "if 'p2tmaker.exe' is in the project directory")
        add2Log(entry:=Leadingstring & "it can be used without these two cmd line args")
        add2Log(entry:=Leadingstring & "-------------------------------------------------")
        add2Log(entry:=Leadingstring & "To skip warm up years, default = 0")
        add2Log(entry:=Leadingstring & "warmup:=0")
        add2Log(entry:=Leadingstring & "-------------------------------------------------")
        add2Log(entry:=Leadingstring & "To set the mean residence time in days, default = 20days")
        add2Log(entry:=Leadingstring & "mrt:=20")
        add2Log(entry:=Leadingstring & "-------------------------------------------------")
        add2Log(entry:=Leadingstring & "To set the max. precipitation per hour for")
        add2Log(entry:=Leadingstring & "calculation of event duration in mm, default = 2mm")
        add2Log(entry:=Leadingstring & "maxPREC:=2")
        add2Log(entry:=Leadingstring & "-------------------------------------------------")
        add2Log(entry:=Leadingstring & "GW_discharge calc. with exponential ")
        add2Log(entry:=Leadingstring & "discharge formula, std. = false")
        add2Log(entry:=Leadingstring & "exp:=true/false")
        add2Log(entry:=Leadingstring & "-------------------------------------------------")
        add2Log(entry:=Leadingstring)
        add2Log(entry:=Leadingstring)

        For Each member As String In log
            Console.WriteLine(member)
        Next


    End Sub

#End Region


    ''' <summary>
    ''' List of cmd line arguments
    ''' </summary>
    ''' <remarks></remarks>
    Private arguments As New List(Of String)

    ''' <summary>
    ''' Mean residence time in days, std. = 20d
    ''' Used for
    ''' GW_discharge = (1 / MRT) * GW_storage
    ''' </summary>
    ''' <remarks></remarks>
    Private MRT As Double = 20

    ''' <summary>
    ''' use monthly average for INFL, std. = false
    ''' </summary>
    Private monthlyAverage As Boolean = False

    ''' <summary>
    ''' compound names given by user
    ''' </summary>
    Private userCMPNames As String = Nothing

    ''' <summary>
    ''' only old FOCUS season, std. = false
    ''' </summary>
    Private seasonOnly As Boolean = False

    ''' <summary>
    ''' Max precipitation per hour in mm
    ''' used to calc. the event duration, std. = 2mm
    ''' Math.Round(PRCP / MaxPRECperHour, digits:=0, mode:=MidpointRounding.AwayFromZero)
    ''' </summary>
    ''' <remarks></remarks>
    Private MaxPRECperHour As Double = 2

    ''' <summary>
    ''' Warm up years to skip for p2t output, std. = 0 years
    ''' </summary>
    ''' <remarks></remarks>
    Private WarmUp As Integer = 0

    ''' <summary>
    ''' add Description to the p2t file header
    ''' </summary>
    Private description As String = String.Empty

    ''' <summary>
    ''' GW_discharge  calc. with exponential discharge formula, std. = false
    ''' </summary>
    ''' <remarks></remarks>
    Private Exp As Boolean = False

    ''' <summary>
    ''' get zts file recursively
    ''' </summary>
    ''' <remarks></remarks>
    Private recursive As Boolean = True


    Private PRZMRunDir As String = String.Empty
    Private OldPRZMRunDir As String = String.Empty
    Private ArgPathRecurisve As String = String.Empty

    Private ZTSfiles2go As New List(Of String)
    Private SWASHNos As New List(Of Integer)

    Private Parent As String = String.Empty
    Private Met01 As String = String.Empty
    Private Met02 As String = String.Empty


    Private SWASHno As Integer = -1
    Private Scenario As String = String.Empty
    Private Crop As String = String.Empty
    Private OldCrop As String = String.Empty


    Private SimStart As Date
    Private SimEnd As Date

    Private seasonStart As New Date
    Private seasonEnd As New Date

    Private applns As New List(Of String)
    Private applnsSeason As New List(Of String)

    Private p2tHeader As New List(Of String)
    Private p2tDataParent As New List(Of String)
    Private p2tDataMet01 As New List(Of String)
    Private p2tDataMet02 As New List(Of String)

    Private HeavyRain As New List(Of String)
    Private irrigation As New List(Of String)

#Region "    get and parse CMD line arguments"

    Const argZTS As String = "zts:="
    Const argPath As String = "path:="
    Const argWarmup As String = "warmup:="
    Const argMRT As String = "mrt:="
    Const argMonthlyAverage As String = "monthlyAverage:="
    Const argUserCMPname As String = "UserCMPname:="
    Const argSeasonOnly As String = "seasonOnly:="
    Const argMaxPRECperHour As String = "maxPREC:="
    Const argEXP As String = "exp:="
    Const argRecursive As String = "recursive:="
    Const argDescription As String = "description:="

    Private Sub getCMDArgs()

        Dim PathOrZTS As New List(Of String)

        Try

            With arguments

                .AddRange(Environment.GetCommandLineArgs())
                .RemoveAt(index:=0)

            End With

            If Filter(
                Source:=arguments.ToArray,
                Match:="?", Include:=True,
                Compare:=CompareMethod.Text).Count <> 0 OrElse
               Filter(
                   Source:=arguments.ToArray,
                   Match:="help",
                   Include:=True,
                   Compare:=CompareMethod.Text).Count <> 0 Then

                writeUsage()

                Process.Start(fileName:=logFileName)
                End

            Else

                PathOrZTS.AddRange(
                    Filter(
                        Source:=arguments.ToArray,
                        Match:=argPath,
                        Include:=True,
                        Compare:=CompareMethod.Text))

                PathOrZTS.AddRange(
                   Filter(
                       Source:=arguments.ToArray,
                       Match:=argZTS,
                       Include:=True,
                       Compare:=CompareMethod.Text))

                If PathOrZTS.Count = 0 Then
                    'no arguments => path = actual exec. path
                    arguments.Add(argPath & Environment.CurrentDirectory)
                End If


            End If

        Catch ex As Exception

            add2Log(entry:="Error parsing cmd line arguments " & vbCrLf &
                      ex.Message & vbCrLf &
                      Join(
                          SourceArray:=Environment.GetCommandLineArgs(),
                          Delimiter:=vbCrLf))

            Process.Start(fileName:=logFileName)
            End

        End Try

    End Sub

    Private Sub getDescription()

        Dim tempFilter As String() = {}
        Dim std As Boolean = True

        tempFilter =
            Filter(
                Source:=arguments.ToArray,
                Match:=argDescription,
                Include:=True,
                Compare:=CompareMethod.Text)

        If tempFilter.Count = 1 Then

            Try

                description =
                    Replace(
                        Expression:=tempFilter.First,
                        Find:=argDescription,
                        Replacement:="",
                        Compare:=CompareMethod.Text)

                std = False

            Catch ex As Exception

                add2Log(entry:=
                    "Error parsing cmd line for " & argDescription & vbCrLf &
                     tempFilter.First & vbCrLf &
                     ex.Message)

                writeUsage()

                Process.Start(fileName:=logFileName)
                End

            End Try

        End If

    End Sub

    Private Sub getWarmUp()

        Dim tempFilter As String() = {}
        Dim std As Boolean = True

        tempFilter =
            Filter(
                Source:=arguments.ToArray,
                Match:=argWarmup,
                Include:=True,
                Compare:=CompareMethod.Text)

        If tempFilter.Count = 1 Then

            Try

                WarmUp =
                    CInt(Replace(
                        Expression:=tempFilter.First,
                        Find:=argWarmup,
                        Replacement:="",
                        Compare:=CompareMethod.Text))

                std = False

            Catch ex As Exception

                add2Log(entry:=
                    "Error parsing cmd line for " & argWarmup & vbCrLf &
                     tempFilter.First & vbCrLf &
                     ex.Message)

                writeUsage()

                Process.Start(fileName:=logFileName)
                End

            End Try

        End If

        add2Log(
             entry:=((argWarmup).PadLeft(logLen) & WarmUp.ToString & " years").PadRight(stdPos) &
                        IIf(std, " (std.)", " *** user def. ***").ToString)

    End Sub

    Private Sub getMRT()

        Dim tempFilter As String() = {}
        Dim std As Boolean = True

        tempFilter =
            Filter(
                Source:=arguments.ToArray,
                Match:=argMRT,
                Include:=True,
                Compare:=CompareMethod.Text)

        If tempFilter.Count = 1 Then

            Try

                MRT =
                    CDbl(Replace(
                        Expression:=tempFilter.First,
                        Find:=argMRT,
                        Replacement:="",
                        Compare:=CompareMethod.Text))

                std = False

            Catch ex As Exception

                add2Log(entry:=
                    "Error parsing cmd line for " & argMRT & vbCrLf &
                     tempFilter.First & vbCrLf &
                     ex.Message)

                writeUsage()

                Process.Start(fileName:=logFileName)
                End

            End Try

        End If

        add2Log(
            entry:=((argMRT).PadLeft(logLen) & MRT.ToString & "days").PadRight(stdPos) &
                        IIf(std, " (std.)", " *** user def. ***").ToString)

    End Sub

    Private Sub getMonthlyAverage()

        Dim tempFilter As String() = {}
        Dim std As Boolean = True

        tempFilter =
            Filter(
                Source:=arguments.ToArray,
                Match:=argMonthlyAverage,
                Include:=True,
                Compare:=CompareMethod.Text)

        If tempFilter.Count = 1 Then

            Try

                monthlyAverage =
                    CBool(Replace(
                            Expression:=tempFilter.First,
                            Find:=argMonthlyAverage,
                            Replacement:="",
                            Compare:=CompareMethod.Text))

                std = False

            Catch ex As Exception

                add2Log(entry:=
                    "Error parsing cmd line for " & argMonthlyAverage & vbCrLf &
                     tempFilter.First & vbCrLf &
                     ex.Message)

                writeUsage()

                Process.Start(fileName:=logFileName)
                End

            End Try

        End If

        add2Log(
            entry:=((argMonthlyAverage).PadLeft(logLen) & monthlyAverage.ToString).PadRight(stdPos) &
                        IIf(std, " (std.)", " *** user def. ***").ToString)

    End Sub

    Private Sub getUserCMPnames()

        Dim tempFilter As String() = {}
        Dim std As Boolean = True

        tempFilter =
            Filter(
                Source:=arguments.ToArray,
                Match:=argMonthlyAverage,
                Include:=True,
                Compare:=CompareMethod.Text)

        If tempFilter.Count = 1 Then

            Try

                userCMPNames =
                    CStr(Replace(
                            Expression:=tempFilter.First,
                            Find:=argUserCMPname,
                            Replacement:="",
                            Compare:=CompareMethod.Text))

                std = False

            Catch ex As Exception

                add2Log(entry:=
                    "Error parsing cmd line for " & argUserCMPname & vbCrLf &
                     tempFilter.First & vbCrLf &
                     ex.Message)

                writeUsage()

                Process.Start(fileName:=logFileName)
                End

            End Try

        End If

        add2Log(
            entry:=((argUserCMPname).PadLeft(logLen) & userCMPNames.ToString).PadRight(stdPos) &
                        IIf(std, " (std.)", " *** user def. ***").ToString)

    End Sub

    Private Sub getSeasonOnly()

        Dim tempFilter As String() = {}
        Dim std As Boolean = True

        tempFilter =
            Filter(
                Source:=arguments.ToArray,
                Match:=argSeasonOnly,
                Include:=True,
                Compare:=CompareMethod.Text)

        If tempFilter.Count = 1 Then

            Try

                seasonOnly =
                    CBool(Replace(
                            Expression:=tempFilter.First,
                            Find:=argSeasonOnly,
                            Replacement:="",
                            Compare:=CompareMethod.Text))

                std = False

            Catch ex As Exception

                add2Log(entry:=
                    "Error parsing cmd line for " & argSeasonOnly & vbCrLf &
                     tempFilter.First & vbCrLf &
                     ex.Message)

                writeUsage()

                Process.Start(fileName:=logFileName)
                End

            End Try

        End If

        add2Log(
            entry:=((argSeasonOnly).PadLeft(logLen) & seasonOnly.ToString).PadRight(stdPos) &
                        IIf(std, " (std.)", " *** user def. ***").ToString)

    End Sub

    Private Sub getmaxPREC()

        Dim tempFilter As String() = {}
        Dim std As Boolean = True

        tempFilter =
            Filter(
                Source:=arguments.ToArray,
                Match:=argMaxPRECperHour,
                Include:=True,
                Compare:=CompareMethod.Text)

        If tempFilter.Count = 1 Then

            Try

                MaxPRECperHour =
                    CDbl(Replace(
                        Expression:=tempFilter.First,
                        Find:=argMaxPRECperHour,
                        Replacement:="",
                        Compare:=CompareMethod.Text))

                std = False

            Catch ex As Exception

                add2Log(entry:=
                    "Error parsing cmd line for " & argMaxPRECperHour & vbCrLf &
                     tempFilter.First & vbCrLf &
                     ex.Message)

                writeUsage()

                Process.Start(fileName:=logFileName)
                End

            End Try

        End If

        add2Log(
           entry:=((argMaxPRECperHour).PadLeft(logLen) & MaxPRECperHour.ToString & "mm").PadRight(stdPos) &
                       IIf(std, " (std.)", " *** user def. ***").ToString)

    End Sub

    Private Sub getEXP()

        Dim tempFilter As String() = {}
        Dim std As Boolean = True

        tempFilter =
            Filter(
                Source:=arguments.ToArray,
                Match:=argEXP,
                Include:=True,
                Compare:=CompareMethod.Text)

        If tempFilter.Count = 1 Then

            Try

                Exp =
                    CBool(Replace(
                        Expression:=tempFilter.First,
                        Find:=argEXP,
                        Replacement:="",
                        Compare:=CompareMethod.Text))

                std = False

            Catch ex As Exception

                add2Log(entry:=
                    "Error parsing cmd line for " & argEXP & vbCrLf &
                     tempFilter.First & vbCrLf &
                     ex.Message)

                writeUsage()

                Process.Start(fileName:=logFileName)
                End

            End Try

        End If

        add2Log(
            entry:=((argEXP).PadLeft(logLen) & Exp.ToString).PadRight(stdPos) &
                        IIf(std, " (std.)", " *** user def. ***").ToString)

    End Sub

    Private Sub getSingleZTS()

        Dim tempFilter As String() = {}

        'zts file
        tempFilter =
            Filter(
                Source:=arguments.ToArray,
                Match:=argZTS,
                Include:=True,
                Compare:=CompareMethod.Text)

        If tempFilter.Count = 1 Then

            Try

                ZTSfiles2go.Add(
                   Trim(
                       Replace(
                        Expression:=tempFilter.First,
                        Find:=argZTS,
                        Replacement:="",
                        Compare:=CompareMethod.Text)))

                add2Log(entry:=(argZTS).PadLeft(logLen) & ZTSfiles2go.Last)

            Catch ex As Exception

                add2Log(entry:=
                    "Error parsing cmd line for " & argZTS & vbCrLf &
                     tempFilter.First & vbCrLf &
                     ex.Message)

                writeUsage()

                Process.Start(fileName:=logFileName)
                End

            End Try

        End If

    End Sub

    Private Sub getAllZTSinDir()

        Dim argFilter As String() = {}
        Dim ZTSFilePath As String = String.Empty
        Dim ZTSFiles As String() = {}

        'get path from arguments
        argFilter =
            Filter(
                Source:=arguments.ToArray,
                Match:=argPath,
                Include:=True,
                Compare:=CompareMethod.Text)

        If argFilter.Count = 0 Then
            add2Log(entry:=
                       "No path for *.zts files in cmd line arguments ")

            writeUsage()

            Process.Start(fileName:=logFileName)
            End

        Else

            ZTSFilePath =
                Replace(
                    Expression:=argFilter.First,
                    Find:=argPath,
                    Replacement:="",
                    Compare:=CompareMethod.Text)

            ArgPathRecurisve = ZTSFilePath

        End If

        'get *.zts files from path
        Try
            If recursive Then
                ZTSFiles =
                    Directory.GetFiles(
                        path:=ZTSFilePath,
                        searchPattern:="*.zts",
                        searchOption:=SearchOption.AllDirectories)
            Else
                ZTSFiles =
                   Directory.GetFiles(
                       path:=ZTSFilePath,
                       searchPattern:="*.zts",
                       searchOption:=SearchOption.TopDirectoryOnly)

            End If



            If ZTSFiles.Count = 0 Then
                add2Log(entry:=
                   "No *.zts files in path " & vbCrLf &
                    ZTSFilePath)

                writeUsage()

                Process.Start(fileName:=logFileName)
                End

            Else

                ZTSfiles2go.AddRange(ZTSFiles)

                add2Log(
                    entry:=(argPath).PadLeft(logLen) & ZTSFilePath)

                If Not recursive Then

                    add2Log(
                        entry:=("ZTS files:=").PadLeft(logLen))


                    For Each ZTSFile As String In ZTSfiles2go

                        add2Log(entry:=(" :=").PadLeft(logLen) &
                                Path.GetFileName(ZTSFile))

                    Next

                Else

                    add2Log(
                        entry:=("ZTS files:=").PadLeft(logLen) & ZTSfiles2go.Count.ToString("000"))

                End If


            End If

        Catch ex As Exception

            add2Log(entry:=
                   "Error getting *.zts files from path  " & vbCrLf &
                    ZTSFilePath & vbCrLf &
                    ex.Message)

            writeUsage()

            Process.Start(fileName:=logFileName)
            End

        End Try

    End Sub

    Private Sub getRecursive()

        Dim tempFilter As String() = {}
        Dim std As Boolean = True

        tempFilter =
            Filter(
                Source:=arguments.ToArray,
                Match:=argRecursive,
                Include:=True,
                Compare:=CompareMethod.Text)

        If tempFilter.Count = 1 Then

            Try

                recursive =
                    CBool(Replace(
                        Expression:=tempFilter.First,
                        Find:=argRecursive,
                        Replacement:="",
                        Compare:=CompareMethod.Text))

                std = True

            Catch ex As Exception

                add2Log(entry:=
                    "Error parsing cmd line for " & argRecursive & vbCrLf &
                     tempFilter.First & vbCrLf &
                     ex.Message)

                writeUsage()

                Process.Start(fileName:=logFileName)
                End

            End Try

        End If

        add2Log(
            entry:=((argRecursive).PadLeft(logLen) & recursive.ToString).PadRight(stdPos) &
                        IIf(std, " (std.)", " *** user def. ***").ToString)

    End Sub

#End Region

#Region "    get header info"

    Private Function getCMPNames(baseName As String,
                   Optional userCMPnames As String = Nothing) As Boolean

        Const Filename As String = "MASTER.FPJ"
        Const CMPSearchString As String = "  Chemical Name:"

        Const inpSearchString As String = "Chemical Input Data:"
        Dim index As Integer

        Dim MasterFPJ As String() = {}
        Dim inpFilePath As String
        Dim inpFile As String() = {}
        Dim compoundNames As String()

        If Not IsNothing(userCMPnames) Then

            Try
                compoundNames =
                    userCMPnames.Split(separator:={"/"c},
                                       options:=StringSplitOptions.RemoveEmptyEntries)
            Catch ex As Exception

                add2Log(entry:=
                            "Fatal Error parsing user given compound names" & vbCrLf &
                            userCMPnames & vbCrLf &
                            "Separator = '/'")

                Process.Start(fileName:=logFileName)
                End

            End Try

            Select Case compoundNames.Count

                Case 0

                    add2Log(entry:=
                            "Fatal Error parsing user given compound names" & vbCrLf &
                            userCMPnames & vbCrLf &
                            "Separator = '/'")

                    Process.Start(fileName:=logFileName)
                    End

                Case 1
                    Parent = compoundNames.First
                    add2Log(
                            entry:=("Parent:=").PadLeft(logLen) & Parent)

                Case 2
                    Parent = compoundNames.First
                    Met01 = compoundNames.Last

                    add2Log(
                            entry:=("Parent:=").PadLeft(logLen) & Parent)
                    add2Log(
                            entry:=("Met 01:=").PadLeft(logLen) & Met01)

                Case 3

                    Parent = compoundNames(0)
                    Met01 = compoundNames(1)
                    Met02 = compoundNames(2)

                    add2Log(
                            entry:=("Parent:=").PadLeft(logLen) & Parent)
                    add2Log(
                            entry:=("Met 01:=").PadLeft(logLen) & Met01)
                    add2Log(
                            entry:=("Met 02:=").PadLeft(logLen) & Met02)

                Case Else

                    add2Log(entry:=
                            "Fatal Error parsing user given compound names" & vbCrLf &
                            userCMPnames & vbCrLf &
                            "Separator = '/'")

                    Process.Start(fileName:=logFileName)

                    End

            End Select

            Return False

        End If

        Try

            'MASTER.FPJ or *.inp
            If File.Exists(
                    Path.Combine(
                        path1:=PRZMRunDir,
                        path2:=Filename)) Then

                add2Log(
                        entry:=("Reading:=").PadLeft(logLen) &
                        Filename & " to get compound name(s)")

                Try
                    MasterFPJ =
                        File.ReadAllLines(
                            path:=Path.Combine(PRZMRunDir,
                                                Filename))
                Catch ex As Exception

                    add2Log(entry:=
                            "Fatal Error reading MASTER.FPJ file to parse for compound names" & vbCrLf &
                            "MASTER.FPJ ?: " & Path.Combine(
                                PRZMRunDir,
                                Filename) & vbCrLf &
                                ex.Message)

                    Process.Start(fileName:=logFileName)
                    End

                End Try


                MasterFPJ =
                    Filter(
                        Source:=MasterFPJ,
                        Match:=CMPSearchString,
                        Include:=True,
                        Compare:=CompareMethod.Text)

                MasterFPJ(0) =
                    Replace(
                        Expression:=MasterFPJ(0),
                        Find:=CMPSearchString,
                        Replacement:="",
                        Compare:=CompareMethod.Text)

                MasterFPJ =
                    MasterFPJ(0).Split(
                        separator:={" "c},
                        options:=StringSplitOptions.RemoveEmptyEntries)

                Select Case MasterFPJ.Count

                    Case 0

                        add2Log(entry:=
                            "Fatal Error reading MASTER.FPJ file to parse for compound names" & vbCrLf &
                            "no parent defined" & vbCrLf &
                            Join(SourceArray:=MasterFPJ, Delimiter:=vbCrLf) & vbCrLf &
                            "MASTER.FPJ : " & Path.Combine(
                                PRZMRunDir,
                                Filename))

                        Process.Start(fileName:=logFileName)
                        End

                    Case 1
                        Parent = MasterFPJ.First
                        add2Log(
                            entry:=("Parent:=").PadLeft(logLen) & Parent)
                    Case 2
                        Parent = MasterFPJ.First
                        Met01 = MasterFPJ.Last

                        add2Log(
                            entry:=("Parent:=").PadLeft(logLen) & Parent)
                        add2Log(
                            entry:=("Met 01:=").PadLeft(logLen) & Met01)

                    Case 3

                        Parent = MasterFPJ(0)
                        Met01 = MasterFPJ(1)
                        Met02 = MasterFPJ(2)

                        add2Log(
                            entry:=("Parent:=").PadLeft(logLen) & Parent)
                        add2Log(
                            entry:=("Met 01:=").PadLeft(logLen) & Met01)
                        add2Log(
                            entry:=("Met 02:=").PadLeft(logLen) & Met02)

                    Case Else
                        add2Log(entry:=
                            "Fatal Error reading MASTER.FPJ file to parse for compound names" & vbCrLf &
                            "more than 2 metabolites defined" & vbCrLf &
                            Join(SourceArray:=MasterFPJ, Delimiter:=vbCrLf) & vbCrLf &
                            "MASTER.FPJ : " & Path.Combine(
                                PRZMRunDir,
                                Filename))

                        Process.Start(fileName:=logFileName)
                        End

                End Select

            Else

                inpFilePath = Path.Combine(
                        path1:=PRZMRunDir,
                        path2:=baseName & ".inp")

                If Not File.Exists(inpFilePath) Then


                    Parent = "Par"
                    Met01 = "Met01"
                    Met02 = "Met02"

                    add2Log(
                        entry:=("Parent:=").PadLeft(logLen) & Parent)
                    add2Log(
                        entry:=("Met 01:=").PadLeft(logLen) & Met01)
                    add2Log(
                        entry:=("Met 02:=").PadLeft(logLen) & Met02)

                    add2Log(entry:=(""))

                    Return False

                End If

                add2Log(
                        entry:=("Reading:=").PadLeft(logLen) &
                        baseName & ".inp to get compound name(s)")

                Try
                    inpFile = File.ReadAllLines(inpFilePath)
                Catch ex As Exception

                    add2Log(entry:=
                            "Fatal Error reading *.inp file to parse for compound names" & vbCrLf &
                            "*.inp      ?: " & Path.Combine(
                                PRZMRunDir,
                                baseName & ".inp ?") & vbCrLf &
                                ex.Message)

                    Process.Start(fileName:=logFileName)
                    End

                End Try

                index = Array.FindIndex(
                array:=inpFile,
                match:=Function(x)
                           Return x.Contains(inpSearchString)
                       End Function)

                If index = -1 Then
                    'error
                End If

                compoundNames =
                    inpFile(index + 2).Split(
                                separator:={" "c},
                                options:=StringSplitOptions.RemoveEmptyEntries)

                Select Case compoundNames.Count

                    Case 0

                        add2Log(entry:=
                            "Fatal Error reading *.inp file to parse for compound names" & vbCrLf &
                            "no parent defined" & vbCrLf &
                            Join(SourceArray:=compoundNames, Delimiter:=vbCrLf) & vbCrLf &
                            "inp file : " & inpFilePath)

                        Process.Start(fileName:=logFileName)
                        End

                    Case 1
                        Parent = compoundNames.First
                        add2Log(
                            entry:=("Parent:=").PadLeft(logLen) & Parent)

                    Case 2
                        Parent = compoundNames.First
                        Met01 = compoundNames.Last

                        add2Log(
                            entry:=("Parent:=").PadLeft(logLen) & Parent)
                        add2Log(
                            entry:=("Met 01:=").PadLeft(logLen) & Met01)

                    Case 3

                        Parent = compoundNames(0)
                        Met01 = compoundNames(1)
                        Met02 = compoundNames(2)

                        add2Log(
                            entry:=("Parent:=").PadLeft(logLen) & Parent)
                        add2Log(
                            entry:=("Met 01:=").PadLeft(logLen) & Met01)
                        add2Log(
                            entry:=("Met 02:=").PadLeft(logLen) & Met02)

                    Case Else

                        add2Log(entry:=
                            "Fatal Error reading *.inp file to parse for compound names" & vbCrLf &
                            "more than 2 metabolites defined" & vbCrLf &
                            Join(SourceArray:=compoundNames, Delimiter:=vbCrLf) & vbCrLf &
                            "inp file : " & inpFilePath)

                        Process.Start(fileName:=logFileName)
                        End

                End Select

            End If

        Catch ex As Exception

            add2Log(entry:=
                       "Fatal Error reading MASTER.FPJ/ *.inp file to parse for compound names" & vbCrLf &
                        Path.Combine(
                            PRZMRunDir,
                            Filename) & vbCrLf &
                        ex.Message)

            Process.Start(fileName:=logFileName)
            End

        End Try

        Return True

    End Function

    Private Function getSWASHNos() As Boolean

        Const Filename As String = "przm.pzm"
        Dim przmpzm As String() = {}
        Dim tempString As String() = {}
        Dim found As Boolean = False

        Try


            If Not File.Exists(
                        path:=Path.Combine(
                            PRZMRunDir,
                            Filename)) Then

                SWASHNos.AddRange(
                    {
                    -99, -99, -99, -99
                    })

                add2Log(entry:=("SWASH #:=").PadLeft(logLen) & " no " & Filename & " found! Use *.inp file name")

                Return False

            End If

            SWASHNos.Clear()

            add2Log(
                entry:=("Reading:=").PadLeft(logLen) &
                Filename & "   to get SWASH #")

            przmpzm =
                File.ReadAllLines(
                    path:=Path.Combine(PRZMRunDir,
                                       Filename))


            add2Log(entry:=("SWASH #:=").PadLeft(logLen))

            For counter As Integer = 1 To 4

                tempString =
                    Filter(
                        Source:=przmpzm,
                        Match:="R" & counter.ToString & "=",
                        Include:=True,
                        Compare:=CompareMethod.Text)

                If tempString.First.Length = 3 Then
                    SWASHNos.Add(-1)
                    add2Log(entry:=("R" & counter.ToString & ":=").PadLeft(logLen) & "-")
                Else
                    SWASHNos.Add(
                        CInt(
                            tempString.First.Split(
                                separator:={"="c},
                                options:=StringSplitOptions.RemoveEmptyEntries).Last))
                    found = True

                    add2Log(entry:=("R" & counter.ToString & ":=").PadLeft(logLen) & SWASHNos.Last)

                End If

            Next

            If Not found Then
                add2Log(entry:=
                    "Fatal Error parsing przm.pzm file for swash run numbers " & vbCrLf &
                    Path.Combine(PRZMRunDir, Filename))

                Process.Start(fileName:=logFileName)
                End

            End If

        Catch ex As Exception

            add2Log(entry:=
                      "Fatal Error parsing przm.pzm file for swash run numbers " & vbCrLf &
                       Path.Combine(PRZMRunDir, Filename) & vbCrLf &
                       ex.Message)

            Process.Start(fileName:=logFileName)
            End

        End Try

        Return True

    End Function

    Private Sub getTPAPpos(ZTSHeaderRow As String)

        Dim tempArray As String() = {}

        posTPAP = -1

        tempArray =
            ZTSHeaderRow.Split(
                separator:={" "c},
                options:=StringSplitOptions.RemoveEmptyEntries)


        For counter As Integer = 0 To tempArray.Count - 1
            If tempArray(counter) = "TPAP" Then
                posTPAP = counter + eZTSHeader.EventDay
                Exit For
            End If

        Next

        If posTPAP = -1 Then
            add2Log(
                entry:="No 'TPAP' found in header, can't get appln. info!")

            Process.Start(fileName:=logFileName)
            End

        End If

    End Sub

    Private Sub getIRRGpos(ZTSHeaderRow As String)

        Dim tempArray As String() = {}

        posIRRG = -1

        tempArray =
            ZTSHeaderRow.Split(
                separator:={" "c},
                options:=StringSplitOptions.RemoveEmptyEntries)

        For counter As Integer = 0 To tempArray.Count - 1
            If tempArray(counter) = "IRRG" Then
                posIRRG = counter + eZTSHeader.EventDay
                Exit For
            End If

        Next

    End Sub

#End Region

    Private Enum eParMet
        Par
        Met01
        Met02
    End Enum

    Private Function createHeader(ParMet As eParMet, ZTSFileName As String) As Boolean

        Dim out As New List(Of String)
        Dim LeadingString As String = "*  "
        Dim ApplnsPerSeason As Integer = -1


        If ParMet = eParMet.Par Then _
            Scenario = getPRZMScenarioFromFilename(ZTSFileName:=ZTSFileName)

        With out

            .AddRange(getApplnInfo)

            'io info
            .Add(LeadingString & "Working directory   : " & Path.GetDirectoryName(ZTSFileName))
            .Add(LeadingString & "PRZM file (*.ZTS)   : " & Path.GetFileName(ZTSFileName))

            'sim info
            .Add(LeadingString & "Crop                : " & Crop)
            .Add(LeadingString & "Scenario            : " & Scenario)

            If description = String.Empty Then
                .Add(LeadingString & "Tier                : " & "Step 03")
            Else
                .Add(LeadingString & "Description         : " & description)
            End If

            If seasonOnly Then
                .Add(LeadingString & "Season start        : " & seasonStart.ToString("dd-MMM-yyyy"))

                .Add(LeadingString & "       end          : " & seasonEnd.ToString("dd-MMM-yyyy"))
                .Add(LeadingString & "Applns per season   : " & applnsSeason.Count.ToString("00"))

            Else
                .Add(LeadingString & "Sim start           : " & SimStart.ToString("dd-MMM-yyyy") &
                                   (IIf(WarmUp <> 0, ", Warm up " & WarmUp & " years", "")).ToString)

                .Add(LeadingString & "      end           : " & SimEnd.ToString("dd-MMM-yyyy"))

                ApplnsPerSeason = CInt(Math.Round(applns.Count / (SimEnd.Year - SimStart.Year + 1), digits:=0))
                .Add(LeadingString & "Applns per season   : " & ApplnsPerSeason.ToString("00"))
            End If

            'parent or met01/met02, must start with '*  Chemical:  ' to fit with txw info
            If ParMet = eParMet.Par Then

                .Add(LeadingString & "Chemical:  " & Parent)
                .Add(LeadingString & "Parent run")

                add2Log(entry:=("Scenario:=").PadLeft(logLen) & Scenario)
                If SWASHno <> -99 Then
                    add2Log(entry:=("SWASH #:=").PadLeft(logLen) & SWASHno.ToString)
                End If

            ElseIf ParMet = eParMet.Met01 Then

                .Add(LeadingString & "Chemical:  " & Met01)
                .Add(LeadingString & "Metabolite run")

            ElseIf ParMet = eParMet.Met02 Then

                .Add(LeadingString & "Chemical:  " & Met02)
                .Add(LeadingString & "2nd Metabolite run!")

            End If

            .Add(LeadingString)
            .Add(LeadingString & "Total number of applns ")

            If seasonOnly Then
                .Add("#    " & applnsSeason.Count)
            Else
                .Add("#    " & applns.Count)
            End If

            .Add(LeadingString)

            .Add(LeadingString)

            'appln info
            .Add(LeadingString & "Number of      Time                       Mass ")
            .Add(LeadingString & "Application    dd-MMM-YYYY-hh:mm          (g ai/ha)")

            If seasonOnly Then
                .AddRange(applnsSeason)
            Else
                .AddRange(applns)
            End If

            .Add(LeadingString)

            .Add("*                       Runoff         Runoff         Erosion        Erosion        Infiltration")
            .Add(
                 "* Time                  Volume         flux           Mass           Flux           " &
                 IIf(
                     Expression:=monthlyAverage,
                     TruePart:="Monthly Average",
                     FalsePart:="Storage/Discharge").ToString)

            .Add(
                 "* dd-MMM-YYYY-hh:mm     (mm/h)         (mg as/m2/h)   (kg/h)         (mg as/m2/h)   (mm/h)")

        End With

        p2tHeader = out

        Return True

    End Function

    Private Function createP2T(ZTSFileName As String) As Boolean

#Region "    Definitions"

        Dim ZTSFile As String() = {}
        Dim headerRow As String = ""
        Dim tempString As String = String.Empty
        Dim tempArray As String() = {}

        Dim EventDate As New Date
        Dim EventDuration As Integer = 0

        Dim RUNF As Double = Double.NaN
        Dim PRCP As Double = Double.NaN

        Dim RFLX1 As Double = Double.NaN
        Dim RFLX2 As Double = Double.NaN
        Dim RFLX3 As Double = Double.NaN

        Dim ESLS As Double = Double.NaN

        Dim EFLX1 As Double = Double.NaN
        Dim EFLX2 As Double = Double.NaN
        Dim EFLX3 As Double = Double.NaN

        Dim INFL As Double = Double.NaN
        Dim IRRG As Double = Double.NaN
        Dim TPAP As Double = Double.NaN

        Dim GW_discharge As Double = 0
        Dim GW_storage As Double = 0
        Dim infiltration As Double = 0
        Dim Timestep As Integer = 1

        Dim monthlyAverageINFL As Double = Double.NaN
        Dim oldDate As Date = Nothing

#End Region

        'get zts file
        Try

            ZTSFile = File.ReadAllLines(ZTSFileName)

            'filter comments
            ZTSFile =
                Filter(Source:=ZTSFile,
                       Match:="* ",
                       Include:=False,
                       Compare:=CompareMethod.Text)

            headerRow = ZTSFile(hearderRowNo)

            If Not ZTSFile(ZTSHeaderRowNo).Contains("RFLX3") Then
                Met02 = String.Empty
            End If
            If Not ZTSFile(ZTSHeaderRowNo).Contains("RFLX2") Then
                Met01 = String.Empty
            End If

            If Met01 = String.Empty Then
                add2Log(
                    entry:=("Parent run:=").PadLeft(logLen) & Parent)
            End If

            'check for valid zts file
            If ZTSFile.Count < 400 Then

                add2Log(entry:=("IO Error:=").PadLeft(logLen) &
                                "IO Error reading ZTS file, file empty or broken" & vbCrLf &
                                ZTSFileName)

                Process.Start(fileName:=logFileName)
                End

            End If

        Catch ex As Exception

            add2Log(entry:=("IO Error:=").PadLeft(logLen) &
                       "IO Error reading ZTS file" & vbCrLf &
                       ZTSFileName & vbCrLf &
                       ex.Message)

            Process.Start(fileName:=logFileName)
            End

        End Try

        getTPAPpos(ZTSHeaderRow:=ZTSFile(ZTSHeaderRowNo))

        'no 'TPAP found : can't continue!
        If posTPAP = -1 Then

            add2Log(entry:="Error reading ZTS file" & vbCrLf &
                       "No appln. date info 'TPAP' found")
            Process.Start(fileName:=logFileName)
            End

        End If

        getIRRGpos(ZTSHeaderRow:=ZTSFile(ZTSHeaderRowNo))

        'no 'IRRG found : set irrigation to 0
        If posIRRG = -1 Then
            add2Log(" ".PadLeft(logLen) & "No 'IRRG' info found, set to 0!")
            IRRG = 0
        End If

        If seasonOnly Then

            seasonStart = getSeasonStart(ZTSfile:=ZTSFile)

            If seasonStart = New Date Then
                seasonOnly = False
                add2Log("Can't get PRZM Season! Switch to 'seasonOnly = False' ")
            Else

                seasonEnd = seasonStart.AddYears(1).AddDays(-1)
                add2Log(
                    "PRZM Season:=".PadLeft(logLen) &
                    seasonStart.ToLongDateString & " to " &
                    seasonEnd.ToLongDateString)

            End If

        End If

        For RowCounter As Integer = ZTSDataStartRowNo To ZTSFile.Count - 1

            tempArray =
                ZTSFile(RowCounter).Split(
                    separator:={" "c},
                    options:=StringSplitOptions.RemoveEmptyEntries)

            'get event date
            Try
                EventDate =
                    New Date(
                        year:=CInt(tempArray(eZTSHeader.EventYear)),
                        month:=CInt(tempArray(eZTSHeader.EventMonth)),
                        day:=CInt(tempArray(eZTSHeader.EventDay)))
            Catch ex As Exception

                add2Log(entry:=("ParseError:=").PadLeft(logLen) &
                        "Can't parse date from" & vbCrLf &
                        ZTSFile(RowCounter) & vbCrLf &
                        ex.Message)

                Process.Start(fileName:=logFileName)
                End

            End Try



            'get sim start / end
            If RowCounter = ZTSDataStartRowNo Then
                SimStart = EventDate
            ElseIf RowCounter = ZTSFile.Count - 1 Then
                SimEnd = EventDate
            End If

            'basic & parent
            Try

                RUNF = CDbl(tempArray(eZTSHeader.RUNF))
                ESLS = CDbl(tempArray(eZTSHeader.ESLS))
                PRCP = CDbl(tempArray(eZTSHeader.PRCP))
                INFL = CDbl(tempArray(eZTSHeader.INFL))

                RFLX1 = CDbl(tempArray(eZTSHeader.RFLX1))
                EFLX1 = CDbl(tempArray(eZTSHeader.EFLX1))

                TPAP = CDbl(tempArray(posTPAP))

                If posIRRG <> -1 Then

                    IRRG = CDbl(tempArray(posIRRG))

                    If IRRG <> 0 Then
                        irrigation.Add(
                            EventDate.ToShortDateString.PadRight("00.00.0000  ".Length) & IRRG & "mm")
                    End If

                Else
                    IRRG = 0
                End If

            Catch ex As Exception

                add2Log(entry:=("ParseError:=").PadLeft(logLen) &
                       "Can't parse data from" & vbCrLf &
                       ZTSFile(RowCounter) & vbCrLf &
                       ex.Message)

                Process.Start(fileName:=logFileName)
                Return False

            End Try

            'add appln
            If TPAP <> 0 Then

                addAppln2List(
                    Eventdate:=EventDate,
                    TPAP:=TPAP)

            End If

            'Met01
            If Met01 <> String.Empty Then

                Try

                    RFLX2 = CDbl(tempArray(eZTSHeader.RFLX2))
                    EFLX2 = CDbl(tempArray(eZTSHeader.EFLX2))

                Catch ex As Exception

                    add2Log(entry:=("ParseError:=").PadLeft(logLen) &
                      "Can't parse Met01 data from" & vbCrLf &
                      ZTSFile(RowCounter) & vbCrLf &
                      ex.Message)

                    Process.Start(fileName:=logFileName)
                    End

                End Try

            End If

            'Met02
            If Met02 <> String.Empty Then

                Try

                    RFLX3 = CDbl(tempArray(eZTSHeader.RFLX3))
                    EFLX3 = CDbl(tempArray(eZTSHeader.EFLX3))

                Catch ex As Exception

                    add2Log(entry:=("ParseError:=").PadLeft(logLen) &
                      "Can't parse Met01 data from" & vbCrLf &
                      ZTSFile(RowCounter) & vbCrLf &
                      ex.Message)

                    Process.Start(fileName:=logFileName)
                    End

                End Try

            End If


#Region "    event duration, snow melt and heavy rain"

            EventDuration =
                getEventDuration(
                        PRCP:=PRCP,
                        IRRG:=IRRG,
                        RUNF:=RUNF)

            If EventDuration = 12 AndAlso PRCP = 0 AndAlso reportSnowMelt Then

                add2Log(entry:=("Snow melt:=").PadLeft(logLen) &
                      ("at " &
                      EventDate.ToString("dd-MMM-yyyy")).PadLeft("           dd-MMM-yyyy".Length))

            ElseIf PRCP >= HeavyRainLimit Then

                HeavyRain.Add(
                    (" ").PadLeft(logLen) &
                     PRCP.ToString("0.0").PadLeft(5) & "mm at " &
                    EventDate.ToString("dd-MMM-yyyy"))

                If PRCP >= 2 * HeavyRainLimit Then
                    twiceHeavyRain += 1
                End If

            End If

            If PRCP >= maxRain Then maxRain = PRCP

#End Region

#Region "    discharge and storage"

            Try

                'MRT  = Mean residence time in days, std. = 20d
                If Exp Then

                    'GW_discharge  calc. with exponential discharge formula
                    ' Q2 = Q1 exp { −A (T2 − T1) } + R [ 1 − exp { −A (T2 − T1) } ]
                    'https://en.wikipedia.org/wiki/Runoff_model_(reservoir)

                    GW_discharge =
                        GW_discharge * Math.Exp(-(1 / MRT) * Timestep) +
                        INFL * (1 - Math.Exp(-(1 / MRT) * Timestep))

                Else

                    'GW_discharge calc. with Stella from nick jarvis          
                    GW_discharge = (1 / MRT) * GW_storage

                End If

                GW_storage += ((INFL - GW_discharge) * Timestep)

            Catch ex As Exception

                add2Log(
                    entry:="Error calc. GW Discharge and Storage at" &
                    EventDate.ToLongDateString & vbCrLf & ex.Message)

                Process.Start(fileName:=logFileName)
                End

            End Try

#End Region

            'check for warm-up years
            If SimStart.AddYears(WarmUp) > EventDate AndAlso
               Not seasonOnly Then Continue For

            If Not seasonOnly OrElse
                    (EventDate >= seasonStart AndAlso
                    EventDate <= seasonEnd) Then

                'calc monthly average INFL
                If monthlyAverage Then

                    If oldDate = New Date OrElse
                   oldDate.Month <> EventDate.Month Then

                        oldDate = EventDate

                        monthlyAverageINFL =
                            getMonthlyAverageINFL(
                            eventDate:=EventDate,
                            ZTSfile:=ZTSFile)

                    End If

                End If

                ' ... and use it if selected
                If monthlyAverage Then
                    infiltration = monthlyAverageINFL
                Else
                    infiltration = GW_discharge
                End If

                p2tDataParent.AddRange(
            createP2TDay(
                EventDate:=EventDate,
                EventDuration:=EventDuration,
                RUNF:=RUNF,
                EFLX:=EFLX1,
                ESLS:=ESLS,
                RFLX:=RFLX1,
                IRRG:=IRRG,
                infiltration:=infiltration))

                If Met01 <> String.Empty Then
                    p2tDataMet01.AddRange(
                createP2TDay(
                    EventDate:=EventDate,
                    EventDuration:=EventDuration,
                    RUNF:=RUNF,
                    EFLX:=EFLX2,
                    ESLS:=ESLS,
                    RFLX:=RFLX2,
                    IRRG:=IRRG,
                    infiltration:=infiltration))

                End If

                If Met02 <> String.Empty Then
                    p2tDataMet02.AddRange(
                createP2TDay(
                    EventDate:=EventDate,
                    EventDuration:=EventDuration,
                    RUNF:=RUNF,
                    EFLX:=EFLX3,
                    ESLS:=ESLS,
                    RFLX:=RFLX3,
                    IRRG:=IRRG,
                    infiltration:=infiltration))

                End If

            End If

        Next

        Return True

    End Function


    Private Function getMonthlyAverageINFL(eventDate As Date,
                                           ZTSfile As String()) As Double

        Const firstPart As String = " TSR PRZ  "
        Dim searchString As String = ""
        Dim targetMonth As String() = {}
        Dim sumOfINFL As Double = 0
        Dim temp As String

        ' TSR PRZ  1975  1  1
        ' TSR PRZ  1975  1 10
        ' TSR PRZ  1975 10  1
        ' TSR PRZ  1975 10 10
        With eventDate
            searchString =
                firstPart &
                .Year.ToString &
                .Month.ToString.PadLeft("  1".Length)
        End With

        targetMonth =
            Filter(
                Source:=ZTSfile,
                Match:=searchString,
                Include:=True,
                Compare:=CompareMethod.Text)

        For Each row As String In targetMonth

            temp = row.Split(
                    separator:={" "c},
                    options:=StringSplitOptions.RemoveEmptyEntries)(eZTSHeader.INFL)

            sumOfINFL +=
               Double.Parse(s:=temp)

        Next

        Return sumOfINFL / targetMonth.Count

    End Function

#Region "    create p2t day"

    Private Function createP2TDay(
                    EventDate As Date,
                    EventDuration As Integer,
                    RUNF As Double,
                    RFLX As Double,
                    ESLS As Double,
                    EFLX As Double,
                    IRRG As Double,
                    infiltration As Double) As String()

        Dim out As New List(Of String)
        Dim tempString As String = String.Empty

        Try

            For HourCounter As Integer = 1 To 24

                If HourCounter <= EventDuration Then

                    tempString =
                        (EventDate.ToString("dd-MMM-yyyy") &
                            "-" & HourCounter.ToString("00") & ":00").PadLeft("  01-Jan-1978-01:00".Length) &
                    ("0" & ((RUNF + IRRG) / EventDuration).ToString(".0000E+00")).PadLeft("     0.0000E+00".Length) &
                    ("0" & (RFLX / EventDuration).ToString(".0000E+00")).PadLeft("     0.0000E+00".Length) &
                    ("0" & (ESLS / EventDuration).ToString(".0000E+00")).PadLeft("     0.0000E+00".Length) &
                    ("0" & (EFLX / EventDuration).ToString(".0000E+00")).PadLeft("     0.0000E+00".Length) &
                    ("0" & (infiltration / 24).ToString(".0000E+00")).PadLeft("     0.0000E+00".Length)

                    out.Add(tempString)

                Else

                    out.Add((EventDate.ToString("dd-MMM-yyyy") &
                            "-" & HourCounter.ToString("00") & ":00").PadLeft("  01-Jan-1978-01:00".Length) &
                          "     0.0000E+00     0.0000E+00     0.0000E+00     0.0000E+00" &
                    ("0" & (infiltration / 24).ToString(".0000E+00")).PadLeft("     0.0000E+00".Length))

                End If

            Next

        Catch ex As Exception

            add2Log(
                entry:="Error creating p2t day" &
                "Date : " & EventDate.ToLongDateString & vbCrLf &
                "RUNF : " & RUNF & vbCrLf &
                "RFLX : " & RFLX & vbCrLf &
                "ESLS : " & ESLS & vbCrLf &
                "EFLX : " & EFLX & vbCrLf &
                "IRRG : " & IRRG & vbCrLf &
                "infiltration   : " & infiltration & vbCrLf &
                "Event Duration : " & EventDuration & vbCrLf &
                ex.Message)

            Process.Start(fileName:=logFileName)
            End

        End Try

        Return out.ToArray

    End Function

    Private Function getEventDuration(
                        PRCP As Double,
                        IRRG As Double,
                        RUNF As Double) As Integer

        Try

            If (PRCP + IRRG) / MaxPRECperHour >= 24 Then
                Return 24
            ElseIf RUNF <> 0 AndAlso PRCP = 0 AndAlso IRRG = 0 Then
                'snow melt => event duration =12h
                Return 12
            Else
                Return CInt(Math.Round(
                    (PRCP + IRRG) / MaxPRECperHour,
                    digits:=0,
                    mode:=MidpointRounding.AwayFromZero))
            End If

        Catch ex As Exception

            add2Log(
               entry:="Error calc. event duration" & vbCrLf &
               "PRCP : " & PRCP & vbCrLf &
               "IRRG : " & IRRG & vbCrLf &
               "RUNF : " & RUNF & vbCrLf &
               "Max PREC per hour : " & MaxPRECperHour & vbCrLf &
                ex.Message)

            Process.Start(fileName:=logFileName)
            End

        End Try


    End Function

#End Region

    Private Function getSeasonStart(ZTSfile As String()) As Date

        Dim temp As String() = {}
        Dim firstApplnDate As Date

        For rowCounter As Integer = ZTSDataStartRowNo To ZTSDataStartRowNo + 365

            temp =
                ZTSfile(rowCounter).Split(
                            separator:={" "c},
                            options:=StringSplitOptions.RemoveEmptyEntries)

            If CDbl(temp(posTPAP)) <> 0 Then

                firstApplnDate =
                    New Date(
                        year:=CInt(temp(eZTSHeader.EventYear)),
                        month:=CInt(temp(eZTSHeader.EventMonth)),
                        day:=CInt(temp(eZTSHeader.EventDay)))

                Return getPRZMSeasonStart(
                            PRZMswScenario:=Scenario.Substring(0, 2),
                            ApplnMonth:=firstApplnDate.Month)

            End If

        Next

        Return New Date

    End Function

    Private Sub addAppln2List(
                    Eventdate As Date,
                    TPAP As Double)

        Dim tempString As String = String.Empty

        Try
            tempString = "#  " & (applns.Count + 1).ToString("00").PadRight("01             ".Length) &
                        (Eventdate.ToString("dd-MMM-yyyy") & "-09:00").PadRight("01-Jan-1975-09:00          ".Length) &
                        (TPAP * 100000000).ToString("0.00")

            applns.Add(tempString)
        Catch ex As Exception

            add2Log(
                entry:="Error adding appln to list" &
                "Date : " & Eventdate.ToLongDateString & vbCrLf &
                "TPAP : " & TPAP & vbCrLf &
                ex.Message)

            Process.Start(fileName:=logFileName)
            End

        End Try


        If seasonOnly AndAlso
                    Eventdate >= seasonStart AndAlso
                    Eventdate <= seasonEnd Then

            Try
                tempString = "#  " & (applnsSeason.Count + 1).ToString("00").PadRight("01             ".Length) &
                        (Eventdate.ToString("dd-MMM-yyyy") & "-09:00").PadRight("01-Jan-1975-09:00          ".Length) &
                        (TPAP * 100000000).ToString("0.00")

                applnsSeason.Add(tempString)
            Catch ex As Exception

                add2Log(
                entry:="Error adding appln to SEASON list" &
                "Date : " & Eventdate.ToLongDateString & vbCrLf &
                "TPAP : " & TPAP & vbCrLf &
                ex.Message)

                Process.Start(fileName:=logFileName)
                End

            End Try

        End If

    End Sub

#Region "    get crop and scenario name from ZTS file name"

    Private Function getPRZMScenarioFromFilename(ZTSFileName As String) As String

        Dim Name2Parse As String = ""
        Dim Out As String = String.Empty

        Try

            Name2Parse = Path.GetFileName(path:=ZTSFileName)
            Name2Parse = Name2Parse.Substring(0, 2)

            Select Case Name2Parse.ToUpper

                Case "R1"
                    SWASHno = SWASHNos(0)
                    Out = "R1, Weiherbach pond/stream"

                Case "R2"
                    SWASHno = SWASHNos(1)
                    Out = "R2, Porto stream"

                Case "R3"
                    SWASHno = SWASHNos(2)
                    Out = "R3, Bologna stream"

                Case "R4"
                    SWASHno = SWASHNos(3)
                    Out = "R4, Roujan stream"

                Case Else
                    SWASHno = -1
                    Out = " ? "
            End Select

            If SWASHno = -1 Then

                If Out <> String.Empty Then
                    add2Log(entry:="Unknown scenario name or can't parse" & vbCrLf &
                            ZTSFileName)
                Else
                    add2Log(entry:="No SWASH number for this scenario" & vbCrLf &
                            Out)
                End If

                Process.Start(fileName:=logFileName)
                End

            End If

            Return Out

        Catch ex As Exception

            add2Log(entry:="Can't parse FOCUSsw runoff scenario from filename" & vbCrLf &
                             ZTSFileName & vbCrLf & ex.Message)

            Process.Start(fileName:=logFileName)
            End

        End Try


    End Function

    Private Function getPRZMCropFromFileName(
                        ZTSFileName As String) As String

        Dim Name2Parse As String = ""
        Dim out As String = String.Empty

        Try

            Name2Parse = Path.GetFileName(path:=ZTSFileName)
            Name2Parse = Name2Parse.Substring(3, 2)

            For CropCounter = 0 To sPRZMswCropShort.Count - 1

                If sPRZMswCropShort(CropCounter) = Name2Parse Then


                    Select Case Name2Parse

                        Case "VB", "VL", "VR"

                            If ZTSFileName.Contains("_2nd") Then
                                out = sPRZMswCropLong(CropCounter) & " 2nd season"
                            Else
                                out = sPRZMswCropLong(CropCounter) & " 1st season"
                            End If

                        Case Else

                            out = sPRZMswCropLong(CropCounter)

                    End Select

                End If

            Next

            If out = String.Empty Then

                add2Log(entry:="Can't parse FOCUSsw runoff crop from filename" & vbCrLf &
                        ZTSFileName)

                Process.Start(fileName:=logFileName)
                End

            End If

            Return out

        Catch ex As Exception

            add2Log(entry:="Can't parse FOCUSsw runoff crop from filename" & vbCrLf &
                               ZTSFileName & vbCrLf &
                               ex.Message)

            Process.Start(fileName:=logFileName)
            End

        End Try

    End Function


    ''' <summary>
    ''' 2 Letter shortcode for PRZM crops
    ''' </summary>
    Private sPRZMswCropShort As String() =
    {
        "CS",
        "CW",
        "CI",
        "FB",
        "GA",
        "HP",
        "LG",
        "MZ",
        "OS",
        "OW",
        "OL",
        "PF",
        "PS",
        "SY",
        "SB",
        "SU",
        "TB",
        "VB",
        "VF",
        "VL",
        "VR",
        "VI"
    }

    ''' <summary>
    ''' Long PRZM crop names
    ''' </summary>
    Private sPRZMswCropLong As String() =
    {
        "Cereals, spring",
        "Cereals, winter",
        "Citrus",
        "Field beans",
        "Grass/alfalfa",
        "Hops",
        "Legumes",
        "Maize",
        "Oil seed rape, spring",
        "Oil seed rape, winter",
        "Olives",
        "Pome/stone fruits",
        "Potatoes",
        "Soybeans",
        "Sugar beets",
        "Sunflowers",
        "Tobacco",
        "Vegetables, bulb",
        "Vegetables, fruiting",
        "Vegetables, leafy",
        "Vegetables, root",
        "Vines"
    }

#End Region


    Public Function getPRZMSeasonStart(
                                        PRZMswScenario As String,
                                        ApplnMonth As Integer) As Date


        Select Case PRZMswScenario

            Case "R1"

                Select Case ApplnMonth

                    Case 3, 4, 5
                        Return New Date(year:=1984,
                                       month:=3,
                                         day:=1)

                    Case 6, 7, 8, 9
                        Return New Date(year:=1978,
                                       month:=6,
                                         day:=1)

                    Case Else
                        Return New Date(year:=1978,
                                       month:=10,
                                         day:=1)

                End Select

            Case "R2"

                Select Case ApplnMonth

                    Case 3, 4, 5
                        Return New Date(year:=1977,
                                       month:=3,
                                         day:=1)

                    Case 6, 7, 8, 9
                        Return New Date(year:=1989,
                                       month:=6,
                                         day:=1)

                    Case Else
                        Return New Date(year:=1977,
                                       month:=10,
                                         day:=1)

                End Select

            Case "R3"

                Select Case ApplnMonth

                    Case 3, 4, 5
                        Return New Date(year:=1980,
                                       month:=3,
                                         day:=1)

                    Case 6, 7, 8, 9
                        Return New Date(year:=1975,
                                       month:=6,
                                         day:=1)

                    Case Else
                        Return New Date(year:=1980,
                                       month:=10,
                                         day:=1)

                End Select

            Case "R4"

                Select Case ApplnMonth

                    Case 3, 4, 5
                        Return New Date(year:=1984,
                                       month:=3,
                                         day:=1)

                    Case 6, 7, 8, 9
                        Return New Date(year:=1985,
                                       month:=6,
                                         day:=1)

                    Case Else
                        Return New Date(year:=1979,
                                       month:=10,
                                         day:=1)

                End Select

            Case Else

                Return New Date

        End Select

    End Function


End Module
