
Imports System.IO

Module m_p2tmaker

    Private log As New List(Of String)
    Private logFileName As String = String.Empty
    Private HeavyRainLimit As Double = 50
    Private reportHeavyRain As Boolean = False
    Private reportSnowMelt As Boolean = False


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

        'get cmd line arguments
        getCMDArgs()

        add2Log("")
        add2Log("")
        add2Log(
            Join(
                SourceArray:=getApplnInfo(Leadingstring:=("   ")),
                Delimiter:=vbCrLf))

        ' and parse them
        getWarmUp()

        getMRT()

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
                If recursive Then

                    add2Log(
                       entry:=(" ").PadLeft(logLen) & " ****************************************************** ")

                    add2Log(
                     entry:=("Actual path:=").PadLeft(logLen) & Path.GetDirectoryName(path:=ZTSFileName))

                    add2Log(
                      entry:=(" ").PadLeft(logLen) & " ****************************************************** ")

                End If

                'get cmp names from master.fpj
                getCMPNames()

                'get SWASH numbers from przm.pzm
                getSWASHNos()

                'get crop from filename
                Try
                    Crop = getPRZMCropFromFileName(ZTSFileName:=ZTSfiles2go.First)
                    add2Log(entry:=("Crop:=").PadLeft(logLen) & Crop)
                Catch ex As Exception
                    add2Log(
                           entry:=("Parsing Error:=").PadLeft(logLen) &
                           "Can' tparse crop name from filename " & vbCrLf &
                           ZTSfiles2go.First & vbCrLf & ex.Message)
                End Try

            Else
                add2Log("")
            End If


            add2Log(
                entry:=(" ").PadLeft(logLen) & " ****************************************************** ")

            add2Log(
                entry:=("ZTS:=").PadLeft(logLen) & ZTSFileName)

            If Met01 <> String.Empty Then
                add2Log(
                    entry:=("Parent run:=").PadLeft(logLen) & Parent)
            End If

            'init
            applns.Clear()
            p2tHeader.Clear()
            p2tDataParent.Clear()
            p2tDataMet01.Clear()
            HeavyRain.Clear()
            out.Clear()

            'get data
            If Not createP2T(
                ZTSFileName:=ZTSFileName) Then

                Continue For

            End If

            createHeader(
                ParMet:=eParMet.Par,
                ZTSFileName:=ZTSFileName)

            out.AddRange(p2tHeader)
            out.AddRange(p2tDataParent)


            Try

                P2TFileName = Path.Combine(
                                PRZMRunDir,
                                SWASHno.ToString("00000") & "-C1.p2t")

                File.WriteAllLines(
                        path:=P2TFileName,
                        contents:=out.ToArray)

                add2Log(
                    entry:=(Parent & " p2t:=").PadLeft(logLen) & P2TFileName)

            Catch ex As Exception
                add2Log(
                    entry:=("IO Error:=").PadLeft(logLen) & ex.Message)
            End Try


            If Met01 <> String.Empty Then


                createHeader(
                    ParMet:=eParMet.Met,
                    ZTSFileName:=ZTSFileName)

                out.Clear()
                out.AddRange(p2tHeader)
                out.AddRange(p2tDataMet01)

                Try

                    P2TFileName = Path.Combine(
                                    PRZMRunDir,
                                    SWASHno.ToString("00000") & "-C2.p2t")

                    File.WriteAllLines(
                            path:=P2TFileName,
                            contents:=out.ToArray)

                    add2Log(
                        entry:=(Met01 & " p2t:=").PadLeft(logLen) & P2TFileName)


                Catch ex As Exception
                    add2Log(
                        entry:=("IO Error:=").PadLeft(logLen) & ex.Message)
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


        Next

    End Sub


#Region "    internal stuff"

    ' row where then data in the zts file starts
    Const ZTSDataStartRowNo As Integer = 3
    ' row where then header in the zts file starts
    Const ZTSHeaderRowNo As Integer = 2
    'format the log file
    Private logLen As Integer = 15
    Private stdPos As Integer = 25


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

    End Enum


    ''' <summary>
    ''' Position of the appln info in zts file
    ''' </summary>
    ''' <remarks></remarks>
    Private posTPAP As Integer = -1

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
        Console.WriteLine(Leadingstring)
        Console.WriteLine(Leadingstring)
        Console.WriteLine(
            Join(
                SourceArray:=getApplnInfo,
                Delimiter:=vbCrLf))
        Console.WriteLine(Leadingstring & "Usage:")
        Console.WriteLine(Leadingstring)
        Console.WriteLine(Leadingstring & "convert a single zts file")
        Console.WriteLine(Leadingstring & "zts:='full zts file path'")
        Console.WriteLine(Leadingstring & "OR")
        Console.WriteLine(Leadingstring & "convert all zts files in a project directory")
        Console.WriteLine(Leadingstring & "path:='full path to directory with zts files'")
        Console.WriteLine(Leadingstring & "-------------------------------------------------")
        Console.WriteLine(Leadingstring & "get ZTS files recursively")
        Console.WriteLine(Leadingstring & "recursive:=true/false")
        Console.WriteLine(Leadingstring & "-------------------------------------------------")
        Console.WriteLine(Leadingstring & "if 'p2tmaker.exe' is in the project directory")
        Console.WriteLine(Leadingstring & "it can be used without these two cmd line args")
        Console.WriteLine(Leadingstring & "-------------------------------------------------")
        Console.WriteLine(Leadingstring & "To skip warm up years, default = 0")
        Console.WriteLine(Leadingstring & "warmup:=0")
        Console.WriteLine(Leadingstring & "-------------------------------------------------")
        Console.WriteLine(Leadingstring & "To set the mean residence time in days, default = 20days")
        Console.WriteLine(Leadingstring & "mrt:=20")
        Console.WriteLine(Leadingstring & "-------------------------------------------------")
        Console.WriteLine(Leadingstring & "To set the max. precipitation per hour for")
        Console.WriteLine(Leadingstring & "calculation of event duration in mm, default = 2mm")
        Console.WriteLine(Leadingstring & "maxPREC:=2")
        Console.WriteLine(Leadingstring & "-------------------------------------------------")
        Console.WriteLine(Leadingstring & "GW_discharge calc. with exponential ")
        Console.WriteLine(Leadingstring & "discharge formula, std. = false")
        Console.WriteLine(Leadingstring & "exp:=true/false")
        Console.WriteLine(Leadingstring & "-------------------------------------------------")
        Console.WriteLine(Leadingstring)
        Console.WriteLine(Leadingstring)
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
    ''' GW_discharge  calc. with exponential discharge formula, std. = false
    ''' </summary>
    ''' <remarks></remarks>
    Private Exp As Boolean = False

    ''' <summary>
    ''' get zts file recursively
    ''' </summary>
    ''' <remarks></remarks>
    Private recursive As Boolean = False



    Private PRZMRunDir As String = String.Empty
    Private OldPRZMRunDir As String = String.Empty
    Private ArgPathRecurisve As String = String.Empty

    Private ZTSfiles2go As New List(Of String)
    Private SWASHNos As New List(Of Integer)




    Private Parent As String = String.Empty
    Private Met01 As String = String.Empty


    Private SWASHno As Integer = -1
    Private Scenario As String = String.Empty
    Private Crop As String = String.Empty
    Private OldCrop As String = String.Empty


    Private SimStart As Date
    Private SimEnd As Date

    Private applns As New List(Of String)
    Private p2tHeader As New List(Of String)
    Private p2tDataParent As New List(Of String)
    Private p2tDataMet01 As New List(Of String)

    Private HeavyRain As New List(Of String)


#Region "    get and parse CMD line arguments"

    Const argZTS As String = "zts:="
    Const argPath As String = "path:="
    Const argWarmup As String = "warmup:="
    Const argMRT As String = "mrt:="
    Const argMaxPRECperHour As String = "maxPREC:="
    Const argEXP As String = "exp:="
    Const argRecursive As String = "recursive:="

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

        End Try

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

            End Try

        End If

        add2Log(
             entry:=((argWarmup).PadLeft(logLen) & WarmUp.ToString & "years").PadRight(stdPos) &
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

            End Try

        End If

        add2Log(
            entry:=((argMRT).PadLeft(logLen) & MRT.ToString & "days").PadRight(stdPos) &
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
            Exit Sub
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
                Exit Sub
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
            Exit Sub

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

                std = False

            Catch ex As Exception

                add2Log(entry:=
                    "Error parsing cmd line for " & argRecursive & vbCrLf &
                     tempFilter.First & vbCrLf &
                     ex.Message)

            End Try

        End If

        add2Log(
            entry:=((argRecursive).PadLeft(logLen) & recursive.ToString).PadRight(stdPos) &
                        IIf(std, " (std.)", " *** user def. ***").ToString)

    End Sub



#End Region

#Region "    get header info"

    Private Sub getCMPNames()

        Const Filename As String = "MASTER.FPJ"
        Const CMPSearchString As String = "  Chemical Name:"

        Dim MasterFPJ As String() = {}

        Try

                add2Log(
                    entry:=("Reading:=").PadLeft(logLen) &
                    Filename)

            MasterFPJ =
                File.ReadAllLines(
                    path:=Path.Combine(PRZMRunDir,
                                       Filename))

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

            End Select

            'at least a parent must be given
            If Parent = String.Empty OrElse
               MasterFPJ.Count = 2 And Met01 = String.Empty Then

                add2Log(entry:=
                    "Fatal Error reading MASTER.FPJ file to parse for compound names" & vbCrLf &
                    Path.Combine(
                        PRZMRunDir,
                        Filename))

                End

            End If

        Catch ex As Exception

            add2Log(entry:=
                       "Fatal Error reading MASTER.FPJ file to parse for compound names" & vbCrLf &
                        Path.Combine(
                            PRZMRunDir,
                            Filename) & vbCrLf &
                        ex.Message)

            End

        End Try



    End Sub

    Private Sub getSWASHNos()

        Const Filename As String = "przm.pzm"
        Dim przmpzm As String() = {}
        Dim tempString As String() = {}
        Dim found As Boolean = False

        Try

            SWASHNos.Clear()

            add2Log(
                entry:=("Reading:=").PadLeft(logLen) &
                Filename)

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

                End

            End If

        Catch ex As Exception

            add2Log(entry:=
                      "Fatal Error parsing przm.pzm file for swash run numbers " & vbCrLf &
                       Path.Combine(PRZMRunDir, Filename) & vbCrLf &
                       ex.Message)

            End

        End Try

    End Sub

    Private Sub getTPAPpos(ZTSHeaderRow As String)

        Dim tempArray As String() = {}

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

        'add2Log(entry:=("TPAPpos:=").PadLeft(logLen) & posTPAP.ToString)

    End Sub

#End Region

    Private Enum eParMet
        Par
        Met
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
            .Add(LeadingString & "Sim start           : " & SimStart.ToString("dd-MMM-yyyy") &
                   (IIf(WarmUp <> 0, ", Warm up " & WarmUp & " years", "")).ToString)

            .Add(LeadingString & "      end           : " & SimEnd.ToString("dd-MMM-yyyy"))


            ApplnsPerSeason = CInt(Math.Round(applns.Count / (SimEnd.Year - SimStart.Year + 1), digits:=0))


            .Add(LeadingString & "Applns per season   : " & ApplnsPerSeason.ToString("00"))

            'parent or metabolite, must start with '*  Chemical:  ' to fit with txw info
            If ParMet = eParMet.Par Then

                .Add(LeadingString & "Chemical:  " & Parent)
                .Add(LeadingString & "Parent run")

                add2Log(entry:=("Scenario:=").PadLeft(logLen) & Scenario)
                add2Log(entry:=("SWASH #:=").PadLeft(logLen) & SWASHno.ToString)

            Else

                .Add(LeadingString & "Chemical:  " & Met01)
                .Add(LeadingString & "Metabolite run")

            End If

            .Add(LeadingString)
            .Add(LeadingString & "Total number of applns ")
            .Add("#    " & applns.Count)
            .Add(LeadingString)

            .Add(LeadingString)

            'appln info
            .Add(LeadingString & "Number of      Time                       Mass ")
            .Add(LeadingString & "Application    dd-MMM-YYYY-hh:mm          (g ai/ha)")
            .AddRange(applns)
            .Add(LeadingString)



            .Add(LeadingString & "                     Runoff         Runoff         Erosion        Erosion ")
            .Add("* " & "Time                  Volume         flux           Mass           Flux           Infiltration")
            .Add("* dd-MMM-YYYY-hh:mm     (mm/h)         (mg as/m2/h)   (kg/h)         (mg as/m2/h)   (mm/h)")

        End With

        p2tHeader = out

        Return True

    End Function

    Private Function createP2T(ZTSFileName As String) As Boolean

        Dim ZTSFile As String() = {}
        Dim tempString As String = String.Empty
        Dim tempArray As String() = {}

        Dim EventDate As New Date
        Dim EventDuration As Integer = 0

        Dim RUNF As Double = Double.NaN
        Dim PRCP As Double = Double.NaN
        Dim RFLX1 As Double = Double.NaN
        Dim RFLX2 As Double = Double.NaN
        Dim ESLS As Double = Double.NaN
        Dim EFLX1 As Double = Double.NaN
        Dim EFLX2 As Double = Double.NaN
        Dim INFL As Double = Double.NaN
        Dim IRRG As Double = Double.NaN
        Dim TPAP As Double = Double.NaN

        Dim GW_discharge As Double = 0
        Dim GW_storage As Double = 0
        Dim Timestep As Integer = 1

        'get zts file
        Try
            ZTSFile = File.ReadAllLines(ZTSFileName)

            'check for valid zts file
            If ZTSFile.Count < 400 Then
                add2Log(entry:=("IO Errorr:=").PadLeft(logLen) &
                                       "IO Error reading ZTS file, file empty or broken" & vbCrLf &
                                       ZTSFileName)
                Process.Start(fileName:=logFileName)
                Return False
            End If

        Catch ex As Exception
            add2Log(entry:=("IO Errorr:=").PadLeft(logLen) &
                       "IO Error reading ZTS file" & vbCrLf &
                       ZTSFileName & vbCrLf &
                       ex.Message)
            Process.Start(fileName:=logFileName)
            Return False
        End Try

        getTPAPpos(ZTSHeaderRow:=ZTSFile(ZTSHeaderRowNo))

        For RowCounter As Integer = ZTSDataStartRowNo To ZTSFile.Count - 1

            tempArray =
                ZTSFile(RowCounter).Split(
                    separator:={" "c},
                    options:=StringSplitOptions.RemoveEmptyEntries)

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
                Return False

            End Try


            If RowCounter = ZTSDataStartRowNo Then
                SimStart = EventDate
            ElseIf RowCounter = ZTSFile.Count - 1 Then
                SimEnd = EventDate
            End If

            Try

                RUNF = CDbl(tempArray(eZTSHeader.RUNF))
                ESLS = CDbl(tempArray(eZTSHeader.ESLS))
                PRCP = CDbl(tempArray(eZTSHeader.PRCP))
                INFL = CDbl(tempArray(eZTSHeader.INFL))

                RFLX1 = CDbl(tempArray(eZTSHeader.RFLX1))
                EFLX1 = CDbl(tempArray(eZTSHeader.EFLX1))

                TPAP = CDbl(tempArray(posTPAP))
                IRRG = CDbl(tempArray.Last)

            Catch ex As Exception

                add2Log(entry:=("ParseError:=").PadLeft(logLen) &
                       "Can't parse data from" & vbCrLf &
                       ZTSFile(RowCounter) & vbCrLf &
                       ex.Message)
                Process.Start(fileName:=logFileName)
                Return False

            End Try



            If TPAP <> 0 Then

                addAppln2List(
                    Eventdate:=EventDate,
                    TPAP:=TPAP)

                'add2Log(entry:=("Appln:=").PadLeft(logLen) & applns.Last)

            End If


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

            EventDuration = getEventDuration(PRCP:=PRCP, RUNF:=RUNF)
            If EventDuration = 12 And PRCP = 0 And reportSnowMelt Then
                add2Log(entry:=("Snow melt:=").PadLeft(logLen) &
                      ("at " & EventDate.ToString("dd-MMM-yyyy")).PadLeft("           dd-MMM-yyyy".Length))
            ElseIf PRCP >= HeavyRainLimit Then
                HeavyRain.Add(
                    (" ").PadLeft(logLen) &
                     PRCP.ToString("0.0").PadLeft(5) & "mm at " &
                    EventDate.ToString("dd-MMM-yyyy"))
            End If


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
                GW_storage = GW_storage + ((INFL - GW_discharge) * Timestep)

            End If


            'check for warm-up years
            If SimStart.AddYears(WarmUp) > EventDate Then Continue For

            p2tDataParent.AddRange(
                createP2TDay(
                    EventDate:=EventDate,
                    EventDuration:=EventDuration,
                    RUNF:=RUNF,
                    EFLX:=EFLX1,
                    ESLS:=ESLS,
                    RFLX:=RFLX1,
                    GW_discharge:=GW_discharge))

            If Met01 <> String.Empty Then
                p2tDataMet01.AddRange(
                   createP2TDay(
                       EventDate:=EventDate,
                       EventDuration:=EventDuration,
                       RUNF:=RUNF,
                       EFLX:=EFLX2,
                       ESLS:=ESLS,
                       RFLX:=RFLX2,
                       GW_discharge:=GW_discharge))
            End If

        Next

        Return True

    End Function


    Private Function createP2TDay(
                    EventDate As Date,
                    EventDuration As Integer,
                    RUNF As Double,
                    RFLX As Double,
                    ESLS As Double,
                    EFLX As Double,
                    GW_discharge As Double) As String()

        Dim out As New List(Of String)
        Dim tempString As String = String.Empty


        For HourCounter As Integer = 1 To 24

            If HourCounter <= EventDuration Then

                tempString =
                    (EventDate.ToString("dd-MMM-yyyy") &
                        "-" & HourCounter.ToString("00") & ":00").PadLeft("  01-Jan-1978-01:00".Length) &
                ("0" & (RUNF / EventDuration).ToString(".0000E+00")).PadLeft("     0.0000E+00".Length) &
                ("0" & (RFLX / EventDuration).ToString(".0000E+00")).PadLeft("     0.0000E+00".Length) &
                ("0" & (ESLS / EventDuration).ToString(".0000E+00")).PadLeft("     0.0000E+00".Length) &
                ("0" & (EFLX / EventDuration).ToString(".0000E+00")).PadLeft("     0.0000E+00".Length) &
                ("0" & (GW_discharge / 24).ToString(".0000E+00")).PadLeft("     0.0000E+00".Length)

                out.Add(tempString)

            Else

                out.Add((EventDate.ToString("dd-MMM-yyyy") &
                        "-" & HourCounter.ToString("00") & ":00").PadLeft("  01-Jan-1978-01:00".Length) &
                      "     0.0000E+00     0.0000E+00     0.0000E+00     0.0000E+00" &
                ("0" & (GW_discharge / 24).ToString(".0000E+00")).PadLeft("     0.0000E+00".Length))

            End If

        Next

        Return out.ToArray

    End Function


    Private Function getEventDuration(
                        PRCP As Double,
                        RUNF As Double) As Integer

        If PRCP / MaxPRECperHour > 24 Then
            Return 24
        ElseIf RUNF <> 0 AndAlso PRCP = 0 Then          
            Return 12
        Else
            Return CInt(Math.Round(
                PRCP / MaxPRECperHour,
                digits:=0,
                mode:=MidpointRounding.AwayFromZero))
        End If

    End Function

    Private Sub addAppln2List(
                    Eventdate As Date,
                    TPAP As Double)

        Dim tempString As String = String.Empty


        tempString = "#  " & (applns.Count + 1).ToString("00").PadRight("01             ".Length) &
            (Eventdate.ToString("dd-MMM-yyyy") & "-09:00").PadRight("01-Jan-1975-09:00          ".Length) &
            (TPAP * 100000000).ToString("0.00")

        applns.Add(tempString)

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
                    Out = " ? "
            End Select

            If SWASHno = -1 Then
                add2Log(entry:="No SWASH number for this scenario" & vbCrLf & Out)
                End
            End If

            Return Out

        Catch ex As Exception

            add2Log(entry:="Can't parse FOCUSsw runoff scenario from filename" & vbCrLf &
                             ZTSFileName & vbCrLf & ex.Message)
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
                add2Log(entry:="Can't parse crop from filename" & vbCrLf & ZTSFileName)
                End
            End If

            Return out

        Catch ex As Exception

            add2Log(entry:="Can't parse FOCUSsw runoff crop from filename" & vbCrLf &
                               ZTSFileName & vbCrLf & ex.Message)
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

End Module
