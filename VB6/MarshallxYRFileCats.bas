Attribute VB_Name = "MarshallxYRFileCats"
Option Explicit

Function GameChecksum(ByVal sFile As String, ByVal bTFD As Boolean) As String
    Select Case sFile
    Case "GAME"
        If bTFD Then
            GameChecksum = "6a3902386b8696069f59cc86e177daea"
        Else
            GameChecksum = "11178639bbb218df62a10Fa46157b035"
        End If
    Case "GAMEMD"
        If bTFD Then
            GameChecksum = "6604fd6f4ab71a11c708ea423dc0a29e"
        Else
            GameChecksum = "937b14db32f07ee47e92a8114ebdbdde"
        End If
    Case "RA2"
        If bTFD Then
            GameChecksum = "fcde1e5108b1bf48723adece5a7394ef"
        Else
            GameChecksum = "301a0Dbcf4ab8910aea1730Ecb0C997c"
        End If
    Case "RA2MD"
        If bTFD Then
            GameChecksum = "c4bbc74d1b9218f36da5bf7f9c3488eb"
        Else
            GameChecksum = "37c68043f478394cc1eb60b5c0b3f1ca"
        End If
    Case "YURI"
        If bTFD Then
            GameChecksum = "a7710A14956b611830eb72a763f9c520"
        Else
            GameChecksum = "22fea7ded6038e37ec78586b97f6d94d"
        End If
    End Select
End Function

Function FileIsAresComponent(ByVal FileName As String) As Boolean
    FileIsAresComponent = (UCase$(FileName) = "ARES.INI")
End Function

Function FileIsDestructive(ByVal FileName As String) As Boolean
    Dim iCounter As Integer
    Dim FileNameArray(79) As String
    FileNameArray(0) = "00000000.256" 'TFD only? what is this?
    FileNameArray(1) = "00000409.016" 'what are these?
    FileNameArray(2) = "00000409.256"
    FileNameArray(3) = "AMAZON.MMX"
    FileNameArray(4) = "BINKW32.DLL"
    FileNameArray(5) = "BLOWFISH.DLL"
    FileNameArray(6) = "BLOWFISH.TLB"
    FileNameArray(7) = "CONQUER.DAT"
    FileNameArray(8) = "CONQUERMD.DAT"
    FileNameArray(9) = "DRVMGT.DLL"
    FileNameArray(10) = "EB1.MMX"
    FileNameArray(11) = "EB2.MMX"
    FileNameArray(12) = "EB3.MMX"
    FileNameArray(13) = "EB4.MMX"
    FileNameArray(14) = "EB5.MMX"
    FileNameArray(15) = "EXCEPT.TXT"
    FileNameArray(16) = "EXPANDMD01.MIX"
    FileNameArray(17) = "GAME.EXE"
    FileNameArray(18) = "GAMEMD.EXE"
    FileNameArray(19) = "GAMEMD.EXE.BAK"
    FileNameArray(20) = "INVASION.MMX"
    FileNameArray(21) = "LANGMD.MIX"
    FileNameArray(22) = "LANGUAGE.MIX"
    FileNameArray(23) = "LAUNCHER.BMP"
    FileNameArray(24) = "LAUNCHER.TXT"
    FileNameArray(25) = "LAUNCHERMD.BMP"
    FileNameArray(26) = "MAPS01.MIX"
    FileNameArray(27) = "MAPS02.MIX"
    FileNameArray(28) = "MAPSMD03.MIX"
    FileNameArray(29) = "MOVIES01.MIX"
    FileNameArray(30) = "MOVIES02.MIX"
    FileNameArray(31) = "MOVMD03.MIX"
    FileNameArray(32) = "MPH.EXE"
    FileNameArray(33) = "MPHMD.EXE"
    FileNameArray(34) = "MULTI.MIX"
    FileNameArray(35) = "MULTIMD.MIX"
    FileNameArray(36) = "NOTES.ICO"
    FileNameArray(37) = "PATCH.DOC"
    FileNameArray(38) = "PATCHGET.DAT"
    FileNameArray(39) = "PATCHGETMD.DAT"
    FileNameArray(40) = "PATCHW32.DLL"
    FileNameArray(41) = "PREVIEW.BIN"
    FileNameArray(42) = "RA2.EXE"
    FileNameArray(43) = "RA2.ICO"
    FileNameArray(44) = "RA2.INI"
    FileNameArray(45) = "RA2.LCF"
    FileNameArray(46) = "RA2.MIX"
    FileNameArray(47) = "RA2.TLB"
    FileNameArray(48) = "RA2MD UPDATE.ICO"
    FileNameArray(49) = "RA2MD.EXE"
    FileNameArray(50) = "RA2MD.ICO"
    FileNameArray(51) = "RA2MD.INI"
    FileNameArray(52) = "RA2MD.LCF"
    FileNameArray(53) = "RA2MD.MIX"
    FileNameArray(54) = "RANDMAP.IMG"
    FileNameArray(55) = "RANDMAP.SED"
    FileNameArray(56) = "README.DOC"
    FileNameArray(57) = "README.TXT"
    FileNameArray(58) = "REGISTER.EXE"
    FileNameArray(59) = "REGISTER.INI"
    FileNameArray(60) = "SECDRV.SYS"
    FileNameArray(61) = "SYNC0.TXT"
    FileNameArray(62) = "SYNC1.TXT"
    FileNameArray(63) = "THEME.MIX"
    FileNameArray(64) = "THEMEMD.MIX"
    FileNameArray(65) = "UNINST.EXE"
    FileNameArray(66) = "UNINST.WSU"
    FileNameArray(67) = "UNINSTLL.EXE"
    FileNameArray(68) = "UNINSTMD.WSU"
    FileNameArray(69) = "UNINSTRA.WSU"
    FileNameArray(70) = "WOLDATA.KEY"
    FileNameArray(71) = "WOLINFO.INI"
    FileNameArray(72) = "WOLINFOMD.INI"
    FileNameArray(73) = "YR1.DSK" 'TFD only?
    FileNameArray(74) = "YURI.EXE"
    FileNameArray(75) = "YURI.LCF"
    FileNameArray(76) = "KEYBOARD.INI"
    FileNameArray(77) = "KEYBOARDMD.INI"
    FileNameArray(78) = "WDT.MIX" 'TFD only?
    FileNameArray(79) = "WSOCK32.DLL" 'third party wsock32.dll replaces the Windows System one to convert IPX to UDP for YR.
    FileIsDestructive = False
    FileName = UCase$(GetFileName(FileName))
    iCounter = UBound(FileNameArray)
    Do While iCounter <> -1
        If FileName = FileNameArray(iCounter) Then
            FileIsDestructive = True
            Exit Do
        End If
        iCounter = iCounter - 1
    Loop
End Function

Function FileIsFA2File(ByVal FileName As String) As Boolean
    Dim Counter As Integer
    Dim Ok As Boolean
    Dim SizeFileNameArray As Integer
    SizeFileNameArray = 3
    Dim FileNameArray(3) As String
    FileNameArray(1) = "FADATA.INI"
    FileNameArray(2) = "FALANGUAGE.INI"
    FileNameArray(3) = "MARBLE.MIX"
    Ok = False
    For Counter = 1 To SizeFileNameArray
        If UCase$(FileName) = FileNameArray(Counter) Then
            Ok = True
            Counter = SizeFileNameArray
        End If
    Next Counter
    FileIsFA2File = Ok
End Function

Function FileIsTXINI(ByVal FileName As String) As Boolean
    Dim Counter As Integer
    Dim Ok As Boolean
    Dim SizeFileNameArray As Integer
    SizeFileNameArray = 6
    Dim FileNameArray(6) As String
    FileNameArray(1) = "DESERTMD.INI"
    FileNameArray(2) = "LUNARMD.INI"
    FileNameArray(3) = "SNOWMD.INI"
    FileNameArray(4) = "TEMPERATMD.INI"
    FileNameArray(5) = "URBANMD.INI"
    FileNameArray(6) = "URBANNMD.INI"
    Ok = False
    For Counter = 1 To SizeFileNameArray
        If UCase$(FileName) = FileNameArray(Counter) Then
            Ok = True
            Counter = SizeFileNameArray
        End If
    Next Counter
    FileIsTXINI = Ok
End Function

Function FileIsReservedMix(ByVal FileName As String)
    Dim Counter As Integer
    Dim Ok As Boolean
    Dim SizeFileNameArray As Integer
    SizeFileNameArray = 9
    Dim FileNameArray(9) As String
    FileNameArray(1) = "EXPANDMD00.MIX"
    FileNameArray(2) = "EXPANDMD02.MIX"
    FileNameArray(3) = "EXPANDMD03.MIX"
    FileNameArray(4) = "EXPANDMD04.MIX"
    FileNameArray(5) = "EXPANDMD05.MIX"
    FileNameArray(6) = "EXPANDMD07.MIX"
    FileNameArray(7) = "EXPANDMD08.MIX"
    FileNameArray(8) = "EXPANDMD09.MIX"
    FileNameArray(9) = "EXPANDMD10.MIX"
    Ok = False
    FileName = UCase$(FileName)
    For Counter = 1 To SizeFileNameArray
        If FileName = FileNameArray(Counter) Then
            Ok = True
            Counter = SizeFileNameArray
        End If
    Next Counter
    FileIsReservedMix = Ok
End Function

Function FileIsModfile(ByVal FileName As String)
    'PRECONDITIONS:
    'Not FileIsDestructive
    'Not FileIsReserved
    'Not FileIsOfficialMapPackMap
    Dim Counter As Integer
    Dim Ok As Boolean
    Dim SizeFileNameArray, SizeFileNameArray2 As Integer
    Dim FileNameArray(59) As String
    Dim FileNameArray2(23) As String
    SizeFileNameArray = 59
    SizeFileNameArray2 = 23
    Ok = False
    FileNameArray(1) = "AI.TLB"
    FileNameArray(2) = "AUDIO.BAG"
    FileNameArray(3) = "AUDIO.IDX"
    FileNameArray(4) = "CREDITS.TXT"
    FileNameArray(5) = "CREDITSMD.TXT"
    FileNameArray(6) = "MOUSE.SHA"
    FileNameArray(7) = "RA2.CSF"
    FileNameArray(8) = "RA2MD.CSF"
    FileNameArray(9) = "SUBTITLE.TXT"
    FileNameArray(10) = "SUBTITLEMD.TXT"
    FileNameArray(11) = "VOXELS.VPA"
    'Inside ra2.mix
    FileNameArray(12) = "CACHE.MIX"
    FileNameArray(13) = "CACHEMD.MIX"
    FileNameArray(14) = "CONQUER.MIX"
    FileNameArray(15) = "CONQMD.MIX"
    FileNameArray(16) = "DES.MIX"
    FileNameArray(17) = "DESERT.MIX"
    FileNameArray(18) = "GENERIC.MIX"
    FileNameArray(19) = "GENERMD.MIX"
    FileNameArray(20) = "ISODES.MIX"
    FileNameArray(21) = "ISODESMD.MIX"
    FileNameArray(22) = "ISOGEN.MIX"
    FileNameArray(23) = "ISOGENMD.MIX"
    FileNameArray(24) = "ISOLUNMD.MIX"
    FileNameArray(25) = "ISOSNOMD.MIX"
    FileNameArray(26) = "ISOSNOW.MIX"
    FileNameArray(27) = "ISOTEMMD.MIX"
    FileNameArray(28) = "ISOTEMP.MIX"
    FileNameArray(29) = "ISOUBN.MIX"
    FileNameArray(30) = "ISOUBNMD.MIX"
    FileNameArray(31) = "ISOURB.MIX"
    FileNameArray(32) = "ISOURBMD.MIX"
    FileNameArray(33) = "LOAD.MIX"
    FileNameArray(34) = "LOADMD.MIX"
    FileNameArray(35) = "LOCAL.MIX"
    FileNameArray(36) = "LOCALMD.MIX"
    FileNameArray(37) = "LUN.MIX"
    FileNameArray(38) = "LUNAR.MIX"
    FileNameArray(39) = "NEUTRAL.MIX"
    FileNameArray(40) = "NTRLMD.MIX"
    FileNameArray(41) = "SIDEC01.MIX"
    FileNameArray(42) = "SIDEC02.MIX"
    FileNameArray(43) = "SIDEC01MD.MIX"
    FileNameArray(44) = "SIDEC02MD.MIX"
    FileNameArray(45) = "SIDENC01.MIX"
    FileNameArray(46) = "SIDENC02.MIX"
    FileNameArray(47) = "SNO.MIX"
    FileNameArray(48) = "SNOW.MIX"
    FileNameArray(49) = "SNOWMD.MIX"
    FileNameArray(50) = "TEM.MIX"
    FileNameArray(51) = "TEMPERAT.MIX"
    FileNameArray(52) = "UBN.MIX"
    FileNameArray(53) = "URB.MIX"
    FileNameArray(54) = "URBAN.MIX"
    FileNameArray(55) = "URBANN.MIX"
    'Inside language.mix
    FileNameArray(56) = "AUDIO.MIX"
    FileNameArray(57) = "AUDIOMD.MIX"
    FileNameArray(58) = "CAMEO.MIX"
    FileNameArray(59) = "CAMEOMD.MIX"
    'file types
    FileNameArray2(1) = "AUD"
    FileNameArray2(2) = "BIK"
    FileNameArray2(3) = "DES"
    FileNameArray2(4) = "FNT"
    FileNameArray2(5) = "HVA"
    FileNameArray2(6) = "INI"
    FileNameArray2(7) = "LUN"
    FileNameArray2(8) = "MIX"
    FileNameArray2(9) = "PAL"
    FileNameArray2(10) = "PCX"
    FileNameArray2(11) = "PKT"
    FileNameArray2(12) = "SHP"
    FileNameArray2(13) = "SNO"
    FileNameArray2(14) = "TEM"
    FileNameArray2(15) = "UBN"
    FileNameArray2(16) = "URB"
    FileNameArray2(17) = "VXL"
    FileNameArray2(18) = "MAP"
    FileNameArray2(19) = "MMX"
    FileNameArray2(20) = "MPR"
    FileNameArray2(21) = "WAV"
    FileNameArray2(22) = "YRM"
    FileNameArray2(23) = "YRO"
    FileName = UCase$(FileName)
    'If Not FileIsDestructive(FileName) and Not FileIsReserved Then
        If Len(FileName) > 4 Then
            For Counter = 1 To SizeFileNameArray2
                If FileType(FileName) = FileNameArray2(Counter) Then
                    If FileNameArray2(Counter) = "PCX" Then
                        If Len(FileName) >= 12 And Mid$(FileName, 1, 4) = "SCRN" Then
                            Ok = False
                        Else
                            Ok = True
                        End If
                    Else
                        If FileNameArray2(Counter) = "MIX" Then
                            If Len(FileName) >= 10 Then
                                If Mid$(FileName, 1, 6) = "EXPAND" Then
                                    Ok = True
                                ElseIf Mid$(FileName, 1, 6) = "ECACHE" Then
                                    Ok = True
                                ElseIf Mid$(FileName, 1, 6) = "ELOCAL" Then
                                    Ok = True
                                ElseIf Mid$(FileName, 1, 4) = "SIDE" Then
                                    Ok = True
                                End If
                            End If
                        Else
                            Ok = True
                        End If
                    End If
                    Counter = SizeFileNameArray2
                End If
            Next Counter
            If Ok = False Then
                For Counter = 1 To SizeFileNameArray
                    If FileName = FileNameArray(Counter) Then
                        Ok = True
                        Counter = SizeFileNameArray
                    End If
                Next Counter
            End If
        End If
    'End If
    FileIsModfile = Ok
End Function

Function FileIsSoundtrack(ByVal FileName As String)
    Dim Counter As Integer
    Dim Ok As Boolean
    Dim SizeFileNameArray As Integer
    SizeFileNameArray = 18
    Dim FileNameArray(18) As String
    FileNameArray(1) = "ECACHETHEMERA2.MIX"
    FileNameArray(2) = "ECACHETHEMETD.MIX"
    FileNameArray(3) = "ECACHETHEMECO.MIX"
    FileNameArray(4) = "ECACHETHEMERA.MIX"
    FileNameArray(5) = "ECACHETHEMERED1.MIX"
    FileNameArray(6) = "ECACHETHEMERED2.MIX"
    FileNameArray(7) = "ECACHETHEMEAM.MIX"
    FileNameArray(8) = "ECACHETHEMECS.MIX"
    FileNameArray(9) = "ECACHETHEMETS.MIX"
    FileNameArray(10) = "ECACHETHEMETS1.MIX"
    FileNameArray(11) = "ECACHETHEMETS2.MIX"
    FileNameArray(12) = "ECACHETHEMEFS.MIX"
    FileNameArray(13) = "ECACHETHEMEREN.MIX"
    FileNameArray(14) = "ECACHETHEMEREN1.MIX"
    FileNameArray(15) = "ECACHETHEMEREN2.MIX"
    FileNameArray(16) = "ECACHETHEMEUSER.MIX"
    FileNameArray(17) = "ECACHETHEMEEU.MIX"
    FileNameArray(18) = "ECACHETHEMERA2X.MIX"
    Ok = False
    FileName = UCase$(FileName)
    For Counter = 1 To SizeFileNameArray
        If FileName = FileNameArray(Counter) Then
            Ok = True
            Counter = SizeFileNameArray
        End If
    Next Counter
    FileIsSoundtrack = Ok
End Function

Function FileIsUserTheme(ByVal FileName As String)
    Dim Ok As Boolean
    Ok = False
    If Len(FileName) = 10 Then
        If FileType(FileName) = "WAV" Then
            FileName = UCase$(Left$(FileName, 6))
            If Left$(FileName, 4) = "USER" Then
                FileName = Right$(FileName, 2)
                If Len(StripNonNumbers(FileName)) = Len(FileName) Then
                    If Val(FileName) Then
                        Ok = True
                    End If
                End If
            End If
        End If
    End If
    FileIsUserTheme = Ok
End Function

Function FileIsCustomPlaylist(ByVal FileName As String, Optional ByVal FileFull As String = "")
    If UCase$(FileName) = "THEMEMD.INI" Or FileType(FileName) = "YPL" Then
        If Len(FileFull) <> 0 Then
            FileIsCustomPlaylist = Len(ReadINIStr("YRPMOPTS", "Music", FileFull)) <> 0
        Else
            FileIsCustomPlaylist = True
        End If
    Else
        FileIsCustomPlaylist = False
    End If
End Function

Function FileIsCustomMap(ByVal FileName As String)
    Dim Ok As Boolean
    Select Case FileType(FileName)
    Case "MAP", "MPR", "YRM"
        Ok = Not FileIsOfficialMap(FileName)
    End Select
    FileIsCustomMap = Ok
End Function

Function FileIsMap(ByVal FileName As String)
    Select Case FileType(FileName)
    Case "MAP", "MPR", "YRM", "MMX", "YRO"
        FileIsMap = True
    Case Else
        FileIsMap = False
    End Select
End Function

Function FileIsSeed(ByVal FileName As String, Optional ByVal FilePathIfCheckInside As String = "")
    Dim Ok As Boolean
    Ok = True
    If FileType(FileName) <> "SED" Then
        Ok = False
    Else
        If Len(FilePathIfCheckInside) <> 0 Then
            If FileExists(FilePathIfCheckInside) Then
                If ReadINIStr("RandomMap", "Seed", FilePathIfCheckInside) = "" Then
                    Ok = False
                End If
            Else
                Ok = False
            End If
        End If
    End If
    FileIsSeed = Ok
End Function

Function FileIsOfficialMap(ByVal FileName As String)
    Dim Counter As Integer
    Dim Ok As Boolean
    Dim SizeFileNameArray As Integer
    SizeFileNameArray = 328
    Dim FileNameArray(328) As String
    'missionsmd.pkt
    FileNameArray(1) = "XMP29U2.MAP"
    FileNameArray(2) = "XMP21S2.MAP"
    FileNameArray(3) = "BldFeud.MAP"
    FileNameArray(4) = "XDeadman.MAP"
    FileNameArray(5) = "DunePatr.MAP"
    FileNameArray(6) = "XDustbowl.MAP"
    FileNameArray(7) = "XMP14T2.MAP"
    FileNameArray(8) = "XHailMary.MAP"
    FileNameArray(9) = "XXmas.MAP"
    FileNameArray(10) = "HidVally.MAP"
    FileNameArray(11) = "HillBtwn.MAP"
    FileNameArray(12) = "XMP02T2.MAP"
    FileNameArray(13) = "Fight.MAP"
    FileNameArray(14) = "XMP08T2.MAP"
    FileNameArray(15) = "PcOfDune.MAP"
    FileNameArray(16) = "XBarrel.MAP"
    FileNameArray(17) = "XMP31S2.MAP"
    FileNameArray(18) = "XGrinder.MAP"
    FileNameArray(19) = "XNewHghts.MAP"
    FileNameArray(20) = "XMP11T2.MAP"
    FileNameArray(21) = "SaharaMi.MAP"
    FileNameArray(22) = "XEB3.MAP"
    FileNameArray(23) = "2Peaks.MAP"
    FileNameArray(24) = "XMP06T2.MAP"
    FileNameArray(25) = "XBreak.MAP"
    FileNameArray(26) = "XMP09T3.MAP"
    FileNameArray(27) = "XMP18S3.MAP"
    FileNameArray(28) = "XKiller.MAP"
    FileNameArray(29) = "XYuriPlot.MAP"
    FileNameArray(30) = "XMP03T4.MAP"
    FileNameArray(31) = "XArena.MAP"
    FileNameArray(32) = "XDisaster.MAP"
    FileNameArray(33) = "XCarville.MAP"
    FileNameArray(34) = "XEB1.MAP"
    FileNameArray(35) = "XEB4.MAP"
    FileNameArray(36) = "XMP34U4.MAP"
    FileNameArray(37) = "XMP10S4.MAP"
    FileNameArray(38) = "DoubleTrouble.MAP"
    FileNameArray(39) = "downtown.MAP"
    FileNameArray(40) = "DryHeat.MAP"
    FileNameArray(41) = "FaceDown.MAP"
    FileNameArray(42) = "FourCorners.MAP"
    FileNameArray(43) = "Frstbite.MAP"
    FileNameArray(44) = "XMP19T4.MAP"
    FileNameArray(45) = "GroundZe.MAP"
    FileNameArray(46) = "XMP23T4.MAP"
    FileNameArray(47) = "XHills.MAP"
    FileNameArray(48) = "XMP05T4.MAP"
    FileNameArray(49) = "XInvasion.MAP"
    FileNameArray(50) = "IsleOfOades.MAP"
    FileNameArray(51) = "XMP12S4.MAP"
    FileNameArray(52) = "XLostlake.MAP"
    FileNameArray(53) = "Manhatta.MAP"
    FileNameArray(54) = "XMP13S4.MAP"
    FileNameArray(55) = "XEB5.MAP"
    FileNameArray(56) = "XNoRest.MAP"
    FileNameArray(57) = "NoWimps.MAP"
    FileNameArray(58) = "XOceansid.MAP"
    FileNameArray(59) = "OffenseDefense.MAP"
    FileNameArray(60) = "OttersRevenge.MAP"
    FileNameArray(61) = "XPacific.MAP"
    FileNameArray(62) = "XMP33U4.MAP"
    FileNameArray(63) = "XRockets.MAP"
    FileNameArray(64) = "XRound.MAP"
    FileNameArray(65) = "RushHr.MAP"
    FileNameArray(66) = "XShrapnel.MAP"
    FileNameArray(67) = "XEB2.MAP"
    FileNameArray(68) = "XMP15S4.MAP"
    FileNameArray(69) = "XMP16S4.MAP"
    FileNameArray(70) = "XMP01T4.MAP"
    FileNameArray(71) = "XTanyas.MAP"
    FileNameArray(72) = "XTOWER.MAP"
    FileNameArray(73) = "XTsunami.MAP"
    FileNameArray(74) = "XValley.MAP"
    FileNameArray(75) = "Arena33Forever.MAP"
    FileNameArray(76) = "XBayOPigs.MAP"
    FileNameArray(77) = "XMP26S6.MAP"
    FileNameArray(78) = "BridgeGap.MAP"
    FileNameArray(79) = "XMP25T6.MAP"
    FileNameArray(80) = "EastVsBest.MAP"
    FileNameArray(81) = "XKaliforn.MAP"
    FileNameArray(82) = "XMP17T6.MAP"
    FileNameArray(83) = "MonumentValley.MAP"
    FileNameArray(84) = "SedonaPass.MAP"
    FileNameArray(85) = "XMP30S6.MAP"
    FileNameArray(86) = "XGoldSt.MAP"
    FileNameArray(87) = "TourOfEgypt.MAP"
    FileNameArray(88) = "XMP20T6.MAP"
    FileNameArray(89) = "XDeath.MAP"
    FileNameArray(90) = "XMP32S8.MAP"
    FileNameArray(91) = "XMP27T8.MAP"
    FileNameArray(92) = "XPowdrKeg.MAP"
    FileNameArray(93) = "Xroulette.MAP"
    FileNameArray(94) = "TrailerPark.MAP"
    FileNameArray(95) = "C1A01MD.MAP"
    FileNameArray(96) = "C1A02MD.MAP"
    FileNameArray(97) = "C1A03MD.MAP"
    FileNameArray(98) = "C2S01MD.MAP"
    FileNameArray(99) = "C2S02MD.MAP"
    FileNameArray(100) = "C2S03MD.MAP"
    FileNameArray(101) = "C3Y01MD.MAP"
    FileNameArray(102) = "C3Y02MD.MAP"
    FileNameArray(103) = "C3Y03MD.MAP"
    FileNameArray(104) = "C4W01MD.MAP"
    FileNameArray(105) = "XMP09DU.MAP"
    FileNameArray(106) = "XMP24DU.MAP"
    FileNameArray(107) = "XMP18DU.MAP"
    FileNameArray(108) = "XMP05DU.MAP"
    FileNameArray(109) = "XMP13DU.MAP"
    FileNameArray(110) = "XMP15DU.MAP"
    FileNameArray(111) = "XMP01DU.MAP"
    FileNameArray(112) = "XMP25DU.MAP"
    FileNameArray(113) = "XMP17DU.MAP"
    FileNameArray(114) = "XMP27DU.MAP"
    FileNameArray(115) = "XMP32DU.MAP"
    FileNameArray(116) = "XDustbowlmw.MAP"
    FileNameArray(117) = "XMP14MW.MAP"
    FileNameArray(118) = "XMP08MW.MAP"
    FileNameArray(119) = "SaharaMimw.MAP"
    FileNameArray(120) = "XMP29MW.MAP"
    FileNameArray(121) = "XMP06MW.MAP"
    FileNameArray(122) = "XEB1mw.MAP"
    FileNameArray(123) = "DeathValleyGirlmw.MAP"
    FileNameArray(124) = "DryHeatmw.MAP"
    FileNameArray(125) = "FourCornersmw.MAP"
    FileNameArray(126) = "Groundzemw.MAP"
    FileNameArray(127) = "XMP23MW.MAP"
    FileNameArray(128) = "XMP05MW.MAP"
    FileNameArray(129) = "XMP13MW.MAP"
    FileNameArray(130) = "XPacificmw.MAP"
    FileNameArray(131) = "XMP15MW.MAP"
    FileNameArray(132) = "XMP16MW.MAP"
    FileNameArray(133) = "XTowermw.MAP"
    FileNameArray(134) = "XValleymw.MAP"
    FileNameArray(135) = "XMP25MW.MAP"
    FileNameArray(136) = "XMP17MW.MAP"
    FileNameArray(137) = "MonumentValleymw.MAP"
    FileNameArray(138) = "SedonaPassmw.MAP"
    FileNameArray(139) = "XMP30MW.MAP"
    FileNameArray(140) = "XMP27MW.MAP"
    FileNameArray(141) = "XMP22MW.MAP"
    FileNameArray(142) = "XMP32MW.MAP"
    FileNameArray(143) = "TopOTheHill.MAP"
    FileNameArray(144) = "AustinTX.MAP"
    FileNameArray(145) = "DeathValleyGirl.MAP"
    FileNameArray(146) = "XSeaofIso.MAP"
    FileNameArray(147) = "XAMAZON01.MAP"
    FileNameArray(148) = "TurfWar.MAP"
    FileNameArray(149) = "MountMoras.MAP"
    FileNameArray(150) = "XPotomac.MAP"
    FileNameArray(151) = "XBermuda.MAP"
    FileNameArray(152) = "TripleCrossed.MAP"
    FileNameArray(153) = "XMP22S8.MAP"
    FileNameArray(154) = "NearOreF.MAP"
    FileNameArray(155) = "XTerritor.MAP"
    FileNameArray(156) = "XTN02s4.MAP"
    FileNameArray(157) = "XTN01T2.MAP"
    FileNameArray(158) = "XTN04T2.MAP"
    FileNameArray(159) = "XTN02MW.MAP"
    FileNameArray(160) = "XTN01MW.MAP"
    FileNameArray(161) = "XTN04MW.MAP"
    'missions.pkt
    FileNameArray(162) = "MP02T2.MAP"
    FileNameArray(163) = "MP06T2.MAP"
    FileNameArray(164) = "MP11T2.MAP"
    FileNameArray(165) = "MP08T2.MAP"
    FileNameArray(166) = "MP21S2.MAP"
    FileNameArray(167) = "MP14T2.MAP"
    FileNameArray(168) = "MP29U2.MAP"
    FileNameArray(169) = "MP31S2.MAP"
    FileNameArray(170) = "MP18S3.MAP"
    FileNameArray(171) = "MP09T3.MAP"
    FileNameArray(172) = "MP01T4.MAP"
    FileNameArray(173) = "MP03T4.MAP"
    FileNameArray(174) = "MP05T4.MAP"
    FileNameArray(175) = "MP10S4.MAP"
    FileNameArray(176) = "MP12S4.MAP"
    FileNameArray(177) = "MP13S4.MAP"
    FileNameArray(178) = "MP19T4.MAP"
    FileNameArray(179) = "MP15S4.MAP"
    FileNameArray(180) = "MP16S4.MAP"
    FileNameArray(181) = "MP23T4.MAP"
    FileNameArray(182) = "MP33U4.MAP"
    FileNameArray(183) = "MP34U4.MAP"
    FileNameArray(184) = "MP17T6.MAP"
    FileNameArray(185) = "MP20T6.MAP"
    FileNameArray(186) = "MP25T6.MAP"
    FileNameArray(187) = "MP26S6.MAP"
    FileNameArray(188) = "MP30S6.MAP"
    FileNameArray(189) = "MP22S8.MAP"
    FileNameArray(190) = "MP27T8.MAP"
    FileNameArray(191) = "MP32S8.MAP"
    FileNameArray(192) = "MP06MW.MAP"
    FileNameArray(193) = "MP08MW.MAP"
    FileNameArray(194) = "MP14MW.MAP"
    FileNameArray(195) = "MP29MW.MAP"
    FileNameArray(196) = "MP05MW.MAP"
    FileNameArray(197) = "MP13MW.MAP"
    FileNameArray(198) = "MP15MW.MAP"
    FileNameArray(199) = "MP16MW.MAP"
    FileNameArray(200) = "MP23MW.MAP"
    FileNameArray(201) = "MP17MW.MAP"
    FileNameArray(202) = "MP25MW.MAP"
    FileNameArray(203) = "MP30MW.MAP"
    FileNameArray(204) = "MP22MW.MAP"
    FileNameArray(205) = "MP27MW.MAP"
    FileNameArray(206) = "MP32MW.MAP"
    FileNameArray(207) = "MP09DU.MAP"
    FileNameArray(208) = "MP01DU.MAP"
    FileNameArray(209) = "MP05DU.MAP"
    FileNameArray(210) = "MP13DU.MAP"
    FileNameArray(211) = "MP15DU.MAP"
    FileNameArray(212) = "MP18DU.MAP"
    FileNameArray(213) = "MP24DU.MAP"
    FileNameArray(214) = "MP17DU.MAP"
    FileNameArray(215) = "MP25DU.MAP"
    FileNameArray(216) = "MP27DU.MAP"
    FileNameArray(217) = "MP32DU.MAP"
    FileNameArray(218) = "C1M1A.MAP"
    FileNameArray(219) = "C1M1B.MAP"
    FileNameArray(220) = "C1M1C.MAP"
    FileNameArray(221) = "C1M2A.MAP"
    FileNameArray(222) = "C1M2B.MAP"
    FileNameArray(223) = "C1M2C.MAP"
    FileNameArray(224) = "C1M3A.MAP"
    FileNameArray(225) = "C1M3B.MAP"
    FileNameArray(226) = "C1M3C.MAP"
    FileNameArray(227) = "C1M4A.MAP"
    FileNameArray(228) = "C1M4B.MAP"
    FileNameArray(229) = "C1M4C.MAP"
    FileNameArray(230) = "C1M5A.MAP"
    FileNameArray(231) = "C1M5B.MAP"
    FileNameArray(232) = "C1M5C.MAP"
    FileNameArray(233) = "C2M1A.MAP"
    FileNameArray(234) = "C2M1B.MAP"
    FileNameArray(235) = "C2M1C.MAP"
    FileNameArray(236) = "C2M2A.MAP"
    FileNameArray(237) = "C2M2B.MAP"
    FileNameArray(238) = "C2M2C.MAP"
    FileNameArray(239) = "C2M3A.MAP"
    FileNameArray(240) = "C2M3B.MAP"
    FileNameArray(241) = "C2M3C.MAP"
    FileNameArray(242) = "C2M4A.MAP"
    FileNameArray(243) = "C2M4B.MAP"
    FileNameArray(244) = "C2M4C.MAP"
    FileNameArray(245) = "C2M5A.MAP"
    FileNameArray(246) = "C2M5B.MAP"
    FileNameArray(247) = "C2M5C.MAP"
    FileNameArray(248) = "C3M1A.MAP"
    FileNameArray(249) = "C3M1B.MAP"
    FileNameArray(250) = "C3M1C.MAP"
    FileNameArray(251) = "C3M2A.MAP"
    FileNameArray(252) = "C3M2B.MAP"
    FileNameArray(253) = "C3M2C.MAP"
    FileNameArray(254) = "C3M3A.MAP"
    FileNameArray(255) = "C3M3B.MAP"
    FileNameArray(256) = "C3M3C.MAP"
    FileNameArray(257) = "C3M4A.MAP"
    FileNameArray(258) = "C3M4B.MAP"
    FileNameArray(259) = "C3M4C.MAP"
    FileNameArray(260) = "C3M5A.MAP"
    FileNameArray(261) = "C3M5B.MAP"
    FileNameArray(262) = "C3M5C.MAP"
    FileNameArray(263) = "C4M1A.MAP"
    FileNameArray(264) = "C4M1B.MAP"
    FileNameArray(265) = "C4M1C.MAP"
    FileNameArray(266) = "C4M2A.MAP"
    FileNameArray(267) = "C4M2B.MAP"
    FileNameArray(268) = "C4M2C.MAP"
    FileNameArray(269) = "C4M3A.MAP"
    FileNameArray(270) = "C4M3B.MAP"
    FileNameArray(271) = "C4M3C.MAP"
    FileNameArray(272) = "C4M4A.MAP"
    FileNameArray(273) = "C4M4B.MAP"
    FileNameArray(274) = "C4M4C.MAP"
    FileNameArray(275) = "C4M5A.MAP"
    FileNameArray(276) = "C4M5B.MAP"
    FileNameArray(277) = "C4M5C.MAP"
    FileNameArray(278) = "C5M1A.MAP"
    FileNameArray(279) = "C5M1B.MAP"
    FileNameArray(280) = "C5M1C.MAP"
    FileNameArray(281) = "C5M2A.MAP"
    FileNameArray(282) = "C5M2B.MAP"
    FileNameArray(283) = "C5M2C.MAP"
    FileNameArray(284) = "C5M3A.MAP"
    FileNameArray(285) = "C5M3B.MAP"
    FileNameArray(286) = "C5M3C.MAP"
    FileNameArray(287) = "C5M4A.MAP"
    FileNameArray(288) = "C5M4B.MAP"
    FileNameArray(289) = "C5M4C.MAP"
    FileNameArray(290) = "C5M5A.MAP"
    FileNameArray(291) = "C5M5B.MAP"
    FileNameArray(292) = "C5M5C.MAP"
    FileNameArray(293) = "TN01T2.MAP"
    FileNameArray(294) = "TN01MW.MAP"
    FileNameArray(295) = "TN04T2.MAP"
    FileNameArray(296) = "TN04MW.MAP"
    FileNameArray(297) = "TN02s4.MAP"
    FileNameArray(298) = "TN02MW.MAP"
    'YR SP maps
    FileNameArray(299) = "ALL01UMD.MAP"
    FileNameArray(300) = "ALL02UMD.MAP"
    FileNameArray(301) = "ALL03UMD.MAP"
    FileNameArray(302) = "ALL04DMD.MAP"
    FileNameArray(303) = "ALL05UMD.MAP"
    FileNameArray(304) = "ALL06UMD.MAP"
    FileNameArray(305) = "ALL07SMD.MAP"
    FileNameArray(306) = "SOV01UMD.MAP"
    FileNameArray(307) = "SOV02SMD.MAP"
    FileNameArray(308) = "SOV03UMD.MAP"
    FileNameArray(309) = "SOV04DMD.MAP"
    FileNameArray(310) = "SOV05UMD.MAP"
    FileNameArray(311) = "SOV06LMD.MAP"
    FileNameArray(312) = "SOV07TMD.MAP"
    'RA2 SP maps
    FileNameArray(313) = "ALL01T.MAP"
    FileNameArray(314) = "E31.MAP"
    FileNameArray(315) = "E32.MAP"
    FileNameArray(316) = "SOV01T.MAP"
    FileNameArray(317) = "SOV02T.MAP"
    FileNameArray(318) = "SOV03U.MAP"
    FileNameArray(319) = "SOV04S.MAP"
    FileNameArray(320) = "SOV05U.MAP"
    FileNameArray(321) = "SOV06T.MAP"
    FileNameArray(322) = "SOV07S.MAP"
    FileNameArray(323) = "SOV08U.MAP"
    FileNameArray(324) = "SOV09U.MAP"
    FileNameArray(325) = "SOV10T.MAP"
    FileNameArray(326) = "SOV11S.MAP"
    FileNameArray(327) = "SOV12S.MAP"
    FileNameArray(328) = "SOV1U.MAP"
    Ok = False
    For Counter = 1 To SizeFileNameArray
        If UCase$(FileName) = UCase$(FileNameArray(Counter)) Then
            Ok = True
            Counter = SizeFileNameArray
        End If
    Next Counter
    FileIsOfficialMap = Ok
End Function

Function FileIsOfficialMapPackMap(ByVal FileName As String)
    FileIsOfficialMapPackMap = FileIsRA2MapPackMap(FileName) Or FileIsYRMapPackMap(FileName)
End Function

Function FileIsRA2MapPackMap(ByVal FileName As String)
    Dim Counter As Integer
    Dim LocalBoolean As Boolean
    Dim SizeLocalFileArray As Integer
    SizeLocalFileArray = 33
    Dim LocalFileArray(33) As String
    LocalFileArray(1) = "ARENA.MMX"
    LocalFileArray(2) = "BARREL.MMX"
    LocalFileArray(3) = "BAYOPIGS.MMX"
    LocalFileArray(4) = "BERMUDA.MMX"
    LocalFileArray(5) = "BREAK.MMX"
    LocalFileArray(6) = "CARVILLE.MMX"
    LocalFileArray(7) = "DEADMAN.MMX"
    LocalFileArray(8) = "DEATH.MMX"
    LocalFileArray(9) = "DISASTER.MMX"
    LocalFileArray(10) = "DUSTBOWL.MMX"
    LocalFileArray(11) = "GOLDST.MMX"
    LocalFileArray(12) = "GRINDER.MMX"
    LocalFileArray(13) = "HAILMARY.MMX"
    LocalFileArray(14) = "HILLS.MMX"
    LocalFileArray(15) = "KALIFORN.MMX"
    LocalFileArray(16) = "KILLER.MMX"
    LocalFileArray(17) = "LOSTLAKE.MMX"
    LocalFileArray(18) = "NEWHGHTS.MMX"
    LocalFileArray(19) = "OCEANSID.MMX"
    LocalFileArray(20) = "PACIFIC.MMX"
    LocalFileArray(21) = "POTOMAC.MMX"
    LocalFileArray(22) = "POWDRKEG.MMX"
    LocalFileArray(23) = "ROCKETS.MMX"
    LocalFileArray(24) = "ROULETTE.MMX"
    LocalFileArray(25) = "ROUND.MMX"
    LocalFileArray(26) = "SEAOFISO.MMX"
    LocalFileArray(27) = "SHRAPNEL.MMX"
    LocalFileArray(28) = "TANYAS.MMX"
    LocalFileArray(29) = "TOWER.MMX"
    LocalFileArray(30) = "TSUNAMI.MMX"
    LocalFileArray(31) = "VALLEY.MMX"
    LocalFileArray(32) = "XMAS.MMX"
    LocalFileArray(33) = "YURIPLOT.MMX"
    LocalBoolean = False
    For Counter = 1 To SizeLocalFileArray
        If UCase$(FileName) = LocalFileArray(Counter) Then
            LocalBoolean = True
            Counter = SizeLocalFileArray
        End If
    Next Counter
    FileIsRA2MapPackMap = LocalBoolean
End Function

Function FileIsYRMapPackMap(ByVal FileName As String)
    Dim Counter As Integer
    Dim LocalBoolean As Boolean
    Dim SizeLocalFileArray As Integer
    SizeLocalFileArray = 13
    Dim LocalFileArray(13) As String
    LocalFileArray(1) = "CRCTBRD.YRO"
    LocalFileArray(2) = "DEEPFRZE.YRO"
    LocalFileArray(3) = "HIGHEXPR.YRO"
    LocalFileArray(4) = "ICE_AGE.YRO"
    LocalFileArray(5) = "IRVINECA.YRO"
    LocalFileArray(6) = "ISLELAND.YRO"
    LocalFileArray(7) = "MOJOSPRT.YRO"
    LocalFileArray(8) = "MONSTERM.YRO"
    LocalFileArray(9) = "MOONPATR.YRO"
    LocalFileArray(10) = "RIVERRAM.YRO"
    LocalFileArray(11) = "SINKSWIM.YRO"
    LocalFileArray(12) = "TRANSYLV.YRO"
    LocalFileArray(13) = "UNREPENT.YRO"
    LocalBoolean = False
    For Counter = 1 To SizeLocalFileArray
        If UCase$(FileName) = LocalFileArray(Counter) Then
            LocalBoolean = True
            Counter = SizeLocalFileArray
        End If
    Next Counter
    FileIsYRMapPackMap = LocalBoolean
End Function

Function FileIsAssaultMapPack(ByVal FileName As String)
    Dim Ok As Boolean
    Ok = False
    If Len(FileName) >= 19 Then
        If FileType(FileName) = "MIX" Then
            If UCase$(Left$(FileName, 13)) = "ECACHEASSAULT" Then
                FileName = Mid$(FileName, 14, (Len(FileName) - 17))
                If Len(StripNonNumbers(FileName)) = Len(FileName) Then
                    Ok = True
                End If
            End If
        End If
    End If
    FileIsAssaultMapPack = Ok
End Function

Public Sub DisectScrnFormat(ByVal ScrnFormat As String, ByRef ReturnPre As String, ByRef ReturnPost As String, ByRef ReturnNumFormat As Long)
    Dim sNumFormat As String
    Dim PercentPos As Integer
    Dim dPos As Integer
    ReturnPre = ""
    ReturnPost = ""
    ReturnNumFormat = -1
    sNumFormat = ""
    PercentPos = InStr(1, ScrnFormat, "%")
    If PercentPos <> 0 Then
        dPos = InStr(PercentPos + 1, ScrnFormat, "d")
        If dPos > PercentPos + 1 Then 'must be at least one character between % and d
            sNumFormat = Mid$(ScrnFormat, PercentPos + 1, (dPos - PercentPos) - 1)
            If StripNonNumbers(sNumFormat) = sNumFormat Then
                ReturnNumFormat = Val(sNumFormat)
                If PercentPos > 1 Then ReturnPre = Left$(ScrnFormat, PercentPos - 1)
                If dPos < Len(ScrnFormat) Then ReturnPost = Mid$(ScrnFormat, dPos + 1)
            End If
        End If
    End If
    'If ReturnPost is empty then there is nothing after the "%Nd" or there is no "%Nd"
    'If ReturnPre is empty then there is nothing before the "%Nd" or there is no "%Nd"
    'If ReturnNumFormat is -1 then there is no "%Nd"
End Sub

Function ConfirmScrnFormat(ByVal FileName As String, ByVal ScrnFormat As String) As Boolean
    Dim ReturnPre As String
    Dim ReturnPost As String
    Dim ReturnNumFormat As Long
    Call DisectScrnFormat(ScrnFormat, ReturnPre, ReturnPost, ReturnNumFormat)
    If ReturnNumFormat = -1 Then
        ConfirmScrnFormat = False
    Else
        FileName = UCase$(FileName)
        ReturnPre = UCase$(ReturnPre)
        ReturnPost = UCase$(ReturnPost)
        If Left$(FileName, Len(ReturnPre)) = ReturnPre Then
            If Right$(FileName, Len(ReturnPost)) = ReturnPost Then
                'we're not returning this anymore because nothing needed it.
                'ReturnNum = Val(Mid$(FileName, Len(ReturnPre) + 1, Len(FileName) - Len(ReturnPre) - Len(ReturnPost)))
                ConfirmScrnFormat = True
            End If
        End If
    End If
End Function

Function ConfirmScrnFormatReturnNum(ByVal FileName As String, ByVal ScrnFormat As String, ByRef ReturnNum As Long) As Boolean
    Dim ReturnPre As String
    Dim ReturnPost As String
    Dim ReturnNumFormat As Long
    Call DisectScrnFormat(ScrnFormat, ReturnPre, ReturnPost, ReturnNumFormat)
    If ReturnNumFormat = -1 Then
        ConfirmScrnFormatReturnNum = False
    Else
        FileName = UCase$(FileName)
        ReturnPre = UCase$(ReturnPre)
        ReturnPost = UCase$(ReturnPost)
        If Left$(FileName, Len(ReturnPre)) = ReturnPre Then
            If Right$(FileName, Len(ReturnPost)) = ReturnPost Then
                'and then something needed it!
                ReturnNum = Val(Mid$(FileName, Len(ReturnPre) + 1, Len(FileName) - Len(ReturnPre) - Len(ReturnPost)))
                ConfirmScrnFormatReturnNum = True
            End If
        End If
    End If
End Function

Function FileIsSaveGame(ByVal FileName As String)
    If FileType(FileName) = "SAV" Or UCase$(FileName) = "COOPSAVE.INI" Then
        FileIsSaveGame = True
    Else
        FileIsSaveGame = False
    End If
End Function

Function FileIsOfficialTaunt(ByVal FileName As String)
    Dim Counter As Integer
    Dim Ok As Boolean
    Dim SizeLocalFileArray As Integer
    SizeLocalFileArray = 80
    Dim LocalFileArray(80) As String
    LocalFileArray(1) = "TAUAM01.WAV"
    LocalFileArray(2) = "TAUAM02.WAV"
    LocalFileArray(3) = "TAUAM03.WAV"
    LocalFileArray(4) = "TAUAM04.WAV"
    LocalFileArray(5) = "TAUAM05.WAV"
    LocalFileArray(6) = "TAUAM06.WAV"
    LocalFileArray(7) = "TAUAM07.WAV"
    LocalFileArray(8) = "TAUAM08.WAV"
    LocalFileArray(9) = "TAUBR01.WAV"
    LocalFileArray(10) = "TAUBR02.WAV"
    LocalFileArray(11) = "TAUBR03.WAV"
    LocalFileArray(12) = "TAUBR04.WAV"
    LocalFileArray(13) = "TAUBR05.WAV"
    LocalFileArray(14) = "TAUBR06.WAV"
    LocalFileArray(15) = "TAUBR07.WAV"
    LocalFileArray(16) = "TAUBR08.WAV"
    LocalFileArray(17) = "TAUCU01.WAV"
    LocalFileArray(18) = "TAUCU02.WAV"
    LocalFileArray(19) = "TAUCU03.WAV"
    LocalFileArray(20) = "TAUCU04.WAV"
    LocalFileArray(21) = "TAUCU05.WAV"
    LocalFileArray(22) = "TAUCU06.WAV"
    LocalFileArray(23) = "TAUCU07.WAV"
    LocalFileArray(24) = "TAUCU08.WAV"
    LocalFileArray(25) = "TAUFR01.WAV"
    LocalFileArray(26) = "TAUFR02.WAV"
    LocalFileArray(27) = "TAUFR03.WAV"
    LocalFileArray(28) = "TAUFR04.WAV"
    LocalFileArray(29) = "TAUFR05.WAV"
    LocalFileArray(30) = "TAUFR06.WAV"
    LocalFileArray(31) = "TAUFR07.WAV"
    LocalFileArray(32) = "TAUFR08.WAV"
    LocalFileArray(33) = "TAUGE01.WAV"
    LocalFileArray(34) = "TAUGE02.WAV"
    LocalFileArray(35) = "TAUGE03.WAV"
    LocalFileArray(36) = "TAUGE04.WAV"
    LocalFileArray(37) = "TAUGE05.WAV"
    LocalFileArray(38) = "TAUGE06.WAV"
    LocalFileArray(39) = "TAUGE07.WAV"
    LocalFileArray(40) = "TAUGE08.WAV"
    LocalFileArray(41) = "TAUIR01.WAV"
    LocalFileArray(42) = "TAUIR02.WAV"
    LocalFileArray(43) = "TAUIR03.WAV"
    LocalFileArray(44) = "TAUIR04.WAV"
    LocalFileArray(45) = "TAUIR05.WAV"
    LocalFileArray(46) = "TAUIR06.WAV"
    LocalFileArray(47) = "TAUIR07.WAV"
    LocalFileArray(48) = "TAUIR08.WAV"
    LocalFileArray(49) = "TAUKO01.WAV"
    LocalFileArray(50) = "TAUKO02.WAV"
    LocalFileArray(51) = "TAUKO03.WAV"
    LocalFileArray(52) = "TAUKO04.WAV"
    LocalFileArray(53) = "TAUKO05.WAV"
    LocalFileArray(54) = "TAUKO06.WAV"
    LocalFileArray(55) = "TAUKO07.WAV"
    LocalFileArray(56) = "TAUKO08.WAV"
    LocalFileArray(57) = "TAULI01.WAV"
    LocalFileArray(58) = "TAULI02.WAV"
    LocalFileArray(59) = "TAULI03.WAV"
    LocalFileArray(60) = "TAULI04.WAV"
    LocalFileArray(61) = "TAULI05.WAV"
    LocalFileArray(62) = "TAULI06.WAV"
    LocalFileArray(63) = "TAULI07.WAV"
    LocalFileArray(64) = "TAULI08.WAV"
    LocalFileArray(65) = "TAURU01.WAV"
    LocalFileArray(66) = "TAURU02.WAV"
    LocalFileArray(67) = "TAURU03.WAV"
    LocalFileArray(68) = "TAURU04.WAV"
    LocalFileArray(69) = "TAURU05.WAV"
    LocalFileArray(70) = "TAURU06.WAV"
    LocalFileArray(71) = "TAURU07.WAV"
    LocalFileArray(72) = "TAURU08.WAV"
    LocalFileArray(73) = "TAUYU01.WAV"
    LocalFileArray(74) = "TAUYU02.WAV"
    LocalFileArray(75) = "TAUYU03.WAV"
    LocalFileArray(76) = "TAUYU04.WAV"
    LocalFileArray(77) = "TAUYU05.WAV"
    LocalFileArray(78) = "TAUYU06.WAV"
    LocalFileArray(79) = "TAUYU07.WAV"
    LocalFileArray(80) = "TAUYU08.WAV"
    Ok = False
    For Counter = 1 To SizeLocalFileArray
        If UCase$(FileName) = LocalFileArray(Counter) Then
            Ok = True
            Counter = SizeLocalFileArray
        End If
    Next Counter
    FileIsOfficialTaunt = Ok
End Function

Function FileIsRPTaunt(ByVal FileName As String)
    Dim Ok As Boolean
    Ok = False
    If Len(FileName) = 11 Then
        If FileType(FileName) = "WAV" Then
            If UCase$(Left$(FileName, 3)) = "TAU" Then
                Select Case Mid$(FileName, 6, 2)
                Case "01", "02", "03", "04", "05", "06", "07", "08"
                FileName = Mid$(FileName, 4, 2)
                If Len(StripNonNumbers(FileName)) = Len(FileName) Then Ok = True
                End Select
            End If
        End If
    End If
    FileIsRPTaunt = Ok
End Function


