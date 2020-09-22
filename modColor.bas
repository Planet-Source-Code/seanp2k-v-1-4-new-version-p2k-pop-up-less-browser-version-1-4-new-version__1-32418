Attribute VB_Name = "modColor"
'*****************************************************************
'*              Color processing subroutines                     *
'*              written by Chavdar Yordanov, 04.2001             *
'*              Email: chavo@beer.com                            *
'*              Please, don't remove this title!                 *
'*****************************************************************
Option Explicit
Public Declare Function CalcGradient Lib "bfcolor1.dll" (ByVal hDC As Long, ByVal lMaxColor As Long, ByVal bShowHue As Byte, ByVal ColDepth As Byte) As Long

Public RGBs() As String
Public SafeCol(224) As Long
Public iColorDepth As Integer

'dividers for GetColorByte function
Public Const clr24Bit = 1
Public Const clr16Bit = 8
Public Const clrWebSafe = 51

Public Function HexToLong(sHexColor As String) As Long
    Dim lCol As Long, i, N
    If Left(sHexColor, 1) = "#" Then sHexColor = Mid(sHexColor, 2)
    sHexColor = UCase(sHexColor)
    
    For i = 1 To Len(sHexColor) Step 2
        lCol = lCol + Dec(Mid(sHexColor, i, 2)) * 256 ^ N
        N = N + 1
    Next i
    HexToLong = lCol
End Function

Public Function RgbToLong(sRgbColor As String) As Long
    Dim vCol, i, lCol As Long, N
    vCol = Split(sRgbColor, ",")
    For i = LBound(vCol) To UBound(vCol)
        lCol = lCol + Val(vCol(i)) * 256 ^ N
        N = N + 1
    Next i
    RgbToLong = lCol
End Function

Public Function Dec(ByVal sHex As String) As Long 'Converts Hex to Decimal
    Const HVal = "0123456789ABCDEF"
    Dim iPos As Byte, i As Integer, lDec As Long
    Dim L As Integer, X As Byte
    L = Len(sHex)
    If L > 255 Then Exit Function
    lDec = 0
    For i = L To 1 Step -1
        X = InStr(1, HVal, Mid(sHex, i, 1), vbTextCompare)
        If X = 0 Then Exit Function Else X = X - 1
        lDec = lDec + X * 16 ^ (L - i)
    Next i
    Dec = lDec
End Function

Public Function Invert(ByVal iCol As Long) As Long
    Dim bCol() As Byte    'Byte values
    SplitIntoBytes iCol, 3, bCol()
    Invert = RGB(255 - bCol(1), 255 - bCol(2), 255 - bCol(3))
End Function

Public Sub SplitIntoBytes(ByVal lNumber As Long, bSize As Byte, ByRef bBytes() As Byte, Optional bRedim = True)
    Dim i As Long
    Dim KF As Long
    If bRedim Then ReDim bBytes(1 To bSize)
    For i = bSize To 1 Step -1
        KF = 256 ^ (i - 1)
        bBytes(i) = lNumber \ KF
        lNumber = lNumber - bBytes(i) * KF
    Next i
End Sub

Public Sub GetWebColors(ByRef cbWeb As ComboBox)
    cbWeb.AddItem "[not in the list]"
    cbWeb.ItemData(cbWeb.NewIndex) = -1
    cbWeb.AddItem "aliceblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 16775408
    cbWeb.AddItem "antiquewhite"
    cbWeb.ItemData(cbWeb.NewIndex) = 14150650
    cbWeb.AddItem "aqua"
    cbWeb.ItemData(cbWeb.NewIndex) = 16776960
    cbWeb.AddItem "aquamarine"
    cbWeb.ItemData(cbWeb.NewIndex) = 13959039
    cbWeb.AddItem "azure"
    cbWeb.ItemData(cbWeb.NewIndex) = 16777200
    cbWeb.AddItem "beige"
    cbWeb.ItemData(cbWeb.NewIndex) = 14480885
    cbWeb.AddItem "bisque"
    cbWeb.ItemData(cbWeb.NewIndex) = 12903679
    cbWeb.AddItem "black"
    cbWeb.ItemData(cbWeb.NewIndex) = 0
    cbWeb.AddItem "blanchedalmond"
    cbWeb.ItemData(cbWeb.NewIndex) = 13495295
    cbWeb.AddItem "blue"
    cbWeb.ItemData(cbWeb.NewIndex) = 16711680
    cbWeb.AddItem "blueviolet"
    cbWeb.ItemData(cbWeb.NewIndex) = 14822282
    cbWeb.AddItem "brown"
    cbWeb.ItemData(cbWeb.NewIndex) = 2763429
    cbWeb.AddItem "burlywood"
    cbWeb.ItemData(cbWeb.NewIndex) = 8894686
    cbWeb.AddItem "cadetblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 10526303
    cbWeb.AddItem "chartreuse"
    cbWeb.ItemData(cbWeb.NewIndex) = 65407
    cbWeb.AddItem "chocolate"
    cbWeb.ItemData(cbWeb.NewIndex) = 1993170
    cbWeb.AddItem "coral"
    cbWeb.ItemData(cbWeb.NewIndex) = 5275647
    cbWeb.AddItem "cornflower"
    cbWeb.ItemData(cbWeb.NewIndex) = 15570276
    cbWeb.AddItem "cornsilk"
    cbWeb.ItemData(cbWeb.NewIndex) = 14481663
    cbWeb.AddItem "crimson"
    cbWeb.ItemData(cbWeb.NewIndex) = 3937500
    cbWeb.AddItem "cyan"
    cbWeb.ItemData(cbWeb.NewIndex) = 16776960
    cbWeb.AddItem "darkblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 9109504
    cbWeb.AddItem "darkcyan"
    cbWeb.ItemData(cbWeb.NewIndex) = 9145088
    cbWeb.AddItem "darkgoldenrod"
    cbWeb.ItemData(cbWeb.NewIndex) = 755384
    cbWeb.AddItem "darkgray"
    cbWeb.ItemData(cbWeb.NewIndex) = 11119017
    cbWeb.AddItem "darkgreen"
    cbWeb.ItemData(cbWeb.NewIndex) = 25600
    cbWeb.AddItem "darkkhaki"
    cbWeb.ItemData(cbWeb.NewIndex) = 7059389
    cbWeb.AddItem "darkmagenta"
    cbWeb.ItemData(cbWeb.NewIndex) = 9109643
    cbWeb.AddItem "darkolivegreen"
    cbWeb.ItemData(cbWeb.NewIndex) = 3107669
    cbWeb.AddItem "darkorange"
    cbWeb.ItemData(cbWeb.NewIndex) = 36095
    cbWeb.AddItem "darkorchid"
    cbWeb.ItemData(cbWeb.NewIndex) = 13382297
    cbWeb.AddItem "darkred"
    cbWeb.ItemData(cbWeb.NewIndex) = 139
    cbWeb.AddItem "darksalmon"
    cbWeb.ItemData(cbWeb.NewIndex) = 8034025
    cbWeb.AddItem "darkseagreen"
    cbWeb.ItemData(cbWeb.NewIndex) = 9157775
    cbWeb.AddItem "darkslateblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 9125192
    cbWeb.AddItem "darkslategray"
    cbWeb.ItemData(cbWeb.NewIndex) = 5197615
    cbWeb.AddItem "darkturquoise"
    cbWeb.ItemData(cbWeb.NewIndex) = 13749760
    cbWeb.AddItem "darkviolet"
    cbWeb.ItemData(cbWeb.NewIndex) = 13828244
    cbWeb.AddItem "deeppink"
    cbWeb.ItemData(cbWeb.NewIndex) = 9639167
    cbWeb.AddItem "deepskyblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 16760576
    cbWeb.AddItem "dimgray"
    cbWeb.ItemData(cbWeb.NewIndex) = 6908265
    cbWeb.AddItem "dodgerblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 16748574
    cbWeb.AddItem "firebrick"
    cbWeb.ItemData(cbWeb.NewIndex) = 2237106
    cbWeb.AddItem "floralwhite"
    cbWeb.ItemData(cbWeb.NewIndex) = 15792895
    cbWeb.AddItem "forestgreen"
    cbWeb.ItemData(cbWeb.NewIndex) = 2263842
    cbWeb.AddItem "fuchia"
    cbWeb.ItemData(cbWeb.NewIndex) = 16711935
    cbWeb.AddItem "gainsboro"
    cbWeb.ItemData(cbWeb.NewIndex) = 14474460
    cbWeb.AddItem "ghostwhite"
    cbWeb.ItemData(cbWeb.NewIndex) = 16775416
    cbWeb.AddItem "gold"
    cbWeb.ItemData(cbWeb.NewIndex) = 55295
    cbWeb.AddItem "goldenrod"
    cbWeb.ItemData(cbWeb.NewIndex) = 2139610
    cbWeb.AddItem "gray"
    cbWeb.ItemData(cbWeb.NewIndex) = 8421504
    cbWeb.AddItem "green"
    cbWeb.ItemData(cbWeb.NewIndex) = 32768
    cbWeb.AddItem "greenyellow"
    cbWeb.ItemData(cbWeb.NewIndex) = 3145645
    cbWeb.AddItem "honeydew"
    cbWeb.ItemData(cbWeb.NewIndex) = 15794160
    cbWeb.AddItem "hotpink"
    cbWeb.ItemData(cbWeb.NewIndex) = 11823615
    cbWeb.AddItem "indianred"
    cbWeb.ItemData(cbWeb.NewIndex) = 6053069
    cbWeb.AddItem "indigo"
    cbWeb.ItemData(cbWeb.NewIndex) = 8519755
    cbWeb.AddItem "ivory"
    cbWeb.ItemData(cbWeb.NewIndex) = 15794175
    cbWeb.AddItem "khaki"
    cbWeb.ItemData(cbWeb.NewIndex) = 9234160
    cbWeb.AddItem "lavender"
    cbWeb.ItemData(cbWeb.NewIndex) = 16443110
    cbWeb.AddItem "lavenderblush"
    cbWeb.ItemData(cbWeb.NewIndex) = 16118015
    cbWeb.AddItem "lawngreen"
    cbWeb.ItemData(cbWeb.NewIndex) = 64636
    cbWeb.AddItem "lemonchiffon"
    cbWeb.ItemData(cbWeb.NewIndex) = 13499135
    cbWeb.AddItem "lightblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 15128749
    cbWeb.AddItem "lightcoral"
    cbWeb.ItemData(cbWeb.NewIndex) = 8421616
    cbWeb.AddItem "lightcyan"
    cbWeb.ItemData(cbWeb.NewIndex) = 16777184
    cbWeb.AddItem "lightgoldenrodyellow"
    cbWeb.ItemData(cbWeb.NewIndex) = 13826810
    cbWeb.AddItem "lightgreen"
    cbWeb.ItemData(cbWeb.NewIndex) = 9498256
    cbWeb.AddItem "lightgrey"
    cbWeb.ItemData(cbWeb.NewIndex) = 13882323
    cbWeb.AddItem "lightpink"
    cbWeb.ItemData(cbWeb.NewIndex) = 12695295
    cbWeb.AddItem "lightsalmon"
    cbWeb.ItemData(cbWeb.NewIndex) = 8036607
    cbWeb.AddItem "lightseagreen"
    cbWeb.ItemData(cbWeb.NewIndex) = 11186720
    cbWeb.AddItem "lightskyblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 16436871
    cbWeb.AddItem "lightslategray"
    cbWeb.ItemData(cbWeb.NewIndex) = 10061943
    cbWeb.AddItem "lightsteelblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 14599344
    cbWeb.AddItem "lightyellow"
    cbWeb.ItemData(cbWeb.NewIndex) = 14745599
    cbWeb.AddItem "lime"
    cbWeb.ItemData(cbWeb.NewIndex) = 65280
    cbWeb.AddItem "limegreen"
    cbWeb.ItemData(cbWeb.NewIndex) = 3329330
    cbWeb.AddItem "linen"
    cbWeb.ItemData(cbWeb.NewIndex) = 15134970
    cbWeb.AddItem "magenta"
    cbWeb.ItemData(cbWeb.NewIndex) = 16711935
    cbWeb.AddItem "maroon"
    cbWeb.ItemData(cbWeb.NewIndex) = 128
    cbWeb.AddItem "mediumaquamarine"
    cbWeb.ItemData(cbWeb.NewIndex) = 11193702
    cbWeb.AddItem "mediumblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 13434880
    cbWeb.AddItem "mediumorchid"
    cbWeb.ItemData(cbWeb.NewIndex) = 13850042
    cbWeb.AddItem "mediumpurple"
    cbWeb.ItemData(cbWeb.NewIndex) = 14381203
    cbWeb.AddItem "mediumseagreen"
    cbWeb.ItemData(cbWeb.NewIndex) = 7451452
    cbWeb.AddItem "mediumslateblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 15624315
    cbWeb.AddItem "mediumspringgreen"
    cbWeb.ItemData(cbWeb.NewIndex) = 10156544
    cbWeb.AddItem "mediumturquoise"
    cbWeb.ItemData(cbWeb.NewIndex) = 13422920
    cbWeb.AddItem "mediumvioletred"
    cbWeb.ItemData(cbWeb.NewIndex) = 8721863
    cbWeb.AddItem "midnightblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 7346457
    cbWeb.AddItem "mintcream"
    cbWeb.ItemData(cbWeb.NewIndex) = 16449525
    cbWeb.AddItem "mistyrose"
    cbWeb.ItemData(cbWeb.NewIndex) = 14804223
    cbWeb.AddItem "moccasin"
    cbWeb.ItemData(cbWeb.NewIndex) = 11920639
    cbWeb.AddItem "navajowhite"
    cbWeb.ItemData(cbWeb.NewIndex) = 11394815
    cbWeb.AddItem "navy"
    cbWeb.ItemData(cbWeb.NewIndex) = 8388608
    cbWeb.AddItem "oldlace"
    cbWeb.ItemData(cbWeb.NewIndex) = 15136253
    cbWeb.AddItem "olive"
    cbWeb.ItemData(cbWeb.NewIndex) = 32896
    cbWeb.AddItem "olivedrab"
    cbWeb.ItemData(cbWeb.NewIndex) = 2330219
    cbWeb.AddItem "orange"
    cbWeb.ItemData(cbWeb.NewIndex) = 42495
    cbWeb.AddItem "orangered"
    cbWeb.ItemData(cbWeb.NewIndex) = 17919
    cbWeb.AddItem "orchid"
    cbWeb.ItemData(cbWeb.NewIndex) = 14053594
    cbWeb.AddItem "palegoldenrod"
    cbWeb.ItemData(cbWeb.NewIndex) = 11200750
    cbWeb.AddItem "palegreen"
    cbWeb.ItemData(cbWeb.NewIndex) = 10025880
    cbWeb.AddItem "paleturquoise"
    cbWeb.ItemData(cbWeb.NewIndex) = 15658671
    cbWeb.AddItem "palevioletred"
    cbWeb.ItemData(cbWeb.NewIndex) = 9662683
    cbWeb.AddItem "papayawhip"
    cbWeb.ItemData(cbWeb.NewIndex) = 14020607
    cbWeb.AddItem "peachpuff"
    cbWeb.ItemData(cbWeb.NewIndex) = 12180223
    cbWeb.AddItem "peru"
    cbWeb.ItemData(cbWeb.NewIndex) = 4163021
    cbWeb.AddItem "pink"
    cbWeb.ItemData(cbWeb.NewIndex) = 13353215
    cbWeb.AddItem "plum"
    cbWeb.ItemData(cbWeb.NewIndex) = 14524637
    cbWeb.AddItem "powderblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 15130800
    cbWeb.AddItem "purple"
    cbWeb.ItemData(cbWeb.NewIndex) = 8388736
    cbWeb.AddItem "red"
    cbWeb.ItemData(cbWeb.NewIndex) = 255
    cbWeb.AddItem "rosybrown"
    cbWeb.ItemData(cbWeb.NewIndex) = 9408444
    cbWeb.AddItem "royalblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 14772545
    cbWeb.AddItem "saddlebrown"
    cbWeb.ItemData(cbWeb.NewIndex) = 1262987
    cbWeb.AddItem "salmon"
    cbWeb.ItemData(cbWeb.NewIndex) = 7504122
    cbWeb.AddItem "sandybrown"
    cbWeb.ItemData(cbWeb.NewIndex) = 6333684
    cbWeb.AddItem "seagreen"
    cbWeb.ItemData(cbWeb.NewIndex) = 5737262
    cbWeb.AddItem "seashell"
    cbWeb.ItemData(cbWeb.NewIndex) = 15660543
    cbWeb.AddItem "sienna"
    cbWeb.ItemData(cbWeb.NewIndex) = 2970272
    cbWeb.AddItem "silver"
    cbWeb.ItemData(cbWeb.NewIndex) = 12632256
    cbWeb.AddItem "skyblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 15453831
    cbWeb.AddItem "slateblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 13458026
    cbWeb.AddItem "slategray"
    cbWeb.ItemData(cbWeb.NewIndex) = 9470064
    cbWeb.AddItem "snow"
    cbWeb.ItemData(cbWeb.NewIndex) = 16448255
    cbWeb.AddItem "springgreen"
    cbWeb.ItemData(cbWeb.NewIndex) = 8388352
    cbWeb.AddItem "steelblue"
    cbWeb.ItemData(cbWeb.NewIndex) = 11829830
    cbWeb.AddItem "tan"
    cbWeb.ItemData(cbWeb.NewIndex) = 9221330
    cbWeb.AddItem "teal"
    cbWeb.ItemData(cbWeb.NewIndex) = 8421376
    cbWeb.AddItem "thistle"
    cbWeb.ItemData(cbWeb.NewIndex) = 14204888
    cbWeb.AddItem "tomato"
    cbWeb.ItemData(cbWeb.NewIndex) = 4678655
    cbWeb.AddItem "turquoise"
    cbWeb.ItemData(cbWeb.NewIndex) = 13688896
    cbWeb.AddItem "violet"
    cbWeb.ItemData(cbWeb.NewIndex) = 15631086
    cbWeb.AddItem "wheat"
    cbWeb.ItemData(cbWeb.NewIndex) = 11788021
    cbWeb.AddItem "white"
    cbWeb.ItemData(cbWeb.NewIndex) = 16777215
    cbWeb.AddItem "whitesmoke"
    cbWeb.ItemData(cbWeb.NewIndex) = 16119285
    cbWeb.AddItem "yellow"
    cbWeb.ItemData(cbWeb.NewIndex) = 65535
    cbWeb.AddItem "yellowgreen"
    cbWeb.ItemData(cbWeb.NewIndex) = 3329434
    SafeCol(1) = 16777215
    SafeCol(2) = 26316
    SafeCol(3) = 10066431
    SafeCol(4) = 3355494
    SafeCol(5) = 16724991
    SafeCol(6) = 13382553
    SafeCol(7) = 6684723
    SafeCol(8) = 13369344
    SafeCol(9) = 16763955
    SafeCol(10) = 13421568
    SafeCol(11) = 6736947
    SafeCol(12) = 3407616
    SafeCol(13) = 26163
    SafeCol(14) = 3381657
    SafeCol(15) = 13434879
    SafeCol(16) = 3368601
    SafeCol(17) = 6711039
    SafeCol(18) = 13408767
    SafeCol(19) = 16711935
    SafeCol(20) = 16724889
    SafeCol(21) = 13408665
    SafeCol(22) = 13382400
    SafeCol(23) = 16763904
    SafeCol(24) = 13434777
    SafeCol(25) = 6736896
    SafeCol(26) = 65280
    SafeCol(27) = 65433
    SafeCol(28) = 39321
    SafeCol(29) = 10092543
    SafeCol(30) = 13158
    SafeCol(31) = 3355647
    SafeCol(32) = 10040319
    SafeCol(33) = 10040268
    SafeCol(34) = 16711833
    SafeCol(35) = 16737894
    SafeCol(36) = 3342336
    SafeCol(37) = 13408563
    SafeCol(38) = 13434726
    SafeCol(39) = 3381504
    SafeCol(40) = 52224
    SafeCol(41) = 3407769
    SafeCol(42) = 26214
    SafeCol(43) = 6750207
    SafeCol(44) = 10079487
    SafeCol(45) = 3342591
    SafeCol(46) = 10027263
    SafeCol(47) = 13408716
    SafeCol(48) = 10027110
    SafeCol(49) = 13395558
    SafeCol(50) = 16737792
    SafeCol(51) = 13408512
    SafeCol(52) = 13434675
    SafeCol(53) = 10092441
    SafeCol(54) = 3394611
    SafeCol(55) = 6737049
    SafeCol(56) = 13421772
    SafeCol(57) = 3407871
    SafeCol(58) = 6724095
    SafeCol(59) = 3342540
    SafeCol(60) = 6684876
    SafeCol(61) = 13395660
    SafeCol(62) = 6697830
    SafeCol(63) = 10053222
    SafeCol(64) = 13395456
    SafeCol(65) = 3355392
    SafeCol(66) = 13434624
    SafeCol(67) = 6750054
    SafeCol(68) = 3394560
    SafeCol(69) = 52377
    SafeCol(70) = 10066329
    SafeCol(71) = 65535
    SafeCol(72) = 26367
    SafeCol(73) = 10066380
    SafeCol(74) = 6697881
    SafeCol(75) = 13369548
    SafeCol(76) = 13395609
    SafeCol(77) = 10040115
    SafeCol(78) = 13395507
    SafeCol(79) = 13421721
    SafeCol(80) = 13421619
    SafeCol(81) = 6750003
    SafeCol(82) = 65382
    SafeCol(83) = 3394713
    SafeCol(84) = 6710886
    SafeCol(85) = 52428
    SafeCol(86) = 3368652
    SafeCol(87) = 6710988
    SafeCol(88) = 3342438
    SafeCol(89) = 13382604
    SafeCol(90) = 16737945
    SafeCol(91) = 6697779
    SafeCol(92) = 16724736
    SafeCol(93) = 13421670
    SafeCol(94) = 10066176
    SafeCol(95) = 6749952
    SafeCol(96) = 3407718
    SafeCol(97) = 39270
    SafeCol(98) = 3355443
    SafeCol(99) = 6737151
    SafeCol(100) = 13209
    SafeCol(101) = 3355596
    SafeCol(102) = 13395711
    SafeCol(103) = 10027161
    SafeCol(104) = 13369446
    SafeCol(105) = 16724787
    SafeCol(106) = 10040064
    SafeCol(107) = 10066227
    SafeCol(108) = 10079334
    SafeCol(109) = 3381555
    SafeCol(110) = 65331
    SafeCol(111) = 6750156
    SafeCol(112) = 0
    SafeCol(113) = 52479
    SafeCol(114) = 102
    SafeCol(115) = 3355545
    SafeCol(116) = 13369599
    SafeCol(117) = 10040217
    SafeCol(118) = 13382502
    SafeCol(119) = 16711731
    SafeCol(120) = 6697728
    SafeCol(121) = 10066278
    SafeCol(122) = 10079283
    SafeCol(123) = 26112
    SafeCol(124) = 52275
    SafeCol(125) = 65484
    SafeCol(126) = 16777215
    SafeCol(127) = 3394815
    SafeCol(128) = 3368703
    SafeCol(129) = 3342489
    SafeCol(130) = 13382655
    SafeCol(131) = 16737996
    SafeCol(132) = 10040166
    SafeCol(133) = 13369395
    SafeCol(134) = 16764057
    SafeCol(135) = 6710784
    SafeCol(136) = 10079232
    SafeCol(137) = 13434828
    SafeCol(138) = 39219
    SafeCol(139) = 3407820
    SafeCol(140) = 16777215
    SafeCol(141) = 39372
    SafeCol(142) = 13260
    SafeCol(143) = 51
    SafeCol(144) = 10053324
    SafeCol(145) = 16724940
    SafeCol(146) = 3342387
    SafeCol(147) = 13382451
    SafeCol(148) = 16750899
    SafeCol(149) = 6710835
    SafeCol(150) = 6723891
    SafeCol(151) = 10079385
    SafeCol(152) = 39168
    SafeCol(153) = 10079436
    SafeCol(154) = 16777215
    SafeCol(155) = 3381708
    SafeCol(156) = 13311
    SafeCol(157) = 10053375
    SafeCol(158) = 10027212
    SafeCol(159) = 16711884
    SafeCol(160) = 16764108
    SafeCol(161) = 10027008
    SafeCol(162) = 16750848
    SafeCol(163) = 16777164
    SafeCol(164) = 6723840
    SafeCol(165) = 6736998
    SafeCol(166) = 10092492
    SafeCol(167) = 3368550
    SafeCol(168) = 16777215
    SafeCol(169) = 26265
    SafeCol(170) = 255
    SafeCol(171) = 6697983
    SafeCol(172) = 6684825
    SafeCol(173) = 13369497
    SafeCol(174) = 16751001
    SafeCol(175) = 6684672
    SafeCol(176) = 13408614
    SafeCol(177) = 16777113
    SafeCol(178) = 3368448
    SafeCol(179) = 6723942
    SafeCol(180) = 6750105
    SafeCol(181) = 13107
    SafeCol(182) = 16777215
    SafeCol(183) = 39423
    SafeCol(184) = 204
    SafeCol(185) = 6684927
    SafeCol(186) = 16764159
    SafeCol(187) = 10053273
    SafeCol(188) = 16724838
    SafeCol(189) = 16750950
    SafeCol(190) = 10053171
    SafeCol(191) = 16777062
    SafeCol(192) = 10092390
    SafeCol(193) = 3368499
    SafeCol(194) = 52326
    SafeCol(195) = 6737100
    SafeCol(196) = 16777215
    SafeCol(197) = 3381759
    SafeCol(198) = 153
    SafeCol(199) = 6697932
    SafeCol(200) = 16751103
    SafeCol(201) = 6684774
    SafeCol(202) = 16711782
    SafeCol(203) = 16737843
    SafeCol(204) = 10053120
    SafeCol(205) = 16777011
    SafeCol(206) = 10092339
    SafeCol(207) = 13056
    SafeCol(208) = 3394662
    SafeCol(209) = 3394764
    SafeCol(210) = 16777215
    SafeCol(211) = 6724044
    SafeCol(212) = 13421823
    SafeCol(213) = 6710937
    SafeCol(214) = 16738047
    SafeCol(215) = 16751052
    SafeCol(216) = 10027059
    SafeCol(217) = 16711680
    SafeCol(218) = 16764006
    SafeCol(219) = 16776960
    SafeCol(220) = 10092288
    SafeCol(221) = 3407667
    SafeCol(222) = 3381606
    SafeCol(223) = 6723993
    SafeCol(224) = 16777215
End Sub

Public Function CalcColorDepth(lColor As Long) As Long
    CalcColorDepth = GetColorByte(lColor And &HFF&)
    CalcColorDepth = CalcColorDepth + CLng(GetColorByte((lColor And &HFF00&) \ &H100&)) * 256
    CalcColorDepth = CalcColorDepth + CLng(GetColorByte((lColor And &HFF0000) \ &H10000)) * 65536
End Function

Public Function GetColorByte(bCol As Long) As Byte
    Dim Z As Long
    Z = iColorDepth * Int((bCol - (bCol > (255 - iColorDepth))) / iColorDepth)
    GetColorByte = CByte(Z + (Z > 255))
End Function

