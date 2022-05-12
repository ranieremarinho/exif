Attribute VB_Name = "ExifModule"
Option Explicit


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Dim ByteFormat As Byte

Dim EXF As New clsEXIF

Public Function GetHoraExif(ByVal CaminhoCompletodaFoto As String) As String

    Dim Hora As String
    
    Hora = Right(LeExif(CaminhoCompletodaFoto), 9)
    '14h25m37s
    Hora = Left(Hora, 2) & ":" & Mid(Hora, 4, 2) & ":" & Mid(Hora, 7, 2)
    
    GetHoraExif = Hora
    
End Function

Public Function GetDataExif_DDMMAA(ByVal CaminhoCompletodaFoto As String, ByVal DataComBarra As Boolean) As String

    Dim Data_DDMMAA As String
    Dim Data_AAMMDD As String
    
    Data_AAMMDD = Left(LeExif(CaminhoCompletodaFoto), 6)
    
    If DataComBarra = True Then
        Data_DDMMAA = Right(Data_AAMMDD, 2) & "/" & Mid(Data_AAMMDD, 3, 2) & "/" & Left(Data_AAMMDD, 2)
    Else
        Data_DDMMAA = Right(Data_AAMMDD, 2) & Mid(Data_AAMMDD, 3, 2) & Left(Data_AAMMDD, 2)
    End If
    
    GetDataExif_DDMMAA = Data_DDMMAA
    
End Function

Public Function GetAllInfo(Caminho As String) As String

    EXF.ImageFile = Caminho    'set the image file property
    GetAllInfo = Trim(EXF.ListInfo)    'list all tags into the text box

End Function

Public Function GetData_e_HoraExif(ByVal CaminhoCompletodaFoto As String, ByRef DataExif As String, ByRef HoraExif As String, ByVal DataComBarra As Boolean) As Boolean

    Dim Hora As String
    Dim Data_DDMMAA As String
    Dim Data_AAMMDD As String
    Dim InfoExif As String
    
    InfoExif = LeExif(CaminhoCompletodaFoto)
    
    If Trim(InfoExif) = "" Then
        DataExif = "Sem informações"
        HoraExif = "Sem informações"
    Else
        Data_AAMMDD = Left(InfoExif, 6)
        If DataComBarra = True Then
            Data_DDMMAA = Right(Data_AAMMDD, 2) & "/" & Mid(Data_AAMMDD, 3, 2) & "/" & Left(Data_AAMMDD, 2)
        Else
            Data_DDMMAA = Right(Data_AAMMDD, 2) & Mid(Data_AAMMDD, 3, 2) & Left(Data_AAMMDD, 2)
        End If
        DataExif = Data_DDMMAA
    
        Hora = Right(InfoExif, 9)
        '14h25m37s
        Hora = Left(Hora, 2) & ":" & Mid(Hora, 4, 2) & ":" & Mid(Hora, 7, 2)
        
        HoraExif = Hora
    End If

End Function

Public Function LeExif(Caminho As String) As String
    
    Dim Cmds As String
    Dim AppDataLen As Integer
    Dim ExifDataChunk As String
    Dim FID As Long
    Dim NoOfDirEntries As Long
    Dim DirEntryInfo As String
    Dim SizeMultiplier As Long
    Dim LenOfTagData As Long
    Dim NxtIFDO As Long
    Dim DataFormat As Long
    Dim tmpStr As String
    Dim NxtExifChunk As Long
    Dim Data_AAMMDD As String
    Dim I As Long
    Dim TagName As String
    Dim ExifResult As String
    Dim AnoStr As String
    Dim MesStr As String
    Dim DiaStr As String
    Dim hhStr As String
    Dim mmStr As String
    Dim ssStr As String
    
    Close #1
    Data_AAMMDD = ""
    EXF.ImageFile = Caminho    'set the image file property

    ExifResult = Trim(EXF.ListInfo)  'list all tags in this variable
    I = InStr(ExifResult, "DateTime: ")
    If I = 0 Then I = InStr(ExifResult, "DateTimeOriginal: ")
    If I = 0 Then I = InStr(ExifResult, "DateTimeDigitized: ")

    If I <> 0 Then
        Data_AAMMDD = Mid(ExifResult, I, InStr(I, ExifResult, vbCrLf) - I)
        Data_AAMMDD = Replace(Data_AAMMDD, "DateTime: ", "")
        Data_AAMMDD = Replace(Data_AAMMDD, "DateTimeOriginal: ", "")
        Data_AAMMDD = Replace(Data_AAMMDD, "DateTimeDigitized: ", "")
    End If

    If Data_AAMMDD = "" Then
        Open Caminho For Binary Access Read As #1
        Cmds = Input(2, 1)

        Cmds = Input(2, 1)
        AppDataLen = B2D(Input(2, 1))

        Cmds = Input(6, 1)

        ExifDataChunk = Input(AppDataLen, 1)
        Select Case Mid$(ExifDataChunk, 1, 2)
            Case H2B("4949"): ByteFormat = 0    ' Reverse bytes "Intel Header Format":
            Case H2B("4D4D"): ByteFormat = 1    '"Motarola Header Format - Might have probs"
        End Select
        FID = B2D(Rev(Mid$(ExifDataChunk, 5, 4)))    'Image File Directory Offset = 8
        NoOfDirEntries = B2D(Rev(Mid$(ExifDataChunk, 9, 2)))
        For I = 0 To 100    'NoOfDirEntries - 1
            DoEvents
            GirosBusca = GirosBusca + 1
            DirEntryInfo = Mid$(ExifDataChunk, (I * 12) + 11, 12)
            TagName = GetTagName(Rev(Mid$(DirEntryInfo, 1, 2)))
            DataFormat = B2D(Rev(Mid$(DirEntryInfo, 3, 2)))    ' Byte, Single, Long...
            SizeMultiplier = B2D(Rev(Mid$(DirEntryInfo, 5, 4)))
            LenOfTagData = CLng(TypeOfTag(DataFormat)) * SizeMultiplier

            If TagName = "ExifOffset" Then NxtExifChunk = ConvertData2Format(DataFormat, Rev(Right$(Mid$(DirEntryInfo, 9, 4), LenOfTagData)))
            If LenOfTagData <= 4 Then    ' No Offset
                If TagName = "DateTimeOriginal" Then
                    Data_AAMMDD = ConvertData2Format(DataFormat, Rev(Right$(Mid$(DirEntryInfo, 9, 4), LenOfTagData)))
                    Exit For
                End If
            Else    ' Offset required
                tmpStr = Mid$(ExifDataChunk, B2D(Rev(Mid$(DirEntryInfo, 9, 4))) + 1, LenOfTagData)
                If TagName = "DateTimeOriginal" Then
                    Data_AAMMDD = ConvertData2Format(DataFormat, tmpStr)
                    Exit For
                End If
            End If
        Next I

        NxtIFDO = B2D(Rev(Mid$(ExifDataChunk, (I * 12) + 11, 4)))

        NoOfDirEntries = B2D(Rev(Mid$(ExifDataChunk, NxtExifChunk + 1, 2)))


        For I = 0 To NoOfDirEntries - 1
            DoEvents
            GirosBusca = GirosBusca + 1

            DirEntryInfo = Mid$(ExifDataChunk, (I * 12) + NxtExifChunk + 11 + 4, 12)

            TagName = GetTagName(Rev(Mid$(DirEntryInfo, 1, 2)))

            DataFormat = B2D(Rev(Mid$(DirEntryInfo, 3, 2)))    ' Byte, Single, Long...
            SizeMultiplier = B2D(Rev(Mid$(DirEntryInfo, 5, 4)))
            LenOfTagData = CLng(TypeOfTag(DataFormat)) * SizeMultiplier

            If LenOfTagData <= 4 Then    ' No Offset
                If TagName = "DateTimeOriginal" Then
                    Data_AAMMDD = ConvertData2Format(DataFormat, Rev(Right$(Mid$(DirEntryInfo, 9, 4), LenOfTagData)))
                    Exit For
                End If
            Else    ' Offset required
                tmpStr = Mid$(ExifDataChunk, B2D(Rev(Mid$(DirEntryInfo, 9, 4))) + 1, LenOfTagData)
                If TagName = "DateTimeOriginal" Then
                    Data_AAMMDD = ConvertData2Format(DataFormat, tmpStr)
                    Exit For
                End If
            End If
        Next I

        If Data_AAMMDD = "" Then
            For I = 0 To 100
            GirosBusca = GirosBusca + 1
                DoEvents

                DirEntryInfo = Mid$(ExifDataChunk, (I * 12) + 11, 12)
                TagName = GetTagName(Rev(Mid$(DirEntryInfo, 1, 2)))
                DataFormat = B2D(Rev(Mid$(DirEntryInfo, 3, 2)))    ' Byte, Single, Long...
                SizeMultiplier = B2D(Rev(Mid$(DirEntryInfo, 5, 4)))
                LenOfTagData = CLng(TypeOfTag(DataFormat)) * SizeMultiplier

                If TagName = "ExifOffset" Then NxtExifChunk = ConvertData2Format(DataFormat, Rev(Right$(Mid$(DirEntryInfo, 9, 4), LenOfTagData)))
                If LenOfTagData <= 4 Then    ' No Offset
                    If TagName = "DateTime" Then
                        Data_AAMMDD = ConvertData2Format(DataFormat, Rev(Right$(Mid$(DirEntryInfo, 9, 4), LenOfTagData)))
                        Exit For
                    End If
                Else    ' Offset required
                    tmpStr = Mid$(ExifDataChunk, B2D(Rev(Mid$(DirEntryInfo, 9, 4))) + 1, LenOfTagData)
                    If TagName = "DateTime" Then
                        Data_AAMMDD = ConvertData2Format(DataFormat, tmpStr)
                        Exit For
                    End If
                End If
            Next I
        End If
    End If

    '06:05:12 10:40:32
    '006:05:12 10:40:32
    '2006:05:12 10:40:32
    Select Case InStr(Data_AAMMDD, ":")
        Case 3: Data_AAMMDD = Left(Data_AAMMDD, 17)
        Case 4: Data_AAMMDD = Left(Data_AAMMDD, 18)
        Case 5: Data_AAMMDD = Left(Data_AAMMDD, 19)
    End Select

    Data_AAMMDD = Replace(Data_AAMMDD, ":", "")

    Data_AAMMDD = Right(Data_AAMMDD, 13)
    Data_AAMMDD = Left(Data_AAMMDD, 9) & "h" & Mid(Data_AAMMDD, 10, 2) & "m" & Right(Data_AAMMDD, 2) & "s"
    If Data_AAMMDD = "hms" Then Data_AAMMDD = ""
    
    If Len(Data_AAMMDD) >= 16 Then '090710 20h47m47s.jpg
        AnoStr = Left(Data_AAMMDD, 2)
        MesStr = Mid(Data_AAMMDD, 3, 2)
        DiaStr = Mid(Data_AAMMDD, 5, 2)
        hhStr = Mid(Data_AAMMDD, 8, 2)
        mmStr = Mid(Data_AAMMDD, 11, 2)
        ssStr = Mid(Data_AAMMDD, 14, 2)
        If ValidaData_com_Hora(AnoStr, MesStr, DiaStr, hhStr, mmStr, ssStr) = False Then
            LeExif = ""
        Else
            LeExif = Data_AAMMDD
        End If
    Else
        LeExif = ""
    End If
    
    Close #1

End Function

Private Function GetTagName(TagNum As String) As String
    Select Case TagNum
        Case H2B("010E"): GetTagName = "ImageDescription"
        Case H2B("010F"): GetTagName = "Make"
        Case H2B("0110"): GetTagName = "Model"
        Case H2B("0112"): GetTagName = "Orientation"
        Case H2B("011A"): GetTagName = "XResolution"
        Case H2B("011B"): GetTagName = "YResolution"
        Case H2B("0128"): GetTagName = "ResolutionUnit"
        Case H2B("0131"): GetTagName = "Software"
        Case H2B("0132"): GetTagName = "DateTime"
        Case H2B("013E"): GetTagName = "WhitePoint"
        Case H2B("013F"): GetTagName = "PrimaryChromaticities"
        Case H2B("0211"): GetTagName = "YCbCrCoefficients"
        Case H2B("0213"): GetTagName = "YCbCrPositioning"
        Case H2B("0214"): GetTagName = "ReferenceBlackWhite"
        Case H2B("8298"): GetTagName = "Copyright"
        Case H2B("8769"): GetTagName = "ExifOffset"

        Case H2B("829A"): GetTagName = "ExposureTime"
        Case H2B("829D"): GetTagName = "FNumber"
        Case H2B("8822"): GetTagName = "ExposureProgram"
        Case H2B("8827"): GetTagName = "ISOSpeedRatings"
        Case H2B("9000"): GetTagName = "ExifVersion"
        Case H2B("9003"): GetTagName = "DateTimeOriginal"
        Case H2B("9004"): GetTagName = "DateTimeDigitized"
        Case H2B("9101"): GetTagName = "ComponentConfiguration"
        Case H2B("9102"): GetTagName = "CompressedBitsPerPixel"
        Case H2B("9201"): GetTagName = "ShutterSpeedValue"
        Case H2B("9202"): GetTagName = "ApertureValue"
        Case H2B("9203"): GetTagName = "BrightnessValue"
        Case H2B("9204"): GetTagName = "ExposureBiasValue"
        Case H2B("9205"): GetTagName = "MaxApertureValue"
        Case H2B("9206"): GetTagName = "SubjectDistance"
        Case H2B("9207"): GetTagName = "MeteringMode"
        Case H2B("9208"): GetTagName = "LightSource"
        Case H2B("9209"): GetTagName = "Flash"
        Case H2B("920A"): GetTagName = "FocalLength"
        Case H2B("927C"): GetTagName = "MakerNote"    ': Stop
        Case H2B("9286"): GetTagName = "UserComment"
        Case H2B("A000"): GetTagName = "FlashPixVersion"
        Case H2B("A001"): GetTagName = "ColorSpace"
        Case H2B("A002"): GetTagName = "ExifImageWidth"
        Case H2B("A003"): GetTagName = "ExifImageHeight"
        Case H2B("A004"): GetTagName = "RelatedSoundFile"
        Case H2B("A005"): GetTagName = "ExifInteroperabilityOffset"
        Case H2B("A20E"): GetTagName = "FocalPlaneXResolution"
        Case H2B("A20F"): GetTagName = "FocalPlaneYResolution"
        Case H2B("A210"): GetTagName = "FocalPlaneResolutionUnit"
        Case H2B("A217"): GetTagName = "SensingMethod"
        Case H2B("A300"): GetTagName = "FileSource"
        Case H2B("A301"): GetTagName = "SceneType"

        Case H2B("0100"): GetTagName = "ImageWidth"
        Case H2B("0101"): GetTagName = "ImageLength"
        Case H2B("0102"): GetTagName = "BitsPerSample"
        Case H2B("0103"): GetTagName = "Compression"
        Case H2B("0106"): GetTagName = "PhotometricInterpretation"
        Case H2B("0111"): GetTagName = "StripOffsets"
        Case H2B("0115"): GetTagName = "SamplesPerPixel"
        Case H2B("0116"): GetTagName = "RowsPerStrip"
        Case H2B("0117"): GetTagName = "StripByteConunts"
        Case H2B("011A"): GetTagName = "XResolution"
        Case H2B("011B"): GetTagName = "YResolution"
        Case H2B("011C"): GetTagName = "PlanarConfiguration"
        Case H2B("0128"): GetTagName = "ResolutionUnit"
        Case H2B("0201"): GetTagName = "JpegIFOffset"
        Case H2B("0202"): GetTagName = "JpegIFByteCount"
        Case H2B("0211"): GetTagName = "YCbCrCoefficients"
        Case H2B("0212"): GetTagName = "YCbCrSubSampling"
        Case H2B("0213"): GetTagName = "YCbCrPositioning"
        Case H2B("0214"): GetTagName = "ReferenceBlackWhite"

        Case H2B("00FE"): GetTagName = "NewSubfileType"
        Case H2B("00FF"): GetTagName = "SubfileType"
        Case H2B("012D"): GetTagName = "TransferFunction"
        Case H2B("013B"): GetTagName = "Artist"
        Case H2B("013D"): GetTagName = "Predictor"
        Case H2B("0142"): GetTagName = "TileWidth"
        Case H2B("0143"): GetTagName = "TileLength"
        Case H2B("0144"): GetTagName = "TileOffsets"
        Case H2B("0145"): GetTagName = "TileByteCounts"
        Case H2B("014A"): GetTagName = "SubIFDs"
        Case H2B("015B"): GetTagName = "JPEGTables"
        Case H2B("828D"): GetTagName = "CFARepeatPatternDim"
        Case H2B("828E"): GetTagName = "CFAPattern"
        Case H2B("828F"): GetTagName = "BatteryLevel"
        Case H2B("83BB"): GetTagName = "IPTC/NAA"
        Case H2B("8773"): GetTagName = "InterColorProfile"
        Case H2B("8824"): GetTagName = "SpectralSensitivity"
        Case H2B("8825"): GetTagName = "GPSInfo"
        Case H2B("8828"): GetTagName = "OECF"
        Case H2B("8829"): GetTagName = "Interlace"
        Case H2B("882A"): GetTagName = "TimeZoneOffset"
        Case H2B("882B"): GetTagName = "SelfTimerMode"
        Case H2B("920B"): GetTagName = "FlashEnergy"
        Case H2B("920C"): GetTagName = "SpatialFrequencyResponse"
        Case H2B("920D"): GetTagName = "Noise"
        Case H2B("9211"): GetTagName = "ImageNumber"
        Case H2B("9212"): GetTagName = "SecurityClassification"
        Case H2B("9213"): GetTagName = "ImageHistory"
        Case H2B("9214"): GetTagName = "SubjectLocation"
        Case H2B("9215"): GetTagName = "ExposureIndex"
        Case H2B("9216"): GetTagName = "TIFF/EPStandardID"
        Case H2B("9290"): GetTagName = "SubSecTime"
        Case H2B("9291"): GetTagName = "SubSecTimeOriginal"
        Case H2B("9292"): GetTagName = "SubSecTimeDigitized"
        Case H2B("A20B"): GetTagName = "FlashEnergy"
        Case H2B("A20C"): GetTagName = "SpatialFrequencyResponse"
        Case H2B("A214"): GetTagName = "SubjectLocation"
        Case H2B("A215"): GetTagName = "ExposureIndex"
        Case H2B("A302"): GetTagName = "CFAPattern"

        Case H2B("0200"): GetTagName = "SpecialMode"
        Case H2B("0201"): GetTagName = "JpegQual"
        Case H2B("0202"): GetTagName = "Macro"
        Case H2B("0203"): GetTagName = "Unknown"
        Case H2B("0204"): GetTagName = "DigiZoom"
        Case H2B("0205"): GetTagName = "Unknown"
        Case H2B("0206"): GetTagName = "Unknown"
        Case H2B("0207"): GetTagName = "SoftwareRelease"
        Case H2B("0208"): GetTagName = "PictInfo"
        Case H2B("0209"): GetTagName = "CameraID"
        Case H2B("0F00"): GetTagName = "DataDump"
            'Case H2B(""): GetTagName = ""
        Case Else: GetTagName = "Unknown"
    End Select
    
End Function

Private Function H2B(InHex As String) As String
    ' Conv Hex to Bytes
    
    Dim I As Long

    For I = 1 To Len(InHex) Step 2
        H2B = H2B & Chr$(CLng("&H" & Mid$(InHex, I, 2)))
        DoEvents
        GirosBusca = GirosBusca + 1
    Next I
    
End Function

Private Function B2D(InBytes As String) As Double
    ' Conv. Bytes to Decimal - Could be > 4 Billion
    
    On Error Resume Next
    
    Dim I As Long
    Dim TMP As String

    For I = 1 To Len(InBytes)
        TMP = TMP & Hex(Format$(Asc(Mid$(InBytes, I, 1)), "00"))
        DoEvents
        GirosBusca = GirosBusca + 1
    Next I
    B2D = "&H" & TMP
    
End Function

Public Function TemExif(Caminho As String) As Boolean

    TemExif = (LeExif(Caminho) <> "")

End Function

Private Function Rev(InBytes As String) As String ' Reverse bytes

    If ByteFormat = 1 Then Exit Function    ' Not needed for Motorola format

    Dim I As Long
    Dim TMP As String

    For I = Len(InBytes) To 1 Step -1
        TMP = TMP & Mid$(InBytes, I, 1)
        DoEvents
        GirosBusca = GirosBusca + 1
    Next I
    Rev = TMP
    
End Function

Private Function TypeOfTag(InDec As Long) As Byte
    
    Select Case InDec
        Case 1: TypeOfTag = 1
        Case 2: TypeOfTag = 1
        Case 3: TypeOfTag = 2
        Case 4: TypeOfTag = 4
        Case 5: TypeOfTag = 8
        Case 6: TypeOfTag = 1
        Case 7: TypeOfTag = 1
        Case 8: TypeOfTag = 2
        Case 9: TypeOfTag = 4
        Case 10: TypeOfTag = 8
        Case 11: TypeOfTag = 4
        Case 12: TypeOfTag = 8
    End Select
    
End Function

Private Function ConvertData2Format(DataFormat As Long, InBytes As String) As String

    ' Read function aboves details
    ' Double check for Motorola format esp. CopyMemory
    
    Dim tmpInt As Integer
    Dim tmpLng As Long
    Dim tmpSng As Single
    Dim tmpDbl As Double
    Dim tmpVal As Long    'inserido por mim
    Dim Convert As Long    'inserido por mim
GirosBusca = GirosBusca + 1
    Select Case DataFormat
        Case 1, 3, 4: ConvertData2Format = B2D(InBytes)
        Case 2, 7: ConvertData2Format = InBytes
        Case 5    ' Kinda Unsigned Fraction
            '        ConvertData2Format = CDbl(B2D(Mid$(InBytes, 1, 4))) / CDbl(B2D(Mid$(InBytes, 5, 4)))
        Case 6
            tmpVal = B2D(InBytes)
            If tmpVal > 127 Then ConvertData2Format = -(tmpVal - 127) Else Convert = tmpVal
        Case 8
            'tmpVal = B2D(InBytes)
            'If tmpVal > 32767 Then ConvertData2Format = -(tmpVal - 32767) Else ConvertData2Format = tmpVal
            CopyMemory tmpInt, InBytes, 2
            ConvertData2Format = tmpInt
        Case 9
            CopyMemory tmpLng, InBytes, 4
            ConvertData2Format = tmpLng
        Case 10    ' Kinda Signed Fraction (Lens Apeture?)
            CopyMemory tmpLng, Mid$(InBytes, 1, 4), 4
            ConvertData2Format = tmpLng
            CopyMemory tmpLng, Mid$(InBytes, 5, 4), 4
            ConvertData2Format = ConvertData2Format / tmpLng
        Case 11
            CopyMemory tmpSng, InBytes, 4
            ConvertData2Format = tmpSng
        Case 12
            CopyMemory tmpDbl, InBytes, 8
            Convert = tmpDbl
    End Select
    
End Function
