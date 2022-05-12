VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEXIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Type tIDF 'main structure, containg tag parameters
    ID As Long
    Name As String
    Format As Long
    Length As Long
    Data As Variant
    StringData As String
End Type

Public Enum enTAG 'enumerator containing the most common metatag names and their respective IDs
    ImageDescription = 270
    Make = 271
    Model = 272
    Orientation = 274
    XResolution = 282
    YResolution = 283
    ResolutionUnit = 296
    Software = 305
    DateTime = 306
    WhitePoint = 318
    PrimaryChromaticities = 319
    YCbCrCoefficients = 529
    YCbCrPositioning = 531
    ReferenceBlackWhite = 532
    Copyright = 33432
    ExifOffset = 34665
    Exposuretime = 33434
    FNumber = 33437
    ExposureProgram = 34850
    ISOSpeedRatings = 34855
    ExifVersion = 36864
    DateTimeOriginal = 36867
    DateTimeDigitized = 36868
    ComponentsConfiguration = 37121
    CompressedBitsPerPixel = 37122
    ShutterSpeedValue = 37377
    ApertureValue = 37378
    BrightnessValue = 37379
    ExposureBiasValue = 37380
    MaxApertureValue = 37381
    SubjectDistance = 37382
    MeteringMode = 37383
    LightSource = 37384
    Flash = 37385
    FocalLength = 37386
    MakerNote = 37500
    UserComment = 37510
    SubsecTime = 37520
    SubsecTimeOriginal = 37521
    SubsecTimeDigitized = 37522
    FlashPixVersion = 40960
    ColorSpace = 40961
    ExifImageWidth = 40962
    ExifImageHeight = 40963
    RelatedSoundFile = 40964
    ExifInteroperabilityOffset = 40965
    FocalPlaneXResolution = 41486
    FocalPlaneYResolution = 41487
    FocalPlaneResolutionUnit = 41488
    ExposureIndex = 41493
    SensingMethod = 41495
    FileSource = 41728
    SceneType = 41729
    CFAPattern = 41730
End Enum

Private Enum enFormat
    unsignedByte = 1
    asciiString = 2
    unsignedShort = 3
    unsignedLong = 4
    unsignedRationale = 5
    signedByte = 6
    undefined = 7
    signedShort = 8
    signedLong = 9
    signedRationale = 10
    singleFloat = 11
    doubleFloat = 12
End Enum

Private m_ImageFile As String    'path to the file on the disk
Private m_Intel As Boolean    'flag containing the byte alignment of the record
Private m_IniFile As String    'path to the exif.ini file containing description of the codes

Private DataLen    'array containing the byte length of each data format
Private IDName    'array containing tag names
Private IDNo    'array containing tag IDs
Private bParsed As Byte    'flag that the file has been parsed succesfuly
Private IFD() As tIDF    'array containing all metatags
'======== Interface part ==========

Public Property Get ImageFile() As Variant
    '-- file containing the jpeg image --
    
    ImageFile = m_ImageFile
    
End Property

Public Property Let ImageFile(ByVal vNewValue As Variant)
    
    m_ImageFile = vNewValue
    bParsed = ReadMetaInfo(m_ImageFile)

End Property

Public Property Get IntelByteAlignment() As Boolean
'-- returns the byte alignment order for the file --

    IntelByteAlignment = m_Intel
    
End Property


Public Function MetaInfo(ByVal l_ID As enTAG, ByRef StringData As String) As Long
    '-- method, which returns the numeric and string values for a single metatag ---
    
    Dim I As Integer
    
    If bParsed = 0 Then
        For I = 1 To UBound(IFD)
            DoEvents
            If IFD(I).ID = l_ID Then
                StringData = IFD(I).StringData
                MetaInfo = IFD(I).Data
                Exit Function
            End If
        Next I
        StringData = "Tag " + CStr(l_ID) + " not found."
    ElseIf bParsed = 1 Then
        Err.Raise 10, "MetaInfo", "Sem informações EXIF."
    ElseIf bParsed = 2 Then
        Err.Raise 11, "MetaInfo", "Erro ao analisar arquivo."
    End If
    
End Function

'-- returns Exif tag name based on its ID --
Private Function GetTagName(ByVal lID As enTAG) As String
    
    Dim I As Integer
    
    For I = 0 To UBound(IDNo)
        DoEvents
        If lID = IDNo(I) Then
            GetTagName = IDName(I)
            Exit Function
        End If
    Next
    
    'tag name unknown; return tag ID
    GetTagName = "Tag #" + CStr(lID)
    
End Function

Function ListInfo() As String
    '--- Lists all metatags found in the header ---
    
    On Error Resume Next
    Dim I As Integer
    
    If bParsed = 0 Then
        For I = 1 To UBound(IFD)
            DoEvents
            ListInfo = ListInfo + IFD(I).Name + ": " + IFD(I).StringData + vbCrLf
        Next I
    ElseIf bParsed = 1 Then
        ListInfo = "Sem informações EXIF."
    ElseIf bParsed = 2 Then
        ListInfo = "Impossível carregar arquivo."
    End If
    
End Function

'========= PARSING FUNCTIONS ==========

Private Function ReadMetaInfo(sFileName As String) As Integer
'-- parses the jpeg header and extracts all Exif information from it --

    Dim sJPEG_Header As String, b() As Byte
    Dim lPos As Long, Offset As Long, HeaderStart As Long
    Dim I As Integer
    Dim NoOfRecs As Integer

    On Error Resume Next
    
    sJPEG_Header = ReadFile(sFileName, 4096)    'may be changed to reflect the actual header size
    If sJPEG_Header = "" Then
        Err.Raise 2, "ReadMetaInfo", "Arquivo não encontrado."
    Else
        HeaderStart = InStr(1, sJPEG_Header, "Exif" + Chr(0), vbBinaryCompare)    'start of EXIF header
        If HeaderStart = 0 Then ReadMetaInfo = 1: Exit Function
        HeaderStart = HeaderStart + 6    'start of data
        lPos = HeaderStart
        m_Intel = Mid(sJPEG_Header, lPos, 2) = "II"    'byte alignment
        lPos = lPos + 4
        Offset = BVal(sJPEG_Header, lPos, 4, m_Intel)    'offset to the first IFD
        ReDim IFD(0)
        'parse the main IFD directory and get the offset to the IFDSubDir
        Offset = ParseIDF(sJPEG_Header, HeaderStart, Offset, ExifOffset)
        Offset = ParseIDF(sJPEG_Header, HeaderStart, Offset, ExifInteroperabilityOffset)
    End If

ExitRead:
    Exit Function

ErrRead:
    Dim S
    S = Err.Description
    Err.Raise 1, "ReadMetaInfo", S
    ReadMetaInfo = 2

End Function

'-- returns formatted string data from an IFD entry ---
Private Sub GetStringData(sJPEG_Header As String, IFD As tIDF, OffsetStart As Long, CurrPosition As Long)
    
    Dim vData As String, X As Variant, sFmt As String
    Dim PixOrder, I As Integer
    
    'extract the data
    If IFD.Length * DataLen(IFD.Format) > 4 Then    'IFD data contains offset to real data
        vData = Mid(sJPEG_Header, OffsetStart + IFD.Data, IFD.Length * DataLen(IFD.Format))
    Else    'IFD record contains the data itself
        vData = Mid(sJPEG_Header, CurrPosition + 8, 4)
    End If
    
    'check the data format and do some conversions if necessary
    Select Case IFD.Format
        Case enFormat.asciiString    'string
            IFD.StringData = vData
        Case enFormat.unsignedRationale    'unsigned rational
            Select Case IFD.ID
                Case enTAG.ApertureValue, enTAG.MaxApertureValue
                    'convert to aperture ratio value
                    X = UnsignedRational(vData, IFD.Data)
                    IFD.StringData = Format(Sqr(2) ^ IFD.Data, "0.0")
                Case enTAG.FNumber
                    X = UnsignedRational(vData, IFD.Data)
                    IFD.StringData = Format(IFD.Data, "0.0")
                Case enTAG.FocalPlaneXResolution, enTAG.FocalPlaneYResolution, enTAG.XResolution, enTAG.YResolution, enTAG.FocalLength, enTAG.SubjectDistance
                    X = UnsignedRational(vData, IFD.Data)
                    IFD.StringData = Format(IFD.Data, "0.0")
                Case enTAG.Exposuretime
                    X = UnsignedRational(vData, IFD.Data)
                    IFD.StringData = Format(1 / IFD.Data, "0")
                Case Else
                    If IFD.Length < 2 Then
                        IFD.StringData = UnsignedRational(vData, IFD.Data)
                    Else
                        For I = 1 To IFD.Length
                            DoEvents
                            Call UnsignedRational(Mid(vData, (I - 1) * 8 + 1, 4), X)
                            IFD.StringData = "/" + IFD.StringData + Format(X, "0.000")
                        Next
                        IFD.StringData = Mid(IFD.StringData, 2)
                    End If
            End Select
        Case enFormat.signedRationale    'signed rational
            Select Case IFD.ID
                Case enTAG.ShutterSpeedValue
                    'convert to shutter speed value
                    X = SignedRational(vData, IFD.Data)
                    IFD.StringData = Format(2 ^ IFD.Data, "0")
                Case enTAG.ExposureBiasValue
                    X = SignedRational(vData, IFD.Data)
                    IFD.StringData = Format(IFD.Data, "0.0")
                Case Else
                    If IFD.Length < 2 Then
                        IFD.StringData = SignedRational(vData, IFD.Data)
                    Else
                        For I = 1 To IFD.Length
                            DoEvents
                            Call SignedRational(Mid(vData, (I - 1) * 8 + 1, 4), X)
                            IFD.StringData = "/" + IFD.StringData + Format(X, "0.000")
                        Next
                        IFD.StringData = Mid(IFD.StringData, 2)
                    End If
            End Select
        Case enFormat.undefined    'undefined
            Select Case IFD.ID
                Case enTAG.MakerNote
                    IFD.StringData = ExtractTextOnly(vData)
                Case enTAG.ComponentsConfiguration
                    PixOrder = Array("", "Y", "Cb", "Cr", "R", "G", "B")
                    For I = 1 To 4
                        DoEvents
                        IFD.StringData = IFD.StringData + PixOrder(Asc(Mid(vData, I, 1)))
                    Next I
                Case enTAG.FileSource, enTAG.SceneType
                    IFD.StringData = CStr(BVal(vData, 1, DataLen(IFD.Format), m_Intel))
                Case enTAG.ExifVersion, enTAG.FlashPixVersion
                    IFD.StringData = CStr(Val(Left(vData, 2))) + "." + Right(vData, 2)
                Case Else
                    IFD.StringData = vData
            End Select
        Case enFormat.signedByte, enFormat.signedLong, enFormat.signedShort    'signed byte, short & long
            IFD.StringData = CStr(BValS(vData, 1, DataLen(IFD.Format), m_Intel))
        Case Else    'all other data types
            IFD.StringData = CStr(BVal(vData, 1, DataLen(IFD.Format), m_Intel))
    End Select
    
    IFD.StringData = Replace(IFD.StringData, Chr(0), "")    'remove null characters
    'try to read description of data from the ini file
    vData = GetProfileString(m_IniFile, IFD.Name, IFD.StringData)
    sFmt = GetProfileString(m_IniFile, IFD.Name, "Format")
    'if description is found, use it instead of raw data
    If vData <> "" Then IFD.StringData = vData
    If sFmt <> "" Then IFD.StringData = Replace(sFmt, "@X", IFD.StringData)

End Sub

Private Function ParseIDF(sJPEG_Header As String, HeaderStart As Long, Offset As Long, ClosingTag As enTAG) As Long
'--- Parses a single EXIF directory record (IFD), returns an offset to the next IFD ---

    Dim lPos As Long, NoOfRecs As Long, I As Integer, u
    
    lPos = HeaderStart + Offset    'get the starting offset position
    NoOfRecs = BVal(sJPEG_Header, lPos, 2, m_Intel)    'get the No of records in the IFD
    u = UBound(IFD)
    ReDim Preserve IFD(u + NoOfRecs)    'redimension the IFD array
    lPos = lPos + 2
    
    'begin retrieving the tags
    For I = u + 1 To u + NoOfRecs
        DoEvents
        IFD(I).ID = BVal(sJPEG_Header, lPos, 2, m_Intel)    'first 2 bytes contain the tag ID
        IFD(I).Format = BVal(sJPEG_Header, lPos + 2, 2, m_Intel)    'next 2 bytes contain the tag data format
        IFD(I).Length = BVal(sJPEG_Header, lPos + 4, 4, m_Intel)    'next 4 bytes contain the No of data components
        IFD(I).Data = BVal(sJPEG_Header, lPos + 8, 4, m_Intel)    'next 4 bytes contain the data or an offset to the data
        IFD(I).Name = GetTagName(IFD(I).ID)    'retrieve the tag name in human readable format
        Call GetStringData(sJPEG_Header, IFD(I), HeaderStart, lPos)    'retrieve the data as a string
        lPos = lPos + 12    'get the offset to the next tag
        If IFD(I).ID = ClosingTag Then
            ParseIDF = Val(IFD(I).StringData)    'return offset to the next sub IFD
        End If
    Next
    
End Function


'======== byte hadling functions ==========

'--- returns calculated unsigned value of a byte sequence ---
Private Function BVal(sData As String, start As Long, ByVal Length As Integer, m_Intel As Boolean) As Variant
    
    On Error Resume Next
    
    Dim I As Long, st As Long, en As Long, step As Integer, N As Long

    If Not m_Intel Then    'Motorola byte alignment
        st = start + Length - 1
        en = start
        step = -1
    Else    'Intel byte alignment
        st = start
        en = start + Length - 1
        step = 1
    End If
    
    For I = st To en Step step
        DoEvents
        BVal = BVal + Asc(Mid(sData, I, 1)) * 256 ^ N
        N = N + 1
    Next

End Function

'--- returns calculated signed value of a byte sequence ---
Private Function BValS(sData As String, start As Long, ByVal Length As Integer, m_Intel As Boolean) As Variant
    
    Dim dn As Variant
    
    dn = 2 ^ (8 * Length - 1)
    BValS = BVal(sData, start, Length, m_Intel)
    If BValS > (dn - 1) Then BValS = Not (BValS - dn)

End Function


Private Function UnsignedRational(ByVal vData As String, ByRef NumValue As Variant) As String
'--- returns a string containing an unsigned rational value in the format Numerator/Denumerator;
'- also calculates its numeric value

    Dim Num As Double, Denum As Double, Rational As Double
    
    Num = BVal(vData, 1, 4, m_Intel)
    Denum = BVal(vData, 5, 4, m_Intel)
    If Denum <> 0 Then NumValue = Num / Denum
    UnsignedRational = CStr(Num) + "/" + CStr(Denum)

End Function


Private Function SignedRational(ByVal vData As String, ByRef NumValue As Variant) As String
'--- returns a string containing a signed rational value in the format Numerator/Denumerator;
'- also calculates its numeric value

    Dim Num As Double, Denum As Double, Rational As Double
    
    Num = BValS(vData, 1, 4, m_Intel)
    Denum = BValS(vData, 5, 4, m_Intel)
    If Denum <> 0 Then NumValue = Num / Denum
    SignedRational = CStr(Num) + "/" + CStr(Denum)

End Function


Private Function ReadFile(ByVal sFilePath As String, Optional iLen = 0) As String
'--- Reads and returns a string from a file on the disk ---

    Dim F As Long
    Dim S As String
    
    On Error Resume Next
    
    If FileLen(sFilePath) < 1 Then
        ReadFile = ""
    Else
        F = FreeFile
        Open sFilePath For Binary Access Read As #F
        If iLen = 0 Then S = Space$(LOF(F)) Else S = Space$(iLen)
        Get #F, , S
        Close #F
        ReadFile = S
        S = ""
    End If
    
End Function

'--- Reads a value from an .INI file ---
Private Function GetProfileString(ByVal sFile As String, ByVal sSection As String, ByVal sKey As String, Optional ByVal DefaultValue = "") As String
    
    Dim sTmp As String, X As Long
    Const StringSize = 1024
    
    sTmp = Space$(StringSize)
    X = GetPrivateProfileString(sSection, sKey, "", sTmp, StringSize, sFile)
    sTmp = Trim$(sTmp)
    sTmp = Left(sTmp, Len(sTmp) - 1)
    If sTmp = "" Then sTmp = DefaultValue
    GetProfileString = sTmp

End Function

'--- Returns only the ascii characters from a byte sequence ---
Private Function ExtractTextOnly(S As String) As String
    
    Dim I As Integer, L As Integer, c As String * 1
    
    L = Len(S)
    
    For I = 1 To L
        c = Mid(S, I, 1)
        If Asc(c) > 31 And Asc(c) < 128 Then ExtractTextOnly = ExtractTextOnly + c
        DoEvents
    Next I
    
End Function

'--- initializes arrays and other variables ---
Private Sub Class_Initialize()
    m_IniFile = App.Path
    If Right(m_IniFile, 1) <> "\" Then m_IniFile = m_IniFile + "\"
    m_IniFile = m_IniFile + "exif.ini"
    DataLen = Array(0, 1, 1, 2, 4, 8, 1, 1, 2, 4, 8, 4, 8)
    IDNo = Array(270, 271, 272, 274, 282, 283, 296, 305, 306, 318, 319, 529, 531, 532, 33432, 34665, 33434, 33437, 34850, 34855, 36864, 36867, 36868, 37121, 37122, 37377, 37378, 37379, 37380, 37381, 37382, 37383, 37384, 37385, 37386, 37500, 37510, 37520, 37521, 37522, 40960, 40961, 40962, 40963, 40964, 40965, 41486, 41487, 41488, 41493, 41495, 41728, 41729, 41730)
    IDName = Array("ImageDescription", "Make", "Model", "Orientation", "XResolution", "YResolution", "ResolutionUnit", "Software", "DateTime", "WhitePoint", "PrimaryChromaticities", "YCbCrCoefficients", "YCbCrPositioning", "ReferenceBlackWhite", "Copyright", "ExifOffset", _
            "ExposureTime", "FNumber", "ExposureProgram", "ISOSpeedRatings", "ExifVersion", "DateTimeOriginal", "DateTimeDigitized", "ComponentsConfiguration", "CompressedBitsPerPixel", "ShutterSpeedValue", "ApertureValue", "BrightnessValue", "ExposureBiasValue", "MaxApertureValue", "SubjectDistance", "MeteringMode", "LightSource", "Flash", "FocalLength", "MakerNote", "UserComment", "SubsecTime", "SubsecTimeOriginal", "SubsecTimeDigitized", "FlashPixVersion", "ColorSpace", "ExifImageWidth", "ExifImageHeight", "RelatedSoundFile", "ExifInteroperabilityOffset", "FocalPlaneXResolution", "FocalPlaneYResolution", "FocalPlaneResolutionUnit", "ExposureIndex", "SensingMethod", "FileSource", "SceneType", "CFAPattern")
End Sub


