Attribute VB_Name = "basFileCHK"
Option Explicit
Sub Main()
Dim FF As Integer, F As Boolean, I As Long, J As Long
Dim K As Long, L As Long, S As String, E As String
Dim P As String, Histo(256) As Long, ASCII As Long
Dim Lnk As String * 16
Dim mpg As String * 8
Dim mpeg As String * 8
Dim mp3 As String * 6
Dim Office As String * 4
Dim OfficeType As String * 64
Dim Hlp As String * 4
Dim zip As String * 4
Dim Asf As String * 16
Dim B As String * 16
Dim B8 As String * 8
Dim B6 As String * 6
Dim B4 As String * 4
Dim B3 As String * 3
Dim B2 As String * 2
    P$ = App.Path: If Right(P$, 1) <> "\" Then P$ = P$ & "\"
    S$ = Dir(P$ & "FILE????.CHK", vbNormal)
    Asf$ = Chr$(&H30) & Chr$(&H26) & Chr$(&HB2) & Chr$(&H75) & Chr$(&H8E) & Chr$(&H66) & Chr$(&HCF) & Chr$(&H11) & Chr$(&HA6) & Chr$(&HD9) & Chr$(&H0) & Chr$(&HAA) & Chr$(&H0) & Chr$(&H62) & Chr$(&HCE) & Chr$(&H6C)
    Lnk$ = Chr$(&H4C) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1) & Chr$(&H14) & Chr$(&H2) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0)
    mpg$ = Chr$(&H0) & Chr$(&H0) & Chr$(&H1) & Chr$(&HB3) & Chr$(&H21) & Chr$(&H0) & Chr$(&H1) & Chr$(&H0)
    mpeg$ = Chr$(&H0) & Chr$(&H0) & Chr$(&H1) & Chr$(&HBA) & Chr$(&H21) & Chr$(&H0) & Chr$(&H1) & Chr$(&H0)
    mp3$ = Chr$(&HFF) & Chr$(&HFB) & Chr$(&HD0) & Chr$(&H4) & Chr$(&H0) & Chr$(&H0)
    Office$ = Chr$(&HD0) & Chr$(&HCF) & Chr$(&H11) & Chr$(&HE0)
    Hlp$ = "?_" & Chr$(&H3) & Chr$(&H0)
    zip$ = "PK" & Chr$(3) & Chr$(4)
    
    Do Until Len(S$) = 0
        FF% = FreeFile
        Open P$ & S$ For Binary As FF%
        If LOF(FF%) > 16 Then
            Get FF%, , B$
            B8$ = Left$(B$, 8)
            B6$ = Left$(B8$, 6)
            B4$ = Left$(B6$, 4)
            B3$ = Left$(B4$, 3)
            B2$ = Left$(B3$, 2)
            Select Case True
                Case B2$ = "MM": E$ = "3ds": Rem 3d Studio
                Case B2$ = "II": E$ = "tif"
                Case B2$ = "MZ": E$ = "exe"
                Case B2$ = "BM": E$ = "bmp"
                Case B3$ = "FWS": E$ = "swf": Rem Macromedia Shockwave
                Case B4$ = "8BPS": E$ = "psd": Rem Adobe Photoshop
                Case B4$ = "%!PS": E$ = "ai": Rem Adobe Illustrator
                Case B4$ = "GIF8": E$ = "gif"
                Case B4$ = "!BDN": E$ = "pst": Rem Office Outlook personal folder
                Case B4$ = "MSCF": E$ = "cab": Rem Microsoft Cabinet file
                Case B4$ = "Rar!": E$ = "rar": Rem RAR Archive
                Case B4$ = "ITSF": E$ = "chm": Rem compiled Help File
                Case B4$ = "MThd": E$ = "mid": Rem MIDI
                Case B4$ = "%PDF": E$ = "pdf": Rem Adobe Acrobat
                Case B4$ = zip$: E$ = "zip": Rem PK/WinZip Archive
                Case B4$ = Hlp$: E$ = "hlp"
                Case B6$ = mp3$: E$ = "mp3"
                Case B6$ = "AC1015": E$ = "dwg": Rem AutoCad Drawing
                Case B8$ = mpg$: E$ = "mpg"
                Case B8$ = mpeg$: E$ = "mpeg"
                Case B$ = Asf$: E$ = "asf": Rem Windows Media
                Case B$ = Lnk$: E$ = "lnk"
                Case B$ = "[InternetShortcu": E$ = "url"
                Rem Case B$ = "<!DOCTYPE HTML P": E$ = "html"
                Rem Case UCase$(Trim$(B4$)) = "<HTM": E$ = "htm"
                Case InStr(UCase$(B$), "HTML") > 0: E$ = "htm"
                Case Mid$(B$, 7, 4) = "JFIF": E$ = "jpg"
                Case Right$(B$, 4) = "DSIG": E$ = "ttf"
                Case Right$(B$, 12) = "Standard Jet": E$ = "mdb": Rem Microsoft Access Database
                Case Mid$(B$, 5, 4) = "moov": Rem Apple QuickTime Movie
                    If Right$(B$, 5) = "lmvhd" Then E$ = "mov" Else E$ = "qt"
                Case B4$ = Office$
                    Get FF%, &H820, OfficeType$
                    OfficeType$ = LCase$(OfficeType$)
                    Select Case True
                    Case InStr(OfficeType$, "word"): E$ = "doc"
                    Case InStr(OfficeType$, "excel"): E$ = "xls"
                    Case Else: Rem E$ = "ppt"
                    End Select
                Case B4$ = "RIFF"
                    B4$ = Mid$(B$, 9, 4)
                    Select Case True
                    Case B4$ = "RMID": E$ = "rmi": Rem MIDI
                    Case B4$ = "WAVE": E$ = "wav"
                    Case B4$ = "AVI ": E$ = "avi"
                    Case B4$ = "CDR8": E$ = "cdr": Rem CorelDraw
                    Case B4$ = "CDR9": E$ = "cdr": Rem CorelDraw
                    End Select
                Case LOF(FF%) < &H100000
                    Do
                        For J& = 1 To Len(B$)
                            K& = Asc(Mid$(B$, J&, 1))
                            Histo&(K&) = Histo&(K&) + 1
                        Next
                        If EOF(FF%) Then Exit Do
                        Get FF%, , B$
                    Loop
                    ASCII& = 0
                    For J& = LBound(Histo&) To UBound(Histo&)
                        Select Case J&
                        Case 9, 10, 13, 32 To 126: ASCII& = ASCII& + Histo&(J&)
                        End Select
                        Histo&(J&) = 0
                    Next
                    If ASCII& > 0 Then If LOF(FF%) / ASCII& < 1.25 Then E$ = "txt"
            End Select: Rem
        End If
        Close
        If Len(E$) > 0 Then Name P$ & S$ As Left$(S$, 9) & E$: E$ = vbNullString
        Do
            S$ = Dir
            If Len(S$) = 0 Then Exit Do
            DoEvents
        Loop Until LCase(Left$(S$, 4)) = "file" And LCase(Right$(S$, 4)) = ".chk"
    Loop
    MsgBox "Done!", vbOKOnly Or vbMsgBoxSetForeground, "FILECHecK"
End Sub


