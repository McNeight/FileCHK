Option Strict Off
Option Explicit On
Module basFileCHK
	Public Sub Main()
		Dim FF As Short
		Dim F As Boolean
		Dim I, J As Integer
		Dim K, L As Integer
		Dim S, E As String
		Dim P As String
		Dim Histo(256) As Integer
		Dim ASCII As Integer
		Dim Lnk As New VB6.FixedLengthString(16)
		Dim mpg As New VB6.FixedLengthString(8)
		Dim mpeg As New VB6.FixedLengthString(8)
		Dim mp3 As New VB6.FixedLengthString(6)
		Dim Office As New VB6.FixedLengthString(4)
		Dim OfficeType As New VB6.FixedLengthString(64)
		Dim Hlp As New VB6.FixedLengthString(4)
		Dim zip As New VB6.FixedLengthString(4)
		Dim Asf As New VB6.FixedLengthString(16)
		Dim B As New VB6.FixedLengthString(16)
		Dim B8 As New VB6.FixedLengthString(8)
		Dim B6 As New VB6.FixedLengthString(6)
		Dim B4 As New VB6.FixedLengthString(4)
		Dim B3 As New VB6.FixedLengthString(3)
		Dim B2 As New VB6.FixedLengthString(2)
		P = My.Application.Info.DirectoryPath : If Right(P, 1) <> "\" Then P = P & "\"
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		S = Dir(P & "FILE????.CHK", FileAttribute.Normal)
		Asf.Value = Chr(&H30) & Chr(&H26) & Chr(&HB2) & Chr(&H75) & Chr(&H8E) & Chr(&H66) & Chr(&HCF) & Chr(&H11) & Chr(&HA6) & Chr(&HD9) & Chr(&H0) & Chr(&HAA) & Chr(&H0) & Chr(&H62) & Chr(&HCE) & Chr(&H6C)
		Lnk.Value = Chr(&H4C) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H1) & Chr(&H14) & Chr(&H2) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&HC0) & Chr(&H0) & Chr(&H0) & Chr(&H0)
		mpg.Value = Chr(&H0) & Chr(&H0) & Chr(&H1) & Chr(&HB3) & Chr(&H21) & Chr(&H0) & Chr(&H1) & Chr(&H0)
		mpeg.Value = Chr(&H0) & Chr(&H0) & Chr(&H1) & Chr(&HBA) & Chr(&H21) & Chr(&H0) & Chr(&H1) & Chr(&H0)
		mp3.Value = Chr(&HFF) & Chr(&HFB) & Chr(&HD0) & Chr(&H4) & Chr(&H0) & Chr(&H0)
		Office.Value = Chr(&HD0) & Chr(&HCF) & Chr(&H11) & Chr(&HE0)
		Hlp.Value = "?_" & Chr(&H3) & Chr(&H0)
		zip.Value = "PK" & Chr(3) & Chr(4)
		
		Do Until Len(S) = 0
			FF = FreeFile
			FileOpen(FF, P & S, OpenMode.Binary)
			If LOF(FF) > 16 Then
				'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileGet(FF, B.Value)
				B8.Value = Left(B.Value, 8)
				B6.Value = Left(B8.Value, 6)
				B4.Value = Left(B6.Value, 4)
				B3.Value = Left(B4.Value, 3)
				B2.Value = Left(B3.Value, 2)
				Select Case True
					Case B2.Value = "MM" : E = "3ds" ' 3d Studio
					Case B2.Value = "II" : E = "tif"
					Case B2.Value = "MZ" : E = "exe"
					Case B2.Value = "BM" : E = "bmp"
					Case B3.Value = "FWS" : E = "swf" ' Macromedia Shockwave
					Case B4.Value = "8BPS" : E = "psd" ' Adobe Photoshop
					Case B4.Value = "%!PS" : E = "ai" ' Adobe Illustrator
					Case B4.Value = "GIF8" : E = "gif"
					Case B4.Value = "!BDN" : E = "pst" ' Office Outlook personal folder
					Case B4.Value = "MSCF" : E = "cab" ' Microsoft Cabinet file
					Case B4.Value = "Rar!" : E = "rar" ' RAR Archive
					Case B4.Value = "ITSF" : E = "chm" ' compiled Help File
					Case B4.Value = "MThd" : E = "mid" ' MIDI
					Case B4.Value = "%PDF" : E = "pdf" ' Adobe Acrobat
					Case B4.Value = zip.Value : E = "zip" ' PK/WinZip Archive
					Case B4.Value = Hlp.Value : E = "hlp"
					Case B6.Value = mp3.Value : E = "mp3"
					Case B6.Value = "AC1015" : E = "dwg" ' AutoCad Drawing
					Case B8.Value = mpg.Value : E = "mpg"
					Case B8.Value = mpeg.Value : E = "mpeg"
					Case B.Value = Asf.Value : E = "asf" ' Windows Media
					Case B.Value = Lnk.Value : E = "lnk"
					Case B.Value = "[InternetShortcu" : E = "url"
						' Case B$ = "<!DOCTYPE HTML P": E$ = "html"
						' Case UCase$(Trim$(B4$)) = "<HTM": E$ = "htm"
					Case InStr(UCase(B.Value), "HTML") > 0 : E = "htm"
					Case Mid(B.Value, 7, 4) = "JFIF" : E = "jpg"
					Case Right(B.Value, 4) = "DSIG" : E = "ttf"
					Case Right(B.Value, 12) = "Standard Jet" : E = "mdb" ' Microsoft Access Database
					Case Mid(B.Value, 5, 4) = "moov" ' Apple QuickTime Movie
						If Right(B.Value, 5) = "lmvhd" Then E = "mov" Else E = "qt"
					Case B4.Value = Office.Value
						'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						FileGet(FF, OfficeType.Value, &H820)
						OfficeType.Value = LCase(OfficeType.Value)
						Select Case True
							Case InStr(OfficeType.Value, "word") : E = "doc"
							Case InStr(OfficeType.Value, "excel") : E = "xls"
							Case Else ' E$ = "ppt"
						End Select
					Case B4.Value = "RIFF"
						B4.Value = Mid(B.Value, 9, 4)
						Select Case True
							Case B4.Value = "RMID" : E = "rmi" ' MIDI
							Case B4.Value = "WAVE" : E = "wav"
							Case B4.Value = "AVI " : E = "avi"
							Case B4.Value = "CDR8" : E = "cdr" ' CorelDraw
							Case B4.Value = "CDR9" : E = "cdr" ' CorelDraw
						End Select
					Case LOF(FF) < &H100000
						Do 
							For J = 1 To Len(B.Value)
								K = Asc(Mid(B.Value, J, 1))
								Histo(K) = Histo(K) + 1
							Next 
							If EOF(FF) Then Exit Do
							'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
							FileGet(FF, B.Value)
						Loop 
						ASCII = 0
						For J = LBound(Histo) To UBound(Histo)
							Select Case J
								Case 9, 10, 13, 32 To 126 : ASCII = ASCII + Histo(J)
							End Select
							Histo(J) = 0
						Next 
						If ASCII > 0 Then If LOF(FF) / ASCII < 1.25 Then E = "txt"
				End Select '
			End If
			FileClose()
			If Len(E) > 0 Then Rename(P & S, Left(S, 9) & E) : E = vbNullString
			Do 
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				S = Dir()
				If Len(S) = 0 Then Exit Do
				System.Windows.Forms.Application.DoEvents()
			Loop Until LCase(Left(S, 4)) = "file" And LCase(Right(S, 4)) = ".chk"
		Loop 
		MsgBox("Done!", MsgBoxStyle.OKOnly Or MsgBoxStyle.MsgBoxSetForeground, "FILECHecK")
	End Sub
End Module