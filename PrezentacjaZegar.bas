Option Explicit

' ============================================================
'  ZEGAR CZASU PREZENTACJI - wersja MAC (bez Application.OnTime)
'  Application.OnTime NIE istnieje w VBA na macOS, wiec zegar
'  aktualizuje sie przy KAZDEJ ZMIANIE SLAJDU (nie co sekunde).
'  Resetuje sie przy wejsciu w nowy segment slajdow.
' ============================================================

Public CurrentSegmentKey As String
Public SegmentStartTime As Date
Public TimerRunning As Boolean
Public MojaKlasa As New KlasaZdarzen

' Uruchamiane RECZNIE raz po otwarciu pliku (Auto_Open nie dziala
' w zwyklych prezentacjach .pptm - tylko w dodatkach .ppam).
' Alt+F8 -> InitEvents -> Uruchom, ZANIM wcisniesz F5.

Sub InitEvents()
    Set MojaKlasa.App = Application
    MsgBox "Zegar podpiety. Mozesz uruchomic pokaz slajdow (F5).", vbInformation
End Sub

' ------------------------------------------------------------
'  TU DEFINIUJESZ SEGMENTY:
'  Kazdy wiersz = (slajd_od, slajd_do, CZAS_TRWANIA_SEGMENTU_w_minutach)
' ------------------------------------------------------------
Sub GetSegmentInfo(slideIdx As Integer, ByRef durationSec As Long, ByRef segKey As String)
    Dim ranges() As Variant
    ranges = Array( _
        Array(1, 5, 1), _
        Array(6, 10, 15), _
        Array(11, 999, 30) _
    )

    Dim i As Integer
    For i = LBound(ranges) To UBound(ranges)
        If slideIdx >= ranges(i)(0) And slideIdx <= ranges(i)(1) Then
            durationSec = ranges(i)(2) * 60
            segKey = ranges(i)(0) & "-" & ranges(i)(1)
            Exit Sub
        End If
    Next i

    durationSec = -1
    segKey = "BRAK"
End Sub

Sub StartClock()
    CurrentSegmentKey = ""   ' wymusza reset przy pierwszym wywolaniu
    TimerRunning = True

    Dim lastTick As Single
    lastTick = Timer   ' sekundy od polnocy

    Do While TimerRunning
        ' zabezpieczenie: jesli pokaz slajdow juz sie zakonczyl, przerwij petle
        If Application.SlideShowWindows.Count = 0 Then
            TimerRunning = False
            Exit Do
        End If

        ' odswiez zegar mniej wiecej raz na sekunde
        If Timer - lastTick >= 1 Or Timer < lastTick Then
            UpdateClock
            lastTick = Timer
        End If

        DoEvents   ' oddaje sterowanie systemowi - pozwala klikac / zmieniac slajdy
    Loop
End Sub

Sub StopClockTimer()
    TimerRunning = False
End Sub

' Wywolywane recznie przy kazdej zmianie slajdu (patrz ThisPresentation)
Sub UpdateClock()
    If Not TimerRunning Then Exit Sub

    Dim sw As SlideShowWindow
    On Error Resume Next
    Set sw = Application.SlideShowWindows(1)
    On Error GoTo 0
    If sw Is Nothing Then Exit Sub

    Dim curSlide As Integer
    curSlide = sw.View.Slide.SlideIndex

    Dim durationSec As Long
    Dim segKey As String
    GetSegmentInfo curSlide, durationSec, segKey

    If segKey <> CurrentSegmentKey Then
        CurrentSegmentKey = segKey
        SegmentStartTime = Now
    End If

    Dim elapsedInSeg As Long
    elapsedInSeg = DateDiff("s", SegmentStartTime, Now)

    Dim txt As String
    Dim isLate As Boolean
    isLate = False

    If durationSec = -1 Then
        Dim m As Integer, s As Integer
        m = elapsedInSeg \ 60
        s = elapsedInSeg Mod 60
        txt = Format(m, "00") & ":" & Format(s, "00")
    Else
        Dim remainSec As Long
        remainSec = durationSec - elapsedInSeg
        If remainSec < 0 Then
            isLate = True
            remainSec = -remainSec
        End If
        Dim mm As Integer, ss As Integer
        mm = remainSec \ 60
        ss = remainSec Mod 60
        If isLate Then
            txt = "+" & Format(mm, "00") & ":" & Format(ss, "00")
        Else
            txt = Format(mm, "00") & ":" & Format(ss, "00")
        End If
    End If

    Dim shp As Shape
    On Error Resume Next
    Set shp = sw.View.Slide.Shapes("ClockBox")
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = sw.View.Slide.Shapes.AddTextbox( _
            Orientation:=msoTextOrientationHorizontal, _
            Left:=sw.View.Slide.Master.Width - 140, _
            Top:=10, _
            Width:=125, _
            Height:=40)
        shp.Name = "ClockBox"
        With shp.TextFrame.TextRange
            .Font.Size = 28
            .Font.Bold = True
            .Font.Name = "Consolas"
            .ParagraphFormat.Alignment = ppAlignRight
        End With
        shp.Fill.Visible = msoFalse
        shp.Line.Visible = msoFalse
    End If

    shp.TextFrame.TextRange.Text = txt

    If isLate Then
        shp.TextFrame.TextRange.Font.Color.RGB = RGB(220, 30, 30)
    ElseIf durationSec <> -1 And (durationSec - elapsedInSeg) <= 30 Then
        shp.TextFrame.TextRange.Font.Color.RGB = RGB(230, 160, 0)
    Else
        shp.TextFrame.TextRange.Font.Color.RGB = RGB(0, 140, 60)
    End If

    ' BRAK Application.OnTime - na Macu ta metoda nie istnieje.
    ' Zegar odswiezy sie przy nastepnej zmianie slajdu (patrz ThisPresentation).
End Sub
