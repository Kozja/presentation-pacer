Attribute VB_Name = "KlasaZdarzen"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Application

Private Sub App_SlideShowBegin(ByVal Wn As SlideShowWindow)
    Call StartClock
End Sub

Private Sub App_SlideShowNextSlide(ByVal Wn As SlideShowWindow)
    Call UpdateClock
End Sub

Private Sub App_SlideShowEnd(ByVal Pres As Presentation)
    Call StopClockTimer
End Sub
