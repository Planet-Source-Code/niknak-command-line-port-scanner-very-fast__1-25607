Attribute VB_Name = "mod_timer"
Option Explicit

'***************************************************************
'PUBLIC API DECLARATIONS
'***************************************************************
    'CREATES AN API TIMER
    Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    'KILLS AN API TIMER
    Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

'***************************************************************
'PRIVATE VARIABLES
'***************************************************************
    'ELAPSED TIME OF CURRENT SCAN
    Public elapsed_time As Currency

'***************************************************************
'TIMER PROCEDURE
'***************************************************************
    'INCREASES ELAPSED TIME BY 0.1 OF A SECOND
    Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
        elapsed_time = elapsed_time + 0.1
        If InStr(1, Str(elapsed_time), ".", vbTextCompare) Then
            frm_main.sta_status.Panels(5).Text = Str(elapsed_time) & " s"
        Else
            frm_main.sta_status.Panels(5).Text = Str(elapsed_time) & ".0 s"
        End If
    End Sub
