VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Command Line Port Scanner"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider sli_timout 
      Height          =   330
      Left            =   60
      TabIndex        =   5
      Top             =   4320
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   582
      _Version        =   393216
      Min             =   1
      Max             =   3000
      SelStart        =   1
      TickFrequency   =   300
      Value           =   1
   End
   Begin MSWinsockLib.Winsock wsk_scanner 
      Index           =   0
      Left            =   660
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tim_timeout 
      Index           =   0
      Left            =   180
      Top             =   2040
   End
   Begin MSComctlLib.StatusBar sta_status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   4725
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "1"
            TextSave        =   "1"
            Object.ToolTipText     =   "IP Address being scanned"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "1"
            TextSave        =   "1"
            Object.ToolTipText     =   "Port being scanned from"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "2"
            TextSave        =   "2"
            Object.ToolTipText     =   "Port being scanned to"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "0"
            TextSave        =   "0"
            Object.ToolTipText     =   "Current port being scanned"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "0.0 s"
            TextSave        =   "0.0 s"
            Object.ToolTipText     =   "Elapsed time of port scan"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd_action 
      Caption         =   "Stop"
      Height          =   495
      Index           =   1
      Left            =   5940
      TabIndex        =   3
      Top             =   4140
      Width           =   1275
   End
   Begin VB.CommandButton cmd_action 
      Caption         =   "Pause"
      Height          =   495
      Index           =   0
      Left            =   4560
      TabIndex        =   2
      Top             =   4140
      Width           =   1275
   End
   Begin VB.Frame fra_results 
      Caption         =   "Results"
      Height          =   3975
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7155
      Begin VB.ListBox lst_results 
         Height          =   3570
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6915
      End
   End
   Begin VB.Label lbl_timout 
      Caption         =   "Thread Timout Length (ms)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   4080
      Width           =   4395
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************
'PRIVATE CONSTANTS
'***************************************************************
    'CMD_ACTIONS INDEX TABLE
    Const cmd_action_pause = 0
    Const cmd_action_stop = 1
    'STA_STATUS INDEX TABLE
    Const sta_status_ipadd = 1
    Const sta_status_startp = 2
    Const sta_status_stopp = 3
    Const sta_status_current = 4
    
'***************************************************************
'PRIVATE VARIABLES
'***************************************************************
    'PARSED COMMAND LINE VARIABLES
    Private ipadd As String
    Private startport As Integer
    Private stopport As Integer
    Private timeout As Integer
    'MISC SCANNER VARIABLES
    Private current_port As Integer
    Private scan_paused As Boolean
    Private scanned_ports As Integer
    Private threads_ports(5) As Integer

'***************************************************************
'PRIVATE SUB ROUTINES
'***************************************************************
    'SCANNER ACTIONS, START - PAUSE - STOP
    Private Sub cmd_action_Click(index As Integer)
        Select Case index
            Case cmd_action_pause
                scan_paused = Not scan_paused
                If scan_paused Then
                    cmd_action(cmd_action_pause).Caption = "Start"
                    KillTimer Me.hwnd, 0
                Else
                    cmd_action(cmd_action_pause).Caption = "Pause"
                    SetTimer Me.hwnd, 0, 100, AddressOf TimerProc
                End If
            Case cmd_action_stop
                Dim retval
                retval = MsgBox("Are you sure?", vbYesNo, "Stop current scan")
                If retval = vbYes Then
                    cmd_action(cmd_action_pause).Visible = False
                    cmd_action(cmd_action_stop).Visible = False
                    lbl_timout.Visible = False
                    sli_timout.Visible = False
                    current_port = stopport
                    elapsed_time = 0
                    KillTimer Me.hwnd, 0
                End If
        End Select
    End Sub

    Private Sub Form_Load()
        On Error GoTo parse_error
        '-----------------------
        'COMMAND LINE PASSING VARIABLES
        '-----------------------
        Dim retval As String    'STORES COMMAND LINE VARIABLES
        Dim startn As Integer   'START POSITION OF CURRENT CHUNK
        Dim stopn As Integer    'STOP POSITION OF CURRENT CHUNK
        '-----------------------
        'RETRIEVE COMMAND LINE VARIABLES
        retval = Command
        If Len(retval) >= 13 Then
            '-----------------------
            'REMOVE IP ADDRESS
            stopn = InStr(1, retval, ",", vbTextCompare)
            ipadd = Mid(retval, 1, stopn - 1)
            '-----------------------
            'REMOVE START PORT
            startn = stopn
            stopn = InStr(startn + 1, retval, ",", vbTextCompare)
            startport = Val(Mid(retval, startn + 1, stopn - 1))
            '-----------------------
            'REMOVE STOP PORT
            startn = stopn
            stopn = InStr(startn + 1, retval, ",", vbTextCompare)
            stopport = Val(Mid(retval, startn + 1, stopn - 1))
            '-----------------------
            'REMOVE TIMOUT
            timeout = Val(Right(retval, Len(retval) - stopn))
            '-----------------------
        Else
            GoTo parse_error
        End If
        '-----------------------
        'REFLECT PARSED VARIABLES IN STATUS BAR
        sta_status.Panels(sta_status_ipadd).Text = ipadd
        If startport < 0 Then startport = 0
        sta_status.Panels(sta_status_startp).Text = startport
        If stopport <= startport Then stopport = startport + 1
        sta_status.Panels(sta_status_stopp).Text = stopport
        If timeout < 1 Or timeout > 3000 Then
            timeout = 1000
        End If
        sli_timout.Value = timeout
        SetTimer Me.hwnd, 0, 100, AddressOf TimerProc
        '-----------------------
        'START SCANNING
        load_threads
        start_scanning
        Exit Sub
parse_error:
        MsgBox "Usage :: Cline_Portscanner.exe <IP ADDRESS>,<START PORT>,<STOP PORT>,<THREAD TIMEOUT>"
        End
    End Sub

'***********************************
'PORT SCANNER
'***********************************
    'LOAD UP ARRAYS OF TIMOUT TIMERS AND WINSOCK CONTROLS
    Private Sub load_threads()
        Dim load_thread As Integer
        For load_thread = 1 To 5
            Load wsk_scanner(load_thread)
            wsk_scanner(load_thread).RemoteHost = ipadd
            Load tim_timeout(load_thread)
            tim_timeout(load_thread).Enabled = False
            tim_timeout(load_thread).Interval = sli_timout.Value
            If load_thread > (stopport - startport) Then Exit For
        Next load_thread
    End Sub

    'START SCANNING
    Private Sub start_scanning()
        Dim start_thread As Integer
        current_port = startport
        For start_thread = 1 To 5
            wsk_scanner(start_thread).RemoteHost = ipadd
            wsk_scanner(start_thread).RemotePort = current_port
            threads_ports(start_thread) = current_port
            wsk_scanner(start_thread).Connect
            tim_timeout(start_thread).Enabled = True
            If current_port >= stopport Then
                Exit Sub
            Else
                current_port = current_port + 1
            End If
        Next start_thread
    End Sub

    'PORT IS OPEN, DISPLAY THE RESULT
    Private Sub wsk_scanner_Connect(index As Integer)
        scanned_ports = scanned_ports + 1
        portstatus_display Str(threads_ports(index)), "open"
        next_port index
    End Sub

    'PORT TIMEDOUT GO ONTO NEXT
    Private Sub tim_timeout_Timer(index As Integer)
        scanned_ports = scanned_ports + 1
        sta_status.Panels(sta_status_current).Text = Str(scanned_ports)
        next_port index
    End Sub

    'SETUP TO SCAN NEXT PORT
    Private Sub next_port(index As Integer)
        While scan_paused
            DoEvents
        Wend
        wsk_scanner(index).Close
        If current_port = stopport Then
            Unload tim_timeout(index)
            Unload wsk_scanner(index)
            sta_status.Panels(sta_status_current).Text = "Complete"
            cmd_action(cmd_action_pause).Visible = False
            cmd_action(cmd_action_stop).Visible = False
            lbl_timout.Visible = False
            sli_timout.Visible = False
            KillTimer Me.hwnd, 0
            Exit Sub
        Else
            current_port = current_port + 1
        End If
        wsk_scanner(index).RemotePort = current_port
        threads_ports(index) = current_port
        wsk_scanner(index).Connect
    End Sub

    'ADDS MESSAGE TO RESULT PANE
    Private Sub portstatus_display(port As Integer, message As String)
        If message = "open" Then
            lst_results.AddItem port & " - open"
        End If
    End Sub
