VERSION 5.00
Begin VB.UserControl DmTimePicker 
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   ScaleHeight     =   2295
   ScaleWidth      =   3330
   ToolboxBitmap   =   "DmTimePicker.ctx":0000
   Begin VB.Timer tmrSup 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2640
      Top             =   0
   End
   Begin VB.Timer tmrSdown 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2640
      Top             =   480
   End
   Begin VB.Timer tmrMdown 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2160
      Top             =   480
   End
   Begin VB.Timer tmrMup 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2160
      Top             =   0
   End
   Begin VB.Timer tmrHdown 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1680
      Top             =   480
   End
   Begin VB.Timer tmrHup 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1680
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   960
   End
   Begin VB.PictureBox picCurrent 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   1335
      TabIndex        =   13
      Top             =   1080
      Width           =   1335
      Begin VB.Label lblCCT 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Current"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   1275
      End
      Begin VB.Label lblCurTime 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "06:02 AM"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   300
         TabIndex        =   14
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.PictureBox picAmPm 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00DC9670&
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   960
      ScaleHeight     =   1005
      ScaleWidth      =   615
      TabIndex        =   11
      Top             =   0
      Width           =   615
      Begin VB.Label lblAMPM 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "AM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   30
         TabIndex        =   12
         Top             =   240
         Width           =   345
      End
   End
   Begin VB.PictureBox PicTime 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   975
      TabIndex        =   1
      Top             =   0
      Width           =   975
      Begin VB.PictureBox btnSdown 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00DC9670&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   615
         ScaleHeight     =   255
         ScaleWidth      =   285
         TabIndex        =   7
         Top             =   720
         Width           =   285
      End
      Begin VB.PictureBox btnSup 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00DC9670&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   615
         ScaleHeight     =   255
         ScaleWidth      =   285
         TabIndex        =   6
         Top             =   240
         Width           =   285
      End
      Begin VB.PictureBox btnHup 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00DC9670&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   15
         ScaleHeight     =   255
         ScaleWidth      =   285
         TabIndex        =   5
         Top             =   240
         Width           =   285
      End
      Begin VB.PictureBox btnHdown 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00DC9670&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   15
         ScaleHeight     =   255
         ScaleWidth      =   285
         TabIndex        =   4
         Top             =   720
         Width           =   285
      End
      Begin VB.PictureBox btnMup 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00DC9670&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   315
         ScaleHeight     =   255
         ScaleWidth      =   285
         TabIndex        =   3
         Top             =   240
         Width           =   285
      End
      Begin VB.PictureBox btnMdown 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00DC9670&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   315
         ScaleHeight     =   255
         ScaleWidth      =   285
         TabIndex        =   2
         Top             =   720
         Width           =   285
      End
      Begin VB.Label lblCH 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DC9670&
         Height          =   240
         Left            =   0
         TabIndex        =   18
         Top             =   15
         Width           =   285
      End
      Begin VB.Label lblCM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DC9670&
         Height          =   240
         Left            =   360
         TabIndex        =   17
         Top             =   15
         Width           =   285
      End
      Begin VB.Label lblCS 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DC9670&
         Height          =   240
         Left            =   600
         TabIndex        =   16
         Top             =   15
         Width           =   285
      End
      Begin VB.Label lblS 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DC9670&
         Height          =   240
         Left            =   615
         TabIndex        =   10
         Top             =   480
         Width           =   285
      End
      Begin VB.Label lblM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DC9670&
         Height          =   240
         Left            =   315
         TabIndex        =   9
         Top             =   480
         Width           =   285
      End
      Begin VB.Label lblH 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DC9670&
         Height          =   240
         Left            =   15
         TabIndex        =   8
         Top             =   480
         Width           =   285
      End
   End
   Begin VB.Label lblTmp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC9670&
      Height          =   240
      Left            =   2880
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "DmTimePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Programmer:        Donckers Frank
'                    DarkManSoft@Gmail.com
'
' Description:       User Control TimePicker

'=====================================================
' Enum languages
'=====================================================
Public Enum Languages
    [Englisch] = 0
    [Nederlands] = 1
    [Francais] = 2
    [Deutch] = 3
    [Italiano] = 4
    [Espagnol] = 5
End Enum

'=====================================================
' Enum TimeFormats
'=====================================================
Public Enum TimeFormats
    [24Hours] = 0
    [12Hours] = 1
End Enum

'=====================================================
' Enum styles
'=====================================================
Public Enum Styles
    [Flat] = 0
    [3D] = 1
End Enum

'=====================================================
' Events
'=====================================================
Event Change()

'=====================================================
' Default Property Values
'=====================================================
' Time
Const m_def_TimeFormat = 0
Const m_def_TimeForeColor = &HEA692E
Const m_def_TimeBackColor = &H80000005
Const m_def_TimeBorderColor = &HEA692E
' Current time
Const m_def_ShowCurrentTime = True
Const m_def_CurrentTimeForeColor = &H80FF&
Const m_def_CurrentTimeBackColor = &HEA692E
Const m_def_CurrentTimeBorderColor = &H80FF&
' Am/Pm
Const m_def_AMPMForeColor = &HC000&
Const m_def_AMPMBorderColor = &HEA692E
Const m_def_AMPMBackColor = &HEA692E
' Buttons
Const m_def_HourBtnsForeColor = &H80000005
Const m_def_HourBtnsBackColor = &HEA692E
Const m_def_HourBtnsBorderColor = &HEA692E
Const m_def_MinutesBtnsForeColor = &H80000005
Const m_def_MinutesBtnsBackColor = &HEA692E
Const m_def_MinutesBtnsBorderColor = &HEA692E
Const m_def_SecondsBtnsForeColor = &H80000005
Const m_def_SecondsBtnsBackColor = &HEA692E
Const m_def_SecondsBtnsBorderColor = &HEA692E
' Disabled
Const m_def_DisabledForeColor = &H81BECB
Const m_def_DisabledBackColor = &HCFF0F2
Const m_def_DisabledBorderColor = &H81BECB
' Return Selected
Const m_def_SelectedTime = "0"
Const m_def_SelectedHours = 0
Const m_def_SelectedMinutes = 0
Const m_def_SelectedAmPm = 0
' Rest
Const m_def_Style = 0
Const m_def_Enabled = True
Const m_def_Language = 0
Const m_def_SpinBtnInterval = 250

'=====================================================
' Property Variables
'=====================================================
' Time
Dim m_Starttime                 As String
Dim m_TimeFormat                As TimeFormats
Dim m_TimeForeColor             As OLE_COLOR
Dim m_TimeBackColor             As OLE_COLOR
Dim m_TimeBorderColor           As OLE_COLOR
' Current time
Dim m_ShowCurrentTime           As Boolean
Dim m_CurrentTimeForeColor      As OLE_COLOR
Dim m_CurrentTimeBackColor      As OLE_COLOR
Dim m_CurrentTimeBorderColor    As OLE_COLOR
' Am/Pm
Dim m_AMPMForeColor             As OLE_COLOR
Dim m_AMPMBorderColor           As OLE_COLOR
Dim m_AMPMBackColor             As OLE_COLOR
' Buttons
Dim m_HourBtnsForeColor         As OLE_COLOR
Dim m_HourBtnsBackColor         As OLE_COLOR
Dim m_HourBtnsBorderColor       As OLE_COLOR
Dim m_MinutesBtnsForeColor      As OLE_COLOR
Dim m_MinutesBtnsBackColor      As OLE_COLOR
Dim m_MinutesBtnsBorderColor    As OLE_COLOR
Dim m_SecondsBtnsForeColor      As OLE_COLOR
Dim m_SecondsBtnsBackColor      As OLE_COLOR
Dim m_SecondsBtnsBorderColor    As OLE_COLOR
' Disabled
Dim m_DisabledForeColor         As OLE_COLOR
Dim m_DisabledBackColor         As OLE_COLOR
Dim m_DisabledBorderColor       As OLE_COLOR
' Fonts
Dim m_FontTime As Font
Dim m_FontCurrentTime As Font
Dim m_FontAmPm As Font
' Return Selected
Dim m_SelectedTime              As String
Dim m_SelectedHours             As String
Dim m_SelectedMinutes           As String
Dim m_SelectedAmPm              As String
' Rest
Dim m_Style                     As Styles
Dim m_Enabled                   As Boolean
Dim m_Language                  As Languages
Dim m_SpinBtnInterval           As Integer

'=====================================================
' Other program Variables
'=====================================================
Dim OldScaleMode As Byte
Dim cControl As Control
Dim StartCol As Double, EndCol As Double
Dim RedI As Single, BlueI As Single, GreenI As Single
Dim RedStart As Integer, GreenStart As Integer, BlueStart As Integer
Dim RedEnd As Double, GreenEnd As Double, BlueEnd As Double
Dim i, ii, iii As Integer
Dim NewColor As Single
Dim MidX, MidY As Integer
Dim HoursNow As Byte
Dim MinutesNow As Byte
Dim SecondsNow As Byte
Dim CurrentTime As String
Dim CurrentTimeLabel As String
Dim MaxHour As Byte
Dim MinWidth As Long
Dim MinHeight As Long
Dim Rounded As Long
Dim DevidedHeight As Long

'=====================================================
' Spinbutton Hours up
'=====================================================
Private Sub btnHup_Click()
    If Val(lblH.Caption) + 1 > MaxHour - 1 Then
        lblH.Caption = "00"
        If lblAMPM.Caption = "AM" Then
            lblAMPM = "PM"
        Else
            lblAMPM = "AM"
        End If
    Else
        If Val(lblH.Caption) + 1 > 9 Then
            lblH.Caption = Val(lblH.Caption) + 1
        Else
            lblH.Caption = "0" & Val(lblH.Caption) + 1
        End If
    End If
End Sub
Private Sub btnHup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnHup, ShiftColors(m_HourBtnsBackColor, 170), m_HourBtnsBackColor, False, True, m_HourBtnsForeColor, m_HourBtnsBorderColor
    tmrHup.Enabled = True
End Sub
Private Sub btnHup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnHup, ShiftColors(m_HourBtnsBackColor, 170), m_HourBtnsBackColor, True, True, m_HourBtnsForeColor, m_HourBtnsBorderColor
    tmrHup.Enabled = False
End Sub
Private Sub tmrHup_Timer()
    If Val(lblH.Caption) + 1 > MaxHour - 1 Then
        lblH.Caption = "00"
        If lblAMPM.Caption = "AM" Then
            lblAMPM = "PM"
        Else
            lblAMPM = "AM"
        End If
    Else
        If Val(lblH.Caption) + 1 > 9 Then
            lblH.Caption = Val(lblH.Caption) + 1
        Else
            lblH.Caption = "0" & Val(lblH.Caption) + 1
        End If
    End If
End Sub

'=====================================================
' Spinbutton Hours Down
'=====================================================
Private Sub btnHdown_Click()
    If Val(lblH.Caption) - 1 < 0 Then
        lblH.Caption = MaxHour - 1
        If lblAMPM.Caption = "AM" Then
            lblAMPM = "PM"
        Else
            lblAMPM = "AM"
        End If
    Else
        If Val(lblH.Caption) - 1 > 9 Then
            lblH.Caption = Val(lblH.Caption) - 1
        Else
            lblH.Caption = "0" & Val(lblH.Caption) - 1
        End If
    End If
End Sub
Private Sub btnHdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnHdown, ShiftColors(m_HourBtnsBackColor, 170), m_HourBtnsBackColor, False, False, m_HourBtnsForeColor, m_HourBtnsBorderColor
    tmrHdown.Enabled = True
End Sub
Private Sub btnHdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnHdown, ShiftColors(m_HourBtnsBackColor, 170), m_HourBtnsBackColor, True, False, m_HourBtnsForeColor, m_HourBtnsBorderColor
    tmrHdown.Enabled = False
End Sub
Private Sub tmrHdown_Timer()
    If Val(lblH.Caption) - 1 < 0 Then
        lblH.Caption = MaxHour - 1
         If lblAMPM.Caption = "AM" Then
            lblAMPM = "PM"
        Else
            lblAMPM = "AM"
        End If
   Else
        If Val(lblH.Caption) - 1 > 9 Then
            lblH.Caption = Val(lblH.Caption) - 1
        Else
            lblH.Caption = "0" & Val(lblH.Caption) - 1
        End If
    End If
End Sub

'=====================================================
' Spinbutton Minutes Up
'=====================================================
Private Sub btnMup_Click()
    If Val(lblM.Caption) + 1 > 59 Then
        lblM.Caption = "00"
    Else
        If Val(lblM.Caption) + 1 > 9 Then
            lblM.Caption = Val(lblM.Caption) + 1
        Else
            lblM.Caption = "0" & Val(lblM.Caption) + 1
        End If
    End If
End Sub
Private Sub btnMup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnMup, ShiftColors(m_HourBtnsBackColor, 170), m_HourBtnsBackColor, False, True, m_HourBtnsForeColor, m_HourBtnsBorderColor
    tmrMup.Enabled = True
End Sub
Private Sub btnMup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnMup, ShiftColors(m_HourBtnsBackColor, 170), m_HourBtnsBackColor, True, True, m_HourBtnsForeColor, m_HourBtnsBorderColor
    tmrMup.Enabled = False
End Sub
Private Sub tmrMup_Timer()
    If Val(lblM.Caption) + 1 > 59 Then
        lblM.Caption = "00"
    Else
        If Val(lblM.Caption) + 1 > 9 Then
            lblM.Caption = Val(lblM.Caption) + 1
        Else
            lblM.Caption = "0" & Val(lblM.Caption) + 1
        End If
    End If
End Sub

'=====================================================
' Spinbutton Minutes Down
'=====================================================
Private Sub btnMdown_Click()
    If Val(lblM.Caption) - 1 < 0 Then
        lblM.Caption = MaxHour - 1
    Else
        If Val(lblM.Caption) - 1 > 9 Then
            lblM.Caption = Val(lblM.Caption) - 1
        Else
            lblM.Caption = "0" & Val(lblM.Caption) - 1
        End If
    End If
End Sub
Private Sub btnMdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnMdown, ShiftColors(m_HourBtnsBackColor, 170), m_HourBtnsBackColor, False, False, m_HourBtnsForeColor, m_HourBtnsBorderColor
    tmrMdown.Enabled = True
End Sub
Private Sub btnMdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnMdown, ShiftColors(m_HourBtnsBackColor, 170), m_HourBtnsBackColor, True, False, m_HourBtnsForeColor, m_HourBtnsBorderColor
    tmrMdown.Enabled = False
End Sub
Private Sub tmrMdown_Timer()
    If Val(lblM.Caption) - 1 < 0 Then
        lblM.Caption = MaxHour - 1
    Else
        If Val(lblM.Caption) - 1 > 9 Then
            lblM.Caption = Val(lblM.Caption) - 1
        Else
            lblM.Caption = "0" & Val(lblM.Caption) - 1
        End If
    End If
End Sub

'=====================================================
' Spinbutton Seconds Up
'=====================================================
Private Sub btnSup_Click()
    If Val(lblS.Caption) + 1 > 59 Then
        lblS.Caption = "00"
    Else
        If Val(lblS.Caption) + 1 > 9 Then
            lblS.Caption = Val(lblS.Caption) + 1
        Else
            lblS.Caption = "0" & Val(lblS.Caption) + 1
        End If
    End If
End Sub
Private Sub btnSup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnSup, ShiftColors(m_HourBtnsBackColor, 170), m_HourBtnsBackColor, False, True, m_HourBtnsForeColor, m_HourBtnsBorderColor
    tmrSup.Enabled = True
End Sub
Private Sub btnSup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnSup, ShiftColors(m_HourBtnsBackColor, 170), m_HourBtnsBackColor, True, True, m_HourBtnsForeColor, m_HourBtnsBorderColor
    tmrSup.Enabled = False
End Sub
Private Sub tmrSup_Timer()
    If Val(lblS.Caption) + 1 > 59 Then
        lblS.Caption = "00"
    Else
        If Val(lblS.Caption) + 1 > 9 Then
            lblS.Caption = Val(lblS.Caption) + 1
        Else
            lblS.Caption = "0" & Val(lblS.Caption) + 1
        End If
    End If
End Sub
'=====================================================
' Spinbutton Seconds Down
'=====================================================
Private Sub btnSdown_Click()
    If Val(lblS.Caption) - 1 < 0 Then
        lblS.Caption = MaxHour - 1
    Else
        If Val(lblS.Caption) - 1 > 9 Then
            lblS.Caption = Val(lblS.Caption) - 1
        Else
            lblS.Caption = "0" & Val(lblS.Caption) - 1
        End If
    End If
End Sub
Private Sub btnSdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnSdown, ShiftColors(m_HourBtnsBackColor, 170), m_HourBtnsBackColor, False, False, m_HourBtnsForeColor, m_HourBtnsBorderColor
    tmrSdown.Enabled = True
End Sub
Private Sub btnSdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton btnSdown, ShiftColors(m_HourBtnsBackColor, 170), m_HourBtnsBackColor, True, False, m_HourBtnsForeColor, m_HourBtnsBorderColor
    tmrSdown.Enabled = False
End Sub
Private Sub tmrSdown_Timer()
    If Val(lblS.Caption) - 1 < 0 Then
        lblS.Caption = MaxHour - 1
    Else
        If Val(lblS.Caption) - 1 > 9 Then
            lblS.Caption = Val(lblS.Caption) - 1
        Else
            lblS.Caption = "0" & Val(lblS.Caption) - 1
        End If
    End If
End Sub

'=====================================================
' Doubleclick on current time sets time to current
'=====================================================
Private Sub lblCCT_DblClick()
    If m_TimeFormat = [12Hours] Then
        HoursNow = Val(Left$(Format(Now, "hh:mm:ss"), 2))
        If HoursNow > 12 Then
            lblAMPM = "PM"
            lblH = HoursNow - 12
        Else
            lblAMPM = "AM"
            lblH = HoursNow
        End If
            If Val(lblH) < 10 Then lblH = "0" & lblH
    Else
        lblH = Left$(Format(Now, "hh:mm:ss"), 2)
    End If
    lblM = Mid$(Format(Now, "hh:mm:ss"), 4, 2)
    lblS = Right$(Format(Now, "hh:mm:ss"), 2)
End Sub
Private Sub lblCurTime_Click()
    lblCCT_DblClick
End Sub


'=====================================================
' Raise event change on change date
'=====================================================
Private Sub lblM_Change()
    RaiseEvent Change
End Sub
Private Sub lblH_Change()
    RaiseEvent Change
End Sub

Private Sub lblS_Change()
    RaiseEvent Change
End Sub
Private Sub lblAMPM_Change()
    RaiseEvent Change
End Sub

'=====================================================
' Draw all timepickercontrols
'=====================================================
Private Sub DrawControls()
    Dim NewForeColor, NewBackColor, NewBorderColor As OLE_COLOR
    lblH.Font.Name = m_FontTime.Name
    lblH.Font.Size = m_FontTime.Size
    lblH.Font.Bold = m_FontTime.Bold
    lblM.Font.Name = m_FontTime.Name
    lblM.Font.Size = m_FontTime.Size
    lblM.Font.Bold = m_FontTime.Bold
    lblS.Font.Name = m_FontTime.Name
    lblS.Font.Size = m_FontTime.Size
    lblS.Font.Bold = m_FontTime.Bold
    lblCH.Font.Name = m_FontTime.Name
    lblCH.Font.Size = m_FontTime.Size
    lblCH.Font.Bold = m_FontTime.Bold
    lblCM.Font.Name = m_FontTime.Name
    lblCM.Font.Size = m_FontTime.Size
    lblCM.Font.Bold = m_FontTime.Bold
    lblCS.Font.Name = m_FontTime.Name
    lblCS.Font.Size = m_FontTime.Size
    lblCS.Font.Bold = m_FontTime.Bold
    lblTmp.Font.Name = m_FontTime.Name '\
    lblTmp.Font.Size = m_FontTime.Size ' -used to calculate minwidth
    lblTmp.Font.Bold = m_FontTime.Bold '/
    lblAMPM.Font.Name = m_FontAmPm.Name
    lblAMPM.Font.Size = m_FontAmPm.Size
    lblAMPM.Font.Bold = m_FontAmPm.Bold
    lblCCT.Font.Name = m_FontCurrentTime.Name
    lblCCT.Font.Size = m_FontCurrentTime.Size
    lblCCT.Font.Bold = m_FontCurrentTime.Bold
    lblCurTime.Font.Name = m_FontCurrentTime.Name
    lblCurTime.Font.Size = m_FontCurrentTime.Size
    lblCurTime.Font.Bold = m_FontCurrentTime.Bold
    If m_Enabled = False Then
        lblH.ForeColor = m_DisabledForeColor
        lblM.ForeColor = m_DisabledForeColor
        lblS.ForeColor = m_DisabledForeColor
        lblCH.ForeColor = m_DisabledForeColor
        lblCM.ForeColor = m_DisabledForeColor
        lblCS.ForeColor = m_DisabledForeColor
        lblCCT.ForeColor = m_DisabledForeColor
        lblCurTime.ForeColor = m_DisabledForeColor
    Else
        lblH.ForeColor = m_TimeForeColor
        lblM.ForeColor = m_TimeForeColor
        lblS.ForeColor = m_TimeForeColor
        lblCH.ForeColor = m_TimeForeColor
        lblCM.ForeColor = m_TimeForeColor
        lblCS.ForeColor = m_TimeForeColor
        lblCCT.ForeColor = m_CurrentTimeForeColor
        lblCurTime.ForeColor = m_CurrentTimeForeColor
    End If
    lblH.Caption = "00"
    lblM.Caption = "00"
    lblS.Caption = "00"
    ' Horizontal alignment
    lblH.Height = lblTmp.Height
    lblM.Height = lblTmp.Height
    lblS.Height = lblTmp.Height
    MinWidth = (3 * lblTmp.Width) + 60
    MinHeight = (4 * lblTmp.Height) + 60
    If m_ShowCurrentTime = True Then
        DrawCurrentTime
        picCurrent.Visible = True
    Else
        picCurrent.Visible = False
    End If
    ' Hours and minutes
    If m_Starttime <> "" Then
        lblH.Caption = Left$(m_Starttime, 2)
        lblM.Caption = Mid$(m_Starttime, 4, 2)
    End If
    If m_TimeFormat = [12Hours] Then
        ' AM/PM
        lblAMPM.Caption = "AM"
        If Right$(m_Starttime, 2) = "PM" Then lblAMPM.Caption = "PM"
        picAmPm.Visible = True
        picAmPm.Width = lblAMPM.Width + 150
        MinWidth = MinWidth + picAmPm.Width
        lblAMPM.Left = 75
        lblAMPM.ZOrder 0
        ' Seconds
        If m_Starttime <> "" And Len(m_Starttime) > 8 Then lblS.Caption = Mid$(m_Starttime, 7, 2)
        PicTime.Width = UserControl.Width - picAmPm.Width
    Else
        ' Seconds
        picAmPm.Visible = False
        If m_Starttime <> "" And Len(m_Starttime) > 6 Then lblS.Caption = Right$(m_Starttime, 2)
        PicTime.Width = UserControl.Width
    End If
    If UserControl.Width < MinWidth Then UserControl.Width = MinWidth
    btnHup.Width = (PicTime.Width / 3) - 15
    Rounded = Round(btnHup.Width / 15)
    btnHup.Width = Rounded * 15
    btnHdown.Width = btnHup.Width
    btnMup.Width = btnHup.Width
    btnMdown.Width = btnHup.Width
    btnSup.Width = btnHup.Width
    btnSdown.Width = btnHup.Width
    btnHup.Left = 15
    btnHdown.Left = 15
    btnMup.Left = btnHup.Left + btnHup.Width + 15
    btnMdown.Left = btnMup.Left
    btnSup.Left = btnMup.Left + btnMup.Width + 15
    btnSdown.Left = btnSup.Left
    lblH.Width = btnHup.Width
    lblM.Width = btnMup.Width
    lblS.Width = btnSup.Width
    lblCH.Width = btnHup.Width
    lblCM.Width = btnMup.Width
    lblCS.Width = btnSup.Width
    lblH.Left = btnHup.Left
    lblM.Left = btnMup.Left
    lblS.Left = btnSup.Left
    lblCH.Left = btnHup.Left
    lblCM.Left = btnMup.Left
    lblCS.Left = btnSup.Left
    If PicTime.Width <> (btnSup.Width + btnSup.Left) + 15 Then PicTime.Width = (btnSup.Width + btnSup.Left) + 15
    If m_ShowCurrentTime = True Then
        picCurrent.Width = UserControl.Width
        lblCurTime.Left = (picCurrent.Width / 2) - (lblCurTime.Width / 2)
        lblCCT.Left = (picCurrent.Width / 2) - (lblCCT.Width / 2)
        lblCCT.Top = 30
        lblCurTime.Top = lblCCT.Top + lblCCT.Height + 30
    End If
    ' Vertical alignment
    If m_ShowCurrentTime = True Then
        If UserControl.Height < (MinHeight + picCurrent.Height) Then
            UserControl.Height = (PicTime.Height + picCurrent.Height)
        Else
            PicTime.Height = UserControl.Height - picCurrent.Height
        End If
    Else
        If PicTime.Height < MinHeight Then
            PicTime.Height = MinHeight
        Else
            PicTime.Height = UserControl.Height
        End If
    End If
    ' Buttons spinup
    DevidedHeight = (PicTime.Height - lblCH.Height - 30) / 3
    btnHup.Top = lblCH.Top + lblCH.Height - 30 '+ 30
    btnMup.Top = btnHup.Top
    btnSup.Top = btnHup.Top
    btnHup.Height = DevidedHeight
    btnMup.Height = DevidedHeight
    btnSup.Height = DevidedHeight
    ' Buttons Spindow
    btnHdown.Top = PicTime.Height - DevidedHeight
    btnMdown.Top = btnHdown.Top
    btnSdown.Top = btnHdown.Top
    btnHdown.Height = DevidedHeight
    btnMdown.Height = DevidedHeight
    btnSdown.Height = DevidedHeight
    ' hh mm ss labels
    lblH.Top = (((btnHup.Top + btnHup.Height) + btnHdown.Top) / 2) - (lblH.Height / 2)
    lblM.Top = lblH.Top
    lblS.Top = lblH.Top
    If m_Enabled = False Then
        DrawButton btnHup, ShiftColors(m_DisabledBackColor, 170), m_DisabledBackColor, True, True, m_DisabledForeColor, m_DisabledBorderColor
        DrawButton btnHdown, ShiftColors(m_DisabledBackColor, 170), m_DisabledBackColor, True, False, m_DisabledForeColor, m_DisabledBorderColor
        DrawButton btnMup, ShiftColors(m_DisabledBackColor, 170), m_DisabledBackColor, True, True, m_DisabledForeColor, m_DisabledBorderColor
        DrawButton btnMdown, ShiftColors(m_DisabledBackColor, 170), m_DisabledBackColor, True, False, m_DisabledForeColor, m_DisabledBorderColor
        DrawButton btnSup, ShiftColors(m_DisabledBackColor, 170), m_DisabledBackColor, True, True, m_DisabledForeColor, m_DisabledBorderColor
        DrawButton btnSdown, ShiftColors(m_DisabledBackColor, 170), m_DisabledBackColor, True, False, m_DisabledForeColor, m_DisabledBorderColor
    Else
        DrawButton btnHup, ShiftColors(m_HourBtnsBackColor, 170), m_HourBtnsBackColor, True, True, m_HourBtnsForeColor, m_HourBtnsBorderColor
        DrawButton btnHdown, ShiftColors(m_HourBtnsBackColor, 170), m_HourBtnsBackColor, True, False, m_HourBtnsForeColor, m_HourBtnsBorderColor
        DrawButton btnMup, ShiftColors(m_MinutesBtnsBackColor, 170), m_MinutesBtnsBackColor, True, True, m_MinutesBtnsForeColor, m_MinutesBtnsBorderColor
        DrawButton btnMdown, ShiftColors(m_MinutesBtnsBackColor, 170), m_MinutesBtnsBackColor, True, False, m_MinutesBtnsForeColor, m_MinutesBtnsBorderColor
        DrawButton btnSup, ShiftColors(m_SecondsBtnsBackColor, 170), m_SecondsBtnsBackColor, True, True, m_SecondsBtnsForeColor, m_SecondsBtnsBorderColor
        DrawButton btnSdown, ShiftColors(m_SecondsBtnsBackColor, 170), m_SecondsBtnsBackColor, True, False, m_SecondsBtnsForeColor, m_SecondsBtnsBorderColor
    End If
    If m_TimeFormat = [12Hours] Then
        picAmPm.Height = UserControl.Height
        picAmPm.Left = PicTime.Width
        picAmPm.Height = PicTime.Height
        If m_Enabled = False Then
            DrawPicBack picAmPm, ShiftColors(m_DisabledBackColor, 170), m_DisabledBackColor, m_DisabledBorderColor, True
        Else
            DrawPicBack picAmPm, ShiftColors(m_AMPMBackColor, 170), m_AMPMBackColor, m_AMPMBorderColor, True
        End If
        If m_Enabled = False Then
            lblAMPM.ForeColor = m_DisabledForeColor
        Else
            lblAMPM.ForeColor = m_AMPMForeColor
        End If
        lblAMPM.Top = (picAmPm.Height / 2) - (lblAMPM.Height / 2)
   End If
    If m_ShowCurrentTime = True Then
        picCurrent.Height = lblCurTime.Top + lblCurTime.Height + 45
        picCurrent.Top = PicTime.Height - 15
        If m_Enabled = False Then
            DrawPicBack picCurrent, ShiftColors(m_DisabledBackColor, 170), m_DisabledBackColor, m_DisabledBorderColor
        Else
            DrawPicBack picCurrent, ShiftColors(m_CurrentTimeBackColor, 170), m_CurrentTimeBackColor, m_CurrentTimeBorderColor
        End If
        ' Update current time only at runtime
        If Ambient.UserMode = True Then Timer1.Enabled = True
    End If
    If m_Enabled = False Then
        DrawPicBack PicTime, ShiftColors(m_DisabledBackColor, 170), m_DisabledBackColor, m_DisabledBorderColor
        PicTime.Line (btnHup.Width + 15, 0)-(btnHup.Width + 15, PicTime.Height), m_DisabledBorderColor
        PicTime.Line (btnSup.Left - 15, 0)-(btnSup.Left - 15, PicTime.Height), m_DisabledBorderColor
    Else
        DrawPicBack PicTime, ShiftColors(m_TimeBackColor, 170), m_TimeBackColor, m_TimeBorderColor
        PicTime.Line (btnHup.Width + 15, 0)-(btnHup.Width + 15, PicTime.Height), m_TimeBorderColor
        PicTime.Line (btnSup.Left - 15, 0)-(btnSup.Left - 15, PicTime.Height), m_TimeBorderColor
    End If
    PicTime.Refresh
End Sub

'=====================================================
' Draw current time
'=====================================================
Private Sub DrawCurrentTime()
    If m_TimeFormat = [12Hours] Then
        HoursNow = Val(Left$(Format(Now, "hh:mm:ss"), 2))
        If HoursNow > 12 Then
            lblCurTime = HoursNow - 12 & ":" & Format(Now, "mm:ss") & " PM"
            If HoursNow < 22 Then lblCurTime = "0" & lblCurTime
        Else
            lblCurTime = Format(Now, "hh:mm:ss") & " AM"
        End If
        If MinWidth < lblCurTime.Width Then MinWidth = lblCurTime.Width
    Else
        lblCurTime = Format(Now, "hh:mm:ss")
        If MinWidth < lblCurTime.Width Then MinWidth = lblCurTime.Width
    End If
End Sub
Private Sub Timer1_Timer()
    DrawCurrentTime
End Sub

'=====================================================
' Draw buttons with arrows
'=====================================================
Public Sub DrawButton(ctlControl As Control, StartColor As OLE_COLOR, EndColor As OLE_COLOR, btnUp As Boolean, ArrowUp As Boolean, TextColor As OLE_COLOR, BordersColor As OLE_COLOR)  'Horizontal gradient
    On Error Resume Next
    DoEvents
    OldScaleMode = ctlControl.ScaleMode
    ctlControl.ScaleMode = 3
    If m_Style = Flat Or EndColor = &H8000000F Then
        If btnUp = True Then
            ctlControl.BackColor = EndColor
        Else
            ctlControl.BackColor = ShiftColors(EndColor, -50)
        End If
       GoTo DrawArrows
    End If
    If btnUp = True Then
        Call InitializeCol(ctlControl, StartColor, EndColor, False)
    Else
        Call InitializeCol(ctlControl, EndColor, StartColor, False)
    End If
    For i = 0 To ctlControl.ScaleHeight
        NewColor = RGB(RedStart + i * RedI, GreenStart + i * GreenI, BlueStart + i * BlueI)
        ctlControl.Line (0, i)-(ctlControl.ScaleWidth, i), NewColor
    Next
    DoEvents
DrawArrows:
    MidX = Round(ctlControl.ScaleWidth / 2)
    MidY = Round(ctlControl.ScaleHeight / 2)
    If ArrowUp = True Then
        ctlControl.Line (MidX - 3, MidY)-(MidX + 3, MidY), TextColor
        ctlControl.Line (MidX - 2, MidY - 1)-(MidX + 2, MidY - 1), TextColor
        ctlControl.Line (MidX - 1, MidY - 2)-(MidX + 1, MidY - 2), TextColor
        ctlControl.Line (MidX, MidY - 3)-(MidX, MidY - 3), TextColor
    Else
        ctlControl.Line (MidX - 3, MidY - 2)-(MidX + 3, MidY - 2), TextColor
        ctlControl.Line (MidX - 2, MidY - 1)-(MidX + 2, MidY - 1), TextColor
        ctlControl.Line (MidX - 1, MidY)-(MidX + 1, MidY), TextColor
        ctlControl.Line (MidX, MidY + 1)-(MidX, MidY + 1), TextColor
    End If
DrawBorders:
    ctlControl.Line (0, 0)-(ctlControl.ScaleWidth, 0), BordersColor
    ctlControl.Line (0, ctlControl.ScaleHeight - 1)-(ctlControl.ScaleWidth, ctlControl.ScaleHeight - 1), BordersColor
    ctlControl.Line (0, 0)-(0, ctlControl.ScaleHeight), BordersColor
    ctlControl.Line (ctlControl.ScaleWidth - 1, 0)-(ctlControl.ScaleWidth - 1, ctlControl.ScaleHeight), BordersColor
    ctlControl.Refresh
    ctlControl.ScaleMode = OldScaleMode
End Sub

'=====================================================
' Draw Backgrounds
'=====================================================
Public Sub DrawPicBack(ctlControl As Control, StartColor As OLE_COLOR, EndColor As OLE_COLOR, BordersColor As OLE_COLOR, Optional NoLeftLine As Boolean, Optional NoTopLine As Boolean)
    On Error Resume Next
    DoEvents
    OldScaleMode = ctlControl.ScaleMode
    ctlControl.ScaleMode = 3
    If m_Style = Flat Or EndColor = &H8000000F Then
       ctlControl.BackColor = EndColor
       GoTo DrawBorders
    End If
    Call InitializeCol(ctlControl, StartColor, EndColor, False)
    For i = 0 To ctlControl.ScaleHeight
        NewColor = RGB(RedStart + i * RedI, GreenStart + i * GreenI, BlueStart + i * BlueI)
        ctlControl.Line (0, i)-(ctlControl.ScaleWidth, i), NewColor
    Next
DrawBorders:
    If NoTopLine <> True Then ctlControl.Line (0, 0)-(ctlControl.ScaleWidth, 0), BordersColor
    ctlControl.Line (0, ctlControl.ScaleHeight - 1)-(ctlControl.ScaleWidth, ctlControl.ScaleHeight - 1), BordersColor
    If NoLeftLine <> True Then ctlControl.Line (0, 0)-(0, ctlControl.ScaleHeight), BordersColor
    ctlControl.Line (ctlControl.ScaleWidth - 1, 0)-(ctlControl.ScaleWidth - 1, ctlControl.ScaleHeight), BordersColor
    ctlControl.Refresh
    ctlControl.ScaleMode = OldScaleMode
    DoEvents
End Sub

'=====================================================
' Initialize colors for controls
'=====================================================
Function InitializeCol(ctlControl As Control, StartColor As OLE_COLOR, EndColor As OLE_COLOR, Clear As Boolean)
    StartCol = StartColor
    EndCol = EndColor
    RedStart = StartCol Mod 256
    RedEnd = EndCol Mod 256
    RedI = (RedEnd - RedStart) / (ctlControl.ScaleHeight)
    GreenStart = (StartCol And &HFF00FF00) / 256
    GreenEnd = (EndCol And &HFF00FF00) / 256
    GreenI = (GreenEnd - GreenStart) / (ctlControl.ScaleHeight)
    BlueStart = (StartCol And &HFFFF0000) / (65536)
    BlueEnd = (EndCol And &HFFFF0000) / (65536)
    BlueI = (BlueEnd - BlueStart) / (ctlControl.ScaleHeight)
    If Clear = True Then ctlControl.Cls
End Function

'=====================================================
' Shift colors within colorrange
'=====================================================
Private Function ShiftColors(ByVal MyColor As Long, ByVal Base As Long) As Long
    Dim R As Long, G As Long, B As Long, Delta As Long
    R = (MyColor And &HFF)
    G = ((MyColor \ &H100) Mod &H100)
    B = ((MyColor \ &H10000) Mod &H100)
    Delta = &HFF - Base
    B = Base + B * Delta \ &HFF
    G = Base + G * Delta \ &HFF
    R = Base + R * Delta \ &HFF
    If R > 255 Then R = 255
    If G > 255 Then G = 255
    If B > 255 Then B = 255
    ShiftColors = R + 256& * G + 65536 * B
End Function


'=============================================================================================
' Usercontrol properties
'=============================================================================================

'=====================================================
' InitProperties
'=====================================================
Private Sub UserControl_InitProperties()
    ' Time
    m_Starttime = ""
    m_TimeFormat = m_def_TimeFormat
    m_TimeForeColor = m_def_TimeForeColor
    m_TimeBackColor = m_def_TimeBackColor
    m_TimeBorderColor = m_def_TimeBorderColor
    ' Current Time
    m_ShowCurrentTime = m_def_ShowCurrentTime
    m_CurrentTimeForeColor = m_def_CurrentTimeForeColor
    m_CurrentTimeBackColor = m_def_CurrentTimeBackColor
    m_CurrentTimeBorderColor = m_def_CurrentTimeBorderColor
    ' Am/Pm
    m_AMPMForeColor = m_def_AMPMForeColor
    m_AMPMBackColor = m_def_AMPMBackColor
    m_AMPMBorderColor = m_def_AMPMBorderColor
    ' Hour
    m_HourBtnsForeColor = m_def_HourBtnsForeColor
    m_HourBtnsBackColor = m_def_HourBtnsBackColor
    m_HourBtnsBorderColor = m_def_HourBtnsBorderColor
    ' Minutes
    m_MinutesBtnsForeColor = m_def_MinutesBtnsForeColor
    m_MinutesBtnsBackColor = m_def_MinutesBtnsBackColor
    m_MinutesBtnsBorderColor = m_def_MinutesBtnsBorderColor
    ' Seconds
    m_SecondsBtnsForeColor = m_def_SecondsBtnsForeColor
    m_SecondsBtnsBackColor = m_def_SecondsBtnsBackColor
    m_SecondsBtnsBorderColor = m_def_SecondsBtnsBorderColor
    ' DisabledColor
    m_DisabledForeColor = m_def_DisabledForeColor
    m_DisabledBackColor = m_def_DisabledBackColor
    m_DisabledBorderColor = m_def_DisabledBorderColor
    'Fonts
    Set m_FontTime = Ambient.Font
    Set m_FontCurrentTime = Ambient.Font
    Set m_FontAmPm = Ambient.Font
    ' Selected to return
    m_SelectedTime = m_def_SelectedTime
    m_SelectedHours = m_def_SelectedHours
    m_SelectedMinutes = m_def_SelectedMinutes
    m_SelectedAmPm = m_def_SelectedAmPm
    ' Rest
    m_Enabled = m_def_Enabled
    m_Style = m_def_Style
    m_Language = m_def_Language
    m_SpinBtnInterval = m_def_SpinBtnInterval
End Sub

'=====================================================
' ReadProperties
'=====================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' Time
    m_Starttime = PropBag.ReadProperty("Starttime", "")
    m_TimeFormat = PropBag.ReadProperty("TimeFormat", m_def_TimeFormat)
    m_TimeForeColor = PropBag.ReadProperty("TimeForeColor", m_def_TimeForeColor)
    m_TimeBackColor = PropBag.ReadProperty("TimeBackColor", m_def_TimeBackColor)
    m_TimeBorderColor = PropBag.ReadProperty("TimeBorderColor", m_def_TimeBorderColor)
    ' Current Time
    m_ShowCurrentTime = PropBag.ReadProperty("ShowCurrentTime", m_def_ShowCurrentTime)
    m_CurrentTimeForeColor = PropBag.ReadProperty("CurrentTimeForeColor", m_def_CurrentTimeForeColor)
    m_CurrentTimeBackColor = PropBag.ReadProperty("CurrentTimeBackColor", m_def_CurrentTimeBackColor)
    m_CurrentTimeBorderColor = PropBag.ReadProperty("CurrentTimeBorderColor", m_def_CurrentTimeBorderColor)
    ' Am/Pm
    m_AMPMForeColor = PropBag.ReadProperty("AMPMForeColor", m_def_AMPMForeColor)
    m_AMPMBorderColor = PropBag.ReadProperty("AMPMBorderColor", m_def_AMPMBorderColor)
    m_AMPMBackColor = PropBag.ReadProperty("AMPMBackColor", m_def_AMPMBackColor)
    ' Hour
    m_HourBtnsForeColor = PropBag.ReadProperty("HourBtnsForeColor", m_def_HourBtnsForeColor)
    m_HourBtnsBackColor = PropBag.ReadProperty("HourBtnsBackColor", m_def_HourBtnsBackColor)
    m_HourBtnsBorderColor = PropBag.ReadProperty("HourBtnsBorderColor", m_def_HourBtnsBorderColor)
    ' Minutes
    m_MinutesBtnsForeColor = PropBag.ReadProperty("MinutesBtnsForeColor", m_def_MinutesBtnsForeColor)
    m_MinutesBtnsBackColor = PropBag.ReadProperty("MinutesBtnsBackColor", m_def_MinutesBtnsBackColor)
    m_MinutesBtnsBorderColor = PropBag.ReadProperty("MinutesBtnsBorderColor", m_def_MinutesBtnsBorderColor)
    ' Seconds
    m_SecondsBtnsForeColor = PropBag.ReadProperty("SecondsBtnsForeColor", m_def_SecondsBtnsForeColor)
    m_SecondsBtnsBackColor = PropBag.ReadProperty("SecondsBtnsBackColor", m_def_SecondsBtnsBackColor)
    m_SecondsBtnsBorderColor = PropBag.ReadProperty("SecondsBtnsBorderColor", m_def_SecondsBtnsBorderColor)
    ' DisabledColor
    m_DisabledBackColor = PropBag.ReadProperty("DisabledBackColor", m_def_DisabledBackColor)
    m_DisabledBorderColor = PropBag.ReadProperty("DisabledBorderColor", m_def_DisabledBorderColor)
    m_DisabledForeColor = PropBag.ReadProperty("DisabledForeColor", m_def_DisabledForeColor)
    ' Fonts
    Set m_FontTime = PropBag.ReadProperty("FontTime", Ambient.Font)
    Set m_FontCurrentTime = PropBag.ReadProperty("FontCurrentTime", Ambient.Font)
    Set m_FontAmPm = PropBag.ReadProperty("FontAmPm", Ambient.Font)
    ' Selected to return
    m_SelectedTime = PropBag.ReadProperty("SelectedTime", m_def_SelectedTime)
    m_SelectedHours = PropBag.ReadProperty("SelectedHours", m_def_SelectedHours)
    m_SelectedMinutes = PropBag.ReadProperty("SelectedMinutes", m_def_SelectedMinutes)
    m_SelectedAmPm = PropBag.ReadProperty("SelectedAmPm", m_def_SelectedAmPm)
    ' Rest
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_Language = PropBag.ReadProperty("Language", m_def_Language)
    m_SpinBtnInterval = PropBag.ReadProperty("SpinBtnInterval", m_def_SpinBtnInterval)
    ' Set labels to language
    lblCCT = "Current"
    lblCH = "HH"
    lblCM = "MM"
    lblCS = "SS"
    Select Case m_Language
        Case 0
            lblCCT = "Current"
        Case 1
            lblCCT = "Huidig"
            lblCH = "UU"
        Case 2
            lblCCT = "Actuel"
        Case 3
            lblCCT = "Aktuell"
            lblCH = "SS"
        Case 4
            lblCCT = "Attuale"
            lblCH = "OO"
        Case 5
            lblCCT = "Actual"
    End Select
    MaxHour = 12
    If m_TimeFormat = [24Hours] Then MaxHour = 24
    DrawControls
    tmrHup.Interval = m_SpinBtnInterval
    tmrHdown.Interval = m_SpinBtnInterval
    tmrMup.Interval = m_SpinBtnInterval
    tmrMdown.Interval = m_SpinBtnInterval
    tmrSup.Interval = m_SpinBtnInterval
    tmrSdown.Interval = m_SpinBtnInterval
    UserControl.Enabled = m_Enabled
End Sub

Private Sub UserControl_Resize()
    DrawControls
End Sub

'=====================================================
' WriteProperties
'=====================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' Time
    Call PropBag.WriteProperty("StartTime", m_Starttime, "")
    Call PropBag.WriteProperty("TimeFormat", m_TimeFormat, m_def_TimeFormat)
    Call PropBag.WriteProperty("TimeForeColor", m_TimeForeColor, m_def_TimeForeColor)
    Call PropBag.WriteProperty("TimeBackColor", m_TimeBackColor, m_def_TimeBackColor)
    Call PropBag.WriteProperty("TimeBorderColor", m_TimeBorderColor, m_def_TimeBorderColor)
    ' Current time
    Call PropBag.WriteProperty("ShowCurrentTime", m_ShowCurrentTime, m_def_ShowCurrentTime)
    Call PropBag.WriteProperty("CurrentTimeForeColor", m_CurrentTimeForeColor, m_def_CurrentTimeForeColor)
    Call PropBag.WriteProperty("CurrentTimeBackColor", m_CurrentTimeBackColor, m_def_CurrentTimeBackColor)
    Call PropBag.WriteProperty("CurrentTimeBorderColor", m_CurrentTimeBorderColor, m_def_CurrentTimeBorderColor)
    ' Am/Pm
    Call PropBag.WriteProperty("AMPMForeColor", m_AMPMForeColor, m_def_AMPMForeColor)
    Call PropBag.WriteProperty("AMPMBackColor", m_AMPMBackColor, m_def_AMPMBackColor)
    Call PropBag.WriteProperty("AMPMBorderColor", m_AMPMBorderColor, m_def_AMPMBorderColor)
    ' Hours
    Call PropBag.WriteProperty("HourBtnsForeColor", m_HourBtnsForeColor, m_def_HourBtnsForeColor)
    Call PropBag.WriteProperty("HourBtnsBackColor", m_HourBtnsBackColor, m_def_HourBtnsBackColor)
    Call PropBag.WriteProperty("HourBtnsBorderColor", m_HourBtnsBorderColor, m_def_HourBtnsBorderColor)
    ' Minutes
    Call PropBag.WriteProperty("MinutesBtnsForeColor", m_MinutesBtnsForeColor, m_def_MinutesBtnsForeColor)
    Call PropBag.WriteProperty("MinutesBtnsBackColor", m_MinutesBtnsBackColor, m_def_MinutesBtnsBackColor)
    Call PropBag.WriteProperty("MinutesBtnsBorderColor", m_MinutesBtnsBorderColor, m_def_MinutesBtnsBorderColor)
    ' Seconds
    Call PropBag.WriteProperty("SecondsBtnsForeColor", m_SecondsBtnsForeColor, m_def_SecondsBtnsForeColor)
    Call PropBag.WriteProperty("SecondsBtnsBackColor", m_SecondsBtnsBackColor, m_def_SecondsBtnsBackColor)
    Call PropBag.WriteProperty("SecondsBtnsBorderColor", m_SecondsBtnsBorderColor, m_def_SecondsBtnsBorderColor)
    ' Disabled
    Call PropBag.WriteProperty("DisabledForeColor", m_DisabledForeColor, m_def_DisabledForeColor)
    Call PropBag.WriteProperty("DisabledBackColor", m_DisabledBackColor, m_def_DisabledBackColor)
    Call PropBag.WriteProperty("DisabledBorderColor", m_DisabledBorderColor, m_def_DisabledBorderColor)
    ' Fonts
    Call PropBag.WriteProperty("FontTime", m_FontTime, Ambient.Font)
    Call PropBag.WriteProperty("FontCurrentTime", m_FontCurrentTime, Ambient.Font)
    Call PropBag.WriteProperty("FontAmPm", m_FontAmPm, Ambient.Font)
    ' Selected to return
    Call PropBag.WriteProperty("SelectedTime", m_SelectedTime, m_def_SelectedTime)
    Call PropBag.WriteProperty("SelectedHours", m_SelectedHours, m_def_SelectedHours)
    Call PropBag.WriteProperty("SelectedMinutes", m_SelectedMinutes, m_def_SelectedMinutes)
    Call PropBag.WriteProperty("SelectedAmPm", m_SelectedAmPm, m_def_SelectedAmPm)
    ' Rest
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("Language", m_Language, m_def_Language)
    Call PropBag.WriteProperty("SpinBtnInterval", m_SpinBtnInterval, m_def_SpinBtnInterval)
End Sub

'=====================================================
' Get,Set and Let
'=====================================================

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get SelectedTime() As String
    SelectedTime = Trim$(lblH) & ":" & Trim$(lblM) & ":" & Trim$(lblS)
    If m_TimeFormat = 1 Then SelectedTime = SelectedTime & " " & lblAMPM.Caption
End Property
Public Property Let SelectedTime(ByVal New_SelectedTime As String)
    PropertyChanged "SelectedTime"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get SelectedHours() As String
    SelectedHours = Trim$(lblH)
End Property
Public Property Let SelectedHours(ByVal New_SelectedHours As String)
    PropertyChanged "SelectedHours"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get SelectedMinutes() As String
    SelectedMinutes = Trim$(lblM)
End Property
Public Property Let SelectedMinutes(ByVal New_SelectedMinutes As String)
    PropertyChanged "SelectedMinutes"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get SelectedSeconds() As String
    SelectedSeconds = Trim$(lblS)
End Property
Public Property Let SelectedSeconds(ByVal New_SelectedSeconds As String)
    PropertyChanged "SelectedSeconds"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get SelectedAmPm() As String
    SelectedAmPm = Trim$(lblAMPM)
End Property
Public Property Let SelectedAmPm(ByVal New_SelectedAmPm As String)
    PropertyChanged "SelectedAmPm"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TimeBorderColor() As OLE_COLOR
    TimeBorderColor = m_TimeBorderColor
End Property
Public Property Let TimeBorderColor(ByVal New_TimeBorderColor As OLE_COLOR)
    m_TimeBorderColor = New_TimeBorderColor
    PropertyChanged "TimeBorderColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TimeBackColor() As OLE_COLOR
    TimeBackColor = m_TimeBackColor
End Property
Public Property Let TimeBackColor(ByVal New_TimeBackColor As OLE_COLOR)
    m_TimeBackColor = New_TimeBackColor
    PropertyChanged "TimeBackColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TimeForeColor() As OLE_COLOR
    TimeForeColor = m_TimeForeColor
End Property
Public Property Let TimeForeColor(ByVal New_TimeForeColor As OLE_COLOR)
    m_TimeForeColor = New_TimeForeColor
    PropertyChanged "TimeForeColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CurrentTimeBackColor() As OLE_COLOR
    CurrentTimeBackColor = m_CurrentTimeBackColor
End Property
Public Property Let CurrentTimeBackColor(ByVal New_CurrentTimeBackColor As OLE_COLOR)
    m_CurrentTimeBackColor = New_CurrentTimeBackColor
    PropertyChanged "CurrentTimeBackColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CurrentTimeForeColor() As OLE_COLOR
    CurrentTimeForeColor = m_CurrentTimeForeColor
End Property
Public Property Let CurrentTimeForeColor(ByVal New_CurrentTimeForeColor As OLE_COLOR)
    m_CurrentTimeForeColor = New_CurrentTimeForeColor
    PropertyChanged "CurrentTimeForeColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CurrentTimeBorderColor() As OLE_COLOR
    CurrentTimeBorderColor = m_CurrentTimeBorderColor
End Property
Public Property Let CurrentTimeBorderColor(ByVal New_CurrentTimeBorderColor As OLE_COLOR)
    m_CurrentTimeBorderColor = New_CurrentTimeBorderColor
    PropertyChanged "CurrentTimeBorderColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get AMPMForeColor() As OLE_COLOR
    AMPMForeColor = m_AMPMForeColor
End Property
Public Property Let AMPMForeColor(ByVal New_AMPMForeColor As OLE_COLOR)
    m_AMPMForeColor = New_AMPMForeColor
    PropertyChanged "AMPMForeColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get AMPMBorderColor() As OLE_COLOR
    AMPMBorderColor = m_AMPMBorderColor
End Property
Public Property Let AMPMBorderColor(ByVal New_AMPMBorderColor As OLE_COLOR)
    m_AMPMBorderColor = New_AMPMBorderColor
    PropertyChanged "AMPMBorderColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get AMPMBackColor() As OLE_COLOR
    AMPMBackColor = m_AMPMBackColor
End Property
Public Property Let AMPMBackColor(ByVal New_AMPMBackColor As OLE_COLOR)
    m_AMPMBackColor = New_AMPMBackColor
    PropertyChanged "AMPMBackColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ShowCurrentTime() As Boolean
    ShowCurrentTime = m_ShowCurrentTime
End Property
Public Property Let ShowCurrentTime(ByVal New_ShowCurrentTime As Boolean)
    m_ShowCurrentTime = New_ShowCurrentTime
    PropertyChanged "ShowCurrentTime"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Style() As Styles
    Style = m_Style
End Property
Public Property Let Style(ByVal New_Style As Styles)
    m_Style = New_Style
    PropertyChanged "Style"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get TimeFormat() As TimeFormats
    TimeFormat = m_TimeFormat
End Property
Public Property Let TimeFormat(ByVal New_TimeFormat As TimeFormats)
    m_TimeFormat = New_TimeFormat
    PropertyChanged "TimeFormat"
    MaxHour = 12
    If m_TimeFormat = [24Hours] Then MaxHour = 24
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get StartTime() As String
    StartTime = m_Starttime
End Property
Public Property Let StartTime(ByVal New_StartTime As String)
    ' Vallidation on timeformats
    New_StartTime = Trim$(New_StartTime)
    If New_StartTime <> "" Then
        If Mid$(New_StartTime, 3, 1) <> ":" Then GoTo ErrHandling
        If Val(Left$(New_StartTime, 2)) > MaxHour Or Val(Left$(New_StartTime, 2)) < 0 Then GoTo ErrHandling
        If Val(Mid$(New_StartTime, 4, 2)) > 59 Or Val(Left$(New_StartTime, 2)) < 0 Then GoTo ErrHandling
        If TimeFormat = [12Hours] Then
            If Len(New_StartTime) > 8 Then
                If Mid$(New_StartTime, 9, 1) <> " " Then GoTo ErrHandling
                If Val(Mid$(New_StartTime, 7, 2)) > 59 Or Val(Mid$(New_StartTime, 7, 2)) < 0 Then GoTo ErrHandling
            Else
                If Mid$(New_StartTime, 6, 1) <> " " Then GoTo ErrHandling
            End If
            If UCase(Right$(New_StartTime, 2)) <> "AM" And UCase(Right$(New_StartTime, 2)) <> "PM" Then GoTo ErrHandling
        Else
            If Len(New_StartTime) > 8 Then GoTo ErrHandling
            If Len(New_StartTime) > 5 Then
                If Mid$(New_StartTime, 6, 1) <> ":" Then GoTo ErrHandling
                If Val(Right$(New_StartTime, 2)) > 59 Or Val(Right$(New_StartTime, 2)) < 0 Then GoTo ErrHandling
            Else
            End If
        End If
    End If
    m_Starttime = New_StartTime
    PropertyChanged "StartTime"
    DrawControls
    Exit Property
ErrHandling:
    MsgBox "Timeformat must be" & vbCrLf & "'hh:mm' or 'hh:mm:ss' or 'hh:mm AM' or 'hh:mm PM'", vbOKOnly + vbExclamation, "Invallid timeformat"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HourBtnsForeColor() As OLE_COLOR
    HourBtnsForeColor = m_HourBtnsForeColor
End Property
Public Property Let HourBtnsForeColor(ByVal New_HourBtnsForeColor As OLE_COLOR)
    m_HourBtnsForeColor = New_HourBtnsForeColor
    PropertyChanged "HourBtnsForeColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HourBtnsBackColor() As OLE_COLOR
    HourBtnsBackColor = m_HourBtnsBackColor
End Property
Public Property Let HourBtnsBackColor(ByVal New_HourBtnsBackColor As OLE_COLOR)
    m_HourBtnsBackColor = New_HourBtnsBackColor
    PropertyChanged "HourBtnsBackColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HourBtnsBorderColor() As OLE_COLOR
    HourBtnsBorderColor = m_HourBtnsBorderColor
End Property
Public Property Let HourBtnsBorderColor(ByVal New_HourBtnsBorderColor As OLE_COLOR)
    m_HourBtnsBorderColor = New_HourBtnsBorderColor
    PropertyChanged "HourBtnsBorderColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MinutesBtnsBackColor() As OLE_COLOR
    MinutesBtnsBackColor = m_MinutesBtnsBackColor
End Property
Public Property Let MinutesBtnsBackColor(ByVal New_MinutesBtnsBackColor As OLE_COLOR)
    m_MinutesBtnsBackColor = New_MinutesBtnsBackColor
    PropertyChanged "MinutesBtnsBackColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MinutesBtnsBorderColor() As OLE_COLOR
    MinutesBtnsBorderColor = m_MinutesBtnsBorderColor
End Property
Public Property Let MinutesBtnsBorderColor(ByVal New_MinutesBtnsBorderColor As OLE_COLOR)
    m_MinutesBtnsBorderColor = New_MinutesBtnsBorderColor
    PropertyChanged "MinutesBtnsBorderColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MinutesBtnsForeColor() As OLE_COLOR
    MinutesBtnsForeColor = m_MinutesBtnsForeColor
End Property
Public Property Let MinutesBtnsForeColor(ByVal New_MinutesBtnsForeColor As OLE_COLOR)
    m_MinutesBtnsForeColor = New_MinutesBtnsForeColor
    PropertyChanged "MinutesBtnsForeColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SecondsBtnsBackColor() As OLE_COLOR
    SecondsBtnsBackColor = m_SecondsBtnsBackColor
End Property
Public Property Let SecondsBtnsBackColor(ByVal New_SecondsBtnsBackColor As OLE_COLOR)
    m_SecondsBtnsBackColor = New_SecondsBtnsBackColor
    PropertyChanged "SecondsBtnsBackColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SecondsBtnsBorderColor() As OLE_COLOR
    SecondsBtnsBorderColor = m_SecondsBtnsBorderColor
End Property
Public Property Let SecondsBtnsBorderColor(ByVal New_SecondsBtnsBorderColor As OLE_COLOR)
    m_SecondsBtnsBorderColor = New_SecondsBtnsBorderColor
    PropertyChanged "SecondsBtnsBorderColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SecondsBtnsForeColor() As OLE_COLOR
    SecondsBtnsForeColor = m_SecondsBtnsForeColor
End Property
Public Property Let SecondsBtnsForeColor(ByVal New_SecondsBtnsForeColor As OLE_COLOR)
    m_SecondsBtnsForeColor = New_SecondsBtnsForeColor
    PropertyChanged "SecondsBtnsForeColor"
    DrawControls
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DisabledBackColor() As OLE_COLOR
    DisabledBackColor = m_DisabledBackColor
End Property
Public Property Let DisabledBackColor(ByVal New_DisabledBackColor As OLE_COLOR)
    m_DisabledBackColor = New_DisabledBackColor
    PropertyChanged "DisabledBackColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DisabledBorderColor() As OLE_COLOR
    DisabledBorderColor = m_DisabledBorderColor
End Property
Public Property Let DisabledBorderColor(ByVal New_DisabledBorderColor As OLE_COLOR)
    m_DisabledBorderColor = New_DisabledBorderColor
    PropertyChanged "DisabledBorderColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DisabledForeColor() As OLE_COLOR
    DisabledForeColor = m_DisabledForeColor
End Property
Public Property Let DisabledForeColor(ByVal New_DisabledForeColor As OLE_COLOR)
    m_DisabledForeColor = New_DisabledForeColor
    PropertyChanged "DisabledForeColor"
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get Language() As Languages
    Language = m_Language
End Property
Public Property Let Language(ByVal New_Language As Languages)
    m_Language = New_Language
    PropertyChanged "Language"
    ' Set labels to language
    lblCCT = "Current"
    lblCH = "HH"
    lblCM = "MM"
    lblCS = "SS"
    Select Case m_Language
        Case 0
            lblCCT = "Current"
        Case 1
            lblCCT = "Huidig"
            lblCH = "UU"
        Case 2
            lblCCT = "Actuel"
        Case 3
            lblCCT = "Aktuell"
            lblCH = "SS"
        Case 4
            lblCCT = "Attuale"
            lblCH = "OO"
        Case 5
            lblCCT = "Actual"
    End Select
    DrawControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get FontTime() As Font
    Set FontTime = m_FontTime
End Property
Public Property Set FontTime(ByVal New_FontTime As Font)
    Set m_FontTime = New_FontTime
    PropertyChanged "FontTime"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get FontCurrentTime() As Font
    Set FontCurrentTime = m_FontCurrentTime
End Property
Public Property Set FontCurrentTime(ByVal New_FontCurrentTime As Font)
    Set m_FontCurrentTime = New_FontCurrentTime
    PropertyChanged "FontCurrentTime"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get FontAmPm() As Font
    Set FontAmPm = m_FontAmPm
End Property
Public Property Set FontAmPm(ByVal New_FontAmPm As Font)
    Set m_FontAmPm = New_FontAmPm
    PropertyChanged "FontAmPm"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get SpinBtnInterval() As Integer
    SpinBtnInterval = m_SpinBtnInterval
End Property
Public Property Let SpinBtnInterval(ByVal New_SpinBtnInterval As Integer)
    m_SpinBtnInterval = New_SpinBtnInterval
    PropertyChanged "SpinBtnInterval"
    tmrHup.Interval = m_SpinBtnInterval
    tmrHdown.Interval = m_SpinBtnInterval
    tmrMup.Interval = m_SpinBtnInterval
    tmrMdown.Interval = m_SpinBtnInterval
    tmrSup.Interval = m_SpinBtnInterval
    tmrSdown.Interval = m_SpinBtnInterval
End Property


