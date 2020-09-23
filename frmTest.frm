VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DM TimePicker"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Enabled = False"
      Height          =   4695
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   2655
      Begin Project1.DmTimePicker DmTimePicker8 
         Height          =   1515
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   2672
         StartTime       =   "11:59:59 PM"
         TimeFormat      =   1
         TimeForeColor   =   49152
         TimeBackColor   =   14737632
         TimeBorderColor =   33023
         AMPMForeColor   =   33023
         AMPMBorderColor =   33023
         BeginProperty FontTime {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontCurrentTime {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontAmPm {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin Project1.DmTimePicker DmTimePicker7 
         Height          =   1515
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   2672
         StartTime       =   "11:59:59 PM"
         TimeFormat      =   1
         TimeForeColor   =   49152
         TimeBackColor   =   14737632
         TimeBorderColor =   33023
         AMPMForeColor   =   33023
         AMPMBorderColor =   33023
         BeginProperty FontTime {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontCurrentTime {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontAmPm {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Style           =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3D"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flat"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   17
         Top             =   2760
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Showcurrenttime = False      Timeformat = 24 Hours"
      Height          =   2295
      Left            =   2880
      TabIndex        =   11
      Top             =   5040
      Width           =   7455
      Begin Project1.DmTimePicker DmTimePicker3 
         Height          =   1515
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   2672
         StartTime       =   "11:59:59"
         TimeForeColor   =   33023
         TimeBackColor   =   12648384
         TimeBorderColor =   192
         ShowCurrentTime =   0   'False
         CurrentTimeForeColor=   192
         CurrentTimeBackColor=   32768
         CurrentTimeBorderColor=   192
         AMPMForeColor   =   15362350
         AMPMBackColor   =   32768
         AMPMBorderColor =   192
         HourBtnsBackColor=   32768
         HourBtnsBorderColor=   32768
         MinutesBtnsBackColor=   32768
         MinutesBtnsBorderColor=   32768
         SecondsBtnsForeColor=   16777215
         SecondsBtnsBackColor=   32768
         SecondsBtnsBorderColor=   32768
         BeginProperty FontTime {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontCurrentTime {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontAmPm {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin Project1.DmTimePicker DmTimePicker6 
         Height          =   1515
         Left            =   4800
         TabIndex        =   22
         Top             =   600
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   2672
         StartTime       =   "11:59:59"
         TimeForeColor   =   33023
         TimeBackColor   =   12648384
         TimeBorderColor =   192
         ShowCurrentTime =   0   'False
         CurrentTimeForeColor=   192
         CurrentTimeBackColor=   32768
         CurrentTimeBorderColor=   192
         AMPMForeColor   =   15362350
         AMPMBackColor   =   32768
         AMPMBorderColor =   192
         HourBtnsBackColor=   32768
         HourBtnsBorderColor=   32768
         MinutesBtnsBackColor=   32768
         MinutesBtnsBorderColor=   32768
         SecondsBtnsForeColor=   16777215
         SecondsBtnsBackColor=   32768
         SecondsBtnsBorderColor=   32768
         BeginProperty FontTime {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontCurrentTime {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontAmPm {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   2760
         TabIndex        =   14
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3D"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flat"
         Height          =   195
         Index           =   8
         Left            =   4800
         TabIndex        =   12
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Showcurrenttime = True      Timeformat = 24 Hours"
      Height          =   2295
      Left            =   2880
      TabIndex        =   4
      Top             =   2640
      Width           =   7455
      Begin Project1.DmTimePicker DmTimePicker5 
         Height          =   1515
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   2672
         StartTime       =   "11:59:59"
         TimeForeColor   =   192
         TimeBackColor   =   12632256
         TimeBorderColor =   192
         CurrentTimeForeColor=   14737632
         CurrentTimeBackColor=   0
         CurrentTimeBorderColor=   192
         AMPMForeColor   =   0
         AMPMBackColor   =   0
         AMPMBorderColor =   192
         HourBtnsForeColor=   192
         HourBtnsBackColor=   0
         HourBtnsBorderColor=   0
         MinutesBtnsForeColor=   192
         MinutesBtnsBackColor=   0
         MinutesBtnsBorderColor=   0
         SecondsBtnsForeColor=   192
         SecondsBtnsBackColor=   0
         SecondsBtnsBorderColor=   0
         BeginProperty FontTime {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontCurrentTime {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontAmPm {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin Project1.DmTimePicker DmTimePicker4 
         Height          =   1515
         Left            =   4800
         TabIndex        =   23
         Top             =   600
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   2672
         StartTime       =   "11:59:59"
         TimeForeColor   =   192
         TimeBackColor   =   12632256
         TimeBorderColor =   192
         CurrentTimeForeColor=   14737632
         CurrentTimeBackColor=   0
         CurrentTimeBorderColor=   192
         AMPMForeColor   =   0
         AMPMBackColor   =   0
         AMPMBorderColor =   192
         HourBtnsForeColor=   192
         HourBtnsBackColor=   0
         HourBtnsBorderColor=   0
         MinutesBtnsForeColor=   192
         MinutesBtnsBackColor=   0
         MinutesBtnsBorderColor=   0
         SecondsBtnsForeColor=   192
         SecondsBtnsBackColor=   0
         SecondsBtnsBorderColor=   0
         BeginProperty FontTime {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontCurrentTime {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontAmPm {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flat"
         Height          =   195
         Index           =   3
         Left            =   4800
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3D"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   210
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   2760
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Showcurrenttime = True      Timeformat = 12 Hours"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin Project1.DmTimePicker DmTimePicker1 
         Height          =   1455
         Left            =   2880
         TabIndex        =   1
         Top             =   720
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   2566
         StartTime       =   "11:59:59 PM"
         TimeFormat      =   1
         TimeForeColor   =   49152
         TimeBackColor   =   14737632
         TimeBorderColor =   33023
         AMPMForeColor   =   33023
         AMPMBorderColor =   33023
         BeginProperty FontTime {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontCurrentTime {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontAmPm {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
         SpinBtnInterval =   10
      End
      Begin Project1.DmTimePicker DmTimePicker2 
         Height          =   1515
         Left            =   7560
         TabIndex        =   2
         Top             =   720
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   2672
         StartTime       =   "11:59:59 PM"
         TimeFormat      =   1
         TimeForeColor   =   49152
         TimeBackColor   =   14737632
         TimeBorderColor =   33023
         AMPMForeColor   =   33023
         AMPMBorderColor =   33023
         BeginProperty FontTime {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontCurrentTime {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontAmPm {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Check speed of interval ==>  for spinbuttons"
         Height          =   555
         Index           =   5
         Left            =   720
         TabIndex        =   24
         Top             =   950
         Width           =   2145
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Click on current time sets ==>  time to current "
         Height          =   555
         Index           =   4
         Left            =   720
         TabIndex        =   20
         Top             =   1680
         Width           =   2145
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flat"
         Height          =   195
         Index           =   1
         Left            =   7560
         TabIndex        =   8
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3D"
         Height          =   195
         Index           =   0
         Left            =   2880
         TabIndex        =   7
         Top             =   480
         Width           =   210
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   5520
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Programmer:        Donckers Frank
'                    DarkManSoft@Gmail.com
'
' Description:       Testform Control TimePicker

Private Sub DmTimePicker1_Change()
    Label1 = "Time: " & DmTimePicker1.SelectedTime & vbCrLf
    Label1 = Label1 & "Hours: " & DmTimePicker1.SelectedHours & vbCrLf
    Label1 = Label1 & "Minutes: " & DmTimePicker1.SelectedMinutes & vbCrLf
    Label1 = Label1 & "Seconds: " & DmTimePicker1.SelectedSeconds & vbCrLf
    If DmTimePicker1.TimeFormat = [12Hours] Then
        Label1 = Label1 & "AM/PM: " & DmTimePicker1.SelectedAmPm & vbCrLf
    End If
End Sub

Private Sub DmTimePicker3_Change()
    Label6 = "Time: " & DmTimePicker3.SelectedTime & vbCrLf
    Label6 = Label6 & "Hours: " & DmTimePicker3.SelectedHours & vbCrLf
    Label6 = Label6 & "Minutes: " & DmTimePicker3.SelectedMinutes & vbCrLf
    Label6 = Label6 & "Seconds: " & DmTimePicker3.SelectedSeconds & vbCrLf
    If DmTimePicker3.TimeFormat = [12Hours] Then
        Label6 = Label6 & "AM/PM: " & DmTimePicker3.SelectedAmPm & vbCrLf
    End If

End Sub

Private Sub DmTimePicker5_Change()
    Label5 = "Time: " & DmTimePicker5.SelectedTime & vbCrLf
    Label5 = Label5 & "Hours: " & DmTimePicker5.SelectedHours & vbCrLf
    Label5 = Label5 & "Minutes: " & DmTimePicker5.SelectedMinutes & vbCrLf
    Label5 = Label5 & "Seconds: " & DmTimePicker5.SelectedSeconds & vbCrLf
    If DmTimePicker5.TimeFormat = [12Hours] Then
        Label5 = Label5 & "AM/PM: " & DmTimePicker5.SelectedAmPm & vbCrLf
    End If
End Sub
