VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   360
   ClientLeft      =   7425
   ClientTop       =   3900
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleWidth      =   3255
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   60
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   15
      Width           =   2310
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Synch"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2445
      TabIndex        =   0
      Top             =   15
      Width           =   780
   End
   Begin MSWinsockLib.Winsock StinkySock 
      Left            =   990
      Top             =   375
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'coded by :: em ::® 2000
'http://www.em.f2s.com
'Synchronise your system time with an NTP server
'Made for CZ of www.freevbcode.com - check it out
'SyncClock Sub found online - by Paul Hews - hacked up by
'me though >=)
'hope you learn something =)

'The 32bit time stamp returned by NTP servers is the number of
'seconds since 0000 (midnight) 1 January 1900 GMT, such that the
'time "1" is 12:00:01 am on 1 January 1900 GMT; this base will
'serve until the year 2036.

Private Declare Function SetSystemTime Lib "kernel32" _
   (lpSystemTime As SYSTEMTIME) As Long
   
Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Dim sNTP As String ' the 32bit time stamp returned by the server
Dim TimeDelay As Single 'the time between the acknowledgement of
                        'the connection and the data received.
                        'we compensate by adding half of the round
                        'trip latency

Private Sub Command1_Click()
StinkySock.Close
sNTP = Empty
StinkySock.RemoteHost = Combo1.Text
StinkySock.RemotePort = 37 'NTP servers port
StinkySock.Connect
End Sub

Private Sub Form_Load()
Combo1.AddItem "ntp.cs.mu.oz.au"
Combo1.AddItem "tock.usno.navy.mil"
Combo1.AddItem "tick.usno.navy.mil"
Combo1.AddItem "swisstime.ethz.ch"
Combo1.AddItem "ntp-cup.external.hp.com"
Combo1.AddItem "ntp1.fau.de"
Combo1.AddItem "ntps1-0.cs.tu-berlin.de"
Combo1.AddItem "time.ien.it"
Combo1.AddItem "ntps1-1.rz.Uni-Osnabrueck.DE"
Combo1.AddItem "tempo.cstv.to.cnr.it"
Combo1.ListIndex = 0
End Sub

Private Sub StinkySock_DataArrival(ByVal bytesTotal As Long)
Dim Data As String

StinkySock.GetData Data, vbString
sNTP = sNTP & Data
End Sub

Private Sub StinkySock_Connect()
TimeDelay = Timer
End Sub

Private Sub StinkySock_Close()
On Error Resume Next
Do Until StinkySock.State = sckClosed
 StinkySock.Close
 DoEvents
Loop
TimeDelay = ((Timer - TimeDelay) / 2)
Call SyncClock(sNTP)
End Sub

Private Sub SyncClock(tStr As String)
Dim NTPTime As Double
Dim UTCDATE As Date
Dim LngTimeFrom1990 As Long
Dim ST As SYSTEMTIME
     
tStr = Trim(tStr)
If Len(tStr) <> 4 Then
 MsgBox "NTP Server returned an invalid response.", vbCritical, "Invalid Response"
 Exit Sub
End If

NTPTime = Asc(Left$(tStr, 1)) * 256 ^ 3 + Asc(Mid$(tStr, 2, 1)) * 256 ^ 2 + Asc(Mid$(tStr, 3, 1)) * 256 ^ 1 + Asc(Right$(tStr, 1))
      
LngTimeFrom1990 = NTPTime - 2840140800#

UTCDATE = DateAdd("s", CDbl(LngTimeFrom1990 + CLng(TimeDelay)), #1/1/1990#)

ST.wYear = Year(UTCDATE)
ST.wMonth = Month(UTCDATE)
ST.wDay = Day(UTCDATE)
ST.wHour = Hour(UTCDATE)
ST.wMinute = Minute(UTCDATE)
ST.wSecond = Second(UTCDATE)

Call SetSystemTime(ST)
MsgBox "Clock synchronised succesfully.", vbInformation, "Tick Tock"

End Sub
