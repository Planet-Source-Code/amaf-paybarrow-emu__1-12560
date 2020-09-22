VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "paybarrow emu example"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   3360
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   1200
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Left            =   2160
      Top             =   2160
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3480
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "by amaf"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "not active."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "pw:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "id#:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-paybarrow emu example
'-by amaf
'-url: www.arnaf.org or email: amaf@email.com
'-
'-this example uses inet, and it has both a recover
'-agent and login check. i created this in about 10
'-minutes max. but i still except to see my name
'-somewhere in the credits for releasing the source.
'-this is organized by using different functions for
'-different operations. i will be releasing my .ocx
'-which will emulate all of the popular companies.

Sub LoginAccount()
retry:
'-cancel operation------]
Inet1.Cancel
ChangeCaption "logging in."
'-login/get string------]
LoginString$ = Inet1.OpenURL("http://www.paybarrow.com/wbAddr2/login.php3?id=" & Text1.Text & "&&password=" & Text2.Text & "&&vers=9")
'-time program out------]
TimeOut 1
If InStr(LoginString$, "0") Then
'-login good, next------]
ChangeCaption "logged in."
TimeOut 1
'-ok, find ping now------]
FindPing
Else
'-login bad, retry------]
ChangeCaption "retrying."
GoTo retry
End If
End Sub
Sub ChangeCaption(lData$)
Label5.Caption = lData$
End Sub
Sub FindPing()
retry2:
'-cancel operation------]
Inet1.Cancel
ChangeCaption "ping wait."
'-ping/get ping data------]
Ping$ = Inet1.OpenURL("http://www.paybarrow.com/wbAddr2/ping.txt")
TimeOut 1
If InStr(Ping$, "pong") Then
'-ping good, next------]
ChangeCaption "ping caught."
TimeOut 1
AdCount = 0
'-start emulation------]
StartEmulation
Else
'-ping bad, retry------]
ChangeCaption "retrying."
GoTo retry2
End If
End Sub
Sub StartEmulation()
'-turn on ad viewer------]
Timer1.Enabled = True   ']
Timer1.Interval = 10000 ']
'------------------------]
ChangeCaption "emulating bar."
'-turn on click-thru-----]
Timer2.Enabled = True   ']
Timer2.Interval = 60000 ']
'------------------------]
End Sub
Sub StopEmulation()
Timer1.Enabled = False
Timer1.Interval = 0
Timer2.Enabled = False
Timer2.Interval = 0
End Sub
Private Sub Command1_Click()
If Command1.Caption = "start" Then
LoginAccount
ChangeCaption "started.."
Command1.Caption = "stop"
Else
StopEmulation
ChangeCaption "stopped."
Command1.Caption = "start"
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub Timer1_Timer()
'-this will view all the ads, you can set the
'-speed by changing the interval.
'-run through ads------]
retry3:
On Error GoTo retry3
Inet1.Cancel
On Error GoTo retry3:
Inet1.OpenURL ("http://www.paybarrow.com/wbAddr2/upd.php3?id=" & Text1.Text & "&txt=" & Text1.Text & "&adcount=" & AdCount)
On Error GoTo retry3:
AdCount = AdCount + 1
ChangeCaption "ad count [" & AdCount & "]"
End Sub
Private Sub Timer2_Timer()
'-will click ad every 60 minutes
'-change interval for faster or slower speed
'-click ad------]
retry4:
On Error GoTo retry4
If TimerSet = 30 Then
On Error GoTo retry4
Inet1.Cancel
On Error GoTo retry4
Inet1.OpenURL ("http://www.paybarrow.com/wbAddr2/clickthru.php3?id=" & Text1.Text)
On Error GoTo retry4
ChangeCaption "click thru!"
End If
TimerSet = TimerSet + 1
End Sub
