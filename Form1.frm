VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remaining Time Calculator"
   ClientHeight    =   1605
   ClientLeft      =   3930
   ClientTop       =   1620
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   6405
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Stop"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "500"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3120
      Top             =   1080
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label4 
      Caption         =   "ProgressBar value: 0"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "ProgressBar Max:"
      Height          =   435
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Percentage 0%"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "TimeLeft 00:00:00"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************
'
'   Time Remaining Calculator
'
'    by Kendall Cain
'    gabberhed@usa.net
'
'   This calculates the average time taken
'   for a progressbar to get to its current
'   percentage and calculates how much time
'   it will take for the progressbar to
'   reach 100%
'
'   All you do is:
'   1.) Declare a variable in the (General)
'       area of your form as a string
'   2.) When the progressbar starts call 'BeginProgress'
'   3.) When ever you make the progressbars value increase
'       get the percentage with the 'getPercentage' function
'       then call the 'TimeRemaining' function
'   4.) Voila! You will have the approxamate remaining time
'       for your progressbar.
'
'  *note: If you get your stopwatch out and see that it isnt
'         going second by second, its just like a download
'         time calculator, the closer you get to 100% the
'         more accurate it gets, also it depends on the
'         programmers competentcy to increment their progressbar
'         equally from 0 to 100. Meaning you dont increase it by
'         1 percent for some small process and then 10 percent
'         for an equally small process.   =)
'         Enjoy!
'**********************

Dim Percentage As String


Private Sub Command1_Click()
On Error GoTo Handler

'Reset progressbar value to zero
ProgressBar1.Value = 0

'set progressbar max value to the value in the text box
ProgressBar1.Max = Text1.Text


'Duh!
Timer1.Enabled = True

'assign variable the current percentage value for the progressbar
Percentage = getPercentage(ProgressBar1.Value, ProgressBar1.Max)

'call the BeginProgress to set the start time for the progressbar
BeginProgress

'Call the TimeRemaining function and display it in label1
Label1.Caption = "TimeLeft " & TimeRemaining(Percentage)

Exit Sub

Handler:
MsgBox "Error #" & Err.Number & vbCrLf & Err.Description, vbOKOnly, "Error"


End Sub

Private Sub Command2_Click()
'Halt Procedure
Timer1.Enabled = False
ProgressBar1.Value = 0
Label2.Caption = "Percentage 0%"
Label1.Caption = "TimeLeft 00:00:00"
Label4.Caption = "ProgressBar value: 0"

End Sub


Private Sub Timer1_Timer()

'increase progressbar value
ProgressBar1.Value = Val(ProgressBar1.Value + 1)

'disable timer when progressbar hits 100%
If ProgressBar1.Value = ProgressBar1.Max Then Timer1.Enabled = False

'display percentage value for the progressbar on label2
Label2.Caption = "Percentage " & getPercentage(ProgressBar1.Value, ProgressBar1.Max) & "%"

'display current progressbar value
Label4.Caption = "ProgressBar value: " & ProgressBar1.Value

'assign variable the current percentage value for the progressbar
Percentage = getPercentage(ProgressBar1.Value, ProgressBar1.Max)

'Call the TimeRemaining function and display it in label1
Label1.Caption = "TimeLeft " & TimeRemaining(Percentage)


End Sub
