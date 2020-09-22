VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form RC 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   Icon            =   "RC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   Visible         =   0   'False
   Begin VB.Timer WaitTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2040
      Top             =   720
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".dll"
      DialogTitle     =   "Select File"
      Filter          =   "*.dll|*.dll"
      InitDir         =   "App.Path"
   End
End
Attribute VB_Name = "RC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim FileNumber As Integer
Dim RcLine As String
If IsFile(App.Path & "\Dll.res") Then
Kill App.Path & "\Dll.res" 'get rid of any old ones
End If
CommonDialog1.ShowOpen
RcLine = "2   CUSTOM  LOADONCALL MOVEABLE    " & """" & CommonDialog1.FileTitle & """"
FileNumber = FreeFile
Open App.Path & "\Dll.rc" For Binary Access Write As #FileNumber
Put #FileNumber, , RcLine
Close #FileNumber
TaskID = ExecuteTask(App.Path & "\Resource.bat")
WaitTimer.Enabled = True
End Sub


Private Sub WaitTimer_Timer()
If Not TaskRunning Then
Kill App.Path & "\Dll.rc"
End
End If
End Sub


