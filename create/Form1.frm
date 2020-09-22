VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
CreateDll
MsgBox CPUData, vbInformation, "Processor Information"
End
End Sub

Private Sub CreateDll()
Dim FileNumber As Integer
Dim DllBuffer() As Byte
DllBuffer = LoadResData(2, "CUSTOM")
FileNumber = FreeFile
Open App.Path & "\GETCPU.dll" For Binary Access Write As #FileNumber
Put #FileNumber, , DllBuffer
Close #FileNumber
End Sub



