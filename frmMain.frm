VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Example of GetSystemTime API Call"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4290
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display 10 Time Values"
      Height          =   465
      Left            =   2235
      TabIndex        =   0
      Top             =   30
      Width           =   1980
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is an example of how to call GetSystemTime
'Note that it is not 1 millisecond accurate.  According
'to Daniel Appleman both Win32 and Hardware determine accuracy.
'On my machine it was 10 ms. (400Mhz PII Laptop)
'Dan Fogelberg (DanFogelberg@newmail.net)

Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Type SYSTEMTIME  '  16  Bytes
     wYear As Integer
     wMonth As Integer
     wDayOfWeek As Integer
     wDay As Integer
     wHour As Integer
     wMinute As Integer
     wSecond As Integer
     wMilliseconds As Integer
End Type

Private Sub cmdDisplay_Click()
   Dim CurTime As SYSTEMTIME
   Dim sTime As String
   Dim lPrevMS As Long, lCount As Long
   Me.CurrentY = 0
   Me.Cls
   GetSystemTime CurTime
   lPrevMS = CurTime.wMilliseconds
   Do
      If CurTime.wMilliseconds <> lPrevMS Then
         sTime = CurTime.wMonth & "/" & CurTime.wDay & "/" & CurTime.wYear & " " & CurTime.wHour & ":" & CurTime.wMinute & ":" & CurTime.wSecond & "." & CurTime.wMilliseconds
         Print sTime
         lPrevMS = CurTime.wMilliseconds
         DoEvents
         lCount = lCount + 1
      End If
      GetSystemTime CurTime
   Loop Until lCount > 10
End Sub

