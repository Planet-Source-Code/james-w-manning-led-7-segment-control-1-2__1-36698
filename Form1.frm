VERSION 5.00
Object = "*\ALED7Seg.vbp"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LED7Seg Control - You set the color and the character it uses, and voila! This is a control array to illustrate the capabilities."
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin LED7Segctl.LED7Seg LED7Seg2 
      Height          =   1020
      Left            =   2130
      TabIndex        =   8
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   1
      Left            =   8550
      TabIndex        =   0
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   2
      Left            =   7830
      TabIndex        =   1
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   3
      Left            =   7110
      TabIndex        =   2
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   4
      Left            =   6390
      TabIndex        =   3
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   5
      Left            =   5670
      TabIndex        =   4
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   6
      Left            =   4950
      TabIndex        =   5
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   7
      Left            =   4230
      TabIndex        =   6
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   8
      Left            =   3510
      TabIndex        =   7
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'    Dim i As Integer
'    Dim j As Integer
'    Const Alpha As String = "fedcba9876543210"
'    j = 0
'    For i = 0 To LED7Seg1.UBound
'        If j <= 5 Then
'            LED7Seg1(i).Color = j
'        Else
'            j = -1
'        End If
'        LED7Seg1(i).DrawLED Mid$(Alpha, i + 1, 1)
'        If i = 0 Then
'            LED7Seg1(i).DrawLED Mid$(Alpha, i + 1, 1) & "."
'        End If
'        j = j + 1
'    Next i
    Dim i As Integer
    Dim Tmp(1 To 8) As String
    i = 1
    Tmp(8) = "10000000" 'Top Right
    Tmp(7) = "01000000" 'Top
    Tmp(6) = "00100000" 'Top Left
    Tmp(5) = "00010000" 'Bottom Left
    Tmp(4) = "00001000" 'Bottom
    Tmp(3) = "00000100" 'Bottom Right
    Tmp(2) = "00000010" 'Middle
    Tmp(1) = "00000001" 'Decimal
    For i = 1 To 8
        LED7Seg1(i).DrawSpecLED Tmp(i)
    Next i
    LED7Seg2.DrawLED "8."
End Sub
