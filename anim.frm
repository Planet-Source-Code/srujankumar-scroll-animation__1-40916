VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "Animation  Form"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   347
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   518
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5880
      Top             =   1260
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Option4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   870
      TabIndex        =   3
      Top             =   2325
      Width           =   2175
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   840
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Option3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1845
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   840
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Option2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   1365
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   840
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   885
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   840
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   " Sub Option1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   900
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   " Sub Option2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   900
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   " Sub Option3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   1380
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   " Sub Option4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   1380
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   " Sub Option5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   1890
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   " Sub Option6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   1890
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   " Sub Option7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   2340
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   " Sub Option8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   2340
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   4905
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4275
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   8745
      Left            =   -60
      Top             =   -120
      Width           =   2595
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
'*******************************************************'
'                                                       '
'       CREATED BY SRUJANKUMAR                          '
'                                                       '
'                                                       '
'                                                       '
'*******************************************************'
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 49 Then
Label1_Click
ElseIf KeyAscii = 50 Then
Label2_Click
ElseIf KeyAscii = 51 Then
Label3_Click
ElseIf KeyAscii = 52 Then
Label4_Click
End If
End Sub
Private Sub Label1_Click()
If Label6.Visible = True Then
i = 1
Else
i = 0
End If
Timer1.Enabled = True
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbRed
Label1.FontBold = True
Shape1.BorderColor = vbRed
End Sub

Private Sub Label2_Click()
If Label8.Visible = True Then
i = 3
Else
i = 2
End If
Timer1.Enabled = True
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlue
Label2.FontBold = True
Shape2.BorderColor = vbBlue
End Sub

Private Sub Label3_Click()
If Label10.Visible = True Then
i = 5
Else
i = 4
End If
Timer1.Enabled = True
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = RGB(30, 120, 120)
Label3.FontBold = True
Shape3.BorderColor = RGB(30, 120, 120)

End Sub

Private Sub Label4_Click()
If Label13.Visible = True Then
i = 7
Else
i = 6
End If
Timer1.Enabled = True
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = RGB(200, 120, 20)
Label4.FontBold = True
Shape4.BorderColor = RGB(200, 120, 20)
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbBlack
Label2.ForeColor = vbBlack
Label3.ForeColor = vbBlack
Label4.ForeColor = vbBlack
Label5.ForeColor = vbBlack
Label6.ForeColor = vbBlack
Label7.ForeColor = vbBlack
Label8.ForeColor = vbBlack
Label9.ForeColor = vbBlack
Label10.ForeColor = vbBlack
Label11.ForeColor = vbBlack
Label12.ForeColor = vbBlack
Label13.ForeColor = vbBlack
Shape1.BorderColor = vbBlack
Shape2.BorderColor = vbBlack
Shape3.BorderColor = vbBlack
Shape4.BorderColor = vbBlack
Label1.FontBold = False
Label2.FontBold = False
Label3.FontBold = False
Label4.FontBold = False
End Sub

Private Sub Label6_Click()
MsgBox "You have currently clicked sub option1", vbInformation + vbYesNo, "Animation"
End Sub
Private Sub Label7_Click()
MsgBox "You have currently clicked sub option2", vbInformation + vbYesNo, "Animation"
End Sub
Private Sub Label8_Click()
MsgBox "You have currently clicked sub option3", vbInformation + vbYesNo, "Animation"
End Sub
Private Sub Label9_Click()
MsgBox "You have currently clicked sub option4", vbInformation + vbYesNo, "Animation"
End Sub
Private Sub Label10_Click()
MsgBox "You have currently clicked sub option5", vbInformation + vbYesNo, "Animation"
End Sub
Private Sub Label11_Click()
MsgBox "You have currently clicked sub option6", vbInformation + vbYesNo, "Animation"
End Sub
Private Sub Label12_Click()
MsgBox "You have currently clicked sub option7", vbInformation + vbYesNo, "Animation"
End Sub
Private Sub Label13_Click()
MsgBox "You have currently clicked sub option8", vbInformation + vbYesNo, "Animation"
End Sub
Private Sub Timer1_Timer()
Select Case i
Case 0
    If Label8.Visible = True Then
    option2up
    ElseIf Label10.Visible = True Then
    option3up
    ElseIf Label13.Visible = True Then
    option4up
    Else
    option1dn
    End If
Case 1
    option1up
Case 2
    If Label6.Visible = True Then
    option1up
    ElseIf Label10.Visible = True Then
    option3up
    ElseIf Label13.Visible = True Then
    option4up
    Else
    option2dn
    End If
Case 3
    option2up
Case 4
    If Label6.Visible = True Then
    option1up
    ElseIf Label8.Visible = True Then
    option2up
    ElseIf Label13.Visible = True Then
    option4up
    Else
    option3dn
    End If
Case 5
    option3up
Case 6
    
    If Label6.Visible = True Then
    option1up
    ElseIf Label8.Visible = True Then
    option2up
    ElseIf Label10.Visible = True Then
    option3up
    Else
    option4dn
    End If
 Case 7
    option4up
End Select
End Sub
Sub option1dn()
If Label4.Top >= 230 Then
Timer1.Enabled = False
Else
Label2.Top = Label2.Top + 5
Shape2.Top = Shape2.Top + 5
Label3.Top = Label3.Top + 5
Shape3.Top = Shape3.Top + 5
Label4.Top = Label4.Top + 5
Shape4.Top = Shape4.Top + 5
Label6.Visible = True
Label7.Visible = True
Label6.Top = Label6.Top + 2.5
Label7.Top = Label7.Top + 4.5
End If
End Sub
Sub option1up()
If Label4.Top <= 155 Then
Label6.Visible = False
Label7.Visible = False
Else
Label2.Top = Label2.Top - 5
Shape2.Top = Shape2.Top - 5
Label3.Top = Label3.Top - 5
Shape3.Top = Shape3.Top - 5
Label4.Top = Label4.Top - 5
Shape4.Top = Shape4.Top - 5
Label6.Top = Label6.Top - 2.5
Label7.Top = Label7.Top - 4.5
End If
End Sub
Sub option2dn()
If Label4.Top >= 230 Then
Timer1.Enabled = False
Else
Label3.Top = Label3.Top + 5
Shape3.Top = Shape3.Top + 5
Label4.Top = Label4.Top + 5
Shape4.Top = Shape4.Top + 5
Label8.Visible = True
Label9.Visible = True
Label8.Top = Label8.Top + 2.5
Label9.Top = Label9.Top + 4.5
End If
End Sub
Sub option2up()
If Label4.Top <= 155 Then
Label8.Visible = False
Label9.Visible = False
Else
Label3.Top = Label3.Top - 5
Shape3.Top = Shape3.Top - 5
Label4.Top = Label4.Top - 5
Shape4.Top = Shape4.Top - 5
Label8.Top = Label8.Top - 2.5
Label9.Top = Label9.Top - 4.5
End If
End Sub
Sub option3dn()
If Label4.Top >= 230 Then
Timer1.Enabled = False
Else
Label10.Visible = True
Label11.Visible = True
Label10.Top = Label10.Top + 2.5
Label11.Top = Label11.Top + 4.5
Label4.Top = Label4.Top + 5
Shape4.Top = Shape4.Top + 5
End If
End Sub
Sub option3up()
If Label4.Top <= 155 Then
Label10.Visible = False
Label11.Visible = False
Else
Label4.Top = Label4.Top - 5
Shape4.Top = Shape4.Top - 5
Label10.Top = Label10.Top - 2.5
Label11.Top = Label11.Top - 4.5
End If
End Sub
Sub option4dn()
If Label13.Top >= 220 Then
Timer1.Enabled = False
Else
Label12.Visible = True
Label13.Visible = True
Label12.Top = Label12.Top + 2.5
Label13.Top = Label13.Top + 4.5
End If
End Sub
Sub option4up()
If Label13.Top <= 158 Then
Label12.Visible = False
Label13.Visible = False
Else
Label12.Top = Label12.Top - 2.5
Label13.Top = Label13.Top - 4.5
End If
End Sub
