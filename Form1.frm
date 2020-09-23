VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GAMEMAP EXAMPLE"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Height          =   5840
      Left            =   0
      ScaleHeight     =   5775
      ScaleWidth      =   7200
      TabIndex        =   0
      Top             =   0
      Width           =   7260
      Begin VB.Image LeftMove 
         Height          =   480
         Left            =   1560
         Picture         =   "Form1.frx":0000
         Top             =   960
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Right 
         Height          =   480
         Left            =   960
         Picture         =   "Form1.frx":0C42
         Top             =   960
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Up 
         Height          =   480
         Left            =   480
         Picture         =   "Form1.frx":1884
         Top             =   960
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Down 
         Height          =   480
         Left            =   0
         Picture         =   "Form1.frx":24C6
         Top             =   960
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Bush 
         Height          =   480
         Left            =   1440
         Picture         =   "Form1.frx":3108
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image picTree 
         Height          =   480
         Left            =   960
         Picture         =   "Form1.frx":3D4A
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Grass 
         Height          =   480
         Left            =   480
         Picture         =   "Form1.frx":498C
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Blue 
         Height          =   480
         Left            =   0
         Picture         =   "Form1.frx":55CE
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ObjPic 
         Height          =   480
         Index           =   0
         Left            =   0
         Top             =   480
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   960
         Picture         =   "Form1.frx":6210
         Top             =   960
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Block As Integer

Private Sub Form_Load()
Dim Hallo As String
ArrNum = 1
Open App.Path & "\MainMap.map" For Input As #1
Do While Not EOF(1)
Line Input #1, a$
Call Addline(Picture1, a$)
DoEvents
Loop
Close #1
Picture1.ZOrder vbBringToFront
Form1.Caption = Picture1.Width & " - " & Picture1.Width
End Sub


Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDown Then
Block = 0
Image1.Picture = Down.Picture
For i = 1 To ObjPic.Count - 1
Form1.Caption = i
If Collision(Image1.Left, Image1.Top, ObjPic(i).Left, ObjPic(i).Top - ObjPic(i).Height, Image1.Height) = True Then
Block = 1
Exit For
Else
Block = 0
End If
Next i
If Block = 0 Then
Image1.Top = Val(Image1.Top) + 480
Picture1.Top = Val(Picture1.Top) - 480
End If
End If
If KeyCode = vbKeyUp Then
Block = 0
Image1.Picture = Up.Picture
For i = 1 To ObjPic.Count - 1
Form1.Caption = i
If Collision(Image1.Left, Image1.Top, ObjPic(i).Left, ObjPic(i).Top + Image1.Height, ObjPic(i).Height) = True Then
Block = 1
Exit For
Else
Block = 0
End If
Next i
If Block = 0 Then
Image1.Top = Val(Image1.Top) - 480
Picture1.Top = Val(Picture1.Top) + 480
End If
End If
If KeyCode = vbKeyRight Then
Image1.Picture = Right.Picture
For i = 1 To ObjPic.Count - 1
Form1.Caption = i
If Collision(Image1.Left, Image1.Top, ObjPic(i).Left - Image1.Width, ObjPic(i).Top, ObjPic(i).Height) = True Then
Block = 1
Exit For
Else
Block = 0
End If
Next i
If Block = 0 Then
Image1.Left = Val(Image1.Left) + 480
Picture1.Left = Val(Picture1.Left) - 480
End If
End If
If KeyCode = vbKeyLeft Then
Image1.Picture = LeftMove.Picture
For i = 1 To ObjPic.Count - 1
Form1.Caption = i
If Collision(Image1.Left, Image1.Top, ObjPic(i).Left + Image1.Width, ObjPic(i).Top, ObjPic(i).Height) = True Then
Block = 1
Exit For
Else
Block = 0
End If
Next i
If Block = 0 Then
Image1.Left = Val(Image1.Left) - 480
Picture1.Left = Val(Picture1.Left) + 480
End If
End If
End Sub

Public Function Collision(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, Radius As Single) As Boolean

 'Check for the collision
 
 'Radius would be the Width or Height of your sprite(or whatever you're using)
 If X1 > X2 - Radius And X1 < X2 + Radius And _
    Y1 > Y2 - Radius And Y1 < Y2 + Radius Then
  Collision = True
 Else
  Collision = False
 End If

End Function

