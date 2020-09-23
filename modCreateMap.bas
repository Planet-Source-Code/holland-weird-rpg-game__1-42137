Attribute VB_Name = "modCreateMap"
Public CurrX As Integer, CurrY As Integer
Public ArrNum As Integer
Sub Addline(Picture As PictureBox, GGG As String)
For i = 1 To Len(GGG)
bbb = Mid$(GGG, i, 1)
Select Case UCase(bbb)
Case "W"
Load Form1.ObjPic(ArrNum)
Form1.ObjPic(ArrNum).Left = CurrX
Form1.ObjPic(ArrNum).Top = CurrY
Form1.ObjPic(ArrNum).Visible = True
Form1.ObjPic(ArrNum).Picture = Form1.Blue.Picture
ArrNum = Val(ArrNum) + 1
Case "G"
Load Form1.ObjPic(ArrNum)
Form1.ObjPic(ArrNum).Left = CurrX
Form1.ObjPic(ArrNum).Top = CurrY
Form1.ObjPic(ArrNum).Visible = True
Form1.ObjPic(ArrNum).Picture = Form1.Grass.Picture
ArrNum = Val(ArrNum) + 1
Case "T"
Load Form1.ObjPic(ArrNum)
Form1.ObjPic(ArrNum).Left = CurrX
Form1.ObjPic(ArrNum).Top = CurrY
Form1.ObjPic(ArrNum).Visible = True
Form1.ObjPic(ArrNum).Picture = Form1.picTree.Picture
ArrNum = Val(ArrNum) + 1
Case "B"
Load Form1.ObjPic(ArrNum)
Form1.ObjPic(ArrNum).Left = CurrX
Form1.ObjPic(ArrNum).Top = CurrY
Form1.ObjPic(ArrNum).Visible = True
Form1.ObjPic(ArrNum).Picture = Form1.Bush.Picture
ArrNum = Val(ArrNum) + 1
End Select
CurrX = Val(CurrX) + 480
Picture.CurrentX = CurrX
Next i
CurrY = Val(CurrY) + 480
Picture.CurrentY = CurrY
CurrX = 0
End Sub
