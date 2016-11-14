Attribute VB_Name = "ModString"
Public Const Onelong = 720
Public TheRestart


Public NumLevel

Public Sub RedrawPic(Target As PictureBox, Source As StdPicture)
Target.PaintPicture Source, 0, 0, Target.Width, Target.Height
End Sub
