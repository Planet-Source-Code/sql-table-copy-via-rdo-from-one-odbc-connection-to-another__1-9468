Attribute VB_Name = "modMain"
Option Explicit

Public Function SetOn(TabX As SSTab, Index As Integer, FormX As Form)
Dim ii As Integer

With TabX
    'Set arrows to grey
    For ii = 0 To .Tabs - 1
        .TabPicture(ii) = FormX.imgOff.Picture
    Next ii
    
    'Set arrow you Yellow
    .TabPicture(Index) = FormX.imgOn.Picture
End With
DoEvents
End Function
