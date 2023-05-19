Set ovladani = Createobject("wMPlayer.OCX.7" )
Set mechaniky = ovladani.cdromCollection
If mechaniky.count Then
For i = O To mechaniky.Count - 1
mechaniky.Item(i).Eject
MsgBox "Vložte disk! / Vyjmìte disk! "
mechaniky.Item(i).Eject
Next
End If


