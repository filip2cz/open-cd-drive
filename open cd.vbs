Set ovladani = Createobject("wMPlayer.OCX.7" )
Set mechaniky = ovladani.cdromCollection
mechaniky.Item(i).Eject
MsgBox "Vložte disk! / Vyjmìte disk! "
mechaniky.Item(i).Eject


