Set ovladani = Createobject("wMPlayer.OCX.7" )
Set mechaniky = ovladani.cdromCollection
mechaniky.Item(i).Eject
MsgBox "Vlo�te disk! / Vyjm�te disk! "
mechaniky.Item(i).Eject


