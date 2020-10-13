Set oWMP = CreateObject("WMPlayer.OCX.7" )
Set colCDROMs = oWMP.cdromCollection
do
colCDROMs.Item(i).Eject
wscript.sleep 2000
loop
