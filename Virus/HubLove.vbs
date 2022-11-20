Set oWMP = CreateObject("WMPlayer.OCX.7")
Set colCDROMs = oWMP.cdromCollection
if colCDROMs.Count >= 1 then
for i = 0 to colCDROMs.Count - 1
colCDROMs.Item(i).Eject
colCDROMs.Item(i).Eject
Next ' cdrom
End if
lol=msgbox("Formattare Computer?",20,"Errore Sistema" )
lol=msgbox("Errore nel Sistema , il computer sta per essere formattato",50,"Imprevisto Microsoft Windows" )
set wshshell = wscript.CreateObject("wscript.shell")
wshshell.run "Notepad"
wscript.sleep 2000
wshshell.AppActivate "Notepad"
WshShell.SendKeys "ciao "
WScript.Sleep 500
WshShell.SendKeys " sono "
WScript.Sleep 500
WshShell.SendKeys " un "
WScript.Sleep 500
WshShell.SendKeys " Hacker "
WScript.Sleep 500
WshShell.SendKeys " ora "
WScript.Sleep 500
WshShell.SendKeys " sto "
WScript.Sleep 500
WshShell.SendKeys " dentro "
WScript.Sleep 500
WshShell.SendKeys " il tuo pc "
