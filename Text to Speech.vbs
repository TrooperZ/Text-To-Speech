Dim sapi
Set WshShell = WScript.CreateObject("WScript.Shell")
Set sapi=CreateObject("sapi.spvoice")
x=msgbox("Welcome to the Text to Speech converter made by Trooper Z!", 0+64, "Text to Speech")
sapi.Speak "Welcome to the Text to Speech converter made by Trooper Z!"
do
message=InputBox("Please type your text here. If you want to exit, press OK with an empty box.", "Text to Speech")
wscript.sleep 10
sapi.Speak message
loop Until message=""
x=msgbox("Thanks for using the Text to Speech analyzer. We hope you use it again!", 0+64, "Text to Speech")
sapi.Speak "Thanks for using the Text to Speech analyzer. We hope you use it again!"