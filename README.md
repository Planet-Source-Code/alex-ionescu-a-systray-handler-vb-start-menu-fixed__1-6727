<div align="center">

## A Systray Handler/VB Start Menu \*\*\*FIXED\*\*\*


</div>

### Description

A lot of people have been asking me for this. I've finally managed to create this program and I hope you will all enjoy it. It is a System Tray Handler, like Litestep has. This is great for shells... It will make a small picture box and load the system tray buttons that usually show up in the Windows Explorer... If you open ICQ and then minimize it, it will show it's icon and you can fully interact with it. You can specify height (I recommend 25), XPOS, YPOS, whether it should auto-hide and if you just want an empty one. But wait! There's EVEN MORE! It scans your registry for Run/RunOnce entries and runs the startup programs in the registry at startup (it's explorer that does this, not windows itself... If you have a replacement shell, it won't run the apps by itself). You think that's not enough? Well, I've also included a Dynamic Start Menu that you can use! Just click the button on the form, and it will show you a menu (with icons/backround pic/office style selection) that shows all your groups. When you go over a group, it'll show you the programs inside. I've just finished this so I haven't implemented everything. Click on a program in the start menu won't run the program (yet) and the icons for the programs will all be the same (for now). I'm very excited about this and I hope you are too. VB Shells should start popping up everywhere now!
 
### More Info
 
hInst, height, XPos, YPos, autoHiding, useEmpty

Directory to scan for program groups/links

There are 2 dlls and 1 ocx. Here's what to do:

Copy the tray.dll in your system folder.

Copy the ocx's and ssubtmr.dll to your sys folder.

Go to a run box and type "regsvr32 cpopmenu.ocx".

Go to a run box and type "regsvr32 ssubtmr.dll"

Copy the tray.exe file in your windows folder.

Open your system.ini file with notepad

Replace "shell=Explorer.exe" with "shell=Tray.exe"

Reboot your computer. Hope you like it!

A great systray handler.

Extreme fascination.


<span>             |<span>
---                |---
**Submitted On**   |2000-03-22 07:08:36
**By**             |[Alex Ionescu](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/alex-ionescu.md)
**Level**          |Advanced
**User Rating**    |3.1 (44 globes from 14 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD41403222000\.zip](https://github.com/Planet-Source-Code/alex-ionescu-a-systray-handler-vb-start-menu-fixed__1-6727/archive/master.zip)








