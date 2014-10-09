WinACE - Windows Alternate Client for Empire 
Original Copyright 1997-2001 by James A. Simons. All rights reserved.
Changes in Version 2.2.0 between 2.3.0 Copyright Daniel R Kiracofe 2002.

Readme file for Version 2.5.68 released June 14, 2012.

IMPORTANT NOTICE
If you have a previous version of WinACE installed on 
your machine, see below for special instructions

WinACE is a Windows 95/98/2000/NT/XP client for connecting to a 
server running an empire game.  It will not run under 
Window 3.1.  It is recommended you have at least a 
PentiumII 550 MHz and 64 Mb of memory to get good performance 
from this client, but the program should run on smaller 
machines.

This client is written in Microsoft Visual Basic 6.0, 
and the installation program contains both the program 
executable as well as a number of system DLL's for controls 
common to most Visual Basic programs.  These controls will 
be installed on your machine as part of the WinACE 
installation process.  Because of the included controls, 
the installation program makes for a rather large download 
(about 7 MB) 

WinACE's main feature is its graphical user interface.
Display of game information centers around a 
graphical map rather than being entirely text based.
Commands can be entered through a point and click 
interface using prompt boxes.  A command line interface
is also provided.  In addition, WinACE keeps 
a local database of observed information, and contains
some tools to simplify routine tasks.

WinACE was originally written by Jim Simons, who usally plays as 
Escher.  Jim Simons has retired from playing Empire (at least for now).
Daniel Kiracofe (Bwian) maintain WinACE from July 2002 until September 2003.
Ron Koenderink took over in October 2003. He is person to annoy with bug reports 
and enhancement suggestions at either www.sourceforge.net/projects/winace or
rkoenderink@yahoo.ca.

IMPORTANT NOTICE
If you have a previous version of WinACE that it 2.4.2 or older installed on 
your machine, you will need to recreate your game databases from scratch.
If you are actively playing any games,
and want to save your current intelligence and telegrams
in order to use them with the databases created by the 
new version do the following : 1) Start WinACE; 
2) Load your game;  3) Go to the File menu and take the 
"Export" option.  This will export your intelligence and 
telegrams for that game to a text file.

It is good idea to delete your old copy of WinACE but not necessary, go to your Windows
Control Panel, click on Add/Remove Program, and 
select WinACE and click the Remove button.  This will
remove the WinACE components installed by setup, but 
not the .egp and .mdb files you created by running games.
GAME DATABASES CREATED BY VERSIONS 2.3.x AND EARLIER OF 
WINACE ARE NOT 100% COMAPATABLE WITH THIS VERSION, AND TRYING 
TO USE THEM WILL CAUSE WINACE TO CRASH.  To avoid possible 
problems, it is recommended you use Windows Explorer 
to delete any .mdb files left in the installation directory 
(default is usually C:\Program Files\Empire\WinACE). 
The .egp files created by earlier versions ARE compatable, 
and may be left.  If there are .txt files there, these are
probably your exported game information.  Leave these.

Do a full install of the current version of WinACE (2.5.0 or newer).
To restore your current game information after installing 
the new version of WinACE:   1) Start WinACE;
2) Load the your game - let WinACE recreate the database when
it asks.  3) Go to the file menu and choose "Import Telegram"
followed by "Import Intelligence".   

WinACE includes http proxy using XmlRPC.  To use this feature
you must download the XML Core Services 4.0 from Microsoft site.

PROBLEM INSTALLING
------------------
On XP, sometimes to you can get setup loop.  The easiest solution
is to instead of installing the package, to just extract all files
from the CAB in the full install package into the WinACE directory.
You can use WinZip or similar packages to extract the files from the
CAB file.

Disclamers
----------
The damage assesment tool that predicts the new efficency may not 
work correctly with DEF_INFRA option.  (it assumes the sector is fully
improved)

WinACE's food predictions (as used by the famine relief tool) are 
sometimes one food short of what starvation says is required.  Tests
show WinACE is correct, and no-one will actually starve.  (WinACE always
give them the extra food anyway using the avoid starvation option)

When sectors are lost to che at the update, the server sometimes does
not add them to the "lost" file.  This means that WinACE will not pick
them up as lost and will still show them as yours.  You will need to
do an Update Sectors off the Database menu post update to be sure.

Blitz setup was designed for the old vampire blitz (now defunct) and
may not work with the newer blitzs.  It doesn't know about such things
as GO_RENEW and BIG_CITIES

The production summary tool assumes that all thresholds will be met 
when doing distribution point mobility calculations.  For this reason
it is usually conservative.  It does not count stockpiles in the 
distribution center as "on-hand" if the distribution center has a 
different distribution point.

+++ End of File ++++
