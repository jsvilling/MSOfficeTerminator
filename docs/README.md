# OfficeKiller

## Overview

OfficeKiller is a small application written in C#. The application can be used to kill all running word, powerpoint and excel processes. I used Costura.Fody to build the application into a standalone .exe file containing all required assemblies.  

## Usage

Build the .exe and execute it. This will start the application and display an icon in the system tray. To trigger the termination process you can eiter double click the icon or right click on the icon and choose "Terminate all office programs". To exit the application right click the system tray icon and choose "Exit".
