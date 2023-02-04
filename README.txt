======================================================================
•▀			    eTweetXL Setup Guide		            
======================================================================
Written by: anz7re (André)

----------------------------------------------------------------------
Latest Revision:

2/1/2023

----------------------------------------------------------------------
Version:

1.9.0

----------------------------------------------------------------------
Developer(s): 

anz7re (André)

----------------------------------------------------------------------
Contact:

Email: support@etweetxl.xyz | support@autokit.tech | anz7re@autokit.tech
Social: @anz7re | eTweetXL | @AutokitTech 
Web: etweetxl.xyz | autokit.tech/etweetxl
Donate: $donateautokitdevs

(Don't hesitate to reach out if you're having any issues!)

----------------------------------------------------------------------
License Information:

Copyright (C) 2022-present, Autokit Technology.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

----------------------------------------------------------------------
Application Description:

eTweetXL is a basic Excel(VBA) based Windows program for automating, and managing tweets for Twitter on desktop.

(This application mainly uses Windows scripting languages so you'll have to have a Windows operating system in order to run this!)
 
----------------------------------------------------------------------
Requirements:

- Excel Version 2107+
- Firefox 88.0+ (not needed if sending w/ Twitter API)
- PowerShell 5.1.19041.1023+
- Python 3.9+ 
- tweepy (Python library)
- xlwings (Python library)
- Windows 10+

(May work w/ previous versions, but not tested)

---------------------------------------------------------------------
Application Installation Guides:

Python Install Guide:

1. Download Python from "www.python.org"

2. Check box to install to "PATH" location (***IMPORTANT)

3. Once install's completed, open the System Command Shell (Command Prompt)

4. Type "pip install tweepy" then press "ENTER" (wait for completion)

5. Type "pip install xlwings" then press "ENTER" (wait for completion)

6. All done!

---------------------------------------------------------------------
Default Install Guide (If this is your first time installing the application):

1. Place the first "eTweetXL v1.9.0 Download" folder from the .zip file on your desktop

2. Open the "setup.xlsm" file (this should automate connecting the application)
***Upon completion you should be prompted w/ a confirmation messagebox, however
files may still be left within the download folder. 
If files are left within the download folder you may need to manully move the "eTweetXL.xlsm" & "shell" folder scripts to their respective locations (See below)

3. All done!

---------------------------------------------------------------------
***Flagged/Manual Install Guide (If the "setup.bat" gets flagged while trying to install):

1. Go to your "C:" drive & find your "Users" folder

2. Select the current home users folder & within that create a directory titled ".z7"

3. Within the ".z7" folder create a folder titled "autokit"

4. Within the "autokit" folder create a folder titled "etweetxl"

5. Within the "etweetxl" folder create 5 folders titled: "app", "debug", "mtsett", "presets", & "shell"

6. Within your shell folder create 2 folders titled: "wb", & "win"

7. Copy all documents from the "shell" folder inside the "eTweetXL v1.9.0 Download" folder & paste into the "win" folder 

8. Copy the "eTweetXL.xlsm" file along w/ "scure_replacement.txt", "twitter.ico", "twitter.jpg" (all .jpg & image files), & paste into the "app" folder

9. Edit this path inside "show_etweetxl.vbs": ("C:\Users\EDITHERE\.z7\autokit\etweetxl\app\eTweetXL.xlsm")

10. Edit this path inside "runtime_error.vbs": ("C:\Users\EDITHERE\.z7\autokit\etweetxl\debug\rt.err", 1)

11. Edit this path inside "runtime_refresh.vbs": ("C:\Users\EDITHERE\.z7\autokit\etweetxl\mtsett\offset.mt", 1)

12. All done!

---------------------------------------------------------------------
Desktop Icon Guide:

1. Go to the "win" folder & right-click the "show_etweetxl.bat" file 

2. Select "Create shortcut"
 
3. Drag shortcut to desktop (you can change the name of the shortcut if you like)

4. Right-click your shortcut & select "Properties"

5. Click "Change Icon...", then select "Ok" 

6. Click "Browse" & find the "twitter.ico" file in your eTweetXL "app" folder

7. Select the "twitter.ico" file, click "Ok", then "Apply"

8. All done! 

----------------------------------------------------------------------
Hope you all have fun w/ this!

Thanks for downloading :)
----------------------------------------------------------------------
