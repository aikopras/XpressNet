# Download and Install the XpressNet DLL (regsvr32)

## Step 1: Create a directory ##

The first step is to open the File Explorer, and create a directory called XpressNetDll. The location and name of the directory should not be changed later. A good place is directly under C:.
<br/><img src="Download/01-Create_Directory.png" alt="Create directory" width="500"/>
<br/><br/>

## Step 2: Go to the GitHub repository ##
Open your browser and go to the [XpressNet repository on Gihub](https://github.com/aikopras/XpressNet).
<br/><img src="Download/02-Goto_GitHub.png" alt="go to GitHub" width="500"/>
<br/><br/>

## Step 3: Double Click XpressNet.dll ##
Go to the directory called Code and double click XpressNet.dll
<br/><img src="Download/03-2A-Double_Click.png" alt="click XpressNet.dll" width="400"/>
<br/><br/>

## Step 4: Download as RAW file ##
Click the download RAW file icon at the right to download XpressNet.dll to your local machine.
<br/><img src="Download/04-2A-Download_raw_file.png" alt="Download RAW" width="500"/>
<br/><br/>

## Step 5: Save XpressNet.dll ##
After the file is downloaded (to your local Downloads directory), move the file to the directory you've created in Step 1 above (like C:/XpressNetDll).
<br/><img src="Regsvr32/01-Move_DLL-File.png" alt="Copy Files" width="600"/>
<br/><br/>

## Step 6: Open the Command prompt as Admin ##
Now that the dll file is downloaded and saved, it is time to register the DLL at your local machine. For that purpose you have to run Command Prompt as Administrator.
Click Windows Start -> Microsoft System followed by a right click on Command Prompt. A windows appears and under More you select Run as Administrator.
<br/><img src="Regsvr32/02-Command_Prompt.png" alt="Open Cmd as Admin" width="600"/>
<br/><br/>

## Step 7: Allow Command as Admin ##
Depending on your version of windows, a window may pop-up, asking you to allow the Windows Command Processor to make changes to your local computer. Click YES.
<br/><img src="Regsvr32/03-Allow_Cmd_As_Admin.png" alt="Allow Cmd as Admin" width="400"/>
<br/><br/>

## Step 8: Move to the DLL directory ##
Move to the directory where you have stored the XpressNet.dll. This directory should not be renamed or moved later.
<br/><img src="Regsvr32/04-Cmd_regsvr32.png" alt="move dir" width="500"/>
<br/><br/>

## Step 9: Run regsvr32 ##
Run regsvr32, using XpressNet.dll as argument.
<br/><img src="Regsvr32/04-Detail_regsvr32.png" alt="Register" width="400"/>
<br/>
If everything goes well, a window pops up, saying that registration of XpressNet.dll succeeded.
<br/><img src="Regsvr32/05-Succeeded.png" alt="Registration Succeeded" width="400"/>
<br/>
If you want to unregister the DLL later, for example because a new version should be registered, you can unregister the DLL using the /u switch.
<br/><img src="Regsvr32/06-unregister.png" alt="Unregister" width="400"/>
<br/><br/>

## Step 10: Reference the XpressNet DLL ##
Before you can use the XpressNet DLL library in your project, you must create a reference to the DLL. After opening your project click
Project -> References, and select the XpressNet Interface. See both figures below.
<br/><img src="Reference/01-OpenPreferences.png" alt="Open Project -> Preferences" width="350"/>
<br/><br/>
<br/><img src="Reference/02-SelectXpressNet.png" alt="Select XpressNet" width="500"/>
<br/><br/>
