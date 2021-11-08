This Template presents a WindowsForms solution, that allows COM Interop Communication. The purpose of this project is to allow embedding of WindowsForms Control Libraries into VBA environments like Excel. It will allow to see COM exposed properties and Intellisense will be available as well. 
To create an own project from scratch follow these steps:
1. Create a new Windows Forms Control Library project
2. Get into the properties of your project
3. In the Application Tab, click on the "Assembly Information..." Button
4. Check the "Make assembly COM-Visible" and confirm with OK
5. Navigate to the "Build" tab, scroll down and check the "Register for COM Interop"
6. Navigate to the "Signing" Tab, check the "Sign the assembly" box
7. Under "Choose a strong name key file" is a dropdown menu, click on it and select "New.."
8. Give it a name and optionally a password and confirm with OK
9. Save all changes 
10. Add to your project a new interface
11. Make your interface public
12. Add the following annotations above the interface:
&nbsp;[ComVisible(true)]
&nbsp;[Guid("<Generate a Guid somewhere and paste it in here>")]

13. Add a property with an "DispId" annotation above:
&nbsp;[DispId(1)]
&nbsp;string CustomText { get; set; }

14. Implement this interface into your Forms Control Library Class
15. Give your new implemented property a defaut value, this is very important!
16. Add these annotations for your Control Library class
&nbsp;[ComVisible(true)]
&nbsp;[Guid("<Generate a Guid somewhere and paste it in here>"), ClassInterface(ClassInterfaceType.None)]

For the next part I recommend to copy and paste the ActiveX Control registration functions
17. Copy the ActiveXControlHelper.cs into your project
18. Add the ComRegisterFunction and ComUnregisterFunction to your Control Library class (copy and paste) with annotations
19. Add to your project a manifest file. It should look like the file inside this template project, however you'll have to change the clsid and both the name of the assemblyIdentity and the name of the clrClass respective to your naming
20. Add a new textfile and change its extension to .rc and insert the following into this file
&nbsp;1 RT_MANIFEST <your manifest file>

Before building the project there is one last thing to do:
21. Go to your projects "Properties"
22. Navigate to the "Build Events" tab and insert the following command into "Pre-build event command line"
@echo.
set RCDIR=
IF EXIST "$(FrameworkSDKDir)Bin\rc.exe" (set RCDIR="$(FrameworkSDKDir)Bin\rc.exe")
IF EXIST "$(DevEnvDir)..\..\VC\Bin\rc.exe" (set RCDIR="$(DevEnvDir)..\..\VC\Bin\rc.exe")
IF EXIST "$(DevEnvDir)..\..\SDK\v2.0\Bin\rc.exe" (set RCDIR="$(DevEnvDir)..\..\SDK\v2.0\Bin\rc.exe")
IF EXIST "$(DevEnvDir)..\..\SDK\v3.5\Bin\rc.exe" (set RCDIR="$(DevEnvDir)..\..\SDK\v3.5\Bin\rc.exe")
IF EXIST "$(DevEnvDir)..\..\..\Microsoft SDKs\Windows\v6.0a\bin\rc.exe" (set RCDIR="$(DevEnvDir)..\..\..\Microsoft SDKs\Windows\v6.0a\bin\rc.exe")
IF EXIST "$(DevEnvDir)..\..\..\Microsoft SDKs\Windows\v7.0a\bin\rc.exe" (set RCDIR="$(DevEnvDir)..\..\..\Microsoft SDKs\Windows\v7.0a\bin\rc.exe")
IF EXIST "C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x86\rc.exe" (set RCDIR="C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x86\rc.exe")
if not defined RCDIR (echo "Warning!  Unable to find rc.exe, using default manifest instead.") ELSE (%RCDIR% /r "$(ProjectDir)INSERTYOURRCFILENAMEHERE.rc")
if not defined RCDIR (Exit 0)
@echo.
23. Insert the following command into the "Post-build event command line"
echo register assembly
echo $(frameworkdir)\$(frameworkversion)\regasm.exe "$(TargetPath)" /tlb /codebase

24. Build your project
25. Done!

