CHANGELOG
=========

this is the change log of the source code prior to it being part of GitHub/BitBucket.

--

**Version 1.0 -- May 24 2004**

Initial check in

**Version 1.1 -- Jun 20 2004**

May 24 2004 -- Added HasBorder, adapted from code submitted by "The One"

May 25 2004 -- Added a system hook for an IdleHandler

May 26 2004 -- Added EmptyClipboard, Launch with args

May 28 2004 -- Fixed a leak in ChooseFont, thanks to François Van Lerberghe

Jun 02 2004 -- Added GetPrinters and [Get/Set]DefaultPrinter, thanks to François Van Lerberghe

Jun 14 2004 -- Added ChangeWindowState, thanks to Seth Willits

Jun 20 2004 -- Added IsNAN, IsInf

**Version 1.2 -- Jun 24 2004**

Jun 20 2004 -- Fixed Window.ChangeWindowState function, Added GetCurrentProcessID

Jun 23 2004 -- Added documentation

**Version 1.3 -- Jul 07 2004**

Jun 25 2004 -- Added the Mutex class, more Window extensions, cleaned up project

Jul 01 2004 -- Added GetSpecialFolder, thanks to Stephen Tallent

Jul 02 2004 -- Added ApplicationPriority getter and setter pair

Jul 06 2004 -- Added IsAlphabetic, IsLowerCase, IsUpperCase, IsWhiteSpace, IsPunctuation, IsNumber and IsHexDigit

Jul 07 2004 -- Fixed some missing #if so the project compiles on other platforms

**Version 1.4 -- Jul 30 2004**

Jul 08 2004 -- Added the SystemColors and OSVersionInformation modules, thanks to Brian Rathbone

Jul 10 2004 -- Added SelectMultipleFiles, adapted from code submitted by Brian Rathbone

Jul 11 2004 -- Added the LocaleInformation and SystemMetrics modules, added FillString

Jul 12 2004 -- Added the PressKey method

Jul 23 2004 -- Added the Window.CloseButtonState method

Jul 28 2004 -- Fixed a documentation bug with GetActiveProcessNames

Jul 30 2004 -- Added the ServiceManager module and Service class

**Version 1.5 -- Aug 12 2004**

Aug 06 2004 -- Added the Cryptography module

Aug 07 2004 -- Added the FolderItem.AssociateExtension, IsExtensionAssociated

Aug 08 2004 -- Added CompressData, DecompressData

Aug 09 2004 -- Added the Registry module, fixed some minor ServiceManager bugs

Aug 10 2004 -- Added FolderItem.StartupItem, IsStartupItem, updated the Documentation

Aug 11 2004 -- Added the Progress class

**Version 1.6 -- Sep 04 2004**

Aug 25 2004 -- Added FolderItem.StartWatchingForChanges, HasChanged and StopWatchingForChanges

Aug 26 2004 -- Added Window.FreezeUpdate and UnfreezeUpdate

Aug 27 2004 -- Added Plug and Play notifications

Aug 28 2004 -- Added the Mouse module to allow you to control the mouse

Aug 29 2004 -- Added Window.Topmost and IsToolbarWindow

Sep 02 2004 -- Added the WndProcHelpers module for easy window subclassing

Sep 04 2004 -- Added the TreeView and TreeViewItem classes

**Verion 1.7 -- Sep 27 2004**

Sep 06 2004 -- Added Window.BringToFront, thanks to Tomas Camin

Sep 16 2004 -- Fixed some #ifs that were missing

Sep 18 2004 -- Added the DateTimePicker class

Sep 19 2004 -- Added the HotKeyHelper class so you can register hotkeys for your application

Sep 20 2004 -- Added the Calendar class

Sep 21 2004 -- Added some disclaimed in the source code where the declares are unsupported

Sep 21 2004 -- Updated the documentation to reflect some APIs that may not work in any version but 5.5

Sep 26 2004 -- Added CreateGuid and CreateGuidString to create globally unique identifiers

**Version 1.7.1 -- Sep 28 2004**

Sep 28 2004 -- Fixed download so the project would compile

**Version 1.8 -- Oct 16 2004**

Sep 29 2004 -- Cleaned up the TreeView example code a bit

Sep 29 2004 -- Changed Window.FreezeUpdate so that it uses WinHWND instead of Handle for old version compatibility

Sep 30 2004 -- Added Window.Alpha and Window.Mask, adapted from code by Peter De Berdt

Oct 01 2004 -- Added Window.AnimateWindow, thanks to Peter De Berdt

Oct 09 2004 -- Added OSVersionInformation.IsNT5 and fixed OSVersionInformation.IsNx

Oct 10 2004 -- Added the ToolTip class for complete control over tooltips

Oct 15 2004 -- Renamed Mutex to be Win32Mutex to avoid a conflict with RB 6

Oct 15 2004 -- Renamed the project to be the Windows Functionality Suite

**-------------- NAME AND VERSION NUMBER CHANGE -----------------------**

**Version 1.1 -- Dec 27 2004**

Dec 27 2004 -- Added testing suite for the SystemParameters module

Dec 26 2004 -- Finished adding accessors to the SystemParameters module for accessibility information

Dec 23 2004 -- Added mouse inputs to the SystemParameters module

Dec 21 2004 -- Continued adding more SystemParameters methods for input and power options

Dec 20 2004 -- Added more SystemParameters methods to deal with menu UI

Dec 12 2004 -- Added IconMetrics and LogicalFont, as well as more SystemParameters methods

Dec 10 2004 -- Started adding system information to the SystemParameters module

Dec 05 2004 -- Updated TimeZoneInformation to include information about which set of information the system is currently using

Nov 26 2004 -- Added functions to get the handle to a window based on a partial title, thanks to Brian Rathbone

Nov 22 2004 -- Added the EditFieldExtensions module to add some extra functionality for the EditField control

Nov 21 2004 -- Added the PushButtonExtensions module to add extra features to the PushBotton control

Nov 20 2004 -- Added an optional parameter to the Win32Mutex class constructor to specify a global mutex.

Nov 19 2004 -- Added the TimeZoneInformation class

Nov 19 2004 -- Added LocaleInformation.TimeZoneInformation to get information about the user's timezone

Nov 15 2004 -- Added OSVersionInformation.IsTerminalSession, thanks to Brian Rathbone

Oct 17 2004 -- Added the IPAddress class to let you enter in IP addresses via a standard control

**Version 2.0 -- June 14 2005**

Jun 11 2005 -- Added a new SystemColor for the dark button color, as well as a new test window for the SystemColors module, thanks to Tom Dixon

Jun 10 2005 -- Fixed a bug with loading icons when you pass a .ico file with the LoadIcon method.  Also, changed the parameter to default to -1

Jun 06 2005 -- Fixed SelectMultipleFiles so that it works properly when one file is selected, thanks to Douglas Anderson

Jun 05 2005 -- Changed the WndProcSubclass interface so that it properly handled messages where returning 0 means you've handled the message.  You *must* update your code because of this change.

Jun 03 2005 -- Extended the CaptureScreen method to optionally include the mouse cursor, thanks to Brian Rathbone

May 28 2005 -- Fixed some exceptions that would occur in the IsX string functions of the Win32DeclareLibrary module

May 25 2005 -- Added the DisplaySettings module, thanks to Brian Rathbone

May 20 2005 -- Fixed a bug in WndProcHelpers so that the same window can be subclassed multiple times

May 15 2005 -- Fixed StatusBar so that it subclasses the owner window and handles its own resizing messages.  You no longer need to call StatusBar.Resize manually

May 12 2005 -- Added an extension method to the Window class so you can set its icon.

May 03 2005 -- Modified the test application to allow you to create a new test window via a menu

May 03 2005 -- Fixed a problem with the mutex never closing when the window closed

May 01 2005 -- Fixed some of the examples so that they functioned within a MDI application (such as moving the mouse cursor)

Apr 25 2005 -- Extended the FolderItem class so you can copy and move files with a progress indicator

Apr 17 2005 -- Added ProcessInformation.BringToFront to bring all windows associated with a process to the front.

Apr 11 2005 -- Added ways to tile and cascade children windows of an MIWindow

Apr 09 2005 -- Added the MDIWindowExtensions module to let you set some more properties of App.MDIWindow

Apr 06 2005 -- Added Window.IsMinimized and IsMaximized to the window extensions

Apr 02 2005 -- Added a way to hook keyboard messages for the current application to the Hooks module

Apr 02 2005 -- Added the Hooks.TranslateKeyToString function to help with the keyboard hook

Mar 26 2005 -- Added the Hooks module and added the idle hook there.  Deprecated the idle hook handler from the Win32DeclareLibrary module

Mar 10 2005 -- Fixed Window.Mask and Window.Alpha so they will not throw loader assertions on older versions of Windows

Mar 06 2005 -- Added the PerformanceCounter module for high performance counters

Mar 06 2005 -- Added FolderItem.AddToRecentItems to let you stick a shortcut to the FolderItem in the recent items list

Mar 05 2005 -- Added two new Shutdown methods -- one for shutting down (rebooting or logging off) with UI, and the other for displaying the standard Shutdown dialog.

Mar 03 2005 -- You can now set the screen resolution for the display (width, height and bits per pixel).

Mar 02 2005 -- Added a HexCanvas for viewing process memory, thanks to Jon Johnson

Mar 02 2005 -- Can now view process memory from a HeapEntryInformation block.

Mar 02 2005 -- Added the HeapListInformation and HeapEntryInformation classes and now ProcessInformation instances can now check heap information

Feb 28 2005 -- Added the ThreadInformation class and now ProcessInformation instances can contain a list of loaded threads

Feb 28 2005 -- Added the ModuleInformation class and now ProcessInformation instances can contain a list of loaded modules as well

Feb 26 2005 -- Added MapNetworkDriveDialog and UnmapNetworkDriveDialog for user interaction with network drive mapping

Feb 25 2005 -- Added Map and Unmap network drive support

Feb 20 2005 -- Added the ability to set icons and tooltips in the StatusBar

Feb 19 2005 -- The PlugAndPlay module is now Unicode-savvy.

Feb 17 2005 -- Win32DeclareLibrary.GetActiveProcesses, GetActiveProcessNames and ProcessInformation are now Unicode-savvy

Feb 17 2005 -- The TreeView control is now Unicode-savvy

Feb 16 2005 -- Fixed a bug with the HotKey functionality so that it no longer throws a NilObjectException.

Feb 15 2005 -- Fixed a bug with failing to initialize the Cryptography module properly.

Feb 15 2005 -- The Cryptography and SystemParameters modules are now Unicode-savvy.

Feb 13 2005 -- WndProcHelpers and HotKeyHelpers are now Unicode-savvy.

Feb 11 2005 -- The ServiceManager module is now Unicode-savvy.

Feb 10 2005 -- The LocaleInformation module is now Unicode-savvy.

Feb 07 2005 -- Added new functionality to the StatusBar class for adding and modifying parts.

Feb 05 2005 -- Win32DeclareLibrary.SelectMultipleFiles is now Unicode-savvy.

Feb 03 2005 -- Added Registry.WallpaperStyle to set the style for displaying desktop wallpaper

Jan 29 2005 -- The Progress class is now Unicode-savvy.

Jan 25 2005 -- SetDefaultPrinter is now Unicode-savvy

Jan 24 2005 -- Renamed GetVolumeSerial (returning an Integer) to GetVolumeSerialNumber (fixes a compiler error)

Jan 24 2005 -- Renamed ChooseFont (returning a LogicalFont) to ChooseLogicalFont (fixes a compiler error)

Jan 22 2005 -- LoadIcon and OSVersionString are now Unicode-savvy

Jan 20 2005 -- GetTotalBytes, GetTotalFreeSpace, GetVolumeName, GetVolumeSerial, GetVolumeSerialNumber are now Unicode-savvy

Jan 18 2005 -- GetFreeDiskSpaceForCaller, GetLoggedInUserName, GetSystemName and GetSpecialFolder are now Unicode-savvy

Jan 17 2005 -- Encrypting and Decrypting NTFS files is now Unicode-savvy

Jan 16 2005 -- Finding a window handle from a partial title is now Unicode-savvy

Jan 14 2005 -- Win32DeclareLibrary.FormatErrorMessage, GetDriveStrings, GetDriveType is now Unicode-savvy

Jan 12 2005 -- GetComputerName, GetDefaultPrinter is now Unicode-savvy

Jan 10 2005 -- IPAddress is now Unicode-savvy.

Jan 06 2005 -- LogicalFont and CountWindowsWithPartialTitle are now Unicode-savvy.

Jan 06 2005 -- Added Win32DeclareLibrary.ChooseFont that returns a LogicalFont and is Unicode-savvy

Jan 05 2005 -- Made OSVersionInformation Unicode-savvy

Jan 05 2005 -- Fixed typo, it's now OSSuites instead of OSSuit

Jan 05 2005 -- Added OSVersionInformation.ServicePack

Jan 04 2005 -- Made DateTimePicker and Calender Unicode-savvy

Jan 03 2005 -- Made FolderItemExtensions, Win32Mutex, Service, StatusBar and ToolTip Unicode-savvy

Dec 30 2004 -- Added GetActiveProcesses, GetFrontmostWindowProcessInformation and the ProcessInformation class

Dec 29 2004 -- Added the GetFrontmostWindowHandle method

Dec 28 2004 -- Added a way to capture the screen to a Picture

**Version 2.1 -- Jun 15 2005**

Jun 15 2005 -- Fixed SelectMultipleFiles so that it would properly return the first item of a multiple selection as a file, not a folder.  Thanks to Tom Russell

**Version 2.2 -- Oct 13 2005**

Oct 13 2005 -- Added the ability to query a service for the current state.  Thanks to Richard Laframboise

Sep 03 2005 -- Changed Service.Continue to be Service.Resume so that it doesn't conflict with the new Continue keyword in REALbasic 2005r3

Aug 11 2005 -- Fixed CaptureScreen so that it does not leak memory.

**Version 2.3 -- Mar 03 2006**

Jan 30 2006 -- Added Win32DeclareLibrary.DialogUnitsToPixels

Jan 20 2006 -- Added FolderItem.Reveal

Jan 17 2006 -- Fixed a bug in the FolderItemExtensions.AssociateExtension code, thanks to Mario Buchichio

Jan 05 2006 -- Fixed a compiler error that prevented the suite from compiling in RB2006r1

Dec 30 2005 -- Added the EditField.LineSpacing extension.  Currently supports 1, 1.5 and 2 line spaces

Nov 25 2005 -- Added the InternetSession, FTPSession and FindFile classes

Nov 10 2005 -- Added the Console class to give you more complete control over a console from within a GUI application

Oct 31 2005 -- Fixed Win32DeclareLibrary.CaptureScreen so that it captures transparent windows as well, thanks to Scott Shriver

Oct 31 2005 -- Added Win32DeclareLibrary.GenerateKeyDown and GenerateKeyUp as alternatives to PressKey

Oct 24 2005 -- Added the Win32DeclareLibrary.CPUUsage function, which returns the current CPU load (as a percentage)

Oct 21 2005 -- Added the SystemMemory module

Oct 19 2005 -- Added the InternetConnection module, modified from code by Carlos Martinho

Oct 16 2005 -- Added the FileAssociation class to give you more control over how to associate a file with an application.

**Version 2.4 -- Jul 11 2006**

Jul 10 2006 -- Added a large number of properties and a few methods to the LocaleInformation class.

Jul 09 2006 -- Removed the LocaleInformation module from the SystemInformation template and replaced it with a class of the same name.  Note that you will need to update your code accordingly.

Jul 06 2006 -- The StatusBar class can now be used with MDI windows as well as SDI windows

Jul 03 2006 -- Added GraphicsHelpers.IsSymbolFont as a way to test whether the current font uses symbols

Jul 01 2006 -- Added GraphicsHelpers.CanBeDisplayed to test whether a string can be displayed using the current font

Jun 27 2006 -- Added the ability to temporarily install fonts

Jun 16 2006 -- Added a number of new methods to the Registry module for CPU and BIOS information, thanks to Anthony Cyphers

Jun 10 2006 -- Added the Like function to the VB module, thanks to Mark Nutter

Jun 07 2006 -- Added a new application subclass called KioskApplication to the Misc templates

May 26 2006 -- Added the Font class to give you full access to font creation APIs

May 15 2006 -- Added the WebCam module to the GraphicsHelpersTemplate, adapted from code by Anthony Cyphers

May 03 2006 -- The WndProcHelpers module can now subclass Window, MDIWindow and Integer (HWND handle) instead of just Window

Apr 29 2006 -- Added the AppActivate function to VBCompatTemplate that works off a window title instead of PID

Apr 19 2006 -- Added Pmt, PV and FV functions to the VBCompatTemplate

Apr 18 2006 -- Modified UIExtras.ChooseFont to include setting default style information, thanks to Joe Ranieri

Apr 10 2006 -- Added a VB Compatibility module which gives VB APIs to your REALbasic projects

Mar 04 2006 -- Fixed all places using LongLongToDouble hacks so that they use native datatypes on 2006r1 and higher

Mar 04 2006 -- Fixed all places using a plugin SDK hack to get an HDC for a picture so that it uses Graphics.Handle on 2006r1 and higher

Apr 01 2006 -- Completed refactoring of the project into multiple templates

**Version 2.5 -- Mar 27 2007**

Jul 16 2006 -- Added the ability to create system restore points (ME/XP and up) to the Win32DeclareLibrary module

Jul 20 2006 -- Fixed FTPSession.SetLocalDirectory so that it compiles for platforms other than Windows

Aug 01 2006 -- Added a TCPSocket extension called TransmitFile which will send a file over the network for you.  Note that the stream you pass in to the API must remain open until the send has completed.

Aug 13 2006 -- Updated the SystemParameters module for the new information exposed in Vista

Aug 14 2006 -- Updated the SystemMetrics module for the new information exposed in XP SP 2 and Vista

Sep 10 2006 -- Incorporated several features for the StatusBar class, thanks to Carlos Martinho

Sep 29 2006 -- Fixed MDIWindowExtensions.enumChildProc on non-NT systems, thanks to Carlos Martinho

Nov 28 2006 -- Fixed ToolTip so that it works properly on Windows 2000

Dec 22 2007 -- Added IsVista and IsNT6 to the OSVersionInformation module

Jan 11 2007 -- Fixed WindowExtensions.FreezeUpdade and UnfreezeUpdate so that they no longer do the incorrect thing

Jan 27 2007 -- Added LaunchAsAdministrator to the FolderItemExtensions module for systems which support the "runas" verb

Feb 02 2007 -- Fixed OSVersionInformation so that it works properly on Windows 98

Feb 06 2007 -- Added the ErrorReport class for error reporting in Vista

Feb 10 2007 -- Added the HasShield setter to the PushButtonExtensions module

Feb 18 2007 -- Fixed the Win32Mutex class so that it no longer uses "global" (a newly-reservered keyword) as a parameter name

Mar 01 2007 -- FTPSession no longer has a default property value which is incorrect

Mar 03 2007 -- Added the Terminate method to the ProcessInformation class.  Use this method with caution as it can cause data loss

Mar 08 2007 -- Added the IniFile module to the File Processing Template, thanks to Alessandro Consorti

Mar 09 2007 -- Added the BinaryStreamExtensions module which exposes ways to read CStrings and WStrings

Mar 12 2007 -- Added the ApplicationRecovery module and ApplicationRecoveryCallbackProvider interface for application recovery services on Vista

Mar 18 2007 -- Fixed FileAssociation.Register so that it's only compiled for Windows

Mar 18 2007 -- Fixed FolderItemExtensions.Reveal so that it's only compiled for Windows

Mar 19 2007 -- Updated SystemInformation.GetDefaultPrinter and SetDefaultPrinter to use newer APIs whenever available

Mar 20 2007 -- Fixed GraphicsHelpers.CaptureScreen so that it uses the withCursor parameter

**Version 2.5.1 -- Jan 06 2008**

May 07 2007 -- Added a FolderItem extension method called DeleteOnReboot, thanks to Anthony G. Cyphers

May 17 2007 -- Added the ability to empty the recycle bin for a given drive, or all drives, thanks to Anthony G. Cyphers

May 17 2007 -- Added the ability to query the recyle bin's size and item count for a given drive or all drives, thanks to Anthony G. Cyphers

Jun 12 2007 -- Fixed the Calendar control so that it works at positions other than 0, 0 without an improper border.

Jun 25 2007 -- Added the State method pair to the Progress class to allow you to set normal, paused and error states for a progress indicator.

Aug 01 2007 -- Fixed a compile error with 2007r4

**Version 2.5.2 -- Dec 09 2010**

Dec 09 2010 -- Updated to compile with REAL Studio 2011 Release 1


**Version 2.6.0 -- Dec 22, 2010**

Dec 22, 2010 -- FindFile.Attributes was renamed to FindFile.FileAttributes to avoid conflicts with the Attributes compiler feature

Dec 22, 2010 -- The ToolTip class was renamed to Win32ToolTip to avoid conflicts with the ToolTip framework class

Dec 22, 2010 -- Fixed all of the instances which required Lib "", so the project compiles properly again