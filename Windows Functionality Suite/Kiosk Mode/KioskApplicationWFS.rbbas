#tag Class
Protected Class KioskApplicationWFS
Inherits Application
	#tag Event
		Sub Open()
		  mAppTester = new Mutex( App.ExecutableFile.Name )
		  if mAppTester.TryEnter then
		    Soft Declare Function CreateDesktopW Lib "User32" ( name as WString, device as Integer, devMode as Integer, _
		    flags as Integer, access as Integer, sec as Integer ) as Integer
		    Declare Function GetCurrentThreadId Lib "Kernel32" () as Integer
		    Soft Declare Function GetThreadDesktop Lib "User32" ( id as Integer ) as Integer
		    Soft Declare Function OpenInputDesktop Lib "User32" ( flags as Integer, inherit as Boolean, access as Integer ) as Integer
		    Soft Declare Sub SetThreadDesktop Lib "User32" ( desk as Integer )
		    Soft Declare Sub SwitchDesktop Lib "User32" ( desk as Integer )
		    Soft Declare Sub CloseDesktop Lib "User32" ( desk as Integer )
		    
		    Const GENERIC_ALL = &H10000000
		    Const DESKTOP_SWITCHDESKTOP = &H100
		    
		    // These local variables will hold all of our desktop references
		    Dim hDesk, hOldDesk, hOldInputDesk as Integer
		    
		    // The first thing we need to do is keep track of the current
		    // desktop so that we can switch back to it when we're done
		    hOldDesk = GetThreadDesktop( GetCurrentThreadId )
		    
		    // We also need to get the input desktop
		    hOldInputDesk = OpenInputDesktop( 0, false, DESKTOP_SWITCHDESKTOP )
		    if hOldInputDesk = 0 then
		      KioskModeExceptionWFS.Create( "Could not open the input desktop for switching" )
		      return
		    end if
		    
		    // Now we're ready to make our new desktop.  We're just going to use
		    // the application name as the desktop name.
		    hDesk = CreateDesktopW( App.ExecutableFile.Name, 0, 0, 0, GENERIC_ALL, 0 )
		    if hDesk = 0 then
		      KioskModeExceptionWFS.Create( "Could not create a new desktop" )
		      return
		    end if
		    
		    // Now that we've made a new desktop, let's set it up and
		    // switch over to it
		    SetThreadDesktop( hDesk )
		    SwitchDesktop( hDesk )
		    
		    // Since we're the first instance, we want to launch another instance of
		    // this application, and wait for it to complete
		    App.ExecutableFile.LaunchAndWait( "", App.ExecutableFile.Name )
		    
		    // Now clean everything up
		    if hDesk <> 0 then
		      SwitchDesktop( hOldInputDesk )
		      SetThreadDesktop( hOldDesk )
		      CloseDesktop( hDesk )
		    end if
		    
		    // We're done too
		    Quit
		    
		    return
		  else
		    // Call the user's Open event since this is a "normal" instance
		    // of the kiosk application
		    Open
		  end if
		  
		Exception err as FunctionNotFoundException
		  // Cannot run in true kiosk mode, so do something about it
		  KioskModeExceptionWFS.Create( "Kiosk mode not supported on this operating system" )
		End Sub
	#tag EndEvent


	#tag Hook, Flags = &h0
		Event Open()
	#tag EndHook


	#tag Property, Flags = &h21
		Private mAppTester As Mutex
	#tag EndProperty


	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
