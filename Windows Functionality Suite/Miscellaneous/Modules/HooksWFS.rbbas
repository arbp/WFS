#tag Module
Protected Module HooksWFS
	#tag Method, Flags = &h21
		Private Function IdleHookCallbackFunction(nCode as Integer, wParam as Integer, lParam as Integer) As Integer
		  if nCode >= 0 then
		    if mIdleHandler <> nil then
		      ' Call the idle handler for the user
		      mIdleHandler.Idle
		    end
		  end if
		  
		  #if TargetWin32
		    Soft Declare Function CallNextHookEx Lib "User32" ( hookHandle as Integer, code as Integer, _
		    wParam as Integer, lParam as Integer ) as Integer
		    
		    ' And make sure we call the next hook in the list
		    return CallNextHookEx( mIdleHandlerHook, nCode, wParam, lParam )
		    
		  #else
		    
		    #pragma unused wParam
		    #pragma unused lParam
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub InstallIdleHook(i as IdleHandlerWFS)
		  ' If we already have an idle handler, then we
		  ' cannot install a new one.  Just bail out
		  if mIdleHandlerHook <> 0 then return
		  
		  #if TargetWin32
		    
		    Declare Function SetWindowsHookExA Lib "User32" ( hookType as Integer, proc as Ptr, _
		    instance as Integer, threadID as Integer ) as Integer
		    Declare Function GetCurrentThreadId Lib "Kernel32" () as Integer
		    
		    ' Store the idle handler
		    mIdleHandler = i
		    
		    Const WH_FOREGROUNDIDLE = 11
		    
		    // Well if this isn't about as strange as you can get...  I tried turning this into a
		    // Unicode-savvy function, but couldn't make a go of it.  Using the exact same
		    // code (only the W version of SetWindowsHookEx) causes a crash to occur
		    // immediately after the call returns.  I wasn't able to find any information about
		    // why the crash was happening, and it doesn't make any sense (the Windows is
		    // a Unicode window, so there's no mixed-type calls).  Since this function doesn't
		    // deal with strings, there's no real benefit to making it Unicode-savvy, so I'm leaving
		    // the function as-is.
		    
		    ' And install the handler
		    mIdleHandlerHook= SetWindowsHookExA( WH_FOREGROUNDIDLE, AddressOf IdleHookCallbackFunction, _
		    0, GetCurrentThreadId )
		    
		  #else
		    
		    #pragma unused i
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function InstallKeyboardHook(attachTo as KeyboardHookHandlerWFS) As Boolean
		  ' If we already have an idle handler, then we
		  ' cannot install a new one.  Just bail out
		  if attachTo = nil then return false
		  if mAttached <> nil then return false
		  
		  #if TargetWin32
		    Declare Function SetWindowsHookExA Lib "User32" ( hookType as Integer, proc as Ptr, _
		    instance as Integer, threadID as Integer ) as Integer
		    Declare Function GetCurrentThreadId Lib "Kernel32" () as Integer
		    
		    ' Store the keyboard handler
		    mAttached = attachTo
		    
		    ' And install the handler
		    Const WH_KEYBOARD = 2
		    mKeyboardHookHandle = SetWindowsHookExA( WH_KEYBOARD, AddressOf KeyboardHookCallbackFunction, 0, GetCurrentThreadId )
		  #endif
		  
		  return mKeyboardHookHandle <> 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function KeyboardHookCallbackFunction(nCode as Integer, wParam as Integer, lParam as Integer) As Integer
		  if nCode >= 0 and mAttached <> nil then
		    ' Call the keyboard handler for the user
		    mAttached.KeyboardProc( wParam, lParam )
		  end
		  
		  #if TargetWin32
		    Declare Function CallNextHookEx Lib "User32" ( hookHandle as Integer, code as Integer, _
		    wParam as Integer, lParam as Integer ) as Integer
		    
		    ' And make sure we call the next hook in the list
		    return CallNextHookEx( mKeyboardHookHandle, nCode, wParam, lParam )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub RemoveIdleHook()
		  ' If we don't have an idle handler, then we
		  ' can just bail out
		  if mIdleHandlerHook = 0 then return
		  
		  #if TargetWin32
		    Declare Sub UnhookWindowsHookEx Lib "User32" ( hookHandle as Integer )
		    
		    UnhookWindowsHookEx( mIdleHandlerHook )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub RemoveKeyboardHook()
		  ' If we don't have an keyboard handler, then we
		  ' can just bail out
		  if mKeyboardHookHandle = 0 then return
		  
		  #if TargetWin32
		    Declare Sub UnhookWindowsHookEx Lib "User32" ( hookHandle as Integer )
		    
		    UnhookWindowsHookEx( mKeyboardHookHandle )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function TranslateKeyToString(lParam as Integer) As String
		  #if TargetWin32
		    
		    Soft Declare Function GetKeyNameTextW Lib "User32" ( lParam as Integer, _
		    theStr as Ptr, theStrLen as Integer ) as Integer
		    Soft Declare Function GetKeyNameTextA Lib "User32" ( lParam as Integer, _
		    theStr as Ptr, theStrLen as Integer ) as Integer
		    
		    dim mb as new MemoryBlock( 25 )
		    dim trueLen as Integer
		    
		    if System.IsFunctionAvailable( "GetKeyNameTextW", "User32" ) then
		      trueLen = GetKeyNameTextW( lParam, mb, mb.Size )
		      return mb.WString( 0 )
		    else
		      trueLen = GetKeyNameTextA( lParam, mb, mb.Size )
		      return mb.CString( 0 )
		    end if
		    
		  #else
		    
		    #pragma unused lParam
		    
		  #endif
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mAttached As KeyboardHookHandlerWFS
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIdleHandler As IdleHandlerWFS
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIdleHandlerHook As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mKeyboardHookHandle As Integer
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
