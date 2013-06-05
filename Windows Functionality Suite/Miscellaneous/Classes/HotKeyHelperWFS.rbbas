#tag Class
Protected Class HotKeyHelperWFS
Implements WndProcSubclassWFS
	#tag Method, Flags = &h21
		Private Function ConstructKeyString(lParam as Integer) As String
		  // The low portion of the word is the key modifiers
		  // and the high portion is the virtual key
		  dim lo, high as Integer
		  lo = Bitwise.BitAnd( lParam, &hFFFF )
		  high = Bitwise.ShiftRight( lParam, 16 )
		  
		  dim ret as String
		  if Bitwise.BitAnd( lo, kControlKey) <> 0 then
		    ret = ret + "Control "
		  end
		  if Bitwise.BitAnd( lo, kAltKey) <> 0 then
		    ret = ret + "Alt"
		  end
		  if Bitwise.BitAnd( lo, kShiftKey) <> 0 then
		    ret = ret + "Shift "
		  end
		  if Bitwise.BitAnd( lo, kWindowsKey ) <> 0 then
		    ret = ret + "Window "
		  end
		  
		  // Remove the spaces and put a + in place instead
		  ret = ReplaceAll( ret, " ", "+" )
		  
		  // And be sure to add on the virtual key string
		  #if TargetWin32
		    Soft Declare Function MapVirtualKeyA Lib "User32" ( vk as Integer, type as Integer ) as Integer
		    Soft Declare Function MapVirtualKeyW Lib "User32" ( vk as Integer, type as Integer ) as Integer
		    
		    Soft Declare Function GetKeyNameTextA Lib "User32" ( lParam as Integer, name as Ptr, size as Integer ) as Integer
		    Soft Declare Function GetKeyNameTextW Lib "User32" ( lParam as Integer, name as Ptr, size as Integer ) as Integer
		    
		    // First, map the virtual key into a scan code
		    dim scanCode as Integer
		    if System.IsFunctionAvailable( "MapVirtualKeyW", "User32" ) then
		      scanCode = MapVirtualKeyA( high, 0 )
		    else
		      scanCode = MapVirtualKeyW( high, 0 )
		    end if
		    
		    // The GetKeyNameText function expects the
		    // scan code to occupy bits 16 thru 23, so we
		    // must shift ourselves over to occupy that space
		    scanCode = BItwise.ShiftLeft( scanCode, 16 )
		    
		    // Then get the key text
		    dim keyText as new MemoryBlock( 32 )
		    dim keyTextLen as Integer
		    
		    if System.IsFunctionAvailable( "GetKeyNameTextW", "User32" ) then
		      keyTextLen = GetKeyNameTextW( scanCode, keyText, keyText.SIze )
		      ret = ret + keyText.WString( 0 )
		    else
		      keyTextLen = GetKeyNameTextA( scanCode, keyText, keyText.SIze )
		      ret = ret + keyText.CString( 0 )
		    end if
		  #endif
		  
		  return ret
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  // Just grab the first window we can.  We do this so that
		  // the user can drag this class onto a Window and use it
		  // that way.
		  Constructor( Window( 0 ) )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(forWindow as Window)
		  // Subclass the window that was passed in
		  WndProcHelpersWFS.Subclass( forWindow, self )
		  
		  // Then save the window off
		  mWindow = forWindow
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Destructor()
		  // Be sure to unregister the subclass
		  WndProcHelpersWFS.Unsubclass( mWindow, self )
		  
		  // Go thru the list of remaining atoms and unregister
		  // them all
		  dim i as Integer
		  for each i in mAtomTable
		    UnregisterHotKey( i )
		  next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function RegisterHotKey(modifiers as Integer, virtualKey as Integer) As Integer
		  #if TargetWin32
		    
		    // Come up with a new identifier for the hotkey
		    dim id as Integer
		    Soft Declare Function GlobalAddAtomA Lib "Kernel32" ( name as CString ) as Integer
		    Soft Declare Function GlobalAddAtomW Lib "Kernel32" ( name as WString ) as Integer
		    
		    if System.IsFunctionAvailable( "GlobalAddAtomW", "Kernel32" ) then
		      id = GlobalAddAtomW( "Win32DeclareLibraryAtom" + Str( mNextNum ) )
		    else
		      id = GlobalAddAtomA( "Win32DeclareLibraryAtom" + Str( mNextNum ) )
		    end if
		    
		    mNextNum = mNextNum + 1
		    mAtomTable.Append( id )
		    
		    // Now register the hotkey
		    Declare Function MyRegisterHotKey Lib "User32" Alias "RegisterHotKey" ( hwnd as Integer, _
		    id as Integer, modifiers as Integer, virtualKey as Integer ) as Boolean
		    
		    dim ret as Boolean = MyRegisterHotKey( mWindow.WinHWND, id, modifiers, virtualKey )
		    
		    return id
		    
		  #else
		    
		    #pragma unused modifiers
		    #pragma unused virtualKey
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub UnregisterHotKey(id as Integer)
		  #if TargetWin32
		    
		    // FIrst, unregister the hotkey
		    Declare Function MyUnregisterHotKey Lib "User32" Alias "UnregisterHotKey" ( hwnd as Integer, _
		    id as Integer ) as Boolean
		    
		    Call MyUnregisterHotkey( mWindow.WinHWND, id )
		    
		    // Then delete the atom
		    Declare Sub GlobalDeleteAtom Lib "Kernel32" ( atom as Integer )
		    GlobalDeleteAtom( id )
		    
		    // Find the atom in our table and remove it
		    mAtomTable.Remove( mAtomTable.IndexOf( id ) )
		    
		  #else
		    
		    #pragma unused id
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function WndProc(hWnd as Integer, msg as Integer, wParam as Integer, lParam as Integer, ByRef returnValue as Integer) As Boolean
		  // Check to see if we got an WM_HOTKEY message
		  Const WM_HOTKEY = &h0312
		  
		  if msg = WM_HOTKEY then
		    // w00t!  We got a hot key press, so tell the
		    // user about it
		    if HotKeyPressed( wParam, ConstructKeyString( lParam ) ) then
		      returnValue = 1
		      return true
		    end
		  end
		  
		  return false
		  
		  #pragma unused hWnd
		End Function
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event HotKeyPressed(identifier as Integer, keyString as String) As Boolean
	#tag EndHook


	#tag Property, Flags = &h21
		Private mAtomTable(-1) As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mNextNum As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mWindow As Window
	#tag EndProperty


	#tag Constant, Name = kAltKey, Type = Integer, Dynamic = False, Default = \"&h1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kControlKey, Type = Integer, Dynamic = False, Default = \"&h2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kShiftKey, Type = Integer, Dynamic = False, Default = \"&h4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWindowsKey, Type = Integer, Dynamic = False, Default = \"&h8", Scope = Public
	#tag EndConstant


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
End Class
#tag EndClass
