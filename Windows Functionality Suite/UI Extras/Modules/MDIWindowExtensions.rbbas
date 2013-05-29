#tag Module
Protected Module MDIWindowExtensions
	#tag Method, Flags = &h0
		Sub CascadeChildren(extends w as MDIWindow)
		  #if TargetWin32
		    Declare Sub CascadeWindows Lib "User32" ( parent as Integer, how as Integer, r as Integer, num as Integer, kids as Integer )
		    
		    dim clientHandle as Integer = w.MDIClientHandle
		    dim how as Integer
		    
		    if clientHandle <> 0 then
		      CascadeWindows( clientHandle, 0, 0, 0, 0 )
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ChangeWindowStyle(w as MDIWindow, flag as Integer, set as Boolean)
		  #if TargetWin32
		    Dim oldFlags as Integer
		    Dim newFlags as Integer
		    Dim styleFlags As Integer
		    
		    Const SWP_NOSIZE = &H1
		    Const SWP_NOMOVE = &H2
		    Const SWP_NOZORDER = &H4
		    Const SWP_FRAMECHANGED = &H20
		    
		    Const GWL_STYLE = -16
		    
		    Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (hwnd As Integer,  _
		    nIndex As Integer) As Integer
		    Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (hwnd As Integer, _
		    nIndex As Integer, dwNewLong As Integer) As Integer
		    Declare Function SetWindowPos Lib "user32" (hwnd as Integer, hWndInstertAfter as Integer, _
		    x as Integer, y as Integer, cx as Integer, cy as Integer, flags as Integer) as Integer
		    
		    oldFlags = GetWindowLong(w.Handle, GWL_STYLE)
		    
		    if not set then
		      newFlags = BitwiseAnd( oldFlags, Bitwise.OnesComplement( flag ) )
		    else
		      newFlags = BitwiseOr( oldFlags, flag )
		    end
		    
		    
		    styleFlags = SetWindowLong( w.Handle, GWL_STYLE, newFlags )
		    styleFlags = SetWindowPos( w.Handle, 0, 0, 0, 0, 0, SWP_NOMOVE +_
		    SWP_NOSIZE + SWP_NOZORDER + SWP_FRAMECHANGED )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CloseButtonState(extends w as MDIWindow) As Boolean
		  #if TargetWin32
		    Declare Function GetSystemMenu Lib "User32" ( wnd as Integer, revert as Boolean ) as Integer
		    Declare Function GetMenuState Lib "User32" ( menu as Integer, which as Integer, flags as Integer ) as Integer
		    
		    dim menu as Integer = GetSystemMenu( w.Handle, false )
		    if menu = 0 then return false
		    
		    Const SC_CLOSE = &hF060
		    Const MF_BYCOMMAND = 0
		    Const MF_GRAYED = 1
		    Const MF_ENABLED = 0
		    Const MF_DISABLED = 2
		    
		    dim state as Integer
		    state = GetMenuState( menu, SC_CLOSE, MF_BYCOMMAND )
		    
		    return Bitwise.BitAnd( state, MF_DISABLED + MF_GRAYED ) = 0
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CloseButtonState(extends w as MDIWindow, assigns enabled as Boolean)
		  #if TargetWin32
		    Declare Function GetSystemMenu Lib "User32" ( wnd as Integer, revert as Boolean ) as Integer
		    Declare Sub EnableMenuItem Lib "User32" ( menu as Integer, which as Integer, flags as Integer )
		    
		    dim menu as Integer = GetSystemMenu( w.Handle, false )
		    if menu = 0 then return
		    
		    Const SC_CLOSE = &hF060
		    Const MF_BYCOMMAND = 0
		    Const MF_GRAYED = 1
		    Const MF_ENABLED = 0
		    
		    if not enabled then
		      EnableMenuItem( menu, SC_CLOSE, MF_BYCOMMAND + MF_GRAYED )
		    else
		      EnableMenuItem( menu, SC_CLOSE, MF_BYCOMMAND + MF_ENABLED )
		    end
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function enumChildProc(hwnd as Integer, lParam as Integer) As Boolean
		  #if TargetWin32
		    // We need to figure out what class this window belongs to
		    Soft Declare Function GetClassNameW Lib "User32" ( hwnd as Integer, name as Ptr, count as Integer ) as Integer
		    Soft Declare Function GetClassNameA Lib "User32" ( hwnd as Integer, name as Ptr, count as Integer ) as Integer
		    
		    dim classNamePtr as new MemoryBlock( 256 )
		    dim className as String
		    if System.IsFunctionAvailable( "GetClassNameW", "User32" ) then
		      dim cnt as Integer = GetClassNameW( hwnd, classNamePtr, classNamePtr.Size )
		      
		      className = classNamePtr.WString( 0 )
		    else
		      dim cnt as Integer = GetClassNameA( hwnd, classNamePtr, classNamePtr.Size )
		      
		      className = classNamePtr.CString( 0 )
		    end if
		    
		    // If the name is MDICLIENT, then we're done
		    if className = "MDICLIENT" then
		      mClientHandle = hwnd
		      return false
		    end if
		    
		    return true
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HasMaximizeButton(extends w as MDIWindow) As Boolean
		  Const WS_MAXIMIZEBOX = &h00010000
		  
		  return TestWindowStyle( w, WS_MAXIMIZEBOX )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub HasMaximizeButton(extends w as MDIWindow, assigns set as Boolean)
		  Const WS_MAXIMIZEBOX = &h00010000
		  
		  ChangeWindowStyle( w, WS_MAXIMIZEBOX, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HasMinimizeButton(extends w as MDIWindow) As Boolean
		  Const WS_MINIMIZEBOX = &h00020000
		  
		  return TestWindowStyle( w, WS_MINIMIZEBOX )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub HasMinimizeButton(extends w as MDIWindow, assigns set as Boolean)
		  Const WS_MINIMIZEBOX = &h00020000
		  
		  ChangeWindowStyle( w, WS_MINIMIZEBOX, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HasSystemMenu(extends w as MDIWindow) As Boolean
		  Const WS_SYSMENU = &h00080000
		  
		  return TestWindowStyle( w, WS_SYSMENU )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub HasSystemMenu(extends w as MDIWindow, assigns set as Boolean)
		  Const WS_SYSMENU = &h00080000
		  
		  ChangeWindowStyle( w, WS_SYSMENU, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsMaximized(extends w as MDIWindow) As Boolean
		  #if TargetWin32
		    Declare Function IsZoomed Lib "User32" ( hwnd As Integer ) As Integer
		    
		    return IsZoomed( w.Handle ) <> 0
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsMinimized(extends w as MDIWindow) As Boolean
		  #if TargetWin32
		    Declare Function IsIconic Lib "User32" ( hwnd As Integer ) As Integer
		    
		    return IsIconic( w.Handle ) <> 0
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function MDIClientHandle(extends w as MDIWindow) As Integer
		  // There's two different handles used for an MDI window.  The frame
		  // handle (which is MDIWindow.Handle), and the client handle.  This
		  // gets the client handle, which is used for things like tiling or cascading
		  // child windows.
		  
		  if mClientHandle <> 0 then return mClientHandle
		  
		  #if TargetWin32
		    Declare Sub EnumChildWindows Lib "User32" ( parent as Integer, proc as Ptr, lParam as Integer )
		    
		    mClientHandle = 0
		    
		    // Do the enumeration
		    EnumChildWindows( w.Handle, AddressOf enumChildProc, 0 )
		    
		    // Return the client's handle
		    return mClientHandle
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function TestWindowStyle(w as MDIWindow, flag as Integer) As Boolean
		  #if TargetWin32
		    Dim oldFlags as Integer
		    
		    Const GWL_STYLE = -16
		    
		    Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (hwnd As Integer,  _
		    nIndex As Integer) As Integer
		    
		    oldFlags = GetWindowLong(w.Handle, GWL_STYLE)
		    
		    if Bitwise.BitAnd( oldFlags, flag ) = flag then
		      return true
		    else
		      return false
		    end
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub TileChildren(extends w as MDIWindow, horizontal as Boolean = true)
		  #if TargetWin32
		    Declare Sub TileWindows Lib "User32" ( parent as Integer, how as Integer, r as Integer, num as Integer, kids as Integer )
		    
		    dim clientHandle as Integer = w.MDIClientHandle
		    dim how as Integer
		    
		    Const MDITILE_HORIZONTAL = &h1
		    if horizontal then how = MDITILE_HORIZONTAL
		    
		    if clientHandle <> 0 then
		      TileWindows( clientHandle, how, 0, 0, 0 )
		    end if
		  #endif
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mClientHandle As Integer
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
