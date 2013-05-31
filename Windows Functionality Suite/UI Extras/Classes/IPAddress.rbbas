#tag Class
Protected Class IPAddress
	#tag Method, Flags = &h0
		Function Address() As String
		  Const IPM_GETADDRESS = &h466
		  Const IPM_ISBLANK = &h469
		  
		  if SendMessage( IPM_ISBLANK, 0, 0 ) <> 0 then return ""
		  
		  dim result as new MemoryBlock( 4 )
		  dim numNonBlankField as Integer
		  dim trueResult as Integer
		  
		  numNonBlankField = SendMessage( IPM_GETADDRESS, 0, MemoryBlockToInteger( result ) )
		  
		  trueResult = result.Long( 0 )
		  
		  dim first, second, third, fourth as Integer
		  first = Bitwise.BitAnd( Bitwise.ShiftRight( trueResult, 24 ), &hFF )
		  second = Bitwise.BitAnd( Bitwise.ShiftRight( trueResult, 16 ), &hFF )
		  third = Bitwise.BitAnd( Bitwise.ShiftRight( trueResult, 8 ), &hFF )
		  fourth = Bitwise.BitAnd( trueResult, &hFF )
		  
		  return Str( first ) + "." + Str( second ) + "." + Str( third ) + "." + Str( fourth )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Address(assigns addy as String)
		  dim ip as Integer
		  dim i, numFields as Integer
		  dim fields() as Integer
		  
		  Const IPM_SETADDRESS = &h465
		  Const IPM_CLEARADDRESS = &h464
		  
		  if addy <> "" then
		    // Take the address and make it into a 32-bit integer to set
		    
		    numFields = CountFields( addy, "." )
		    for i = 1 to numFields
		      fields.Append( Val( NthField( addy, ".", i ) ) )
		    next
		    
		    ip = Bitwise.ShiftLeft( fields( 0 ), 24 ) + Bitwise.ShiftLeft( fields( 1 ), 16 ) + Bitwise.ShiftLeft( fields( 2 ), 8 ) + fields( 3 )
		    
		    Call SendMessage( IPM_SETADDRESS, 0, ip )
		  else
		    Call SendMessage( IPM_CLEARADDRESS, 0, 0 )
		  end
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  // Let's initialize the common controls stuff
		  #if TargetWin32
		    Declare Sub InitCommonControlsEx Lib "ComCtl32" ( ex as Ptr )
		    
		    Dim mb as new MemoryBlock( 8 )
		    mb.Long( 0 ) = mb.Size
		    Const ICC_INTERNET_CLASSES = &h800
		    mb.Long( 4 ) = ICC_INTERNET_CLASSES
		    
		    InitCommonControlsEx( mb )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Create(w as Window, x as Integer, y as Integer, width as Integer, height as Integer)
		  #if TargetWin32
		    Soft Declare Function CreateWindowExA Lib "User32" ( _
		    dwExStyle as Integer, lpClassName as CString, lpWindowName as Integer, _
		    dwStyle as Integer, x as Integer, y as Integer, width as Integer, height as Integer, _
		    hWndParent as Integer, hMenu as Integer, hInstance as Integer, lpParam as Integer ) as Integer
		    Soft Declare Function CreateWindowExW Lib "User32" ( _
		    dwExStyle as Integer, lpClassName as WString, lpWindowName as Integer, _
		    dwStyle as Integer, x as Integer, y as Integer, width as Integer, height as Integer, _
		    hWndParent as Integer, hMenu as Integer, hInstance as Integer, lpParam as Integer ) as Integer
		    Declare Function GetModuleHandleA Lib "Kernel32" ( name as Integer ) as Integer
		    
		    // Create the tooltip window
		    Const WC_IPADDRESSA = "SysIPAddress32"
		    
		    Const WS_CHILD = &h40000000
		    Const WS_VISIBLE = &h10000000
		    Const WS_BORDER = &h00800000
		    Const WS_TABSTOP = &h00010000
		    Const WS_EX_NOPARENTNOTIFY = &h4
		    
		    if System.IsFunctionAvailable( "CreateWindowExW", "User32" ) then
		      mWnd = CreateWindowExW( WS_EX_NOPARENTNOTIFY, WC_IPADDRESSA, 0, _
		      Bitwise.BitOr( WS_CHILD, WS_VISIBLE, WS_BORDER, WS_TABSTOP ), x, y, width, _
		      height, w.Handle, 0, GetModuleHandleA( 0 ), 0 )
		    else
		      mWnd = CreateWindowExA( 0, WC_IPADDRESSA, 0, _
		      Bitwise.BitOr( WS_CHILD, WS_VISIBLE, WS_BORDER, WS_TABSTOP ), x, y, width, _
		      height, w.WinHWND, 0, GetModuleHandleA( 0 ), 0 )
		    end if
		    
		    Declare Sub ShowWindow Lib "User32" ( hwnd as Integer, style as Integer )
		    ShowWindow( mWnd, 1 )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Destructor()
		  #if TargetWin32
		    Declare Sub DestroyWindow Lib "User32" ( hwnd as Integer )
		    
		    if mWnd <> 0 then
		      DestroyWindow( mWnd )
		    end
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function MemoryBlockToInteger(mb as MemoryBlock) As Integer
		  dim ret as new MemoryBlock( 4 )
		  ret.Ptr( 0 ) = mb
		  return ret.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function SendMessage(msg as Integer, wParam as Integer, lParam as Integer) As Integer
		  #if TargetWin32
		    Soft Declare Function SendMessageA Lib "User32" ( wnd as Integer, msg as Integer, _
		    wParam as Integer, lParam as Integer ) as Integer
		    Soft Declare Function SendMessageW Lib "User32" ( wnd as Integer, msg as Integer, _
		    wParam as Integer, lParam as Integer ) as Integer
		    
		    if System.IsFunctionAvailable( "SendMessageW", "User32" ) then
		      return SendMessageW( mWnd, msg, wParam, lParam )
		    else
		      return SendMessageA( mWnd, msg, wParam, lParam )
		    end if
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetFieldRange(fieldNumber as Integer = - 1, lowRange as Integer, hiRange as Integer)
		  Const IPM_SETRANGE = &h467
		  
		  dim rangeMash as Integer
		  rangeMash = Bitwise.ShiftLeft( hiRange, 8 ) + lowRange
		  
		  dim i as Integer
		  if fieldNumber = -1 then
		    for i = 0 to 3
		      Call SendMessage( IPM_SETRANGE, i, rangeMash )
		    next
		  else
		    Call SendMessage( IPM_SETRANGE, fieldNumber, rangeMash )
		  end
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetFocus(fieldNumber as Integer)
		  Const IPM_SETFOCUS = &h468
		  Call SendMessage( IPM_SETFOCUS, fieldNumber, 0 )
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected mWnd As Integer
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
End Class
#tag EndClass
