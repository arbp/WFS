#tag Class
Protected Class DateTimePicker
	#tag Method, Flags = &h0
		Sub Constructor()
		  #if TargetWin32
		    // We need to initialize the common controls library and tell it
		    // that we are going to be adding date/time pickers
		    Declare Sub InitCommonControlsEx Lib "comctl32" ( ex as Ptr )
		    
		    dim ex as new MemoryBlock( 8 )
		    ex.Long( 0 ) = ex.Size
		    Const ICC_DATE_CLASSES = &h00000100
		    ex.Long( 4 ) = ICC_DATE_CLASSES
		    
		    InitCommonControlsEx( ex )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Create(onWindow as Window, x as Integer, y as Integer, width as Integer, height as Integer, datePicker as Boolean = true)
		  #if TargetWin32
		    Const DATETIMEPICK_CLASS = "SysDateTimePick32"
		    
		    Soft Declare Function CreateWindowExA Lib "User32" ( ex as Integer, className as CString, _
		    title as Integer, style as Integer, x as Integer, y as Integer, width as Integer, _
		    height as Integer, parent as Integer, menu as Integer, hInstance as Integer, _
		    lParam as Integer ) as Integer
		    
		    Soft Declare Function CreateWindowExW Lib "User32" ( ex as Integer, className as WString, _
		    title as Integer, style as Integer, x as Integer, y as Integer, width as Integer, _
		    height as Integer, parent as Integer, menu as Integer, hInstance as Integer, _
		    lParam as Integer ) as Integer
		    
		    Declare Function GetModuleHandleA Lib "Kernel32" ( name as Integer ) as Integer
		    
		    // First, get the HINSTANCE of this application so we can
		    // use it to create a window
		    dim hInstance as Integer
		    hInstance = GetModuleHandleA( 0 )
		    
		    Const WS_CHILD = &h40000000
		    Const WS_BORDER = &h00800000
		    Const WS_VISIBLE = &h10000000
		    Const DTS_SHOWNONE = &h2
		    Const DTS_UPDOWN = &h1
		    Const DTS_TIMEFORMAT = &h9
		    
		    dim flags as Integer
		    
		    if datePicker then
		      flags = flags + DTS_UPDOWN
		    else
		      flags = flags + DTS_TIMEFORMAT
		    end
		    
		    if System.IsFunctionAvailable( "CreateWindowExW", "User32" ) then
		      mWnd = CreateWindowExW( 0, DATETIMEPICK_CLASS, 0, _
		      WS_CHILD + WS_BORDER + WS_VISIBLE + flags, _
		      x, y, width, height, onWindow.WinHWND, 0, hInstance, 0 )
		    else
		      mWnd = CreateWindowExA( 0, DATETIMEPICK_CLASS, 0, _
		      WS_CHILD + WS_BORDER + WS_VISIBLE + flags, _
		      x, y, width, height, onWindow.WinHWND, 0, hInstance, 0 )
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Destructor()
		  #if TargetWin32
		    Declare Sub DestroyWindow Lib "User32" ( hwnd as Integer )
		    if mWnd <> 0 then
		      DestroyWindow( mWnd )
		      mWnd = 0
		    end
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetDateTime() As Date
		  #if TargetWin32
		    Const DTM_GETSYSTEMTIME = &h1001
		    Soft Declare Function SendMessageA Lib "User32" ( hWnd as Integer, _
		    msg as Integer, wParam as Integer, lParam as Integer ) as Integer
		    Soft Declare Function SendMessageW Lib "User32" ( hWnd as Integer, _
		    msg as Integer, wParam as Integer, lParam as Integer ) as Integer
		    
		    dim ret as Integer
		    dim sysTime as new MemoryBlock( 16 )
		    
		    if System.IsFunctionAvailable( "SendMessageW", "User32" ) then
		      ret = SendMessageW( mWnd, DTM_GETSYSTEMTIME, 0, MemoryBlockToInteger( sysTime ) )
		    else
		      ret = SendMessageA( mWnd, DTM_GETSYSTEMTIME, 0, MemoryBlockToInteger( sysTime ) )
		    end if
		    
		    if ret <> 0 then return nil
		    dim d as new Date
		    d.Year = sysTime.Short( 0 )
		    d.Month = sysTime.Short( 2 )
		    d.Day = sysTime.Short( 6 )
		    d.Hour = sysTime.Short( 8 )
		    d.Minute = sysTime.Short( 10 )
		    d.Second = sysTime.Short( 12 )
		    
		    return d
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function MemoryBlockToInteger(mb as MemoryBlock) As Integer
		  dim ret as new MemoryBlock( 4 )
		  ret.Ptr( 0 ) = mb
		  return ret.Long( 0 )
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mWnd As Integer
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
