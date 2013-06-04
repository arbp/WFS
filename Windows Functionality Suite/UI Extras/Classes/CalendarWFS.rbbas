#tag Class
Protected Class CalendarWFS
	#tag Method, Flags = &h21
		Private Sub ChangeWindowStyle(flag as Integer, set as Boolean)
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
		    
		    oldFlags = GetWindowLong( mWnd, GWL_STYLE )
		    
		    if not set then
		      newFlags = BitwiseAnd( oldFlags, Bitwise.OnesComplement( flag ) )
		    else
		      newFlags = BitwiseOr( oldFlags, flag )
		    end
		    
		    
		    styleFlags = SetWindowLong( mWnd, GWL_STYLE, newFlags )
		    styleFlags = SetWindowPos( mWnd, 0, 0, 0, 0, 0, SWP_NOMOVE +_
		    SWP_NOSIZE + SWP_NOZORDER + SWP_FRAMECHANGED )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CircleToday() As Boolean
		  Const MCS_NOTODAYCIRCLE = &h08
		  return TestWindowStyle( MCS_NOTODAYCIRCLE )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CircleToday(assigns set as Boolean)
		  Const MCS_NOTODAYCIRCLE = &h08
		  ChangeWindowStyle( MCS_NOTODAYCIRCLE, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ColorToInt(c as Color) As Integer
		  dim mb as new MemoryBlock( 4 )
		  mb.Byte( 0 ) = c.Red
		  mb.Byte( 1 ) = c.Green
		  mb.Byte( 2 ) = c.Blue
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  #if TargetWin32
		    // We need to initialize the common controls library and tell it
		    // that we are going to be adding calendar controls
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
		Sub Create(onWindow as Window, x as Integer, y as Integer)
		  #if TargetWin32
		    Const MONTHCAL_CLASS = "SysMonthCal32"
		    
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
		    
		    Const WS_CHILD            = &h40000000
		    Const WS_BORDER = &h00800000
		    Const WS_VISIBLE = &h10000000
		    
		    dim flags as Integer
		    
		    if System.IsFunctionAvailable( "CreateWindowW", "User32" ) then
		      mWnd = CreateWindowExW( 0, MONTHCAL_CLASS, 0, _
		      WS_CHILD + WS_BORDER + WS_VISIBLE + flags, _
		      0, 0, 0, 0, onWindow.WinHWND, 0, hInstance, 0 )
		    else
		      mWnd = CreateWindowExA( 0, MONTHCAL_CLASS, 0, _
		      WS_CHILD + WS_BORDER + WS_VISIBLE + flags, _
		      0, 0, 0, 0, onWindow.WinHWND, 0, hInstance, 0 )
		    end if
		    
		    // Now we should get the size that is required to
		    // show the entire calendar
		    Dim r as new MemoryBlock( 16 )
		    Dim MCM_GETMINREQRECT as Integer = MCM_FIRST + 9
		    Call SendMessage( MCM_GETMINREQRECT, 0, MemoryBlockToInteger( r ) )
		    
		    // Now move the control to the proper place and
		    // give it the proper size
		    Declare Sub SetWindowPos Lib "User32" ( hwnd as Integer, after as Integer, _
		    x as Integer, y as Integer, width as Integer, height as Integer, flags as Integer )
		    Const SWP_NOZORDER = &h4
		    SetWindowPos( mWnd, 0, x, y, r.Long( 8 ), r.Long( 12 ), SWP_NOZORDER )
		    
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
		Function DisplayToday() As Boolean
		  Const MCS_NOTODAY = &h10
		  return TestWindowStyle( MCS_NOTODAY )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DisplayToday(assigns set as Boolean)
		  Const MCS_NOTODAY = &h10
		  ChangeWindowStyle( MCS_NOTODAY, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetColor(which as Integer) As Color
		  #if TargetWin32
		    Dim MCM_GETCOLOR as Integer = MCM_FIRST + 11
		    return IntToColor( SendMessage( MCM_GETCOLOR, which, 0 ) )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IntToColor(i as Integer) As Color
		  dim mb as new MemoryBlock( 4 )
		  
		  mb.Long( 0 ) = i
		  
		  return RGB( mb.Byte( 0 ), mb.Byte( 1 ), mb.Byte( 2 ) )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function MemoryBlockToInteger(mb as MemoryBlock) As Integer
		  dim ret as new MemoryBlock( 4 )
		  ret.Ptr( 0 ) = mb
		  return ret.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RBDateToSystemTime(d as Date) As MemoryBlock
		  dim mb as new MemoryBlock( 16 )
		  mb.Short( 0 ) = d.Year
		  mb.Short( 2 ) = d.Month
		  mb.Short( 6 ) = d.Day
		  
		  return mb
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function SelectedDate() As Date
		  #if TargetWin32
		    Dim MCM_GETCURSEL as Integer = MCM_FIRST + 1
		    
		    dim mb as new MemoryBlock( 16 )
		    Call SendMessage( MCM_GETCURSEL, 0, MemoryBlockToInteger( mb ) )
		    
		    return SystemTimeToRBDate( mb )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SelectedDate(assigns d as Date)
		  #if TargetWin32
		    Dim MCM_SETCURSEL as Integer = MCM_FIRST + 2
		    
		    
		    Call SendMessage( MCM_SETCURSEL, 0, MemoryBlockToInteger( RBDateToSystemTime( d ) ) )
		  #endif
		End Sub
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
		Sub SetColor(which as Integer, assigns c as Color)
		  #if TargetWin32
		    Dim MCM_SETCOLOR as Integer = MCM_FIRST + 10
		    Call SendMessage( MCM_SETCOLOR, which, ColorToInt( c ) )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function SystemTimeToRBDate(mb as MemoryBlock) As Date
		  dim d as new Date
		  d.Year = mb.Short( 0 )
		  d.Month = mb.Short( 2 )
		  d.Day = mb.Short( 6 )
		  
		  return d
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function TestWindowStyle(flag as Integer) As Boolean
		  #if TargetWin32
		    Dim oldFlags as Integer
		    
		    Const GWL_STYLE = -16
		    
		    Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (hwnd As Integer,  _
		    nIndex As Integer) As Integer
		    
		    oldFlags = GetWindowLong( mWnd, GWL_STYLE )
		    
		    if Bitwise.BitAnd( oldFlags, flag ) = flag then
		      return true
		    else
		      return false
		    end
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Today() As Date
		  #if TargetWin32
		    Dim MCM_GETTODAY as Integer = MCM_FIRST + 13
		    dim mb as new MemoryBlock( 16 )
		    Call SendMessage( MCM_GETTODAY, 0, MemoryBlockToInteger( mb ) )
		    
		    return SystemTimeToRBDate( mb )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Today(assigns d as Date)
		  #if TargetWin32
		    Dim MCM_SETTODAY as Integer = MCM_FIRST + 12
		    Call SendMessage( MCM_SETTODAY, 0, MemoryBlockToInteger( RBDateToSystemTime( d ) ) )
		  #endif
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mWnd As Integer
	#tag EndProperty


	#tag Constant, Name = kBackgroundColor, Type = Integer, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kMonthBackgroundColor, Type = Integer, Dynamic = False, Default = \"4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kTextColor, Type = Integer, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kTitleColor, Type = Integer, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kTitleTextColor, Type = Integer, Dynamic = False, Default = \"3", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kTrailingTextColor, Type = Integer, Dynamic = False, Default = \"5", Scope = Public
	#tag EndConstant

	#tag Constant, Name = MCM_FIRST, Type = Integer, Dynamic = False, Default = \"&h1000", Scope = Private
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
