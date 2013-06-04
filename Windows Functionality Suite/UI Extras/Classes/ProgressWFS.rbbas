#tag Class
Protected Class ProgressWFS
	#tag Method, Flags = &h0
		Sub Attach(w as Window)
		  #if TargetWin32
		    
		    Const PROGRESS_CLASS = "msctls_progress32"
		    
		    Soft Declare Function CreateWindowExA Lib "User32" ( ex as Integer, _
		    className as CString, windowName as CString,  style as Integer, _
		    x as Integer, y as Integer, cx as Integer, cy as Integer, _
		    parent as Integer, menu as Integer, hInstance as Integer, _
		    lParam as Integer ) as Integer
		    Soft Declare Function CreateWindowExW Lib "User32" ( ex as Integer, _
		    className as WString, windowName as WString,  style as Integer, _
		    x as Integer, y as Integer, cx as Integer, cy as Integer, _
		    parent as Integer, menu as Integer, hInstance as Integer, _
		    lParam as Integer ) as Integer
		    
		    Declare Function GetModuleHandleA Lib "Kernel32" ( null as Integer ) as Integer
		    
		    dim hInstance as Integer = GetModuleHandleA( 0 )
		    dim hWnd as Integer
		    
		    Const PBS_SMOOTH = &h1
		    Const WS_EX_CLIENTEDGE = &h00000200
		    Const WS_CHILD = &h40000000
		    
		    dim style as Integer
		    
		    if SmoothView then
		      style = PBS_SMOOTH
		    end
		    
		    if System.IsFunctionAvailable( "CreateWindowExW", "User32" ) then
		      hWnd = CreateWindowExW( WS_EX_CLIENTEDGE, PROGRESS_CLASS, "", _
		      style+ WS_CHILD, _
		      X, Y, Width, Height, w.WinHWND, 0, hInstance, 0 )
		    else
		      hWnd = CreateWindowExA( WS_EX_CLIENTEDGE, PROGRESS_CLASS, "", _
		      style+ WS_CHILD, _
		      X, Y, Width, Height, w.WinHWND, 0, hInstance, 0 )
		    end if
		    
		    Declare Sub ShowWindow Lib "User32" ( hwnd as Integer, nCmd as Integer )
		    ShowWindow( hWnd, 1 )
		    
		    mWnd = hWnd
		    
		    if mInitialMinValue <> 0 then
		      me.MinValue = mInitialMinValue
		    end
		    
		    if mInitialMaxValue <> 0 then
		      me.MaxValue = mInitialMaxValue
		    end
		    
		    if mInitialValue <> 0 then
		      me.Value = mInitialValue
		    end
		    
		    if mStepValue <> 10 then
		      me.StepValue = mStepValue
		    end
		    
		    if mBarColor <> &c0000FF then
		      me.BarColor = mBarColor
		    end
		    
		    if mBackColor <> &c777777 then
		      me.BackColor = mBackColor
		    end
		    
		  #else
		    
		    #pragma unused w
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function BackColor() As Color
		  return mBackColor
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub BackColor(assigns newVal as Color)
		  mBackColor = newVal
		  
		  if mWnd = 0 then return
		  
		  dim PBM_SETBKCOLOR as Integer
		  PBM_SETBKCOLOR = &h2001
		  
		  Call SendMessage( PBM_SETBKCOLOR, 0, ColorToInt( newVal ) )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function BarColor() As Color
		  return mBarColor
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub BarColor(assigns newVal as Color)
		  mBarColor = newVal
		  
		  if mWnd = 0 then return
		  
		  Const WM_USER = &h400
		  dim PBM_SETBARCOLOR as Integer
		  PBM_SETBARCOLOR = WM_USER + 9
		  
		  Call SendMessage( PBM_SETBARCOLOR, 0, ColorToInt( newVal ) )
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
		    Declare Sub InitCommonControlsEx Lib "comctl32" ( ex as Ptr )
		    
		    Const ICC_PROGRESS_CLASS = &h20
		    
		    dim mb as new MemoryBlock( 8 )
		    mb.Long( 0 ) = mb.SIze
		    mb.Long( 4 ) = ICC_PROGRESS_CLASS
		    
		    InitCommonControlsEx( mb )
		    
		    mStepValue = 10
		    mBarColor = &c0000FF
		    mBackColor = &c777777
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function MaxValue() As Integer
		  Const WM_USER = &h400
		  dim PBM_GETRANGE as Integer
		  PBM_GETRANGE = WM_USER + 7
		  
		  dim range as new MemoryBlock( 8 )
		  dim rangePtr as new MemoryBlock( 4 )
		  rangePtr.Ptr( 0 ) = range
		  Call SendMessage( PBM_GETRANGE, 0, rangePtr.Long( 0 ) )
		  
		  return range.Long( 4 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub MaxValue(assigns newVal as Integer)
		  if mWnd = 0 then
		    mInitialMaxValue = newVal
		    return
		  end
		  
		  Const WM_USER = &h400
		  dim PBM_SETRANGE32 as Integer
		  PBM_SETRANGE32 = WM_USER + 6
		  
		  dim minRange as Integer
		  minRange = me.MinValue
		  
		  Call SendMessage( PBM_SETRANGE32, minValue, newVal )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function MinValue() As Integer
		  Const WM_USER = &h400
		  dim PBM_GETRANGE as Integer
		  PBM_GETRANGE = WM_USER + 7
		  
		  dim range as new MemoryBlock( 8 )
		  dim rangePtr as new MemoryBlock( 4 )
		  rangePtr.Ptr( 0 ) = range
		  Call SendMessage( PBM_GETRANGE, 0, rangePtr.Long( 0 ) )
		  
		  return range.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub MinValue(assigns newVal as Integer)
		  if mWnd = 0 then
		    mInitialMinValue= newVal
		    return
		  end
		  
		  Const WM_USER = &h400
		  dim PBM_SETRANGE32 as Integer
		  PBM_SETRANGE32 = WM_USER + 6
		  
		  dim maxRange as Integer
		  maxRange = me.MaxValue
		  
		  Call SendMessage( PBM_SETRANGE32, newVal, maxRange )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function SendMessage(msg as Integer, wParam as Integer, lParam as Integer) As Integer
		  if mWnd = 0 then return 0
		  
		  #if TargetWin32
		    
		    Soft Declare Function SendMessageA Lib "User32" ( hWnd as Integer, _
		    msg as Integer, wParam as Integer, lParam as Integer ) as Integer
		    Soft Declare Function SendMessageW Lib "User32" ( hWnd as Integer, _
		    msg as Integer, wParam as Integer, lParam as Integer ) as Integer
		    
		    if System.IsFunctionAvailable( "SendMessageW", "User32" ) then
		      return SendMessageW( mWnd, msg, wParam, lParam )
		    else
		      return SendMessageA( mWnd, msg, wParam, lParam )
		    end if
		    
		  #else
		    
		    #pragma unused msg
		    #pragma unused wParam
		    #pragma unused lParam
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function State() As Integer
		  Const PBM_GETSTATE = &h411
		  return SendMessage( PBM_GETSTATE, 0, 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub State(assigns type as Integer)
		  Const PBM_SETSTATE = &h410
		  call SendMessage( PBM_SETSTATE, type, 0 )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub StepOnce()
		  Const WM_USER = &h400
		  dim PBM_STEPIT as Integer
		  PBM_STEPIT = WM_USER + 5
		  
		  Call SendMessage( PBM_STEPIT, 0, 0 )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function StepValue() As Integer
		  return mStepValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub StepValue(assigns newVal as Integer)
		  mStepValue = newVal
		  
		  if mWnd = 0 then return
		  
		  Const WM_USER = &h400
		  dim PBM_SETSTEP as Integer
		  PBM_SETSTEP = WM_USER + 4
		  
		  Call SendMessage( PBM_SETSTEP, newVal, 0 )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Value() As Integer
		  Const WM_USER = &h400
		  dim PBM_GETPOS as Integer
		  PBM_GETPOS = WM_USER + 8
		  
		  return SendMessage( PBM_GETPOS, 0, 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Value(assigns newVal as Integer)
		  if mWnd = 0 then
		    // There's no window yet
		    mInitialValue = newVal
		    return
		  end
		  
		  Const WM_USER = &h400
		  dim PBM_SETPOS as Integer
		  PBM_SETPOS = WM_USER + 2
		  
		  Call SendMessage( PBM_SETPOS, newVal, 0 )
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		Height As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mBackColor As Color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mBarColor As Color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mInitialMaxValue As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mInitialMinValue As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mInitialValue As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mStepValue As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mWnd As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		SmoothView As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		Width As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		X As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Y As Integer
	#tag EndProperty


	#tag Constant, Name = kStateError, Type = Double, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kStateNormal, Type = Double, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kStatePaused, Type = Double, Dynamic = False, Default = \"3", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="Height"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
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
			Name="SmoothView"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
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
		#tag ViewProperty
			Name="Width"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="X"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Y"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
