#tag Module
Protected Module MouseWFS
	#tag Method, Flags = &h1
		Protected Sub LeftButtonDown()
		  #if TargetWin32
		    Declare Sub mouse_event Lib "User32" ( flags as Integer, _
		    dx as Integer, dy as Integer, data as Integer, ByRef extra as Integer )
		    
		    Const MOUSEEVENTF_LEFTDOWN    = &h0002 // left button down
		    
		    dim extra as Integer
		    dim flags as Integer
		    flags = MOUSEEVENTF_LEFTDOWN
		    
		    mouse_event( flags, 0, 0, 0, extra )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LeftButtonUp()
		  #if TargetWin32
		    Declare Sub mouse_event Lib "User32" ( flags as Integer, _
		    dx as Integer, dy as Integer, data as Integer, ByRef extra as Integer )
		    
		    Const MOUSEEVENTF_LEFTUP      = &h0004 // left button up
		    
		    dim extra as Integer
		    dim flags as Integer
		    flags = MOUSEEVENTF_LEFTUP
		    
		    mouse_event( flags, 0, 0, 0, extra )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MiddleButtonDown()
		  #if TargetWin32
		    Declare Sub mouse_event Lib "User32" ( flags as Integer, _
		    dx as Integer, dy as Integer, data as Integer, ByRef extra as Integer )
		    
		    Const MOUSEEVENTF_MIDDLEDOWN  = &h0020 // middle button down
		    
		    dim extra as Integer
		    dim flags as Integer
		    flags = MOUSEEVENTF_MIDDLEDOWN
		    
		    mouse_event( flags, 0, 0, 0, extra )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MiddleButtonUp()
		  #if TargetWin32
		    Declare Sub mouse_event Lib "User32" ( flags as Integer, _
		    dx as Integer, dy as Integer, data as Integer, ByRef extra as Integer )
		    
		    Const MOUSEEVENTF_MIDDLEUP    = &h0040 // middle button up
		    
		    dim extra as Integer
		    dim flags as Integer
		    flags = MOUSEEVENTF_MIDDLEUP
		    
		    mouse_event( flags, 0, 0, 0, extra )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MoveMouse(x as Integer, y as Integer, absolute as Boolean = false)
		  #if TargetWin32
		    Declare Sub mouse_event Lib "User32" ( flags as Integer, _
		    dx as Integer, dy as Integer, data as Integer, ByRef extra as Integer )
		    
		    Const MOUSEEVENTF_MOVE        = &h0001 // mouse move
		    Const MOUSEEVENTF_ABSOLUTE    = &h8000 // absolute move
		    
		    dim extra as Integer
		    dim flags as Integer
		    flags = MOUSEEVENTF_MOVE
		    if absolute then
		      flags = flags + MOUSEEVENTF_ABSOLUTE
		      
		      // If we're in absolute mode, then we need to
		      // map X and Y to proper mapped coordinates.  Windows
		      // maps 0,0 as the upper left, and 65535, 65535 as
		      // the lower right
		      x = (x / Screen( 0 ).Width) * 65535
		      y = (y / Screen( 0 ).Height ) * 65535
		    end
		    mouse_event( flags, x, y, 0, extra )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub RightButtonDown()
		  #if TargetWin32
		    Declare Sub mouse_event Lib "User32" ( flags as Integer, _
		    dx as Integer, dy as Integer, data as Integer, ByRef extra as Integer )
		    
		    Const MOUSEEVENTF_RIGHTDOWN   = &h0008 // right button down
		    
		    dim extra as Integer
		    dim flags as Integer
		    flags = MOUSEEVENTF_RIGHTDOWN
		    
		    mouse_event( flags, 0, 0, 0, extra )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub RightButtonUp()
		  #if TargetWin32
		    Declare Sub mouse_event Lib "User32" ( flags as Integer, _
		    dx as Integer, dy as Integer, data as Integer, ByRef extra as Integer )
		    
		    Const MOUSEEVENTF_RIGHTUP     = &h0010 // right button up
		    
		    dim extra as Integer
		    dim flags as Integer
		    flags = MOUSEEVENTF_RIGHTUP
		    
		    mouse_event( flags, 0, 0, 0, extra )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub SpinMouseWheel(delta as Integer)
		  #if TargetWin32
		    Declare Sub mouse_event Lib "User32" ( flags as Integer, _
		    dx as Integer, dy as Integer, data as Integer, ByRef extra as Integer )
		    
		    Const MOUSEEVENTF_WHEEL       = &h0800 // wheel button rolled
		    
		    dim extra as Integer
		    dim flags as Integer
		    flags = MOUSEEVENTF_WHEEL
		    
		    mouse_event( flags, 0, 0, delta, extra )
		  #endif
		End Sub
	#tag EndMethod


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
