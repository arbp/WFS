#tag Module
Protected Module PushButtonExtensions
	#tag Method, Flags = &h0
		Function BottomJustified(extends p as PushButton) As Boolean
		  Const BS_BOTTOM = &h800
		  
		  return TestWindowStyle( p.Handle, BS_BOTTOM )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub BottomJustified(extends p as PushButton, assigns set as Boolean)
		  Const BS_BOTTOM = &h800
		  
		  ChangeWindowStyle( p.Handle, BS_BOTTOM, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CenterJustified(extends p as PushButton) As Boolean
		  Const BS_CENTER = &h300
		  return TestWindowStyle( p.Handle, BS_CENTER )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CenterJustified(extends p as PushButton, assigns set as Boolean)
		  Const BS_CENTER = &h300
		  ChangeWindowStyle( p.Handle, BS_CENTER, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ChangeWindowStyle(w as Integer, flag as Integer, set as Boolean)
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
		    
		    oldFlags = GetWindowLong( w, GWL_STYLE )
		    
		    if not set then
		      newFlags = BitwiseAnd( oldFlags, Bitwise.OnesComplement( flag ) )
		    else
		      newFlags = BitwiseOr( oldFlags, flag )
		    end
		    
		    
		    styleFlags = SetWindowLong( w, GWL_STYLE, newFlags )
		    styleFlags = SetWindowPos( w, 0, 0, 0, 0, 0, SWP_NOMOVE +_
		    SWP_NOSIZE + SWP_NOZORDER + SWP_FRAMECHANGED )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Flat(extends p as PushButton) As Boolean
		  Const BS_FLAT = &h8000
		  
		  return TestWindowStyle( p.Handle, BS_FLAT )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Flat(extends p as PushButton, assigns set as Boolean)
		  Const BS_FLAT = &h8000
		  
		  ChangeWindowStyle( p.Handle, BS_FLAT, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub HasShield(extends pb as PushButton, assigns set as Boolean)
		  Const BCM_SETSHIELD = &h160C
		  if set then
		    call SendMessage( pb.Handle, BCM_SETSHIELD, 0, 1 )
		  else
		    call SendMessage( pb.Handle, BCM_SETSHIELD, 0, 0 )
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function LeftJustified(extends p as PushButton) As Boolean
		  Const BS_LEFT = &h100
		  
		  return TestWindowStyle( p.Handle, BS_LEFT )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LeftJustified(extends p as PushButton, assigns set as Boolean)
		  Const BS_LEFT = &h100
		  
		  ChangeWindowStyle( p.Handle, BS_LEFT, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Multiline(extends p as PushButton) As Boolean
		  Const BS_MULTILINE = &h2000
		  return TestWindowStyle( p.Handle, BS_MULTILINE )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Multiline(extends p as PushButton, assigns set as Boolean)
		  Const BS_MULTILINE = &h2000
		  ChangeWindowStyle( p.Handle, BS_MULTILINE, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function RightJustified(extends p as PushButton) As Boolean
		  Const BS_RIGHT = &h200
		  return TestWindowStyle( p.Handle, BS_RIGHT )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub RightJustified(extends p as PushButton, assigns set as Boolean)
		  Const BS_RIGHT = &h200
		  ChangeWindowStyle( p.Handle, BS_RIGHT, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function SendMessage(hwnd as Integer, msg as Integer, wParam as Integer, lParam as Integer) As Integer
		  #if TargetWin32
		    Soft Declare Function SendMessageA Lib "User32" ( hwnd as Integer, msg as Integer, wParam as Integer, lParam as Integer ) as Integer
		    Soft Declare Function SendMessageW Lib "User32" ( hwnd as Integer, msg as Integer, wParam as Integer, lParam as Integer ) as Integer
		    
		    if System.IsFunctionAvailable( "SendMessageW", "User32" ) then
		      return SendMessageW( hwnd, msg, wParam, lParam )
		    else
		      return SendMessageA( hwnd, msg, wParam, lParam )
		    end if
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function TestWindowStyle(w as Integer, flag as Integer) As Boolean
		  #if TargetWin32
		    Dim oldFlags as Integer
		    
		    Const GWL_STYLE = -16
		    
		    Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (hwnd As Integer,  _
		    nIndex As Integer) As Integer
		    
		    oldFlags = GetWindowLong( w, GWL_STYLE )
		    
		    if Bitwise.BitAnd( oldFlags, flag ) = flag then
		      return true
		    else
		      return false
		    end
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TopJustified(extends p as PushButton) As Boolean
		  Const BS_TOP = &h400
		  
		  return TestWindowStyle( p.Handle, BS_TOP )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub TopJustified(extends p as PushButton, assigns set as Boolean)
		  Const BS_TOP = &h400
		  
		  ChangeWindowStyle( p.Handle, BS_TOP, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function VerticalCenterJustified(extends p as PushButton) As Boolean
		  Const BS_VCENTER = &hC00
		  
		  return TestWindowStyle( p.Handle, BS_VCENTER )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub VerticalCenterJustified(extends p as PushButton, assigns set as Boolean)
		  Const BS_VCENTER = &hC00
		  
		  ChangeWindowStyle( p.Handle, BS_VCENTER, set )
		End Sub
	#tag EndMethod


	#tag Note, Name = About this module
		1) You may be wondering why there's no HasShield which returns a boolean.  That's
		because Microsoft doesn't expose a message to query for that information.  If they add
		a BCM_GETSHIELD message, then we can implement it.
	#tag EndNote


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
