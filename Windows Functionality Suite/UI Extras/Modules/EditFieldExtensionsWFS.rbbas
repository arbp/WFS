#tag Module
Protected Module EditFieldExtensionsWFS
	#tag Method, Flags = &h21
		Private Sub ChangeWindowStyle(w as Integer, flag as Integer, set as Boolean, ex as Boolean = false)
		  #if TargetWin32
		    
		    Dim oldFlags as Integer
		    Dim newFlags as Integer
		    Dim styleFlags As Integer
		    
		    Const SWP_NOSIZE = &H1
		    Const SWP_NOMOVE = &H2
		    Const SWP_NOZORDER = &H4
		    Const SWP_FRAMECHANGED = &H20
		    
		    Const GWL_STYLE = -16
		    Const GWL_EXSTYLE = -20
		    
		    Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (hwnd As Integer,  _
		    nIndex As Integer) As Integer
		    Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (hwnd As Integer, _
		    nIndex As Integer, dwNewLong As Integer) As Integer
		    Declare Function SetWindowPos Lib "user32" (hwnd as Integer, hWndInstertAfter as Integer, _
		    x as Integer, y as Integer, cx as Integer, cy as Integer, flags as Integer) as Integer
		    
		    dim style as Integer = GWL_STYLE
		    if ex then style = GWL_EXSTYLE
		    
		    oldFlags = GetWindowLong( w, style )
		    
		    if not set then
		      newFlags = BitwiseAnd( oldFlags, Bitwise.OnesComplement( flag ) )
		    else
		      newFlags = BitwiseOr( oldFlags, flag )
		    end
		    
		    
		    styleFlags = SetWindowLong( w, style, newFlags )
		    styleFlags = SetWindowPos( w, 0, 0, 0, 0, 0, SWP_NOMOVE +_
		    SWP_NOSIZE + SWP_NOZORDER + SWP_FRAMECHANGED )
		    
		  #else
		    
		    #pragma unused w
		    #pragma unused flag
		    #pragma unused set
		    #pragma unused ex
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LineSpacingWFS(extends EditField1 as EditField, spacing as Double)
		  #if TargetWin32
		    
		    Dim EM_SETPARAFORMAT as Integer = &h400 + 71
		    Dim EM_HIDESELECTION as Integer = &h400 + 63
		    
		    Dim paraFormat2 as new MemoryBlock( 188 )
		    
		    paraFormat2.Long( 0 ) = paraFormat2.Size
		    
		    Const PFM_LINESPACING = &h100
		    Const PFM_SPACEAFTER = &h80
		    paraFormat2.Long( 4 ) = PFM_LINESPACING
		    
		    dim spacingRule as Integer
		    select case spacing
		    case 1
		      spacingRule = 0
		    case 1.5
		      spacingRule = 1
		    case 2
		      spacingRule = 2
		    end select
		    
		    paraFormat2.Byte( 170 ) = spacingRule
		    
		    Declare Sub SendMessageA Lib "User32" ( hwnd as Integer, msg as Integer, wParam as Integer, lParam as Ptr )
		    
		    // If the user has something selected, then we
		    // want to apply to just that paragraph.  However,
		    // if they don't have anything selected, then we
		    // want to apply to the whole field.
		    dim restoreSelectionPoint as Integer = -1
		    if EditField1.SelLength > 0 then
		      // Just use what the user has
		    else
		      // Hide the selection
		      SendMessageA( EditField1.Handle, EM_HIDESELECTION, 1, nil )
		      
		      restoreSelectionPoint = EditField1.SelStart
		      EditField1.SelStart = 0
		      EditField1.SelLength = -1
		      
		    end if
		    
		    // Set the format
		    SendMessageA( EditField1.Handle, EM_SETPARAFORMAT, 0, paraFormat2 )
		    
		    if restoreSelectionPoint >= 0 then
		      // Restore the selection
		      EditField1.SelStart = restoreSelectionPoint
		      EditField1.SelLength = 0
		      
		      // Show the selection
		      SendMessageA( EditField1.Handle, EM_HIDESELECTION, 0, nil )
		    end if
		    
		  #else
		    
		    #pragma unused EditField1
		    #pragma unused spacing
		    
		  #endif
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function SunkenWFS(extends e as EditField) As Boolean
		  Const ES_SUNKEN = &h4000
		  return TestWindowStyle( e.Handle, ES_SUNKEN )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SunkenWFS(extends e as EditField, assigns set as Boolean)
		  Const ES_SUNKEN = &h4000
		  Const WS_EX_CLIENTEDGE = &h200
		  ChangeWindowStyle( e.Handle, ES_SUNKEN, set )
		  ChangeWindowStyle( e.Handle, WS_EX_CLIENTEDGE, set, true )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function TestWindowStyle(w as Integer, flag as Integer, ex as Boolean = false) As Boolean
		  #if TargetWin32
		    
		    Dim oldFlags as Integer
		    
		    Const GWL_STYLE = -16
		    Const GWL_EXSTYLE = -20
		    
		    Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (hwnd As Integer,  _
		    nIndex As Integer) As Integer
		    
		    dim style as Integer = GWL_STYLE
		    if ex then style = GWL_EXSTYLE
		    oldFlags = GetWindowLong( w, style )
		    
		    if Bitwise.BitAnd( oldFlags, flag ) = flag then
		      return true
		    else
		      return false
		    end
		    
		  #else
		    
		    #pragma unused w
		    #pragma unused flag
		    #pragma unused ex
		    
		  #endif
		End Function
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
