#tag Module
Protected Module WndProcHelpersWFS
	#tag Method, Flags = &h21
		Private Function CallAllSubclases(hWnd as Integer, msg as Integer, wParam as Integer, lParam as Integer, ByRef handled as Boolean) As Integer
		  Dim ret as Integer
		  Dim subclass as WndProcSubclassWFS
		  Dim d as Dictionary
		  
		  handled = false
		  
		  for each d in mSubclasses
		    if d.HasKey( hWnd ) then
		      // We can call this subclass
		      subclass = d.Value( hWnd )
		      // call the subclass and check the return
		      // value
		      if subclass <> nil and subclass.WndProc( hWnd, msg, wParam, lParam, ret ) then
		        handled = true
		        return ret
		      end if
		    end
		  next
		  
		  // No one handled it, so we'll return 0 as well
		  return 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetWndProcToCall(hWnd as Integer) As Integer
		  // We are given an HWND, and we need to figure out
		  // which WndProc it belongs to, which is easy
		  if mWndProcs = nil then return 0
		  
		  return mWndProcs.Lookup( hWnd, 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Subclass(w as Integer, subber as WndProcSubclassWFS)
		  SubclassHelper( w, subber )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Subclass(w as MDIWindow, subber as WndProcSubclassWFS)
		  SubclassHelper( w.Handle, subber )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Subclass(w as Window, subber as WndProcSubclassWFS)
		  SubclassHelper( w.Handle, subber )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub SubclassHelper(w as Integer, subber as WndProcSubclassWFS)
		  #if TargetWin32
		    // First, if we don't have a list of WndProcs to care
		    // about, we need to make a new list
		    if mWndProcs = nil then mWndProcs = new Dictionary
		    
		    // Now, if we don't already have this window in
		    // our list of procs, we should add it
		    if mWndProcs.HasKey( w ) then
		      // We should also make sure that the subclasser is in
		      // our list so it can get notified as well.
		      dim d as new Dictionary
		      d.Value( w ) = subber
		      mSubclasses.Append( d )
		      return
		    end
		    
		    // We need to do the actual subclassing, which
		    // involves getting the old WndProc and replacing
		    // it with the new one
		    Soft Declare Function SetWindowLongA Lib "User32" ( hwnd as Integer, _
		    index as Integer, value as Ptr ) as Integer
		    Soft Declare Function SetWindowLongW Lib "User32" ( hwnd as Integer, _
		    index as Integer, value as Ptr ) as Integer
		    
		    // Do the actual subclassing
		    Dim oldWndProc as Integer
		    Const GWL_WNDPROC = -4
		    
		    if System.IsFunctionAvailable( "SetWindowLongW", "User32" ) then
		      oldWndProc = SetWindowLongW( w, GWL_WNDPROC, AddressOf WndProc )
		    else
		      oldWndProc = SetWindowLongA( w, GWL_WNDPROC, AddressOf WndProc )
		    end if
		    
		    // Now that we're subclassed, add this window to our list
		    mWndProcs.Value( w ) = oldWndProc
		    
		    // Finally, add this delegate to the list of people watching
		    // this window
		    dim d as new Dictionary
		    d.Value( w ) = subber
		    mSubclasses.Append( d )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Unsubclass(w as Integer, subber as WndProcSubclassWFS)
		  UnsubclassHelper( w, subber )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Unsubclass(w as MDIWindow, subber as WndProcSubclassWFS)
		  UnsubclassHelper( w.Handle, subber )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Unsubclass(w as Window, subber as WndProcSubclassWFS)
		  UnsubclassHelper( w.Handle, subber )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub UnsubclassHelper(w as Integer, subber as WndProcSubclassWFS)
		  #if TargetWin32
		    // First, if we don't have a list of WndProcs to care
		    // about, we are done
		    if mWndProcs = nil then return
		    
		    // Now, if we don't already have this window in
		    // our list of procs, we're also done
		    if not mWndProcs.HasKey( w ) then return
		    
		    // We need to do revert the subclassing, which
		    // involves getting the new WndProc and replacing
		    // it with the old one
		    Soft Declare Function SetWindowLongA Lib "User32" ( hwnd as Integer, _
		    index as Integer, value as Integer ) as Integer
		    Soft Declare Function SetWindowLongW Lib "User32" ( hwnd as Integer, _
		    index as Integer, value as Integer ) as Integer
		    
		    // Get the old WndProc
		    Dim oldWndProc as Integer
		    oldWndProc = mWndProcs.Value( w )
		    
		    // And actually revert the subclassing
		    Const GWL_WNDPROC = -4
		    if System.IsFunctionAvailable( "SetWindowLongW", "User32" ) then
		      Call SetWindowLongW( w, GWL_WNDPROC, oldWndProc )
		    else
		      Call SetWindowLongA( w, GWL_WNDPROC, oldWndProc )
		    end if
		    
		    // And now we just need to remove it from the list of
		    // subclassed windows
		    try
		      mWndProcs.Remove( w )
		    end
		    
		    // Remove this delegate from the list of observers
		    dim d as Dictionary
		    dim i, count as Integer
		    count = UBound( mSubclasses )
		    for i = count downto 0
		      d = mSubclasses( i )
		      if d.HasKey( w ) then
		        // We want to remove this one from the array
		        mSubclasses.Remove( i )
		      end
		    next
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function WndProc(hWnd as Integer, msg as Integer, wParam as Integer, lParam as Integer) As Integer
		  #if TargetWin32
		    Soft Declare Function CallWindowProcA Lib "User32" ( wndProc as Integer, hwnd as Integer, _
		    msg as Integer, wParam as Integer, lParam as Integer ) as Integer
		    Soft Declare Function CallWindowProcW Lib "User32" ( wndProc as Integer, hwnd as Integer, _
		    msg as Integer, wParam as Integer, lParam as Integer ) as Integer
		    
		    // Who to tell....who to tell...
		    dim ret as Integer
		    dim handled as Boolean
		    ret = CallAllSubclases( hWnd, msg, wParam, lParam, handled )
		    
		    // If a subclass handled this message, then
		    // we should bail out now
		    if handled then return ret
		    
		    // Get the next window proc to call and actually call it
		    Dim nextWndProc as Integer
		    nextWndProc = GetWndProcToCall( hWnd )
		    if nextWndProc <> 0 then
		      if System.IsFunctionAvailable( "CallWindowProcW", "User32" ) then
		        return CallWindowProcW( nextWndProc, hWnd, msg, wParam, lParam )
		      else
		        return CallWindowProcA( nextWndProc, hWnd, msg, wParam, lParam )
		      end if
		    end if
		  #endif
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mSubclasses(-1) As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mWndProcs As Dictionary
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
