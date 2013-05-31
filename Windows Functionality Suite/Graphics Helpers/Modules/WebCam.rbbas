#tag Module
Protected Module WebCam
	#tag Method, Flags = &h21
		Private Function capCreateCaptureWindow(lpszWindowName as string, dwStyle as Integer, x as Integer, y as Integer, nWidth as Integer, nHeight as Integer, hWndParent as Integer, nID as Integer) As integer
		  #if TargetWin32
		    Soft Declare Function capCreateCaptureWindowA Lib "avicap32" ( lpszWindowName as CString, dwStyle as Integer, x as Integer, y as Integer, nWidth as Integer, nHeight as Integer, hWndParent as Integer, nID as Integer) as Integer
		    Soft Declare Function capCreateCaptureWindowW Lib "avicap32" ( lpszWindowName as CString, dwStyle as Integer, x as Integer, y as Integer, nWidth as Integer, nHeight as Integer, hWndParent as Integer, nID as Integer) as Integer
		    
		    if System.IsFunctionAvailable( "capCreateCaptureWindowW", "avicap32" ) then
		      return capCreateCaptureWindowW( lpszWindowName, dwStyle, x, y, nWidth, nHeight, hWndParent, nID)
		    elseif System.IsFunctionAvailable( "capCreateCaptureWindowA", "avicap32" ) then
		      return capCreateCaptureWindowA( lpszWindowName, dwStyle, x, y, nWidth, nHeight, hWndParent, nID)
		    else
		      return -1
		    end if
		    
		  #Endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function capGetDriverDescription(wDriver as Integer, ByRef lpszName as string, ByRef lpszVer as string) As Boolean
		  #if TargetWin32
		    Soft Declare Function capGetDriverDescriptionW Lib "avicap32" (wDriver as Integer, lpszName as Ptr, cbName as Integer, lpszVer as Ptr, cbVer as Integer) as Boolean
		    Soft Declare Function capGetDriverDescriptionA Lib "avicap32" (wDriver as Integer, lpszName as Ptr, cbName as Integer, lpszVer as Ptr, cbVer as Integer) as Boolean
		    
		    dim name as new MemoryBlock( 2048 )
		    dim vers as new MemoryBlock( 2048 )
		    dim ret as Boolean
		    
		    if System.IsFunctionAvailable( "capGetDriverDescriptionW", "avicap32" ) then
		      ret = capGetDriverDescriptionW( wDriver, name, name.Size, vers, vers.Size )
		      lpszName = name.WString( 0 )
		      lpszVer = vers.WString( 0 )
		    elseif System.IsFunctionAvailable( "capGetDriverDescriptionA", "avicap32" ) then
		      ret = capGetDriverDescriptionA( wDriver, name, name.Size, vers, vers.Size )
		      lpszName = name.CString( 0 )
		      lpszVer = vers.CString( 0 )
		    else
		      ret = false
		    end if
		    
		    return ret
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DestroyWindow()
		  #if TargetWin32
		    Const WM_CAP_DRIVER_DISCONNECT = 1035
		    Declare Sub DestroyWindow Lib "user32"(hndw  as Integer)
		    DestroyWindow( mWnd )
		  #Endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetDeviceList() As String()
		  #if TargetWin32 then
		    Dim strName  as String
		    Dim strVer  as String
		    Dim ret( -1 ) as String
		    Dim i as Integer = 0
		    
		    while capGetDriverDescription( i, strName, strVer )
		      ret.Append( strName )
		      
		      i = i + 1
		    wend
		    
		    return ret
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsPreviewing() As Boolean
		  return mWnd <> 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function SendMessage(wMsg as Integer, wParam as Integer, lparam as memoryblock) As integer
		  #if TargetWin32
		    Declare Function SendMessageA Lib "user32" ( hwnd as Integer, wMsg as Integer, wParam as Integer, lparam as ptr) as integer
		    Declare Function SendMessageW Lib "user32" ( hwnd as Integer, wMsg as Integer, wParam as Integer, lparam as ptr) as integer
		    
		    if System.IsFunctionAvailable( "SendMessageW", "User32" ) then
		      return SendMessageW( mWnd, wMsg, wParam, lparam)
		    else
		      return SendMessageA( mWnd, wMsg, wParam, lparam)
		    end if
		  #Endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function StartPreview(c as Canvas, deviceName as String) As Boolean
		  #if TargetWin32 then
		    Const WS_CHILD = &h40000000
		    Const WS_BORDER = &h00800000
		    Const WS_VISIBLE = &h10000000
		    Const WM_CAP = &H400
		    
		    Const WM_CAP_DRIVER_CONNECT = 1034
		    Const WM_CAP_DRIVER_DISCONNECT = 1035
		    Const WM_CAP_EDIT_COPY = 1054
		    Const WM_CAP_GRAB_FRAME = 1084
		    Const WM_CAP_SET_PREVIEW = 1074
		    Const WM_CAP_SET_PREVIEWRATE = 1076
		    Const WM_CAP_SET_SCALE = 1077
		    Const WM_CAP_DLG_VIDEOFORMAT = 1065
		    Const WM_CAP_DLG_VIDEOSOURCE = 1066
		    Const WM_CLOSE = &H10
		    
		    if mWnd <> 0 then return false
		    
		    mWnd = capCreateCaptureWindow( deviceName, WS_VISIBLE + WS_CHILD, 0, 0, c.Width, c.Height, c.Handle, 0 )
		    if mWnd = -1 then
		      mWnd = 0
		      return false
		    end if
		    
		    dim deviceIndex as Integer = GetDeviceList.IndexOf( deviceName )
		    if deviceIndex = -1 then return false
		    
		    dim p as MemoryBlock
		    p = new MemoryBlock(0)
		    If SendMessage( WM_CAP_DRIVER_CONNECT, deviceIndex, p ) > 0 Then
		      dim ret as Variant
		      ret = SendMessage( WM_CAP_SET_SCALE, 1, p )
		      ret = SendMessage( WM_CAP_SET_PREVIEWRATE, 20, p )
		      ret = SendMessage( WM_CAP_SET_PREVIEW, 1, p )
		      return true
		    Else
		      ' Error connecting to device close window
		      DestroyWindow
		      return false
		    End If
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub StopPreview()
		  if mWnd = 0 then return
		  
		  dim ret as Variant
		  dim p as MemoryBlock
		  Const WM_CAP_DRIVER_DISCONNECT = 1035
		  p = new MemoryBlock(0)
		  ret = SendMessage( WM_CAP_DRIVER_DISCONNECT, mDeviceIndex, p )
		  DestroyWindow
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mDeviceIndex As Integer
	#tag EndProperty

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
End Module
#tag EndModule
