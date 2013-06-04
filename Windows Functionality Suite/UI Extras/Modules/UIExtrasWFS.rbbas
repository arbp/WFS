#tag Module
Protected Module UIExtrasWFS
	#tag Method, Flags = &h1
		Protected Function CountWindowsWithPartialTitle(partial as String) As Integer
		  #if TargetWin32
		    // Return number of matching windows
		    Soft Declare Function FindWindowA Lib "user32.dll" ( lpClassName As integer, lpWindowName As integer ) as integer
		    Soft Declare Function FindWindowW Lib "user32.dll" ( lpClassName As integer, lpWindowName As integer ) as integer
		    Declare Function GetWindow Lib "user32" ( hWnd As integer, wCmd As integer ) As integer
		    Soft Declare Function GetWindowTextA Lib "user32" ( hWnd As integer, lpString As ptr, cch As integer ) As integer
		    Soft Declare Function GetWindowTextW Lib "user32" ( hWnd As integer, lpString As ptr, cch As integer ) As integer
		    
		    Const GW_HWNDNEXT = 2
		    
		    dim ret,wincount as integer
		    dim mb as new MemoryBlock( 255 )
		    dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "FindWindowW", "User32" )
		    
		    if unicodeSavvy then
		      ret = FindWindowW( 0, 0 )
		    else
		      ret = FindWindowA( 0, 0 )
		    end if
		    
		    while ret > 0
		      if unicodeSavvy then
		        if GetWindowTextW( ret, mb, mb.size ) > 0 then
		          if InStr( mb.WString( 0 ), partial ) > 0 then wincount = wincount + 1
		        end if
		      else
		        if GetWindowTextA( ret, mb, mb.size ) > 0 then
		          if InStr( mb.CString( 0 ), partial ) > 0 then wincount = wincount + 1
		        end if
		      end if
		      
		      ret = GetWindow( ret, GW_HWNDNEXT )
		    wend
		    
		    return wincount
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FindWindowHandleFromPartialTitle(partial as String) As Integer
		  #if TargetWin32
		    // Return number of matching windows
		    Soft Declare Function FindWindowA Lib "user32.dll" ( lpClassName As integer, lpWindowName As integer ) as integer
		    Soft Declare Function FindWindowW Lib "user32.dll" ( lpClassName As integer, lpWindowName As integer ) as integer
		    Declare Function GetWindow Lib "user32" ( hWnd As integer, wCmd As integer ) As integer
		    Soft Declare Function GetWindowTextA Lib "user32" ( hWnd As integer, lpString As ptr, cch As integer ) As integer
		    Soft Declare Function GetWindowTextW Lib "user32" ( hWnd As integer, lpString As ptr, cch As integer ) As integer
		    
		    Const GW_HWNDNEXT = 2
		    
		    dim ret,wincount as integer
		    dim mb as new MemoryBlock( 255 )
		    dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "FindWindowW", "User32" )
		    
		    if unicodeSavvy then
		      ret = FindWindowW( 0, 0 )
		    else
		      ret = FindWindowA( 0, 0 )
		    end if
		    
		    while ret > 0
		      if unicodeSavvy then
		        if GetWindowTextW( ret, mb, mb.size ) > 0 then
		          if InStr( mb.WString( 0 ), partial ) > 0 then return ret
		        end if
		      else
		        if GetWindowTextA( ret, mb, mb.size ) > 0 then
		          if InStr( mb.CString( 0 ), partial ) > 0 then return ret
		        end if
		      end if
		      
		      ret = GetWindow( ret, GW_HWNDNEXT )
		    wend
		    
		    return 0
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub TemporarilyInstallFont(fontFile as FolderItem, privateFont as Boolean = true)
		  #if TargetWin32
		    Soft Declare Sub AddFontResourceExW Lib "Gdi32" ( filename as WString, flags as Integer, reserved as Integer )
		    Soft Declare Sub AddFontResourceA Lib "Gdi32" ( filename as CString )
		    Soft Declare Sub AddFontResourceW Lib "Gdi32" ( filename as WString )
		    
		    Const FR_PRIVATE = &h10
		    
		    if privateFont and System.IsFunctionAvailable( "AddFontResourceExW", "Gdi32" ) then
		      // If the user wants to install it as a private font, then we need to
		      // use the Ex APIs.  Otherwise, use the regular APIs.  We know
		      // that AddFontResourceEx is available in Win2k and up, so if
		      // the private flag is specified, we have to check to make sure
		      // we can load the API as well.  We won't bother with the A
		      // version of the call since we know the W version will be there.
		      AddFontResourceExW( fontFile.AbsolutePath, FR_PRIVATE, 0 )
		    else
		      // The user wants to install it as a public font, or they are running
		      // on an OS without the ability to make private fonts
		      if System.IsFunctionAvailable( "AddFontResourceW", "Gdi32" ) then
		        AddFontResourceW( fontFile.AbsolutePath )
		      else
		        AddFontResourceA( fontFile.AbsolutePath )
		      end if
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub UninstallTemporaryFont(fontFile as FolderItem)
		  #if TargetWin32
		    Soft Declare Sub RemoveFontResourceExW Lib "Gdi32" ( filename as WString, flags as Integer, reserved as Integer )
		    Soft Declare Sub RemoveFontResourceA Lib "Gdi32" ( filename as CString )
		    Soft Declare Sub RemoveFontResourceW Lib "Gdi32" ( filename as WString )
		    
		    Const FR_PRIVATE = &h10
		    
		    if System.IsFunctionAvailable( "RemoveFontResourceExW", "Gdi32" ) then
		      RemoveFontResourceExW( fontFile.AbsolutePath, FR_PRIVATE, 0 )
		    end if
		    
		    if System.IsFunctionAvailable( "RemoveFontResourceW", "Gdi32" ) then
		      RemoveFontResourceW( fontFile.AbsolutePath )
		    else
		      RemoveFontResourceA( fontFile.AbsolutePath )
		    end if
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
