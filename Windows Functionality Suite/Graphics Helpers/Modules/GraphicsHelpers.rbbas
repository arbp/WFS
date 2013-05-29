#tag Module
Protected Module GraphicsHelpers
	#tag Method, Flags = &h1
		Protected Function CanBeDisplayed(g as Graphics, text as String) As Boolean
		  // Note, this function is *supposed* to fire a FunctionNotFoundException
		  // in the case that we cannot load the GetGlyphIndices API.  The API is
		  // only supported on Windows 2000 and up.  We could return false if we can't
		  // load the function (or true), but that's incorrect -- the string could be displayed
		  // or not; we simply don't know.  So the onus is on the caller to fall back however
		  // they see fit on older systems.
		  
		  // We're going to ensure that the font is properly setup on the graphics
		  // object by calling StringWidth here.  This selects the font into the object
		  // and ensures code like this works:
		  //     g.TextFont = "Marlett"
		  //     CanBeDisplayed( g, "Testing" )
		  call g.StringWidth( text )
		  
		  Soft Declare Function GetGlyphIndicesW Lib "Gdi32" ( hdc as Integer, lpstr as WString, cch as Integer, retData as Ptr, flags as Integer ) as Integer
		  
		  // Make a MemoryBlock large enough to hold all of the glyph information
		  Dim glyphs as new MemoryBlock( Len( text ) * 2 )
		  Const GGI_MARK_NONEXISTING_GLYPHS = 1
		  Dim numConverted as Integer
		  
		  // Convert the text (automatically into UTF16) into glyph information,
		  // while marking all of the non-available glyphs.
		  numConverted = GetGlyphIndicesW( g.Handle( Graphics.HandleTypeHDC ), _
		  text, Len( text ), glyphs, GGI_MARK_NONEXISTING_GLYPHS )
		  
		  // Now loop over all of the glyphs, and see if we can find any that
		  // aren't displayable.
		  for i as Integer = 0 to numConverted - 1
		    if glyphs.UInt16Value( i * 2 ) = &hFFFF then
		      // We found one, so this string cannot be displayed properly
		      return false
		    end if
		  next i
		  
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function CaptureScreen(withCursor as Boolean = false) As Picture
		  #if TargetWin32
		    // DISCLAIMER!!!  READ THIS!!!!
		    // The following declare uses a very unsupported feature in REALbasic.  This means a few
		    // things.  1) Don't rely on this call working forever -- it's entirely possible that upgrading to
		    // a new version of REALbasic will cause this declare to break.  2) Don't try to use declares
		    // into the plugins APIs yourself.  And 3) Since this is not a supported way to use declares, you
		    // get everything you deserve if you use a declare like this.  You will eventually come down
		    // with the plague and someone will drink all your beer.  Basically... don't use this sort of thing!
		    // The declare will be removed as soon as there is a sanctioned way to get this functionality
		    // from REALbasic.  This declare should (hypothetically) work in 5.5 only (tho it may work in
		    // version 5 as well).  Consider all other versions of REALbasic unsupported.
		    #if RBVersion < 2006.01 then
		      Declare Function REALGraphicsDC Lib "" ( gfx as Graphics ) as Integer
		    #endif
		    Declare Sub BitBlt Lib "GDI32" ( dest as Integer, x as Integer, y as Integer, width as Integer, height as Integer, _
		    src as Integer, srcX as Integer, srcY as Integer, rops as Integer )
		    Declare Function SelectObject Lib "GDI32" ( hdc as Integer, hIcon as Integer ) as Integer
		    Declare Sub DeleteDC Lib "GDI32" ( hdc as Integer )
		    Declare Function GetDC Lib "User32" ( hwnd as Integer ) as Integer
		    Declare Function CreateCompatibleBitmap Lib "Gdi32" ( hdc as Integer, width as Integer, height as Integer ) as Integer
		    Declare Sub GetObjectA Lib "GDI32" ( hBitmap as Integer, size as Integer, struct as Ptr )
		    Declare Function GetCursorInfo Lib "user32.dll" (lpCursor as ptr) As Integer
		    Declare Function DrawIcon Lib "user32" Alias "DrawIcon" (hdc As integer, x As integer, y As integer, hIcon As integer) As integer
		    Declare Function DrawIconEx Lib "user32" Alias "DrawIconEx" (hdc As integer, xLeft As integer, yTop As integer, hIcon As integer, cxWidth As integer, cyWidth As integer, istepIfAniCur As integer, hbrFlickerFreeDraw As integer, diFlags As integer) As integer
		    Declare Sub DeleteObject Lib "Gdi32" ( obj as Integer )
		    
		    Const DI_MASK = &H1
		    
		    // We want to get the screen's DC first
		    dim screenHDC as Integer
		    screenHDC = GetDC( 0 )
		    
		    Dim hBitmap, width, height as Integer
		    width = SystemMetrics.VirtualScreenWidth
		    height = SystemMetrics.VirtualScreenHeight
		    hBitmap = CreateCompatibleBitmap( screenHDC, width, height )
		    
		    if hBitmap = 0 then return nil
		    
		    // Get the bits per pixel of the picture
		    dim bps as Integer
		    dim bitmapInfo as new MemoryBlock( 24 )
		    GetObjectA( hBitmap, 24, bitmapInfo )
		    bps = bitmapInfo.Short( 18 )
		    
		    DeleteObject( hBitmap )
		    
		    ' Now we want to move this handle into a picture object
		    dim ret as Picture
		    ret = NewPicture( width, height, bps )
		    
		    if ret = nil then return nil
		    
		    // Select the bitmap into the picture's HDC
		    dim desthdc, srchdc as Integer
		    #if RBVersion < 2006.01 then
		      desthdc = REALGraphicsDC( ret.Graphics )
		    #else
		      desthdc = ret.Graphics.Handle( Graphics.HandleTypeHDC )
		    #endif
		    
		    // Copy the bitmap data
		    Const CAPTUREBLT = &h40000000
		    Const SRCCOPY = &hCC0020
		    BitBlt( desthdc, 0, 0, width, height, screenHDC, 0, 0, SRCCOPY + CAPTUREBLT )
		    
		    if withCursor then
		      Dim mbCursor as new MemoryBlock( 20 )
		      mbCursor.Long( 0 ) = mbCursor.Size
		      
		      dim cursRet as Integer = GetCursorInfo( mbCursor )
		      if cursRet = 0 then return ret
		      
		      dim cursPict as new Picture( 32, 32, 32 )
		      
		      // Draw the cursor
		      #if RBVersion < 2006.01 then
		        dim desthdc2 as Integer = REALGraphicsDC( cursPict.Graphics )
		      #else
		        dim desthdc2 as Integer = cursPict.Graphics.Handle( Graphics.HandleTypeHDC )
		      #endif
		      
		      cursRet = DrawIcon( desthdc2, 0, 0, mbCursor.Long( 8 ) )
		      if cursRet = 0 then return ret
		      
		      // Draw the cursor's mask
		      #if RBVersion < 2006.01 then
		        dim desthdc3 as Integer = REALGraphicsDC( cursPict.Mask.Graphics )
		      #else
		        dim desthdc3 as Integer = cursPict.Mask.Graphics.Handle( Graphics.HandleTypeHDC )
		      #endif
		      cursRet = DrawIconEx( desthdc3, 0, 0, mbCursor.Long( 8 ), 0, 0, 0, 0, DI_MASK )
		      if cursRet = 0 then return ret
		      
		      // Now let's make sure we draw the cursor at the proper location
		      ret.Graphics.DrawPicture( cursPict, mbCursor.Long( 12 ), mbCursor.Long( 16 ), 32, 32, 0, 0, 32, 32 )
		    end if
		    
		    return ret
		  #endif
		  return nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ChooseFont(owner as Window = nil) As StyleRun
		  #if TargetWin32
		    Soft Declare Function ChooseFontA Lib "ComDlg32" ( lpcf as Ptr ) as Boolean
		    Soft Declare Function ChooseFontW Lib "ComDlg32" ( lpcf as Ptr ) as Boolean
		    Declare Function GetDC Lib "User32" ( hwnd as Integer ) as Integer
		    Declare Sub ReleaseDC Lib "User32" ( dc as Integer )
		    
		    // The size of the CHOOSEFONT structure is the
		    // same whether it's the W version or the A version.  We
		    // just need to treat the pointers differently
		    dim mb as new MemoryBlock( 15 * 4 )
		    dim ret as new StyleRun
		    
		    mb.Long( 0 ) = mb.Size
		    
		    if owner <> nil then
		      mb.Long( 4 ) = owner.WinHWND
		      mb.Long( 8 ) = GetDC( owner.WinHWND )
		    else
		      mb.Long( 8 ) = GetDC( 0 )
		    end
		    
		    dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "ChooseFontW", "ComDlg32" )
		    dim temp as new LogicalFont
		    dim theLogFont as MemoryBlock = temp.ToMemoryBlock( unicodeSavvy )
		    mb.Ptr( 12 ) = theLogFont
		    
		    mb.Long( 20 ) = &h1 + &h100
		    
		    dim success as Boolean
		    if unicodeSavvy then
		      success = ChooseFontW( mb )
		    else
		      success = ChooseFontA( mb )
		    end if
		    
		    if success then
		      if unicodeSavvy then
		        ret.Font = theLogFont.WString( 28 )
		      else
		        ret.Font = theLogFont.CString( 28 )
		      end if
		      ret.Size = mb.Long( 16 ) / 10
		      ret.Bold = (theLogFont.Long( 16 ) > 400)
		      ret.Italic = theLogFont.BooleanValue( 20 )
		      ret.Underline = theLogFont.BooleanValue( 21 )
		      ret.TextColor = mb.ColorValue( 24, 32 )
		      
		      ' Be sure to release the device context
		      ReleaseDC( mb.Long( 8 ) )
		      return ret
		    end
		    
		    ' Be sure to release the device context
		    ReleaseDC( mb.Long( 8 ) )
		    return nil
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ChooseLogicalFont(owner as Window = nil) As LogicalFont
		  #if TargetWin32
		    Soft Declare Function ChooseFontA Lib "ComDlg32" ( lpcf as Ptr ) as Boolean
		    Soft Declare Function ChooseFontW Lib "ComDlg32" ( lpcf as Ptr ) as Boolean
		    Declare Function GetDC Lib "User32" ( hwnd as Integer ) as Integer
		    Declare Sub ReleaseDC Lib "User32" ( dc as Integer )
		    
		    // The size of the CHOOSEFONT structure is the
		    // same whether it's the W version or the A version.  We
		    // just need to treat the pointers differently
		    dim mb as new MemoryBlock( 15 * 4 )
		    dim ret as new LogicalFont
		    
		    mb.Long( 0 ) = mb.Size
		    
		    if owner <> nil then
		      mb.Long( 4 ) = owner.WinHWND
		      mb.Long( 8 ) = GetDC( owner.WinHWND )
		    else
		      mb.Long( 8 ) = GetDC( 0 )
		    end
		    
		    dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "ChooseFontW", "ComDlg32" )
		    dim theLogFont as MemoryBlock = ret.ToMemoryBlock( unicodeSavvy )
		    mb.Ptr( 12 ) = theLogFont
		    
		    mb.Long( 20 ) = &h1 + &h100
		    
		    dim success as Boolean
		    if unicodeSavvy then
		      success = ChooseFontW( mb )
		    else
		      success = ChooseFontA( mb )
		    end if
		    
		    if success then
		      ret = new LogicalFont( theLogFont, unicodeSavvy )
		      
		      ' Be sure to release the device context
		      ReleaseDC( mb.Long( 8 ) )
		      return ret
		    end
		    
		    ' Be sure to release the device context
		    ReleaseDC( mb.Long( 8 ) )
		    return nil
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub DialogUnitsToPixels(g as Graphics, ByRef top as Integer, ByRef left as Integer, ByRef width as Integer, ByRef height as Integer)
		  #if TargetWin32
		    // Get the HDC for the window's Graphics context
		    dim handle as Integer = g.Handle( Graphics.HandleTypeHDC )
		    
		    // Now we can figure out the default width and height
		    Declare Sub GetTextExtentPoint32A Lib "GDI32" ( hdc as Integer, _
		    str as CString, num as Integer, size as Ptr )
		    Declare Sub GetTextMetricsA Lib "GDI32" ( hdc as Integer, metrics as Ptr )
		    Declare Function GetStockObject Lib "GDI32" ( index as Integer ) as Integer
		    Declare Function SelectObject Lib "GDI32" ( hdc as Integer, obj as Integer ) as Integer
		    
		    // We need to do our calculations based on the
		    // default GUI font, not the system font.  This is
		    // because the system font is that old fixed-width
		    // terrible thing from Windows 1.0
		    Const DEFAULT_GUI_FONT = 17
		    dim oldFont as Integer = SelectObject( handle, GetStockObject( DEFAULT_GUI_FONT ) )
		    
		    // Get the text extent and metrics
		    dim size as new MemoryBlock( 8 )
		    dim metrics as new MemoryBlock( 56 )
		    dim str as String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
		    GetTextExtentPoint32A( handle, str, Len( str ), size )
		    GetTextMetricsA( handle, metrics )
		    
		    // Restore the old font
		    call SelectObject( handle, oldFont )
		    
		    // Now we know the height and the width.  But we need
		    // the width of the average character.
		    dim baseUnitX, baseUnitY as Integer
		    baseUnitX = (size.Long( 0 ) / 26 + 1) / 2
		    baseUnitY = metrics.Long( 0 )
		    
		    // Now, the forumal for DLU conversion is:
		    // pixelX = (dialogunitX * baseunitX) / 4
		    // pixelY = (dialogunitY * baseunitY) / 8
		    
		    // Do the conversion
		    top = (top * baseUnitY) / 8
		    left = (left * baseUnitX) / 4
		    width = (width * baseUnitX) / 4
		    height = (height * baseUnitY) / 8
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetStockIcon(id as Integer, smallIcon as Boolean = false, selected as Boolean = false, linkOverlay as Boolean = false) As Picture
		  #if TargetWin32
		    Soft Declare Sub SHGetStockIconInfo Lib "Shell32" ( id as Integer, flags as Integer, info as Ptr )
		    Declare Sub DestroyIcon Lib "User32" (hIcon as Integer )
		    
		    if System.IsFunctionAvailable( "SHGetStockIconInfo", "Shell32" ) then
		      dim info as new MemoryBlock( 536 )
		      info.Long( 0 ) = info.Size
		      
		      dim flags as Integer = &h100
		      if smallIcon then flags = flags + &h1
		      if linkOverlay then flags = flags + &h8000
		      if selected then flags = flags + &h10000
		      
		      SHGetStockIconInfo( id, flags, info )
		      
		      dim ret as Picture = IconHandleToPicture( info.Long( 4 ) )
		      
		      DestroyIcon( info.Long( 4 ) )
		      
		      return ret
		    end if
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetTextCharset(g as Graphics) As Integer
		  Declare Function GetTextCharset Lib "Gdi32" ( hdc as Integer ) as Integer
		  
		  // Select the font into the Graphics object by making a
		  // call to StringWidth.  This way, you can call this
		  // function directly after setting the text font.
		  call g.StringWidth( "X" )
		  
		  return GetTextCharset( g.Handle( Graphics.HandleTypeHDC ) )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IconHandleToPicture(iconHandle as Integer) As Picture
		  #if TargetWin32
		    // DISCLAIMER!!!  READ THIS!!!!
		    // The following declare uses a very unsupported feature in REALbasic.  This means a few
		    // things.  1) Don't rely on this call working forever -- it's entirely possible that upgrading to
		    // a new version of REALbasic will cause this declare to break.  2) Don't try to use declares
		    // into the plugins APIs yourself.  And 3) Since this is not a supported way to use declares, you
		    // get everything you deserve if you use a declare like this.  You will eventually come down
		    // with the plague and someone will drink all your beer.  Basically... don't use this sort of thing!
		    // The declare will be removed as soon as there is a sanctioned way to get this functionality
		    // from REALbasic.  This declare should (hypothetically) work in 5.5 only (tho it may work in
		    // version 5 as well).  Consider all other versions of REALbasic unsupported.
		    #if RBVersion < 2006.01 then
		      Declare Function REALGraphicsDC Lib "" ( gfx as Graphics ) as Integer
		    #endif
		    Declare Sub GetIconInfo Lib "User32" ( hIcon as Integer, iconInfo as Ptr )
		    Declare Sub GetObjectA Lib "GDI32" ( hBitmap as Integer, size as Integer, struct as Ptr )
		    Declare Sub BitBlt Lib "GDI32" ( dest as Integer, x as Integer, y as Integer, width as Integer, height as Integer, _
		    src as Integer, srcX as Integer, srcY as Integer, rops as Integer )
		    Declare Function CreateCompatibleDC Lib "GDI32" ( hdc as Integer ) as Integer
		    Declare Function SelectObject Lib "GDI32" ( hdc as Integer, hIcon as Integer ) as Integer
		    Declare Sub DeleteDC Lib "GDI32" ( hdc as Integer )
		    
		    if iconHandle = 0 then return nil
		    
		    ' Let's get some information about the icon
		    dim iconInfo as new MemoryBlock( 20 )
		    GetIconInfo( iconHandle, iconInfo )
		    
		    dim cx, cy, bps as Integer
		    dim bitmapInfo as new MemoryBlock( 24 )
		    GetObjectA( iconInfo.Long( 16 ), 24, bitmapInfo )
		    
		    cx = bitmapInfo.Long( 4 )
		    cy = bitmapInfo.Long( 8 )
		    bps = bitmapInfo.Short( 18 )
		    
		    ' Now we want to move this handle into a picture object
		    dim ret as Picture
		    ret = NewPicture( cx, cy, bps )
		    
		    dim desthdc, srchdc as Integer
		    #if RBVersion < 2006.01 then
		      desthdc = REALGraphicsDC( ret.Graphics )
		    #else
		      desthdc = ret.Graphics.Handle( Graphics.HandleTypeHDC )
		    #endif
		    srchdc = CreateCompatibleDC( desthdc )
		    
		    Call SelectObject( srchdc, iconInfo.Long( 16 ) )
		    
		    BitBlt( desthdc, 0, 0, cx, cy, srchdc, 0, 0, &hCC0020 )
		    
		    dim desthdc2, srchdc2 as Integer
		    #if RBVersion < 2006.01 then
		      desthdc2 = REALGraphicsDC( ret.Mask.Graphics )
		    #else
		      desthdc2 = ret.Mask.Graphics.Handle( Graphics.HandleTypeHDC )
		    #endif
		    
		    srchdc2 = CreateCompatibleDC( desthdc2 )
		    
		    Call SelectObject( srchdc2, iconInfo.Long( 12 ) )
		    
		    BitBlt( desthdc2, 0, 0, cx, cy, srchdc2, 0, 0, &hCC0020 )
		    
		    DeleteDC( srchdc )
		    DeleteDC( srchdc2 )
		    
		    return ret
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsSymbolFont(g as Graphics) As Boolean
		  return GetTextCharset( g ) = kSymbolFont
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function LoadIcon(appInstance as FolderItem, iconNum as Integer = - 1) As Picture
		  #if TargetWin32
		    Soft Declare Function LoadImageA Lib "User32" ( instance as Integer, iconName as CString, type as Integer, _
		    cxDesired as Integer, cyDesired as Integer, loadFlags as Integer ) as Integer
		    Soft Declare Function LoadImageW Lib "User32" ( instance as Integer, iconName as WString, type as Integer, _
		    cxDesired as Integer, cyDesired as Integer, loadFlags as Integer ) as Integer
		    Soft Declare Function LoadImageA Lib "User32" ( instance as Integer, iconName as Integer, type as Integer, _
		    cxDesired as Integer, cyDesired as Integer, loadFlags as Integer ) as Integer
		    Soft Declare Function GetModuleHandleA Lib "Kernel32" ( absPath as CString ) as Integer
		    Soft Declare Function GetModuleHandleW Lib "Kernel32" ( absPath as WString ) as Integer
		    Declare Sub DestroyIcon Lib "User32" (hIcon as Integer )
		    Soft Declare Function LoadLibraryA Lib "Kernel32" (name as CString ) as Integer
		    Soft Declare Function LoadLibraryW Lib "Kernel32" (name as WString ) as Integer
		    Declare Function SetErrorMode Lib "Kernel32" ( mode as Integer ) as Integer
		    
		    // Turn off the error dialog that LoadLibrary will generate when it's
		    // passed a file that is not a library.
		    Const SEM_FAILCRITICALERRORS = &h1
		    dim oldMode as Integer = SetErrorMode( SEM_FAILCRITICALERRORS )
		    
		    dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "LoadLibraryW", "Kernel32" )
		    
		    dim moduleHandle as Integer = 0
		    if appInstance <> nil then
		      if unicodeSavvy then
		        moduleHandle = GetModuleHandleW( appInstance.AbsolutePath )
		      else
		        moduleHandle = GetModuleHandleA( appInstance.AbsolutePath )
		      end if
		      
		      if moduleHandle = 0 then
		        if unicodeSavvy then
		          moduleHandle = LoadLibraryW( appInstance.AbsolutePath )
		        else
		          moduleHandle = LoadLibraryA( appInstance.AbsolutePath )
		        end if
		      end
		    end
		    
		    // Set the error mode back so that the application behaves as normal
		    Call SetErrorMode( oldMode )
		    
		    dim loadFlags as Integer
		    
		    loadFlags = &h40  ' we want the default sizes
		    
		    dim iconHandle as Integer
		    if iconNum = -1 and appInstance <> nil then
		      ' The user wants us to just load it as a file (usually a .ico file)
		      loadFlags = Bitwise.BitOr( loadFlags, &h10 )
		      if unicodeSavvy then
		        iconHandle = LoadImageW( 0, appInstance.AbsolutePath, 1, 0, 0, loadFlags )
		      else
		        iconHandle = LoadImageA( 0, appInstance.AbsolutePath, 1, 0, 0, loadFlags )
		      end if
		    else
		      ' The user wants us to load the file from a resource
		      iconHandle = LoadImageA( moduleHandle, iconNum, 1, 0, 0, loadFlags )
		    end
		    
		    dim ret as Picture = IconHandleToPicture( iconHandle )
		    
		    DestroyIcon( iconHandle )
		    
		    return ret
		  #endif
		End Function
	#tag EndMethod


	#tag Constant, Name = kSymbolFont, Type = Double, Dynamic = False, Default = \"2", Scope = Private
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
End Module
#tag EndModule
