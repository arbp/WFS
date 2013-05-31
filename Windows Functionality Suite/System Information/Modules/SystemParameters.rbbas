#tag Module
Protected Module SystemParameters
	#tag Method, Flags = &h1
		Protected Function AccessTimeout() As AccessTimeout
		  Const SPI_GETACCESSTIMEOUT = &h3C
		  dim mb as new MemoryBlock( 12 )
		  SystemParametersInfo( SPI_GETACCESSTIMEOUT, mb.Size, mb, 0 )
		  
		  return new AccessTimeout( mb )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AccessTimeout(assigns set as AccessTimeout)
		  Const SPI_SETACCESSTIMEOUT = &h3D
		  dim mb as MemoryBlock
		  mb = set.ToMemoryBlock
		  
		  SystemParametersInfo( SPI_SETACCESSTIMEOUT, mb.Size, mb, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ActiveWindowTracking() As Boolean
		  Const SPI_GETACTIVEWINDOWTRACKING = &h1000
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETACTIVEWINDOWTRACKING, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ActiveWindowTracking(assigns set as Boolean)
		  Const SPI_SETACTIVEWINDOWTRACKING = &h1001
		  SystemParametersInfo( SPI_SETACTIVEWINDOWTRACKING, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ActiveWindowTrackingTimeout() As Integer
		  Const SPI_GETACTIVEWNDTRKTIMEOUT = &h2002
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETACTIVEWNDTRKTIMEOUT, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ActiveWindowTrackingTimeout(assigns set as Integer)
		  Const SPI_SETACTIVEWNDTRKTIMEOUT = &h2003
		  SystemParametersInfo( SPI_SETACTIVEWNDTRKTIMEOUT, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Beep() As Boolean
		  Const SPI_GETBEEP = &h1
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETBEEP, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Beep(assigns set as Boolean)
		  Const SPI_SETBEEP = &h2
		  dim toSet as Integer
		  if set then toSet = 1
		  SystemParametersInfo( SPI_SETBEEP, toSet, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function BlockInputRequests() As Boolean
		  Const SPI_GETBLOCKSENDINPUTRESETS = &h1026
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETBLOCKSENDINPUTRESETS, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub BlockInputRequests(assigns set as Boolean)
		  Const SPI_SETBLOCKSENDINPUTRESETS = &h1027
		  dim toSet as Integer
		  if set then toSet = 1
		  SystemParametersInfo( SPI_SETBLOCKSENDINPUTRESETS, toSet, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function BringTrackedWindowFront() As Boolean
		  Const SPI_GETACTIVEWNDTRKZORDER = &h100C
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETACTIVEWNDTRKZORDER, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub BringTrackedWindowFront(assigns set as Boolean)
		  Const SPI_SETACTIVEWNDTRKZORDER = &h100D
		  SystemParametersInfo( SPI_SETACTIVEWNDTRKZORDER, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function CaretWidth() As Integer
		  Const SPI_GETCARETWIDTH = &h2006
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETCARETWIDTH, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub CaretWidth(assigns set as Integer)
		  Const SPI_SETCARETWIDTH = &h2007
		  SystemParametersInfo( SPI_SETCARETWIDTH, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ClientAreaAnimation() As Boolean
		  Const SPI_GETCLIENTAREAANIMATION = &h1042
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETCLIENTAREAANIMATION, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ClientAreaAnimation(assigns set as Boolean)
		  Const SPI_SETCLIENTAREAANIMATION = &h1043
		  SystemParametersInfo( SPI_SETCLIENTAREAANIMATION, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ComboBoxAnimation() As Boolean
		  Const SPI_GETCOMBOBOXANIMATION = &h1004
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETCOMBOBOXANIMATION, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ComboBoxAnimation(assigns set as Boolean)
		  Const SPI_SETCOMBOBOXANIMATION = &h1005
		  SystemParametersInfo( SPI_SETCOMBOBOXANIMATION, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function CursorShadow() As Boolean
		  Const SPI_GETCURSORSHADOW = &h101A
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETCURSORSHADOW, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub CursorShadow(assigns set as Boolean)
		  Const SPI_SETCURSORSHADOW = &h101B
		  SystemParametersInfo( SPI_SETCURSORSHADOW, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DesktopWallpaper() As String
		  Dim mb as new MemoryBlock( 1024 )
		  Const SPI_GETDESKWALLPAPER = &h73
		  
		  SystemParametersInfo( SPI_GETDESKWALLPAPER, mb.Size, mb )
		  
		  if System.IsFunctionAvailable( "SystemParametersInfoW", "User32" ) then
		    return mb.WString( 0 )
		  else
		    return mb.CString( 0 )
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub DesktopWallpaper(assigns newWallpaper as FolderItem)
		  Const SPI_SETDESKWALLPAPER = &h14
		  Dim mb as new MemoryBlock( 2048 )
		  
		  if newWallpaper = nil then
		    // We want to remove the old wallpaper and bail out
		    SystemParametersInfo( SPI_SETDESKWALLPAPER, 0, mb )
		    return
		  end
		  
		  dim wallpaper as String = newWallpaper.AbsolutePath
		  dim wallpaperPicture as Picture
		  dim saveSpot as FolderItem
		  
		  if right( wallpaper, 4 ) <> ".bmp" then
		    // In order to set the wallpaper properly, it must be
		    // in bitmap form.  How brilliant.  This isn't, so we'll
		    // try to open the file and convert it to a bitmap
		    'wallpaperPicture = newWallpaper.OpenAsPicture
		    wallpaperPicture = Picture.Open( newWallpaper )
		    
		    if wallpaperPicture = nil then
		      // We couldn't open the picture, so just bail out
		      return
		    end
		    
		    // Now that we have the picture, save it as a BMP
		    // in a good place
		    saveSpot = SpecialFolder.Temporary.Child( "Wallpaper1.bmp" )
		    if saveSpot = nil then return
		    
		    // Save to the temp folder
		    'saveSpot.SaveAsPicture( wallpaperPicture, FolderItem.SaveAsWindowsBMP )
		    wallpaperPicture.Save( saveSpot, Picture.SaveAsWindowsBMP )
		    
		    // And be sure to get the new absolute path
		    wallpaper = saveSpot.AbsolutePath
		  end
		  
		  // Set the memory block to point to the new wallpaper
		  if System.IsFunctionAvailable( "SystemParametersInfoW", "User32" ) then
		    mb.WString( 0 ) = wallpaper
		  else
		    mb.CString( 0 ) = wallpaper
		  end if
		  
		  // And set the wallpaper to the new bitmap
		  SystemParametersInfo( SPI_SETDESKWALLPAPER, mb.Size, mb )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DisableOverlappedContent() As Boolean
		  Const SPI_GETDISABLEOVERLAPPEDCONTENT = &h1040
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETDISABLEOVERLAPPEDCONTENT, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub DisableOverlappedContent(assigns set as Boolean)
		  Const SPI_SETDISABLEOVERLAPPEDCONTENT = &h1041
		  SystemParametersInfo( SPI_SETDISABLEOVERLAPPEDCONTENT, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DoubleClickHeight() As Integer
		  Const SM_CYDOUBLECLK = 37
		  return GetSystemMetrics( SM_CYDOUBLECLK )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub DoubleClickHeight(assigns set as Integer)
		  Const SPI_SETDOUBLECLKHEIGHT = &h1E
		  SystemParametersInfo( SPI_SETDOUBLECLKHEIGHT, set, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DoubleClickTime() As Integer
		  #if TargetWin32
		    Declare Function GetDoubleClickTime Lib "User32" () as Integer
		    
		    return GetDoubleClickTime
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub DoubleClickTime(assigns set as Integer)
		  Const SPI_SETDOUBLECLICKTIME = &h20
		  SystemParametersInfo( SPI_SETDOUBLECLICKTIME, set, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DoubleClickWidth() As Integer
		  Const SM_CXDOUBLECLK = 36
		  return GetSystemMetrics( SM_CXDOUBLECLK )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub DoubleClickWidth(assigns set as Integer)
		  Const SPI_SETDOUBLECLKWIDTH = &h1D
		  SystemParametersInfo( SPI_SETDOUBLECLKWIDTH, set, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DragFullWindow() As Boolean
		  Const SPI_GETDRAGFULLWINDOWS = &h26
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETDRAGFULLWINDOWS, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub DragFullWindow(assigns set as Boolean)
		  Const SPI_SETDRAGFULLWINDOWS = &h25
		  dim toSet as Integer
		  if set then toSet = 1
		  
		  SystemParametersInfo( SPI_SETDRAGFULLWINDOWS, toSet, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DragHeight() As Integer
		  Const SM_CYDRAG = 69
		  return GetSystemMetrics( SM_CYDRAG )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub DragHeight(assigns set as Integer)
		  Const SPI_SETDRAGHEIGHT = &h4D
		  SystemParametersInfo( SPI_SETDRAGHEIGHT, set, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DragWidth() As Integer
		  Const SM_CXDRAG = 68
		  return GetSystemMetrics( SM_CXDRAG )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub DragWidth(assigns set as Integer)
		  Const SPI_SETDRAGWIDTH = &h4C
		  SystemParametersInfo( SPI_SETDRAGWIDTH, set, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DropShadow() As Boolean
		  Const SPI_GETDROPSHADOW = &h1024
		  
		  Dim mb as new MemoryBlock( 1 )
		  SystemParametersInfo( SPI_GETDROPSHADOW, 0, mb )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub DropShadow(assigns set as Boolean)
		  Const SPI_SETDROPSHADOW = &h1025
		  SystemParametersInfo( SPI_SETDROPSHADOW, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE  )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FilterKeys() As FilterKeys
		  Const SPI_GETFILTERKEYS = &h32
		  dim mb as new MemoryBlock( 24 )
		  SystemParametersInfo( SPI_GETFILTERKEYS, mb.Size, mb, 0 )
		  
		  return new FilterKeys( mb )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub FilterKeys(assigns set as FilterKeys)
		  Const SPI_SETFILTERKEYS = &h33
		  dim mb as MemoryBlock
		  mb = set.ToMemoryBlock
		  
		  SystemParametersInfo( SPI_SETFILTERKEYS, mb.Size, mb, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FlatMenus() As Boolean
		  Const SPI_GETFLATMENU = &h1022
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETFLATMENU, 0, mb )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub FlatMenus(assigns set as Boolean)
		  Const SPI_SETFLATMENU = &h1023
		  SystemParametersInfo( SPI_SETFLATMENU, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FocusBorderHeight() As Integer
		  Const SPI_GETFOCUSBORDERHEIGHT = &h2010
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETFOCUSBORDERHEIGHT, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub FocusBorderHeight(assigns set as Integer)
		  Const SPI_SETFOCUSBORDERHEIGHT = &h2011
		  SystemParametersInfo( SPI_SETFOCUSBORDERHEIGHT, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FocusBorderWidth() As Integer
		  Const SPI_GETFOCUSBORDERWIDTH = &h200E
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETFOCUSBORDERWIDTH, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub FocusBorderWidth(assigns set as Integer)
		  Const SPI_SETFOCUSBORDERWIDTH = &h200F
		  SystemParametersInfo( SPI_SETFOCUSBORDERWIDTH, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FontSmoothing() As Boolean
		  Const SPI_GETFONTSMOOTHING = &h4A
		  dim mb as new MemoryBlock( 1 )
		  SystemParametersInfo( SPI_GETFONTSMOOTHING, 0, mb )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub FontSmoothing(assigns set as Boolean)
		  Const SPI_SETFONTSMOOTHING = &h4B
		  
		  dim theSet as Integer = 0
		  if set then theSet = 1
		  SystemParametersInfo( SPI_SETFONTSMOOTHING, theSet, nil )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FontSmoothingContrast() As Integer
		  Const SPI_GETFONTSMOOTHINGCONTRAST = &h200C
		  dim mb as new MemoryBlock( 4 )
		  
		  SystemParametersInfo( SPI_GETFONTSMOOTHINGCONTRAST, 0, mb )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub FontSmoothingContrast(assigns set as Integer)
		  Const SPI_SETFONTSMOOTHINGCONTRAST = &h200D
		  dim mb as new MemoryBlock( 4 )
		  
		  if set < 1000 then set = 1000
		  if set > 2200 then set = 2200
		  
		  mb.Long( 0 ) = set
		  
		  SystemParametersInfo( SPI_SETFONTSMOOTHINGCONTRAST, 0, mb, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ForegroundLockTimeout() As Integer
		  Const SPI_GETFOREGROUNDLOCKTIMEOUT = &h2000
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETFOREGROUNDLOCKTIMEOUT, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ForegroundLockTimeout(assigns set as Integer)
		  Const SPI_SETFOREGROUNDLOCKTIMEOUT = &h2001
		  SystemParametersInfo( SPI_SETFOREGROUNDLOCKTIMEOUT, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetSystemMetrics(type as Integer) As Integer
		  #if TargetWin32
		    Declare Function MyGetSystemMetrics Lib "User32" Alias "GetSystemMetrics" ( type as Integer ) as Integer
		    
		    return MyGetSystemMetrics( type )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub GetWorkArea(ByRef left as Integer, ByRef top as Integer, ByRef right as Integer, ByRef bottom as Integer)
		  Const SPI_GETWORKAREA = &h30
		  Dim data as new MemoryBlock( 16 )
		  Dim mb as new MemoryBlock( 4 )
		  mb.Ptr( 0 ) = data
		  
		  SystemParametersInfo( SPI_GETWORKAREA, 0, mb, 0 )
		  
		  left = mb.Ptr( 0 ).Long( 0 )
		  top = mb.Ptr( 0 ).Long( 4 )
		  right = mb.Ptr( 0 ).Long( 8 )
		  bottom = mb.Ptr( 0 ).Long( 12 )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GradientCaptions() As Boolean
		  Const SPI_GETGRADIENTCAPTIONS = &h1008
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETGRADIENTCAPTIONS, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub GradientCaptions(assigns set as Boolean)
		  Const SPI_SETGRADIENTCAPTIONS = &h1009
		  SystemParametersInfo( SPI_SETGRADIENTCAPTIONS, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function HotTracking() As Boolean
		  Const SPI_GETHOTTRACKING = &h100E
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETHOTTRACKING, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub HotTracking(assigns set as Boolean)
		  Const SPI_SETHOTTRACKING = &h100F
		  SystemParametersInfo( SPI_SETHOTTRACKING, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IconHorizontalSpacing() As Integer
		  Const SPI_ICONHORIZONTALSPACING = &hD
		  
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_ICONHORIZONTALSPACING, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub IconHorizontalSpacing(assigns space as Integer)
		  Const SPI_ICONHORIZONTALSPACING = &hD
		  
		  SystemParametersInfo( SPI_ICONHORIZONTALSPACING, space, nil, SPIF_SENDWININICHANGE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IconLogicalFont() As LogicalFont
		  Const SPI_GETICONTITLELOGFONT = &h1F
		  
		  dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "SystemParametersInfoW", "User32" )
		  
		  dim logFont as MemoryBlock
		  if unicodeSavvy then
		    logFont = new MemoryBlock( LogicalFont.kWideCharSize )
		  else
		    logFont = new MemoryBlock( LogicalFont.kANSISize )
		  end if
		  
		  SystemParametersInfo( SPI_GETICONTITLELOGFONT, 0, logFont, 0 )
		  
		  return new LogicalFont( logFont, unicodeSavvy )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub IconLogicalFont(assigns theFont as LogicalFont)
		  Const SPI_SETICONTITLELOGFONT = &h22
		  SystemParametersInfo( SPI_SETICONTITLELOGFONT, 0, theFont.ToMemoryBlock( System.IsFunctionAvailable( "SystemParametersInfoW", "User32" ) ), SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IconMetrics() As IconMetrics
		  dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "SystemParametersInfoW", "User32" )
		  
		  dim mb as MemoryBlock
		  if unicodeSavvy then
		    mb = new MemoryBlock( IconMetrics.kWideCharSize )
		  else
		    mb = new MemoryBlock( IconMetrics.kANSISize )
		  end if
		  
		  Const SPI_GETICONMETRICS = &h2D
		  
		  mb.Long( 0 ) = mb.Size
		  SystemParametersInfo( SPI_GETICONMETRICS, 0, mb, 0 )
		  
		  return new IconMetrics( mb, unicodeSavvy )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub IconMetrics(assigns metrics as IconMetrics)
		  Const SPI_SETICONMETRICS = &h2E
		  
		  SystemParametersInfo( SPI_SETICONMETRICS, 0, metrics.ToMemoryBlock( System.IsFunctionAvailable( "SystemParametersInfoW", "User32" ) ), SPIF_UPDATEINIFILE + SPIF_SENDWININICHANGE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IconTitleWrap() As Boolean
		  Const SPI_GETICONTITLEWRAP =  &h19
		  
		  dim mb as new MemoryBlock( 1 )
		  SystemParametersInfo( SPI_GETICONTITLEWRAP, 0, mb, 0)
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub IconTitleWrap(assigns wrap as Boolean)
		  Const SPI_SETICONTITLEWRAP = &h1A
		  dim toWrap as Integer = 0
		  if wrap then toWrap = 1
		  SystemParametersInfo( SPI_SETICONTITLEWRAP, toWrap, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IconVerticalSpacing() As Integer
		  Const SPI_ICONVERTICALSPACING = &h18
		  
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_ICONVERTICALSPACING, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub IconVerticalSpacing(assigns space as Integer)
		  Const SPI_ICONVERTICALSPACING = &h18
		  
		  SystemParametersInfo( SPI_ICONVERTICALSPACING, space, nil, SPIF_SENDWININICHANGE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsWindowsExtensionInstalled() As Boolean
		  #if TargetWin32
		    Const SPI_GETWINDOWSEXTENSION = &h5C
		    
		    Declare Function SystemParametersInfoA Lib "User32" ( action as Integer, _
		    param1 as Integer, param2 as Integer, change as Integer ) as Boolean
		    
		    return SystemParametersInfoA( SPI_GETWINDOWSEXTENSION, 1, 0, 0 )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function KeyboardCues() As Boolean
		  Const SPI_GETKEYBOARDCUES = &h100A
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETKEYBOARDCUES, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub KeyboardCues(assigns set as Boolean)
		  Const SPI_SETKEYBOARDCUES = &h100B
		  SystemParametersInfo( SPI_SETKEYBOARDCUES,0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function KeyboardDelay() As Integer
		  Const SPI_GETKEYBOARDDELAY = &h16
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETKEYBOARDDELAY, 0, mb, 0 )
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub KeyboardDelay(assigns set as Integer)
		  if set < 0 then
		    set = 0
		  elseif set > 3 then
		    set = 3
		  end
		  
		  Const SPI_SETKEYBOARDDELAY = &h17
		  SystemParametersInfo( SPI_SETKEYBOARDDELAY, set, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function KeyboardPreferred() As Boolean
		  Const SPI_GETKEYBOARDPREF = &h44
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETKEYBOARDPREF, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub KeyboardPreferred(assigns set as Boolean)
		  Const SPI_SETKEYBOARDPREF = &h45
		  dim toSet as Integer
		  if set then toSet = 1
		  SystemParametersInfo( SPI_SETKEYBOARDPREF, toSet, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function KeyboardSpeed() As Integer
		  Const SPI_GETKEYBOARDSPEED = &hA
		  Dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETKEYBOARDSPEED, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub KeyboardSpeed(assigns set as Integer)
		  Const SPI_SETKEYBOARDSPEED = &hB
		  if set < 0 then
		    set = 0
		  elseif set > 31 then
		    set = 31
		  end if
		  
		  SystemParametersInfo( SPI_SETKEYBOARDSPEED, set, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ListBoxSmoothScrolling() As Boolean
		  Const SPI_GETLISTBOXSMOOTHSCROLLING = &h1006
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETLISTBOXSMOOTHSCROLLING, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ListBoxSmoothScrolling(assigns set as Boolean)
		  Const SPI_SETLISTBOXSMOOTHSCROLLING = &h1007
		  SystemParametersInfo( SPI_SETLISTBOXSMOOTHSCROLLING, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function LowPowerActive() As Boolean
		  Const SPI_GETLOWPOWERACTIVE = &h53
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETLOWPOWERACTIVE, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LowPowerActive(assigns set as Boolean)
		  Const SPI_SETLOWPOWERACTIVE = &h55
		  dim toSet as Integer
		  
		  if set then toSet = 1
		  SystemParametersInfo( SPI_SETLOWPOWERACTIVE, toSet, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function LowPowerTimeout() As Integer
		  Const SPI_GETLOWPOWERTIMEOUT = &h4F
		  dim mb as new MemoryBlock( 4 )
		  
		  SystemParametersInfo( SPI_GETLOWPOWERTIMEOUT, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LowPowerTimeout(assigns set as Integer)
		  Const SPI_SETLOWPOWERTIMEOUT = &h51
		  SystemParametersInfo( SPI_SETLOWPOWERTIMEOUT, set, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MenuAnimation() As Boolean
		  Const SPI_GETMENUANIMATION = &h1002
		  dim mb as new MemoryBlock( 4 )
		  
		  SystemParametersInfo( SPI_GETMENUANIMATION, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MenuAnimation(assigns set as Boolean)
		  Const SPI_SETMENUANIMATION = &h1003
		  SystemParametersInfo( SPI_SETMENUANIMATION, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MenuDropLeftAligned() As Boolean
		  Const SPI_GETMENUDROPALIGNMENT = &h1B
		  
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETMENUDROPALIGNMENT, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MenuDropLeftAligned(assigns set as Boolean)
		  Const SPI_SETMENUDROPALIGNMENT = &h1C
		  
		  dim toSet as Integer
		  if set then toSet = 1
		  SystemParametersInfo( SPI_SETMENUDROPALIGNMENT, toSet, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MenuFade() As Boolean
		  Const SPI_GETMENUFADE = &h1012
		  dim mb as new MemoryBlock( 4 )
		  
		  SystemParametersInfo( SPI_GETMENUFADE, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MenuFade(assigns set as Boolean)
		  Const SPI_SETMENUFADE = &h1013
		  SystemParametersInfo( SPI_SETMENUFADE, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MenuShowDelay() As Boolean
		  Const SPI_GETMENUSHOWDELAY = &h6A
		  dim mb as new MemoryBlock( 4 )
		  
		  SystemParametersInfo( SPI_GETMENUSHOWDELAY, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MenuShowDelay(assigns set as Boolean)
		  Const SPI_SETMENUSHOWDELAY = &h6B
		  dim toSet as Integer
		  if set then toSet = 1
		  
		  SystemParametersInfo( SPI_SETMENUSHOWDELAY, toSet, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MessageDuration() As Integer
		  Const SPI_GETMESSAGEDURATION = &h2016
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETMESSAGEDURATION, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MessageDuration(assigns set as Integer)
		  Const SPI_SETMESSAGEDURATION = &h2017
		  SystemParametersInfo( SPI_SETMESSAGEDURATION, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MouseClickLock() As Boolean
		  Const SPI_GETMOUSECLICKLOCK = &h101E
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETMOUSECLICKLOCK, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MouseClickLock(assigns set as Boolean)
		  Const SPI_SETMOUSECLICKLOCK = &h101F
		  dim toSet as Integer
		  if set then toSet = 1
		  
		  SystemParametersInfo( SPI_SETMOUSECLICKLOCK, toSet, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MouseClickLockTime() As Integer
		  Const SPI_GETMOUSECLICKLOCKTIME = &h2008
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETMOUSECLICKLOCKTIME, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MouseClickLockTime(assigns set as Integer)
		  Const SPI_SETMOUSECLICKLOCKTIME = &h2009
		  SystemParametersInfo( SPI_SETMOUSECLICKLOCKTIME, set, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MouseHoverHeight() As Integer
		  Const SPI_GETMOUSEHOVERHEIGHT = &h64
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETMOUSEHOVERHEIGHT, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MouseHoverHeight(assigns set as Integer)
		  Const SPI_SETMOUSEHOVERHEIGHT = &h65
		  SystemParametersInfo( SPI_SETMOUSEHOVERHEIGHT, set, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MouseHoverTime() As Integer
		  Const SPI_GETMOUSEHOVERTIME = &h66
		  Dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETMOUSEHOVERTIME, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MouseHoverTime(assigns set as Integer)
		  Const SPI_SETMOUSEHOVERTIME = &h67
		  SystemParametersInfo( SPI_SETMOUSEHOVERTIME, set, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MouseHoverWidth() As Integer
		  Const SPI_GETMOUSEHOVERWIDTH = &h62
		  dim mb as new MemoryBlock( 4 )
		  
		  SystemParametersInfo( SPI_GETMOUSEHOVERWIDTH, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MouseHoverWidth(assigns set as Integer)
		  Const SPI_SETMOUSEHOVERWIDTH = &h63
		  SystemParametersInfo( SPI_SETMOUSEHOVERWIDTH, set, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MouseInformation() As Integer()
		  Const SPI_GETMOUSE = &h3
		  Dim ret(-1) as Integer
		  dim mb as new MemoryBlock( 12 )
		  SystemParametersInfo( SPI_GETMOUSE, 0, mb, 0 )
		  
		  ret.Append( mb.Long( 0 ) )
		  ret.Append( mb.Long( 4 ) )
		  ret.Append( mb.Long( 8 ) )
		  
		  return ret
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MouseInformation(assigns set() as Integer)
		  Const SPI_SETMOUSE = &h4
		  dim mb as new MemoryBlock( 12 )
		  try
		    mb.Long( 0 ) = set( 0 )
		    mb.Long( 4 ) = set( 1 )
		    mb.Long( 8 ) = set( 2 )
		  catch err as OutOfBoundsException
		    // The user didn't specify enough values
		    return
		  end try
		  
		  SystemParametersInfo( SPI_SETMOUSE, 0, mb, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MouseKeys() As MouseKeys
		  Const SPI_GETMOUSEKEYS = &h36
		  dim mb as new MemoryBlock( 28 )
		  SystemParametersInfo( SPI_GETMOUSEKEYS, mb.Size, mb, 0 )
		  
		  return new MouseKeys( mb )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MouseKeys(assigns set as MouseKeys)
		  Const SPI_SETMOUSEKEYS = &h37
		  dim mb as MemoryBlock
		  mb = set.ToMemoryBlock
		  
		  SystemParametersInfo( SPI_SETMOUSEKEYS, mb.Size, mb, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MouseSnapToDefaultButton() As Boolean
		  Const SPI_GETSNAPTODEFBUTTON = &h5F
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETSNAPTODEFBUTTON, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MouseSnapToDefaultButton(assigns set as Boolean)
		  Const SPI_SETSNAPTODEFBUTTON = &h60
		  dim toSet as Integer
		  if set then toSet = 1
		  SystemParametersInfo( SPI_SETSNAPTODEFBUTTON, toSet, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MouseSonar() As Boolean
		  Const SPI_GETMOUSESONAR = &h101C
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETMOUSESONAR, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MouseSonar(assigns set as Boolean)
		  Const SPI_SETMOUSESONAR = &h101D
		  SystemParametersInfo( SPI_SETMOUSESONAR, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MouseSpeed() As Integer
		  Const SPI_GETMOUSESPEED = &h70
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETMOUSESPEED, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MouseSpeed(assigns set as Integer)
		  Const SPI_SETMOUSESPEED = &h71
		  SystemParametersInfo( SPI_SETMOUSESPEED, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MouseTrails() As Integer
		  Const SPI_GETMOUSETRAILS = &h5E
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETMOUSETRAILS, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MouseTrails(assigns numTrails as Integer)
		  Const SPI_SETMOUSETRAILS = &h5D
		  SystemParametersInfo( SPI_SETMOUSETRAILS, numTrails, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MouseVanish() As Boolean
		  Const SPI_GETMOUSEVANISH = &h1020
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETMOUSEVANISH, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MouseVanish(assigns set as Boolean)
		  Const SPI_SETMOUSEVANISH = &h1021
		  SystemParametersInfo( SPI_SETMOUSEVANISH, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MouseWheelLines() As Integer
		  Const SPI_GETWHEELSCROLLLINES = &h68
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETWHEELSCROLLLINES, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MouseWheelLines(assigns set as Integer)
		  Const SPI_SETWHEELSCROLLLINES = &h69
		  SystemParametersInfo( SPI_SETWHEELSCROLLLINES, set, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function PowerOffActive() As Boolean
		  Const SPI_GETPOWEROFFACTIVE = &h54
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETPOWEROFFACTIVE, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub PowerOffActive(assigns set as Boolean)
		  Const SPI_SETPOWEROFFACTIVE = &h56
		  dim toSet as Integer
		  if set then toSet = 1
		  SystemParametersInfo( SPI_SETPOWEROFFACTIVE, toSet, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function PowerOffTimeout() As Integer
		  Const SPI_GETPOWEROFFTIMEOUT = &h50
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETPOWEROFFTIMEOUT, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub PowerOffTimeout(assigns set as Integer)
		  Const SPI_SETPOWEROFFTIMEOUT = &h52
		  SystemParametersInfo( SPI_SETPOWEROFFTIMEOUT, set, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ReloadSystemCursors()
		  Const SPI_SETCURSORS = &h57
		  
		  SystemParametersInfo( SPI_SETCURSORS, 0, nil, SPIF_SENDWININICHANGE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ReloadSystemIcons()
		  Const SPI_SETICONS = &h58
		  
		  SystemParametersInfo( SPI_SETICONS, 0, nil, SPIF_SENDWININICHANGE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ScreenReader() As Boolean
		  Const SPI_GETSCREENREADER = &h46
		  Dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETSCREENREADER, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ScreenReader(assigns set as Boolean)
		  Const SPI_SETSCREENREADER = &h47
		  dim toSet as Integer
		  if set then toSet = 1
		  
		  SystemParametersInfo( SPI_SETSCREENREADER, toSet, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ScreenSaverActive() As Boolean
		  Const SPI_GETSCREENSAVEACTIVE = &h10
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETSCREENSAVEACTIVE, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ScreenSaverActive(assigns set as Boolean)
		  Const SPI_SETSCREENSAVEACTIVE = &h11
		  dim toSet as Integer
		  if set then toSet = 1
		  SystemParametersInfo( SPI_SETSCREENSAVEACTIVE, toSet, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ScreenSaverRunning() As Boolean
		  Const SPI_GETSCREENSAVERRUNNING = &h72
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETSCREENSAVERRUNNING, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ScreenSaverRunning(assigns set as Boolean)
		  Const SPI_SETSCREENSAVERRUNNING = &h61
		  dim toSet as Integer
		  if set then toSet = 1
		  
		  SystemParametersInfo( SPI_SETSCREENSAVERRUNNING, toSet, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ScreenSaverTimeout() As Integer
		  Const SPI_GETSCREENSAVETIMEOUT = &hE
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETSCREENSAVETIMEOUT, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ScreenSaverTimeout(assigns set as Integer)
		  Const SPI_SETSCREENSAVETIMEOUT = &hF
		  SystemParametersInfo( SPI_SETSCREENSAVETIMEOUT, set, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SelectionFade() As Boolean
		  Const SPI_GETSELECTIONFADE = &h1014
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETSELECTIONFADE, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub SelectionFade(assigns set as Boolean)
		  Const SPI_SETSELECTIONFADE = &h1015
		  SystemParametersInfo( SPI_SETSELECTIONFADE, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub SetWorkArea(left as Integer, top as Integer, right as Integer, bottom as Integer)
		  Const SPI_SETWORKAREA = &h2F
		  Dim data as new MemoryBlock( 16 )
		  data.Long( 0 ) = left
		  data.Long( 4 ) = top
		  data.Long( 8 ) = right
		  data.Long( 12 ) = bottom
		  Dim mb as new MemoryBlock( 4 )
		  mb.Ptr( 0 ) = data
		  
		  SystemParametersInfo( SPI_SETWORKAREA, 0, mb, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE  )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ShowIMEUI() As Boolean
		  Const SPI_GETSHOWIMEUI = &h6E
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETSHOWIMEUI, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ShowIMEUI(assigns set as Boolean)
		  Const SPI_SETSHOWIMEUI = &h6F
		  dim toSet as Integer
		  if set then toSet = 1
		  
		  SystemParametersInfo( SPI_SETSHOWIMEUI, toSet, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ShowSounds() As Boolean
		  Const SPI_GETSHOWSOUNDS = &h38
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETSHOWSOUNDS, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ShowSounds(assigns set as Boolean)
		  Const SPI_SETSHOWSOUNDS = &h39
		  dim toSet as Integer
		  if set then toSet = 1
		  
		  SystemParametersInfo( SPI_SETSHOWSOUNDS, toSet, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SpeechRecognitionRunning() As Boolean
		  Const SPI_GETSPEECHRECOGNITION = &h104A
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETSPEECHRECOGNITION, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub SpeechRecognitionRunning(assigns set as Boolean)
		  Const SPI_SETSPEECHRECOGNITION = &h104B
		  SystemParametersInfo( SPI_SETSPEECHRECOGNITION, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub SystemParametersInfo(action as Integer, param1 as Integer, oddSet as Boolean, iniChange as Integer)
		  #if TargetWin32
		    Soft Declare Sub SystemParametersInfoA Lib "User32" ( action as Integer, _
		    param1 as Integer, param2 as Boolean, change as Integer )
		    Soft Declare Sub SystemParametersInfoW Lib "User32" ( action as Integer, _
		    param1 as Integer, param2 as Boolean, change as Integer )
		    
		    if System.IsFunctionAvailable( "SystemParametersInfoW", "User32" ) then
		      SystemParametersInfoW( action, param1, oddSet, iniChange )
		    else
		      SystemParametersInfoA( action, param1, oddSet, iniChange )
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub SystemParametersInfo(action as Integer, param1 as Integer, param2 as Integer, iniChange as Integer)
		  #if TargetWin32
		    Soft Declare Sub SystemParametersInfoA Lib "User32" ( action as Integer, _
		    param1 as Integer, param2 as Integer, change as Integer )
		    Soft Declare Sub SystemParametersInfoW Lib "User32" ( action as Integer, _
		    param1 as Integer, param2 as Integer, change as Integer )
		    
		    if System.IsFunctionAvailable( "SystemParametersInfoW", "User32" ) then
		      SystemParametersInfoW( action, param1, param2, iniChange )
		    else
		      SystemParametersInfoA( action, param1, param2, iniChange )
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub SystemParametersInfo(action as Integer, param1 as Integer, param2 as MemoryBlock, iniChange as Integer = &h2)
		  if param2 = nil then
		    SystemParametersInfo( action, param1, 0, iniChange )
		    return
		  end if
		  
		  #if TargetWin32
		    Soft Declare Function SystemParametersInfoA Lib "User32" ( action as Integer, _
		    param1 as Integer, param2 as Ptr, change as Integer ) as Boolean
		    Soft Declare Function SystemParametersInfoW Lib "User32" ( action as Integer, _
		    param1 as Integer, param2 as Ptr, change as Integer ) as Boolean
		    
		    
		    dim ret as Boolean
		    if System.IsFunctionAvailable( "SystemParametersInfoW", "User32" ) then
		      ret = SystemParametersInfoW( action, param1, param2, iniChange )
		    else
		      ret = SystemParametersInfoA( action, param1, param2, iniChange )
		    end if
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ToggleKeys() As ToggleKeys
		  Const SPI_GETTOGGLEKEYS = &h34
		  dim mb as new MemoryBlock( 8 )
		  SystemParametersInfo( SPI_GETTOGGLEKEYS, mb.Size, mb, 0 )
		  
		  return new ToggleKeys( mb )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ToggleKeys(assigns set as ToggleKeys)
		  Const SPI_SETTOGGLEKEYS = &h35
		  dim mb as MemoryBlock
		  mb = set.ToMemoryBlock
		  
		  SystemParametersInfo( SPI_SETTOGGLEKEYS, mb.Size, mb, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ToolTipAnimation() As Boolean
		  Const SPI_GETTOOLTIPANIMATION = &h1016
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETTOOLTIPANIMATION, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ToolTipAnimation(assigns set as Boolean)
		  Const SPI_SETTOOLTIPANIMATION = &h1017
		  SystemParametersInfo( SPI_SETTOOLTIPANIMATION, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ToolTipFade() As Boolean
		  Const SPI_GETTOOLTIPFADE = &h1018
		  Dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETTOOLTIPFADE, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ToolTipFade(assigns set as Boolean)
		  Const SPI_SETTOOLTIPFADE = &h1019
		  SystemParametersInfo( SPI_SETTOOLTIPFADE, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function UIEffects() As Boolean
		  Const SPI_GETUIEFFECTS = &h103E
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETUIEFFECTS, 0, mb, 0 )
		  
		  return mb.BooleanValue( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub UIEffects(assigns set as Boolean)
		  Const SPI_SETUIEFFECTS = &h103F
		  SystemParametersInfo( SPI_SETUIEFFECTS, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub UpdateDesktopPattern()
		  Const SPI_SETDESKPATTERN = &h15
		  
		  SystemParametersInfo( SPI_SETDESKPATTERN, 0, nil, SPIF_UPDATEINIFILE + SPIF_SENDWININICHANGE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function WheelScrollChars() As Integer
		  Const SPI_GETWHEELSCROLLCHARS = &h6C
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETWHEELSCROLLCHARS, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub WheelScrollChars(assigns set as Integer)
		  Const SPI_SETWHEELSCROLLCHARS = &h6D
		  SystemParametersInfo( SPI_SETWHEELSCROLLCHARS, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function WindowAnimation() As Boolean
		  Const SPI_GETANIMATION = &h48
		  dim anim as new MemoryBlock( 8 )
		  anim.Long( 0 ) = anim.Size
		  SystemParametersInfo( SPI_GETANIMATION, anim.Size, anim, 0 )
		  
		  return anim.Long( 4 ) <> 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub WindowAnimation(assigns set as Boolean)
		  Const SPI_SETANIMATION = &h49
		  dim anim as new MemoryBlock( 8 )
		  anim.Long( 0 ) = anim.Size
		  if set then
		    anim.Long( 4 ) = 1
		  end
		  SystemParametersInfo( SPI_SETANIMATION, anim.Size, anim, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function WindowBorderMultiplier() As Integer
		  Const SPI_GETBORDER = &h5
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETBORDER, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub WindowBorderMultiplier(assigns set as Integer)
		  Const SPI_SETBORDER = &h6
		  SystemParametersInfo( SPI_SETBORDER, set, nil, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function WindowFlashCount() As Integer
		  Const SPI_GETFOREGROUNDFLASHCOUNT = &h2004
		  dim mb as new MemoryBlock( 4 )
		  SystemParametersInfo( SPI_GETFOREGROUNDFLASHCOUNT, 0, mb, 0 )
		  
		  return mb.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub WindowFlashCount(assigns set as Integer)
		  Const SPI_SETFOREGROUNDFLASHCOUNT = &h2005
		  SystemParametersInfo( SPI_SETFOREGROUNDFLASHCOUNT, 0, set, SPIF_SENDWININICHANGE + SPIF_UPDATEINIFILE )
		End Sub
	#tag EndMethod


	#tag Note, Name = ToDo
		SerialKeys class and getter/setter
		HighContrast class and getter/setter
		SoundSentry class and getter/setter
		AudioDescription class and getter/setter
	#tag EndNote


	#tag Constant, Name = SPIF_SENDWININICHANGE, Type = Integer, Dynamic = False, Default = \"&h2", Scope = Private
	#tag EndConstant

	#tag Constant, Name = SPIF_UPDATEINIFILE, Type = Integer, Dynamic = False, Default = \"&h1", Scope = Private
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
