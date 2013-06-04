#tag Module
Protected Module WindowExtensionsWFS
	#tag Method, Flags = &h0
		Function Alpha(extends w as Window) As Single
		  #if TargetWin32
		    Const LWA_ALPHA = 2
		    Soft Declare Sub GetLayeredWindowAttributes Lib "user32" ( hwnd As Integer, thecolor As Integer, ByRef bAlpha As integer, flags As Integer )
		    
		    if not System.IsFunctionAvailable( "GetLayeredWindowAttributes", "User32" ) then return 255.0
		    
		    Dim alpha as Integer = 0
		    GetLayeredWindowAttributes( w.WinHWND, 0 , alpha, LWA_ALPHA )
		    
		    return alpha / 255.0
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Alpha(extends w as Window, assigns alpha as Single)
		  #if TargetWin32
		    // First, check to see if we've set this window up to be layered yet
		    Const WS_EX_LAYERED = &H80000
		    Const LWA_ALPHA = 2
		    
		    if not TestWindowStyleEx( w, WS_EX_LAYERED ) then
		      // The window isn't layered, so make it so
		      ChangeWindowStyleEx( w, WS_EX_LAYERED, true )
		    end
		    
		    // Now we want to set the transparency of the window.  The values range from 0 (totally
		    // transparent) to 255 (totally opaque).
		    dim value as Integer = 255 * alpha
		    
		    Soft Declare Sub SetLayeredWindowAttributes Lib "user32" ( hwnd As Integer, thecolor As Integer, bAlpha As integer, alpha As Integer )
		    
		    if System.IsFunctionAvailable( "SetLayeredWindowAttributes", "User32" ) then
		      SetLayeredWindowAttributes( w.WinHWND, 0 , value, LWA_ALPHA )
		    end if
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AnimateWindow(extends w as Window, microsecs as Integer, flags as Integer)
		  'More info at http://www.developersdomain.com/vb/codesnippets/windows.htm
		  #if targetWin32 then
		    Declare Sub AnimateWindow Lib "user32" ( hwnd As Integer, dwTime As Integer, dwFlags As Integer )
		    
		    AnimateWindow( w.WinHWND, microsecs, flags )
		    
		    if BitwiseAnd( flags, kAnimateWindowActionHide ) <> 0 then
		      w.Hide
		    else
		      w.Refresh( false )
		    end
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub BringToFront(extends w as Window)
		  Dim i, h, r As Integer
		  
		  #if TargetWin32
		    if not w.visible then
		      w.visible = true
		    end if
		    
		    Declare Function SetWindowPos Lib "User32" (hWnd As Integer, hWndInsertAfter As Integer, _
		    X As Integer, Y As Integer, cx As Integer, cy As Integer, wFlags As Integer) As Integer
		    Declare Function ShowWindow Lib "User32" ( hWnd As Integer, nCmdShow As Integer ) As Integer
		    Declare Function BringWindowToTop Lib "User32" ( hWnd As Integer ) As Integer
		    Declare Function SetForegroundWindow Lib "User32" ( hWnd As Integer ) As Integer
		    
		    ' Get the Window Handle
		    h = w.WinHWND
		    r = &h1 + &h2 + &h40
		    
		    ' Activate the Window
		    i = SetWindowPos( h,-1, 0,0,0,0, r )
		    i = ShowWindow( h, 1 )
		    i = BringWindowToTop( h )
		    i = SetForegroundWindow( h )
		    w.FlashWindowEx( 3 )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ChangeWindowState(extends wnd as Window, style as Integer)
		  #if TargetWin32
		    Declare Sub ShowWindow Lib "User32" (wnd As Integer, nCmdShow As Integer)
		    
		    ShowWindow( wnd.WinHWND, style )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ChangeWindowStyle(w as Window, flag as Integer, set as Boolean)
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
		    
		    oldFlags = GetWindowLong(w.WinHWND, GWL_STYLE)
		    
		    if not set then
		      newFlags = BitwiseAnd( oldFlags, Bitwise.OnesComplement( flag ) )
		    else
		      newFlags = BitwiseOr( oldFlags, flag )
		    end
		    
		    
		    styleFlags = SetWindowLong( w.WinHWND, GWL_STYLE, newFlags )
		    styleFlags = SetWindowPos( w.WinHWND, 0, 0, 0, 0, 0, SWP_NOMOVE +_
		    SWP_NOSIZE + SWP_NOZORDER + SWP_FRAMECHANGED )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ChangeWindowStyleEx(w as Window, flag as Integer, set as Boolean)
		  #if TargetWin32
		    Dim oldFlags as Integer
		    Dim newFlags as Integer
		    Dim styleFlags As Integer
		    
		    Const SWP_NOSIZE = &H1
		    Const SWP_NOMOVE = &H2
		    Const SWP_NOZORDER = &H4
		    Const SWP_FRAMECHANGED = &H20
		    
		    Const GWL_EXSTYLE = -20
		    
		    Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (hwnd As Integer,  _
		    nIndex As Integer) As Integer
		    Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (hwnd As Integer, _
		    nIndex As Integer, dwNewLong As Integer) As Integer
		    Declare Function SetWindowPos Lib "user32" (hwnd as Integer, hWndInstertAfter as Integer, _
		    x as Integer, y as Integer, cx as Integer, cy as Integer, flags as Integer) as Integer
		    
		    oldFlags = GetWindowLong(w.WinHWND, GWL_EXSTYLE)
		    
		    if not set then
		      newFlags = BitwiseAnd( oldFlags, Bitwise.OnesComplement( flag ) )
		    else
		      newFlags = BitwiseOr( oldFlags, flag )
		    end
		    
		    
		    styleFlags = SetWindowLong( w.WinHWND, GWL_EXSTYLE, newFlags )
		    styleFlags = SetWindowPos( w.WinHWND, 0, 0, 0, 0, 0, SWP_NOMOVE +_
		    SWP_NOSIZE + SWP_NOZORDER + SWP_FRAMECHANGED )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CloseButtonState(extends w as Window) As Boolean
		  #if TargetWin32
		    Declare Function GetSystemMenu Lib "User32" ( wnd as Integer, revert as Boolean ) as Integer
		    Declare Function GetMenuState Lib "User32" ( menu as Integer, which as Integer, flags as Integer ) as Integer
		    
		    dim menu as Integer = GetSystemMenu( w.WinHWND, false )
		    if menu = 0 then return false
		    
		    Const SC_CLOSE = &hF060
		    Const MF_BYCOMMAND = 0
		    Const MF_GRAYED = 1
		    Const MF_ENABLED = 0
		    Const MF_DISABLED = 2
		    
		    dim state as Integer
		    state = GetMenuState( menu, SC_CLOSE, MF_BYCOMMAND )
		    
		    return Bitwise.BitAnd( state, MF_DISABLED + MF_GRAYED ) = 0
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CloseButtonState(extends w as Window, assigns enabled as Boolean)
		  #if TargetWin32
		    Declare Function GetSystemMenu Lib "User32" ( wnd as Integer, revert as Boolean ) as Integer
		    Declare Sub EnableMenuItem Lib "User32" ( menu as Integer, which as Integer, flags as Integer )
		    
		    dim menu as Integer = GetSystemMenu( w.WinHWND, false )
		    if menu = 0 then return
		    
		    Const SC_CLOSE = &hF060
		    Const MF_BYCOMMAND = 0
		    Const MF_GRAYED = 1
		    Const MF_ENABLED = 0
		    
		    if not enabled then
		      EnableMenuItem( menu, SC_CLOSE, MF_BYCOMMAND + MF_GRAYED )
		    else
		      EnableMenuItem( menu, SC_CLOSE, MF_BYCOMMAND + MF_ENABLED )
		    end
		  #endif
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
		Sub FlashWindow(Extends win as window)
		  #if TargetWin32
		    //works on windows 95+
		    dim res as integer
		    Declare Function FlashWindow Lib "user32" (hwnd As integer, bInvert As integer) As integer
		    Const Invert = 1
		    res = FlashWindow(win.winhwnd,Invert)
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub FlashWindowEx(Extends win as window, x as integer)
		  #if TargetWin32
		    //requires windows 98 or higher
		    Const FLASHW_STOP = 0 'Stop flashing. The system restores the window to its original state.
		    Const FLASHW_CAPTION = &H1 'Flash the window caption.
		    Const FLASHW_TRAY = &H2 'Flash the taskbar button.
		    'Const FLASHW_ALL = (FLASHW_CAPTION Or FLASHW_TRAY) 'Flash both the window caption and taskbar button. This is equivalent to setting the FLASHW_CAPTION Or FLASHW_TRAY flags.
		    Const FLASHW_TIMER = &H4 'Flash continuously, until the FLASHW_STOP flag is set.
		    Const FLASHW_TIMERNOFG = &HC 'Flash continuously until the window comes to the foreground.
		    'Private Type FLASHWINFO
		    'cbSize As Long '0
		    'hwnd As Long '4
		    'dwFlags As Long '8
		    'uCount As Long '12
		    'dwTimeout As Long '16
		    'End Type
		    Declare Function FlashWindowEx Lib "user32" (pfwi As ptr) As integer
		    Dim FlashInfo As memoryblock
		    flashinfo = new MemoryBlock(20)
		    'Specifies the size of the structure.
		    flashinfo.Long(0) = 20
		    'Specifies the flash status
		    flashinfo.long(8) = FLASHW_CAPTION + FLASHW_TRAY
		    'Specifies the rate, in milliseconds, at which the window will be flashed. If dwTimeout is zero, the function uses the default cursor blink rate.
		    FlashInfo.long(16) = 0
		    'Handle to the window to be flashed. The window can be either opened or minimized.
		    FlashInfo.long(4) = win.winhWND
		    'Specifies the number of times to flash the window.
		    FlashInfo.long(12) = x
		    if FlashWindowEx(FlashInfo) =0 then
		      ///succeeded
		    end
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub FreezeUpdate(extends w as Window)
		  #if TargetWin32
		    // This is incorrect and should only be used for drag and drop operations
		    // on ancient versions of Windows that REALbasic doesn't even support
		    'Declare Sub LockWindowUpdate Lib "User32" ( hwnd as Integer )
		    '
		    'LockWindowUpdate( w.WinHWND )
		    
		    // We use the WM_SETREDRAW message instead to lock the redraw
		    // state of the window (note that this can be used for controls as well)
		    Const WM_SETREDRAW = &hB
		    call SendMessage( w.Handle, WM_SETREDRAW, 0, 0 )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub GetWindowRect(w as Window, ByRef trueLeft as Integer, ByRef trueTop as Integer, ByRef trueRight as Integer, ByRef trueBottom as Integer)
		  #if TargetWin32
		    Declare Sub MyGetWindowRect Lib "User32" Alias "GetWindowRect" ( w as WindowPtr, r as Ptr )
		    
		    dim r as new MemoryBlock( 16 )
		    MyGetWindowRect( w, r )
		    
		    trueLeft = r.Long( 0 )
		    trueTop = r.Long( 4 )
		    trueRight = r.Long( 8 )
		    trueBottom = r.Long( 12 )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HasBorder(extends w as Window) As Boolean
		  Const WS_BORDER = &H800000
		  
		  return TestWindowStyle( w, WS_BORDER )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub HasBorder(extends w as Window, assigns border as Boolean)
		  Const WS_BORDER = &H800000
		  ChangeWindowStyle( w, WS_BORDER, border )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HasMaximizeButton(extends w as Window) As Boolean
		  Const WS_MAXIMIZEBOX = &h00010000
		  
		  return TestWindowStyle( w, WS_MAXIMIZEBOX )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub HasMaximizeButton(extends w as Window, assigns set as Boolean)
		  Const WS_MAXIMIZEBOX = &h00010000
		  
		  ChangeWindowStyle( w, WS_MAXIMIZEBOX, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HasMinimizeButton(extends w as Window) As Boolean
		  Const WS_MINIMIZEBOX = &h00020000
		  
		  return TestWindowStyle( w, WS_MINIMIZEBOX )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub HasMinimizeButton(extends w as Window, assigns set as Boolean)
		  Const WS_MINIMIZEBOX = &h00020000
		  
		  ChangeWindowStyle( w, WS_MINIMIZEBOX, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HasSystemMenu(extends w as Window) As Boolean
		  Const WS_SYSMENU = &h00080000
		  
		  return TestWindowStyle( w, WS_SYSMENU )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub HasSystemMenu(extends w as Window, assigns set as Boolean)
		  Const WS_SYSMENU = &h00080000
		  
		  ChangeWindowStyle( w, WS_SYSMENU, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HasTitleBar(extends w as Window) As Boolean
		  Const WS_CAPTION = &h00C00000
		  
		  return TestWindowStyle( w, WS_CAPTION )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub HasTitleBar(extends w as Window, assigns set as Boolean)
		  Const WS_CAPTION = &h00C00000
		  
		  ChangeWindowStyle( w, WS_CAPTION, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function HICONFromRBPicture(theIcon as Picture) As Integer
		  // This code was copy and pasted from the StatusBar class!
		  
		  // If there's no picture, then there's no HICON
		  if theIcon = nil then return 0
		  
		  #if RBVersion < 2010.05 then
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
		    Declare Sub REALLockPictureDescription Lib "" ( pic as Picture, desc As Ptr, request as Integer )
		    Declare Sub unlockPictureDescription Lib "" ( pic as Picture )
		  #endif
		  
		  // We are given an RB Picture object (and it may contain the mask information)
		  // that we want to turn into an HICON.
		  
		  // We will fill out an ICONINFO structure with the proper data, and call
		  // CreateIconIndirect on it, then store the results in a dictionary so that
		  // we don't need to do the conversion multiple times.
		  
		  dim iconInfo as new MemoryBlock( 20 )
		  iconInfo.Long( 0 ) = 1  ' We are an Icon, not a cursor
		  // We can ignore the hotspot information since we're not a cursor
		  
		  dim p as Picture = theIcon
		  
		  #if RBVersion < 2010.05 then
		    // Set the mask HBITMAP
		    dim maskDesc as new MemoryBlock( 21 )
		    REALLockPictureDescription( p.Mask, maskDesc, 7 )
		    if maskDesc.Long( 0 ) = 7 then
		      iconInfo.Long( 12) = maskDesc.Long( 4 )
		    end if
		    
		    // Set the color HBITMAP
		    dim picDesc as new MemoryBlock( 21 )
		    REALLockPictureDescription( p, picDesc, 7 )
		    if picDesc.Long( 0 ) = 7 then
		      iconInfo.Long( 16 ) = picDesc.Long( 4 )
		    end if
		  #else
		    // Since REAL Software decided to break Lib "" declares without giving
		    // any replacement whatsoever for this functionality, we have to go the
		    // slow route.  We will save the bitmaps out to disk, and then load them
		    // back up so we can get the handle.
		    Declare Function LoadImageW Lib "User32" ( hInst as Integer, name as WString, type as Integer, cx as Integer, cy as Integer, load as Integer ) as Integer
		    
		    dim main as FolderItem = SpecialFolder.Temporary.Child( "main" )
		    dim test as Picture
		    test = p.Mask
		    p.Mask = nil
		    
		    p.Save( main, Picture.SaveAsWindowsBMP )
		    Const IMAGE_BITMAP = 0
		    Const LR_LOADFROMFILE = &h10
		    Const LR_CREATEDIBSECTION = &h2000
		    iconInfo.Long( 16 ) = LoadImageW( 0, main.AbsolutePath, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE )
		    
		    dim mask as FolderItem = SpecialFolder.Temporary.Child( "mask" )
		    test.Save( mask, Picture.SaveAsWindowsBMP )
		    iconInfo.Long( 12 ) = LoadImageW( 0, mask.AbsolutePath, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE )
		    
		    p.Mask = test
		  #endif
		  
		  // Now we can create the HICON
		  Declare Function CreateIconIndirect Lib "User32" ( indirect as Ptr ) as Integer
		  dim hicon as Integer = CreateIconIndirect( iconInfo )
		  
		  #if RBVersion < 2010.05 then
		    // And unlock our picture descriptions
		    unlockPictureDescription( p.Mask )
		    unlockPictureDescription( p )
		  #endif
		  
		  // And return the new HICON
		  return hicon
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Icon(extends w as Window, assigns newIcon as Picture)
		  // We want to set the new icon for this window
		  Const WM_SETICON = &h80
		  Const ICON_SMALL = 0
		  
		  #if TargetWin32
		    Soft Declare Sub SendMessageA Lib "User32" ( hwnd as Integer, msg as Integer, wParam as Integer, lParam as Integer )
		    Soft Declare Sub SendMessageW Lib "User32" ( hwnd as Integer, msg as Integer, wParam as Integer, lParam as Integer )
		    
		    dim ret as Integer
		    if System.IsFunctionAvailable( "SendMessageW", "User32" ) then
		      SendMessageW( w.Handle, WM_SETICON, ICON_SMALL, HICONFromRBPicture( newIcon ) )
		    else
		      SendMessageA( w.Handle, WM_SETICON, ICON_SMALL, HICONFromRBPicture( newIcon ) )
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IntToColor(c as Integer) As Color
		  dim mb as new MemoryBlock( 4 )
		  mb.Long( 0 ) = c
		  
		  return RGB( mb.Byte( 0 ), mb.Byte( 1 ), mb.Byte( 2 ) )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsMaximized(extends w as Window) As Boolean
		  #if TargetWin32
		    Declare Function IsZoomed Lib "User32" ( hwnd As Integer ) As Integer
		    
		    return IsZoomed( w.Handle ) <> 0
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsMinimized(extends w as Window) As Boolean
		  #if TargetWin32
		    Declare Function IsIconic Lib "User32" ( hwnd As Integer ) As Integer
		    
		    return IsIconic( w.Handle ) <> 0
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsToolbarWindow(extends w as Window) As Boolean
		  Const WS_EX_TOOLWINDOW = &h00000080
		  return TestWindowStyleEx( w, WS_EX_TOOLWINDOW )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub IsToolbarWindow(extends w as Window, assigns set as Boolean)
		  Const WS_EX_TOOLWINDOW = &h00000080
		  
		  ChangeWindowStyleEx( w, WS_EX_TOOLWINDOW, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Mask(extends w as Window) As Color
		  #if TargetWin32
		    Const LWA_COLOR_KEY = 2
		    Soft Declare Sub GetLayeredWindowAttributes Lib "user32" ( hwnd As Integer, ByRef thecolor As Integer, bAlpha As integer, flags As Integer )
		    
		    if not System.IsFunctionAvailable( "GetLayeredWindowAttributes", "User32" ) then return &cFFFFFF
		    
		    Dim theColor as Integer = 0
		    GetLayeredWindowAttributes( w.WinHWND, theColor, 0, LWA_COLOR_KEY)
		    
		    return IntToColor( theColor )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Mask(extends w as Window, assigns c as Color)
		  // First, check to see if we've set this window up to be layered yet
		  Const WS_EX_LAYERED = 524288 //2^19
		  Const LWA_COLORKEY = 1
		  
		  if not TestWindowStyleEx( w, WS_EX_LAYERED ) then
		    // The window isn't layered, so make it so
		    ChangeWindowStyleEx( w, WS_EX_LAYERED, true )
		  end
		  
		  #if targetWin32 then
		    Soft Declare Sub SetLayeredWindowAttributes Lib "user32" ( hwnd As Integer, thecolor As Integer, bAlpha As integer, alpha As Integer )
		    
		    if System.IsFunctionAvailable( "SetLayeredWindowAttributes", "User32" ) then
		      SetLayeredWindowAttributes( w.WinHWND, ColorToInt( c ), 0, LWA_COLORKEY )
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Resizable(extends w as Window) As Boolean
		  Const WS_SIZEBOX = &h00040000
		  
		  return TestWindowStyle( w, WS_SIZEBOX )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Resizable(extends w as Window, assigns set as Boolean)
		  Const WS_SIZEBOX = &h00040000
		  
		  ChangeWindowStyle( w, WS_SIZEBOX, set )
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
		Private Sub SetWindowPos(w as Window, x as Integer, y as Integer)
		  #if TargetWin32
		    Const SWP_NOSIZE = &H1
		    Const SWP_NOMOVE = &H2
		    Const SWP_NOZORDER = &H4
		    
		    Declare Sub MySetWindowPos Lib "User32" Alias "SetWindowPos" ( hwnd as Integer, hWndInstertAfter as Integer, _
		    x as Integer, y as Integer, cx as Integer, cy as Integer, flags as Integer )
		    
		    MySetWindowPos( w.Handle, 0, x, y, 0, 0, SWP_NOSIZE + SWP_NOZORDER )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function TestWindowStyle(w as Window, flag as Integer) As Boolean
		  #if TargetWin32
		    Dim oldFlags as Integer
		    
		    Const GWL_STYLE = -16
		    
		    Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (hwnd As Integer,  _
		    nIndex As Integer) As Integer
		    
		    oldFlags = GetWindowLong(w.WinHWND, GWL_STYLE)
		    
		    if Bitwise.BitAnd( oldFlags, flag ) = flag then
		      return true
		    else
		      return false
		    end
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function TestWindowStyleEx(w as Window, flag as Integer) As Boolean
		  #if TargetWin32
		    Dim oldFlags as Integer
		    
		    Const GWL_EXSTYLE = -20
		    
		    Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (hwnd As Integer,  _
		    nIndex As Integer) As Integer
		    
		    oldFlags = GetWindowLong(w.WinHWND, GWL_EXSTYLE)
		    
		    if Bitwise.BitAnd( oldFlags, flag ) = flag then
		      return true
		    else
		      return false
		    end
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Topmost(extends w as Window) As Boolean
		  Const WS_EX_TOPMOST = &h00000008
		  return TestWindowStyleEx( w, WS_EX_TOPMOST )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Topmost(extends w as Window, assigns set as Boolean)
		  #if TargetWin32
		    Const WS_EX_TOPMOST = &h00000008
		    ChangeWindowStyleEx( w, WS_EX_TOPMOST, set )
		    
		    Const SWP_NOSIZE = &H1
		    Const SWP_NOMOVE = &H2
		    Const HWND_TOPMOST = -1
		    Const HWND_NOTOPMOST = -2
		    
		    Declare Function SetWindowPos Lib "user32" (hwnd as Integer, hWndInstertAfter as Integer, _
		    x as Integer, y as Integer, cx as Integer, cy as Integer, flags as Integer) as Integer
		    
		    dim after as Integer
		    if set then
		      after = HWND_TOPMOST
		    else
		      after = HWND_NOTOPMOST
		    end
		    Call SetWindowPos( w.WinHWND, after, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TrueBottom(extends w as Window) As Integer
		  dim top, left, right, bottom as Integer
		  GetWindowRect( w, left, top, right, bottom )
		  
		  return bottom
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TrueLeft(extends w as Window) As Integer
		  dim top, left, right, bottom as Integer
		  GetWindowRect( w, left, top, right, bottom )
		  
		  return left
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub TrueLeft(extends w as Window, assigns s as Integer)
		  SetWindowPos( w, s, w.TrueTop )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TrueRight(extends w as Window) As Integer
		  dim top, left, right, bottom as Integer
		  GetWindowRect( w, left, top, right, bottom )
		  
		  return right
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TrueTop(extends w as Window) As Integer
		  dim top, left, right, bottom as Integer
		  GetWindowRect( w, left, top, right, bottom )
		  
		  return top
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub TrueTop(extends w as Window, assigns s as Integer)
		  SetWindowPos( w, w.TrueLeft, s )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub UnfreezeUpdate(extends w as Window)
		  #if TargetWin32
		    // This is incorrect and should only be used for drag and drop operations
		    // on ancient versions of Windows that REALbasic doesn't even support
		    'Declare Sub LockWindowUpdate Lib "User32" ( hwnd as Integer )
		    '
		    'LockWindowUpdate( 0 )
		    
		    // We use the WM_SETREDRAW message instead to unlock the redraw
		    // state of the window (note that this can be used for controls as well)
		    Const WM_SETREDRAW = &hB
		    call SendMessage( w.Handle, WM_SETREDRAW, 1, 0 )
		  #endif
		End Sub
	#tag EndMethod


	#tag Constant, Name = kAnimateWindowActionHide, Type = Integer, Dynamic = False, Default = \"&h10000", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAnimateWindowActionShow, Type = Integer, Dynamic = False, Default = \"&h20000", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAnimateWindowEffectBlend, Type = Integer, Dynamic = False, Default = \"&h80000", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAnimateWindowEffectCenter, Type = Integer, Dynamic = False, Default = \"&h10", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAnimateWindowEffectHorizontalMinus, Type = Integer, Dynamic = False, Default = \"&h2", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAnimateWindowEffectHorizontalPlus, Type = Integer, Dynamic = False, Default = \"&h1", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAnimateWindowEffectSlide, Type = Integer, Dynamic = False, Default = \"&h40000", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAnimateWindowEffectVertialMinus, Type = Integer, Dynamic = False, Default = \"&h8", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAnimateWindowEffectVertialPlus, Type = Integer, Dynamic = False, Default = \"&h4", Scope = Protected
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
