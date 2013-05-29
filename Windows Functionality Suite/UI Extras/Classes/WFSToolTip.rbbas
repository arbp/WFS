#tag Class
Protected Class WFSToolTip
	#tag Method, Flags = &h0
		Sub AdjustPopRectangle(left as Integer, top as Integer, width as Integer, height as Integer)
		  mLeft = left
		  mTop = top
		  mWidth = width
		  mHeight = height
		  
		  Dim size as Integer = 44
		  // Add this back in if supporting bitmaps sometime
		  'if OSVersionInformation.IsNT5 then
		  'size = size + 4
		  'end
		  
		  Dim toolinfo as new MemoryBlock( size )
		  
		  // We want to add a tool mesage to the window now
		  Dim TTM_SETTOOLINFOA as Integer = WM_USER + 9
		  Dim TTM_SETTOOLINFOW as Integer = WM_USER + 54
		  Const TTF_SUBCLASS = &h10
		  Const TTF_IDISHWND = &h1
		  toolinfo.Long( 0 ) = toolinfo.Size
		  toolinfo.Long( 4 ) = TTF_IDISHWND + TTF_SUBCLASS
		  toolinfo.Long( 8 ) = mControlHandle
		  toolinfo.Long( 12 ) = mControlHandle
		  toolinfo.Long( 16 ) = mLeft
		  toolinfo.Long( 20 ) = mTop
		  toolinfo.Long( 24 ) = mLeft + mWidth
		  toolinfo.Long( 28 ) = mTop + mHeight
		  toolinfo.Long( 32 ) = mInstance
		  
		  dim textPtr as MemoryBlock
		  if mUnicodeSavvy then
		    textPtr = new MemoryBlock( (Len( mText ) + 1) * 2 )
		    textPtr.WString( 0 ) = mText
		  else
		    textPtr = new MemoryBlock( Len( mText ) + 1 )
		    textPtr.CString( 0 ) = mText
		  end if
		  toolinfo.Ptr( 36 ) = textPtr
		  
		  dim whatMsg as Integer = TTM_SETTOOLINFOA
		  if mUnicodeSavvy then whatMsg = TTM_SETTOOLINFOW
		  Call SendMessage( whatMsg, 0, MemoryBlockToInteger( toolinfo ) )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function AlwaysTip() As Boolean
		  return mAlwaysTip
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AlwaysTip(assigns set as Boolean)
		  mAlwaysTip = set
		  
		  if mWnd = 0 then return
		  
		  Const TTS_ALWAYSTIP = &h1
		  ChangeWindowStyle( TTS_ALWAYSTIP, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Animate() As Boolean
		  return mAnimate
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Animate(assigns set as Boolean)
		  mAnimate = set
		  
		  if mWnd = 0 then return
		  
		  Const TTS_NOANIMATE = &h10
		  ChangeWindowStyle( TTS_NOANIMATE, not set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function BackgroundColor() As Color
		  if mWnd = 0 then return RGB( 0, 0, 0 )
		  
		  Dim TTM_GETTIPBKCOLOR as Integer = WM_USER + 22
		  
		  dim theColor as Integer
		  theColor = SendMessage( TTM_GETTIPBKCOLOR, 0, 0 )
		  
		  dim mb as new MemoryBlock( 4 )
		  mb.Long( 0 ) = theColor
		  
		  return RGB( mb.Byte( 0 ), mb.Byte( 1 ), mb.Byte( 2 ) )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub BackgroundColor(assigns c as Color)
		  if mWnd = 0 then return
		  
		  Dim TTM_SETTIPBKCOLOR as Integer = WM_USER + 19
		  
		  dim theColor as new MemoryBlock( 4 )
		  theColor.Byte( 0 ) = c.red
		  theColor.Byte( 1 ) = c.green
		  theColor.Byte( 2 ) = c.blue
		  
		  Call SendMessage( TTM_SETTIPBKCOLOR, theColor.Long( 0 ), 0 )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function BalloonStyle() As Boolean
		  return mBalloonStyle
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub BalloonStyle(assigns set as Boolean)
		  mBalloonStyle = set
		  
		  if mWnd = 0 then return
		  
		  Const TTS_BALLOON = &h40
		  ChangeWindowStyle( TTS_BALLOON, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Center() As Boolean
		  return mCenter
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Center(assigns set as Boolean)
		  mCenter = set
		  
		  // Reset the text so that you get
		  // the new style
		  Text = mText
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ChangeWindowStyle(flag as Integer, set as Boolean)
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
		    
		    oldFlags = GetWindowLong( mWnd, GWL_STYLE )
		    
		    if not set then
		      newFlags = BitwiseAnd( oldFlags, Bitwise.OnesComplement( flag ) )
		    else
		      newFlags = BitwiseOr( oldFlags, flag )
		    end
		    
		    
		    styleFlags = SetWindowLong( mWnd, GWL_STYLE, newFlags )
		    styleFlags = SetWindowPos( mWnd, 0, 0, 0, 0, 0, SWP_NOMOVE +_
		    SWP_NOSIZE + SWP_NOZORDER + SWP_FRAMECHANGED )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(c as RectControl, selfReference as Boolean = false)
		  mUnicodeSavvy = System.IsFunctionAvailable( "SendMessageW", "User32" )
		  
		  if selfReference then
		    mSelf = me
		  end
		  
		  InitProperties
		  
		  // We want to pop over this rect control's area
		  CreateToolTip( 0, 0, c.Width, c.Height, c.Handle )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(w as Window, selfReference as Boolean = false)
		  mUnicodeSavvy = System.IsFunctionAvailable( "SendMessageW", "User32" )
		  
		  if selfReference then
		    mSelf = me
		  end
		  
		  InitProperties
		  // We want to pop over this window's client area
		  'CreateToolTip( 0, 0, w.Width, w.Height, w.WinHWND )
		  CreateToolTip( 0, 0, w.Width, w.Height, w.Handle )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(w as Window, left as Integer, top as Integer, width as Integer, height as Integer, selfReference as Boolean = false)
		  mUnicodeSavvy = System.IsFunctionAvailable( "SendMessageW", "User32" )
		  
		  if selfReference then
		    mSelf = me
		  end
		  
		  InitProperties
		  
		  // We are given the rect in of where
		  // we want the tooltip to pop up at relative to the window
		  // we are passed in.  If w is nil, then we assume screen
		  // coordinates
		  if w <> nil then
		    'CreateToolTip( left, top, width, height, w.WinHWND )
		    CreateToolTip( left, top, width, height, w.Handle )
		  else
		    #if TargetWin32
		      Declare Function GetDesktopWindow Lib "User32" () as Integer
		      CreateToolTip( left, top, width, height, GetDesktopWindow )
		    #endif
		  end
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub CreateToolTip(left as Integer, top as Integer, width as Integer, height as Integer, relativeTo as Integer)
		  mControlHandle = relativeTo
		  mLeft = left
		  mTop = top
		  mWidth = width
		  mHeight = height
		  
		  #if TargetWin32
		    // We want to make the tooltip window with all the
		    // proper styles set
		    Soft Declare Function CreateWindowExA Lib "User32" ( _
		    dwExStyle as Integer, lpClassName as CString, lpWindowName as Integer, _
		    dwStyle as Integer, x as Integer, y as Integer, width as Integer, height as Integer, _
		    hWndParent as Integer, hMenu as Integer, hInstance as Integer, lpParam as Integer ) as Integer
		    Soft Declare Function CreateWindowExW Lib "User32" ( _
		    dwExStyle as Integer, lpClassName as WString, lpWindowName as Integer, _
		    dwStyle as Integer, x as Integer, y as Integer, width as Integer, height as Integer, _
		    hWndParent as Integer, hMenu as Integer, hInstance as Integer, lpParam as Integer ) as Integer
		    
		    Declare Function GetModuleHandleA Lib "Kernel32" ( name as Integer ) as Integer
		    Declare Function GetParent Lib "User32" ( hwnd as Integer ) as Integer
		    
		    mParentHandle = GetParent( mControlHandle )
		    mInstance = GetModuleHandleA( 0 )
		    
		    // Create the tooltip window
		    Const WS_EX_TOPMOST = &h8
		    Const WS_POPUP = &H80000000
		    Const CW_USEDEFAULT = &H80000000
		    
		    if mUnicodeSavvy then
		      mWnd = CreateWindowExW( WS_EX_TOPMOST, "tooltips_class32", 0, _
		      WS_POPUP, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, _
		      CW_USEDEFAULT, mControlHandle, 0, mInstance, 0 )
		    else
		      mWnd = CreateWindowExA( WS_EX_TOPMOST, "tooltips_class32", 0, _
		      WS_POPUP, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, _
		      CW_USEDEFAULT, mControlHandle, 0, mInstance, 0 )
		    end if
		    
		    MakeTopmost
		    
		    // Now that we have a window, we can set all the
		    // user's styles properly
		    SetAllUserStyles
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Destructor()
		  #if TargetWin32
		    Declare Sub DestroyWindow Lib "User32" ( hwnd as Integer )
		    
		    if mWnd <> 0 then
		      DestroyWindow( mWnd )
		    end
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DisplayTime() As Integer
		  return mDisplayTime
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DisplayTime(assigns time as Integer)
		  mDisplayTime = time
		  
		  if mWnd = 0 then return
		  
		  Const TTDT_AUTOPOP = 2
		  SetTime( TTDT_AUTOPOP, time )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Fade() As Boolean
		  return mFade
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Fade(assigns set as Boolean)
		  mFade = set
		  
		  if mWnd = 0 then return
		  
		  Const TTS_NOFADE = &h20
		  ChangeWindowStyle( TTS_NOFADE, not set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Hide()
		  Dim TTM_POP as Integer = WM_USER + 28
		  Call SendMessage( TTM_POP, 0, 0 )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HidePrefix() As Boolean
		  Return mHidePrefix
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub HidePrefix(assigns set as Boolean)
		  // Set the prefix
		  mHidePrefix = set
		  
		  // If we're not created yet, bail out
		  if mWnd = 0 then Return
		  
		  // We want to set the new style for the window
		  Const TTS_NOPREFIX = &h2
		  ChangeWindowStyle( TTS_NOPREFIX, set )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Icon() As Integer
		  return mIcon
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Icon(assigns theIcon as Integer)
		  mIcon = theIcon
		  
		  if mWnd = 0 then return
		  
		  // Re-set the title to change the icon
		  Title = mTitle
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function InitialShowTime() As Integer
		  return mInitialShowTime
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub InitialShowTime(assigns time as Integer)
		  mInitialShowTime = time
		  
		  if mWnd = 0 then return
		  
		  Const TTDT_INITIAL = 3
		  SetTime( TTDT_INITIAL, time )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub InitProperties()
		  mFade = true
		  mHidePrefix = true
		  mAlwaysTip = true
		  mAnimate = true
		  mBalloonStyle = false
		  mText = ""
		  mInitialShowTime = -1
		  mDisplayTime = -1
		  mReshowTime = -1
		  mIcon = 0
		  mTitle = ""
		  
		  // This is also a good place to initialize the common
		  // control library as well
		  #if TargetWin32
		    Declare Sub InitCommonControlsEx Lib "ComCtl32" ( ex as Ptr )
		    
		    Dim mb as new MemoryBlock( 8 )
		    mb.Long( 0 ) = mb.Size
		    Const ICC_WIN95_CLASSES = &hFF
		    mb.Long( 4 ) = ICC_WIN95_CLASSES
		    
		    InitCommonControlsEx( mb )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IntegerToMemoryBlock(i as Integer) As MemoryBlock
		  dim mb as new MemoryBlock( 4 )
		  mb.Long( 0 ) = i
		  return mb.Ptr( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MAKELONG(low as Integer, high as Integer) As Integer
		  return Bitwise.ShiftLeft( high, 16 ) + low
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub MakeTopmost()
		  #if TargetWin32
		    // Make sure it shows up as topmost
		    Declare Sub SetWindowPos Lib "User32" ( hwnd as Integer, _
		    after as Integer, x as Integer, y as Integer, width as Integer, _
		    height as Integer, flags as Integer )
		    
		    Const HWND_TOPMOST = -1
		    Const SWP_NOSIZE = &h1
		    Const SWP_NOMOVE = &h2
		    Const SWP_NOACTIVATE = &h10
		    
		    SetWindowPos( mWnd, HWND_TOPMOST, 0, 0, 0, 0, _
		    SWP_NOSIZE + SWP_NOMOVE + SWP_NOACTIVATE )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function MaxWidth() As Integer
		  Dim TTM_GETMAXTIPWIDTH as Integer = WM_USER + 25
		  
		  return SendMessage( TTM_GETMAXTIPWIDTH, 0, 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub MaxWidth(assigns width as Integer)
		  Dim TTM_SETMAXTIPWIDTH as Integer = WM_USER + 24
		  Call SendMessage( TTM_SETMAXTIPWIDTH, 0, width )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function MemoryBlockToInteger(mb as MemoryBlock) As Integer
		  dim ret as new MemoryBlock( 4 )
		  ret.Ptr( 0 ) = mb
		  return ret.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ReshowTime() As Integer
		  return mReshowTime
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ReshowTime(assigns time as Integer)
		  mReshowTime = time
		  
		  if mWnd = 0 then return
		  
		  Const TTDT_RESHOW = 1
		  SetTime( TTDT_RESHOW, time )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function SendMessage(msg as Integer, wParam as Integer, lParam as Integer) As Integer
		  #if TargetWin32
		    Soft Declare Function SendMessageA Lib "User32" ( wnd as Integer, msg as Integer, _
		    wParam as Integer, lParam as Integer ) as Integer
		    Soft Declare Function SendMessageW Lib "User32" ( wnd as Integer, msg as Integer, _
		    wParam as Integer, lParam as Integer ) as Integer
		    
		    if mUnicodeSavvy then
		      return SendMessageW( mWnd, msg, wParam, lParam )
		    else
		      return SendMessageA( mWnd, msg, wParam, lParam )
		    end if
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetActiveState(set as Boolean)
		  Dim TTM_ACTIVATE as Integer = WM_USER + 1
		  dim setInt as Integer
		  
		  if set then setInt = 1
		  Call SendMessage( TTM_ACTIVATE, setInt, 0 )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub SetAllUserStyles()
		  AlwaysTip = mAlwaysTip
		  Animate = mAnimate
		  BalloonStyle = mBalloonStyle
		  Fade = mFade
		  HidePrefix = mHidePrefix
		  if mText <> "" then
		    Text = mText
		  end
		  
		  // This seems like a reasonable default to me...
		  MaxWidth = 256
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub SetAutoTime(time as Integer)
		  // Automatically sets all the times proportional
		  // to one another
		  Const TTDT_AUTOMATIC = 0
		  SetTime( TTDT_AUTOMATIC, time )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub SetTime(field as Integer, time as Integer)
		  Dim TTM_SETDELAYTIME as Integer = WM_USER + 3
		  Call SendMessage( TTM_SETDELAYTIME, field, MAKELONG( time, 0 ) )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Show()
		  Dim TTM_POPUP as Integer = WM_USER + 34
		  call SendMessage( TTM_POPUP, 0, 0 )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Text() As String
		  return mText
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Text(assigns theText as String)
		  mText = theText
		  
		  if mWnd = 0 then return
		  
		  // If we already have a tooltip, then we need
		  // to delete it.
		  Dim TTM_DELTOOLA as Integer = WM_USER + 5
		  Dim TTM_DELTOOLW as Integer = WM_USER + 51
		  Dim size as Integer = 44
		  // Add this back in if supporting bitmaps sometime
		  'if OSVersionInformation.IsNT5 then
		  'size = size + 4
		  'end
		  
		  Dim toolinfo as new MemoryBlock( size )
		  Const TTF_IDISHWND = &h1
		  
		  // We need to remove the tool
		  toolinfo.Long( 0 ) = toolinfo.Size
		  toolinfo.Long( 4 ) = TTF_IDISHWND
		  toolinfo.Long( 8 ) = mParentHandle
		  toolinfo.Long( 12 ) = mControlHandle
		  
		  dim ret as Integer
		  dim whatMsg as Integer = TTM_DELTOOLA
		  if mUnicodeSavvy then whatMsg = TTM_DELTOOLW
		  ret = SendMessage( whatMsg, 0, MemoryBlockToInteger( toolinfo ) )
		  
		  if theText = "" then return
		  
		  // We want to add a tool mesage to the window now
		  Dim TTM_ADDTOOLA as Integer = WM_USER + 4
		  Dim TTM_ADDTOOLW as Integer = WM_USER + 50
		  
		  Const TTF_SUBCLASS = &h10
		  Const TTF_CENTERTIP = &h2
		  
		  toolinfo.Long( 0 ) = toolinfo.Size
		  toolinfo.Long( 4 ) = TTF_IDISHWND + TTF_SUBCLASS
		  
		  if mCenter then
		    toolinfo.Long( 4 ) = toolinfo.Long( 4 ) + TTF_CENTERTIP
		  end
		  
		  toolinfo.Long( 8 ) = mParentHandle
		  toolinfo.Long( 12 ) = mControlHandle
		  toolinfo.Long( 16 ) = mLeft
		  toolinfo.Long( 20 ) = mTop
		  toolinfo.Long( 24 ) = mLeft + mWidth
		  toolinfo.Long( 28 ) = mTop + mHeight
		  toolinfo.Long( 32 ) = mInstance
		  
		  dim textPtr as MemoryBlock
		  if mUnicodeSavvy then
		    whatMsg = TTM_ADDTOOLW
		    textPtr = new MemoryBlock( (Len( theText ) + 1)  * 2 )
		    textPtr.WString( 0 ) = theText
		  else
		    whatMsg = TTM_ADDTOOLA
		    textPtr = new MemoryBlock( Len( theText ) + 1 )
		    textPtr.CString( 0 ) = theText
		  end if
		  toolinfo.Ptr( 36 ) = textPtr
		  
		  ret = SendMessage( whatMsg, 0, MemoryBlockToInteger( toolinfo ) )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TextColor() As Color
		  if mWnd = 0 then return RGB( 0, 0, 0 )
		  
		  Dim TTM_GETTIPTEXTCOLOR as Integer = WM_USER + 23
		  
		  dim theColor as Integer
		  theColor = SendMessage( TTM_GETTIPTEXTCOLOR, 0, 0 )
		  
		  dim mb as new MemoryBlock( 4 )
		  mb.Long( 0 ) = theColor
		  
		  return RGB( mb.Byte( 0 ), mb.Byte( 1 ), mb.Byte( 2 ) )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub TextColor(assigns c as Color)
		  if mWnd = 0 then return
		  
		  Dim TTM_SETTIPTEXTCOLOR as Integer = WM_USER + 20
		  
		  dim theColor as new MemoryBlock( 4 )
		  theColor.Byte( 0 ) = c.red
		  theColor.Byte( 1 ) = c.green
		  theColor.Byte( 2 ) = c.blue
		  
		  Call SendMessage( TTM_SETTIPTEXTCOLOR, theColor.Long( 0 ), 0 )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Title() As String
		  return mTitle
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Title(assigns theTitle as String)
		  if Len( theTitle ) > 100 then
		    // We can't allow this, so raise an exception
		    Raise new UnsupportedFormatException
		    return
		  end
		  
		  mTitle = theTitle
		  
		  if mWnd = 0 then return
		  
		  Dim TTM_SETTITLEA as Integer = WM_USER + 32
		  Dim TTM_SETTITLEW as Integer = WM_USER + 33
		  
		  dim whatMsg as Integer
		  dim theTitlePtr as MemoryBlock
		  
		  if mUnicodeSavvy then
		    whatMsg = TTM_SETTITLEW
		    theTitlePtr = new MemoryBlock( (Len( theTitle ) + 1) * 2 )
		    theTitlePtr.WString( 0 ) = theTitle
		  else
		    whatMsg = TTM_SETTITLEA
		    theTitlePtr = new MemoryBlock( Len( theTitle ) + 1 )
		    theTitlePtr.CString( 0 ) = theTitle
		  end if
		  
		  Call SendMessage( whatMsg, mIcon, MemoryBlockToInteger( theTitlePtr ) )
		End Sub
	#tag EndMethod


	#tag Note, Name = Class Dependencies
		This class relies on the following things:
		
		OSVersionInformation module
	#tag EndNote

	#tag Note, Name = Class Notes
		All time specifiers are in milliseconds.  If the time is 
		at its default value, it will return -1.  If you want to specify
		the time should go back to its default value, then set
		it to be -1.
		
		All UI values are given in pixels for things like MaxWidth, left, right, 
		width, height, etc.
	#tag EndNote


	#tag Property, Flags = &h1
		Protected mAlwaysTip As Boolean
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mAnimate As Boolean
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mBalloonStyle As Boolean
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mCenter As Boolean
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mControlHandle As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mDisplayTime As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mFade As Boolean
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mHidePrefix As Boolean
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mIcon As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mID As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mInitialShowTime As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mInstance As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mLeft As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mParentHandle As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mReshowTime As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mSelf As WFSToolTip
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mText As String
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mTitle As String
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mTop As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mUnicodeSavvy As Boolean
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mWnd As Integer
	#tag EndProperty


	#tag Constant, Name = kErrorIcon, Type = Integer, Dynamic = False, Default = \"3", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kInfoIcon, Type = Integer, Dynamic = False, Default = \"1", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kNoIcon, Type = Integer, Dynamic = False, Default = \"0", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kWarningIcon, Type = Integer, Dynamic = False, Default = \"2", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = WM_USER, Type = Integer, Dynamic = False, Default = \"&h400", Scope = Private
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
End Class
#tag EndClass
