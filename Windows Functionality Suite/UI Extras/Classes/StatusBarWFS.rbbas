#tag Class
Protected Class StatusBarWFS
Implements WndProcSubclassWFS
	#tag Method, Flags = &h0
		Sub AppendPart(width as Integer, percent as Boolean = false)
		  // Append the part.  If we're a percentage, make the width
		  // negative.  We will account for this later.
		  if percent then
		    mParts.Append( -width )
		  else
		    mParts.Append( width )
		  end if
		  
		  SetupParts
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(w as Integer)
		  ConstructorHelper( w, true )
		  
		  // Now we need to subclass the window
		  WndProcHelpersWFS.Subclass( w, me )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(w as MDIWindow)
		  ConstructorHelper( w.MDIClientHandleWFS, true )
		  
		  // Now we need to subclass the window
		  WndProcHelpersWFS.Subclass( w.MDIClientHandleWFS, me )
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(owner as Window)
		  ConstructorHelper( owner.Handle, owner.GrowIcon )
		  
		  // Now we need to subclass the window
		  WndProcHelpersWFS.Subclass( owner, me )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ConstructorHelper(owner as Integer, growIcon as Boolean)
		  if owner = 0 then return
		  mOwner = owner
		  
		  #if TargetWin32
		    
		    Declare Sub InitCommonControlsEx Lib "ComCtl32" ( mb as Ptr )
		    Declare Function GetModuleHandleA Lib "Kernel32" ( zero as Integer ) as Integer
		    
		    dim mb as new MemoryBlock( 8 )
		    mb.Long( 0 ) = 8
		    mb.Long( 4 ) = &hFF
		    
		    InitCommonControlsEx( mb )
		    
		    Soft Declare Function CreateWindowExW Lib "User32" ( exStyle as Integer, _
		    className as WString, windowName as Integer, style as Integer, x as Integer, _
		    y as Integer, cx as Integer, cy as Integer, parent as Integer, menu as Integer, _
		    hInstance as Integer, lParam as Integer ) as Integer
		    
		    Soft Declare Function CreateWindowExA Lib "User32" ( exStyle as Integer, _
		    className as CString, windowName as Integer, style as Integer, x as Integer, _
		    y as Integer, cx as Integer, cy as Integer, parent as Integer, menu as Integer, _
		    hInstance as Integer, lParam as Integer ) as Integer
		    
		    static sNextID as Integer = 5000
		    mID = sNextID
		    sNextID = sNextID + 1
		    
		    Const SB_SETUNICODEFORMAT = &h2005
		    Const SBT_TOOLTIPS = &h800
		    
		    dim style as Integer
		    Const WS_CHILD = &h40000000
		    Const WS_VISIBLE = &h10000000
		    style = WS_CHILD + WS_VISIBLE + SBT_TOOLTIPS
		    if growIcon then
		      style = style + &h100
		    end
		    
		    dim inst as Integer = GetModuleHandleA( 0 )
		    
		    if System.IsFunctionAvailable( "CreateWindowExW", "User32" ) then
		      mUnicodeSavvy = true
		      
		      hWnd = CreateWindowExW( 0, "msctls_statusbar32", 0, style, 0, 0, _
		      OwnerWidth, 20, owner, mID, inst, 0 )
		      SendMessage( SB_SETUNICODEFORMAT, 1, nil )
		    else
		      mUnicodeSavvy = false
		      
		      hWnd = CreateWindowExA( 0, "msctls_statusbar32", 0, style, 0, 0, _
		      OwnerWidth, 20, owner, mID, inst, 0 )
		      SendMessage( SB_SETUNICODEFORMAT, 0, nil )
		    end if
		    
		    '// Now we need to subclass the window
		    'WndProcHelpers.Subclass( owner, me )
		    
		  #else
		    
		    #pragma unused growIcon
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Destructor()
		  // First, unsubclass the window
		  if mOwner <> 0 then
		    WndProcHelpersWFS.Unsubclass( mOwner, me )
		  end if
		  
		  // Then, destroy our child window
		  #if TargetWin32
		    if hWnd <> 0 then
		      Declare Sub DestroyWindow Lib "User32" ( hwnd as Integer )
		      DestroyWindow( hWnd )
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Handle() As Integer
		  // ADDED BY Carlos Martinho (2006-SEP-10)
		  
		  // Returns the status bar Handle
		  Return hWnd
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function HICONFromRBPicture(theIcon as Picture) As Integer
		  #if TargetWin32
		    
		    // If there's no picture, then there's no HICON
		    if theIcon = nil then return 0
		    
		    // Create our cache if we must
		    if mIconCache = nil then mIconCache = new Dictionary
		    
		    // Check to see if the picture exists in our cache
		    if mIconCache.HasKey( theIcon ) then return mIconCache.Value( theIcon )
		    
		    #if RBVersion < 2010.05
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
		    
		    #if RBVersion < 2010.05
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
		    
		    #if RBVersion < 2010.05
		      // And unlock our picture descriptions
		      unlockPictureDescription( p.Mask )
		      unlockPictureDescription( p )
		    #endif
		    
		    // Store the value into our cache
		    mIconCache.Value( p ) = hicon
		    
		    // And return the new HICON
		    return hicon
		    
		  #else
		    
		    #pragma unused theIcon
		    
		  #endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Hide()
		  // ADDED BY Carlos Martinho (2006-SEP-10)
		  // To hide the status bar
		  
		  #if TargetWin32
		    
		    const SW_HIDE = 0
		    Declare Function ShowWindow Lib "user32" (hwnd as Integer, nCmdShow as Integer) as Integer
		    
		    dim e as integer
		    if hWnd <> 0 then
		      e = ShowWindow( hWnd, SW_HIDE )
		    end
		    
		  #endif
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsInSimpleMode() As Boolean
		  // ADDED BY Carlos Martinho (2006-SEP-10)
		  
		  // Returns if the status bar is in simple mode text
		  Return mSimpleMode
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function MAKELONG(a as Integer, b as Integer) As Integer
		  dim ret as Integer
		  
		  ret = Bitwise.BitOr( Bitwise.BitAnd( a, &hFFFF ), Bitwise.ShiftLeft( Bitwise.BitAnd( b, &hFFFF), 16 ) )
		  
		  return ret
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function OwnerHeight() As Integer
		  #if TargetWin32
		    Declare Sub GetClientRect Lib "User32" ( hwnd as Integer, mb as Ptr )
		    
		    dim mb as new MemoryBlock( 16 )
		    GetClientRect( mOwner, mb )
		    
		    return mb.Long( 12 )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function OwnerWidth() As Integer
		  #if TargetWin32
		    Declare Sub GetClientRect Lib "User32" ( hwnd as Integer, mb as Ptr )
		    
		    dim mb as new MemoryBlock( 16 )
		    GetClientRect( mOwner, mb )
		    
		    return mb.Long( 8 )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Resize()
		  #if TargetWin32
		    Declare Sub SetWindowPos Lib "User32" ( hwnd as Integer, prev as Integer, _
		    x as Integer, y as Integer, cx as Integer, cy as Integer, flags as Integer )
		    
		    dim flags as Integer
		    
		    flags = &h2 + &h4 + &h10 + &h200
		    SetWindowPos( hWnd, 0, 0, 0, _
		    OwnerWidth, OwnerHeight, flags )
		    
		    SetupParts
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub SendMessage(msg as Integer, wParam as Integer, lParam as Integer)
		  #if TargetWin32
		    
		    Soft Declare Sub SendMessageA Lib "User32" ( hwnd as Integer, msg as Integer, wParam as Integer, lParam as Integer )
		    Soft Declare Function SendMessageW Lib "User32" ( hwnd as Integer, msg as Integer, wParam as Integer, lParam as Integer ) as Integer
		    
		    if hWnd = 0 then return
		    dim ret as Integer
		    if mUnicodeSavvy then
		      ret = SendMessageW( hWnd, msg, wParam, lParam )
		    else
		      SendMessageA( hWnd, msg, wParam, lParam )
		    end if
		    
		  #else
		    
		    #pragma unused msg
		    #pragma unused wParam
		    #pragma unused lParam
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub SendMessage(msg as Integer, wParam as Integer, lParam as Ptr)
		  #if TargetWin32
		    
		    Soft Declare Sub SendMessageA Lib "User32" ( hwnd as Integer, msg as Integer, wParam as Integer, lParam as Ptr )
		    Soft Declare Function SendMessageW Lib "User32" ( hwnd as Integer, msg as Integer, wParam as Integer, lParam as Ptr ) as Integer
		    
		    if hWnd = 0 then return
		    dim ret as Integer
		    if mUnicodeSavvy then
		      ret = SendMessageW( hWnd, msg, wParam, lParam )
		    else
		      SendMessageA( hWnd, msg, wParam, lParam )
		    end if
		    
		  #else
		    
		    #pragma unused msg
		    #pragma unused wParam
		    #pragma unused lParam
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetIcon(part as Integer, theIcon as Picture)
		  Const SB_SETICON = &h40F
		  
		  dim icon as Integer
		  icon = HICONFromRBPicture( theIcon )
		  
		  SendMessage( SB_SETICON, part, icon )
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetText(str as String, partNum as Integer = 0, raisedText as Boolean = false)
		  Const SB_SETTEXTA = &h401
		  Const SB_SETTEXTW = &h40B
		  Const SBT_POPOUT = &h200
		  
		  dim thewParam as Integer = partNum
		  if raisedText then thewParam = Bitwise.BitOr( thewParam, SBT_POPOUT )
		  
		  if mUnicodeSavvy then
		    SendMessage( SB_SETTEXTW, thewParam, StrToLPStr( str ) )
		  else
		    SendMessage( SB_SETTEXTA, thewParam, StrToLPStr( str ) )
		  end if
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetTextCenter(str as String, partNum as Integer = 0, raisedText as Boolean = false)
		  // ADDED BY Carlos Martinho (2006-SEP-10) (just the part to center text)
		  
		  Const SB_SETTEXTA = &h401
		  Const SB_SETTEXTW = &h40B
		  Const SBT_POPOUT = &h200
		  
		  // To center align text: one tab before text
		  str = Chr(9) + str
		  
		  dim thewParam as Integer = partNum
		  if raisedText then thewParam = Bitwise.BitOr( thewParam, SBT_POPOUT )
		  
		  if mUnicodeSavvy then
		    SendMessage( SB_SETTEXTW, thewParam, StrToLPStr( str ) )
		  else
		    SendMessage( SB_SETTEXTA, thewParam, StrToLPStr( str ) )
		  end if
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetTextOnSimple(str as String, centerText as Boolean = false, raisedText as Boolean = false)
		  // ADDED BY Carlos Martinho (2006-SEP-10)
		  // Used only to set text on a simple mode status bar
		  // The partNum is fixed to 255 - must be 255 to set text on simple mode
		  
		  Const SB_SETTEXTA = &h401
		  Const SB_SETTEXTW = &h40B
		  Const SBT_POPOUT = &h200
		  
		  // Check if is to center text
		  If centerText Then str = Chr(9) + str
		  
		  dim thewParam as Integer = 255 // simple mode
		  if raisedText then thewParam = Bitwise.BitOr( thewParam, SBT_POPOUT )
		  
		  if mUnicodeSavvy then
		    SendMessage( SB_SETTEXTW, thewParam, StrToLPStr( str ) )
		  else
		    SendMessage( SB_SETTEXTA, thewParam, StrToLPStr( str ) )
		  end if
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetTextRight(str as String, partNum as Integer = 0, raisedText as Boolean = false)
		  // ADDED BY Carlos Martinho (2006-SEP-10) (just the part to rigth align text)
		  
		  Const SB_SETTEXTA = &h401
		  Const SB_SETTEXTW = &h40B
		  Const SBT_POPOUT = &h200
		  
		  // To right align text: two tabs before text
		  str = Chr(9) + Chr(9) + str
		  
		  dim thewParam as Integer = partNum
		  if raisedText then thewParam = Bitwise.BitOr( thewParam, SBT_POPOUT )
		  
		  if mUnicodeSavvy then
		    SendMessage( SB_SETTEXTW, thewParam, StrToLPStr( str ) )
		  else
		    SendMessage( SB_SETTEXTA, thewParam, StrToLPStr( str ) )
		  end if
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub SetupParts()
		  // We need to convert the parts into an
		  // array that specifies the right edges of the dividers
		  dim rightEdge, i as Integer
		  dim divArray as new MemoryBlock( (UBound( mParts ) + 1) * 4 )
		  for i = 0 to UBound( mParts )
		    // Figure out the right edge of this divider.  If the part
		    // is a percentage, then we want to handle it special
		    if mParts( i ) < 0 then
		      // Figure this egde out based on the width of the window
		      // minus what we already have for parts laid out.  Then
		      // do a percentage of that
		      dim availWidth as Integer = OwnerWidth - rightEdge
		      rightEdge = rightEdge + (availWidth * (-mParts( i ) / 100.0))
		    elseif mParts( i ) = kRemainder then
		      // If the part is 0, then we should take up the remainder
		      // of the space.  Make sure that we take the grow icon into
		      // account though
		      rightEdge = -1
		    else
		      rightEdge = rightEdge + mParts( i )
		    end if
		    
		    // Add it to our array
		    divArray.Long( i * 4 ) = rightEdge
		  next i
		  
		  // Now set the parts
		  Const SB_SETPARTS = &h404
		  SendMessage( SB_SETPARTS, UBound( mParts ) + 1, divArray )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Show()
		  // ADDED BY Carlos Martinho (2006-SEP-10)
		  // To show the status bar
		  
		  #if TargetWin32
		    
		    const SW_NORMAL = 1
		    
		    Declare Function ShowWindow Lib "user32" (hwnd as Integer, nCmdShow as Integer) as Integer
		    
		    dim e as integer
		    if hWnd <> 0 then
		      e = ShowWindow( hWnd, SW_NORMAL )
		    end
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SimpleMode(simple As Boolean = True)
		  // ADDED BY Carlos Martinho (2006-SEP-10)
		  
		  Const WM_USER = &H400
		  
		  'The SB_SIMPLE message specifies whether a status window
		  'displays simple text or displays all window parts set by
		  'a previous SB_SETPARTS message. The string that a status
		  'bar displays in simple mode is maintained separately
		  'from the strings it displays when it is in multiple-part mode
		  'wParam = Display type flag. If this parameter is TRUE, the window
		  'displays simple text. If it is FALSE, it displays multiple parts.
		  'lParam = 0;
		  'Returns FALSE if an error occurs.
		  'If the status window is being changed from non-simple to simple,
		  'or vice versa, the window is immediately redrawn.
		  'Const SB_SIMPLE = (WM_USER + 9)
		  Dim SB_SIMPLE As Integer
		  SB_SIMPLE = WM_USER + 9
		  
		  'lResult = SendMessage(      // returns LRESULT in lResult
		  '     (HWND) hWndControl,    // handle to destination control
		  '     (UINT) SB_SIMPLE,      // message ID
		  '     (WPARAM) wParam,       // = (WPARAM) (BOOL) fSimple;
		  '     (LPARAM) lParam        // = 0; not used, must be zero );
		  
		  'Parameters
		  ' fSimple: Display type flag. If this parameter is TRUE, the window displays simple text. If it is FALSE, it displays multiple parts.
		  ' lParam: Must be zero.
		  
		  'Return Value
		  ' The return value is not used.
		  
		  
		  If simple Then
		    SendMessage( SB_SIMPLE, 1, 0)
		  Else
		    SendMessage( SB_SIMPLE, 0, 0)
		  End If
		  
		  // Set the mSimpleMode
		  mSimpleMode = simple
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function StrToLPStr(str as String) As MemoryBlock
		  dim strPtr as new MemoryBlock( 1024 )
		  
		  if mUnicodeSavvy then
		    strPtr.WString( 0 ) = ConvertEncoding( str, Encodings.UTF16 )
		  else
		    strPtr.CString( 0 ) = str
		  end if
		  
		  return strPtr
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub TipText(part as Integer = 0, assigns text as String)
		  Const SB_SETTIPTEXTA = &h410
		  Const SB_SETTIPTEXTW = &h411
		  
		  if mUnicodeSavvy then
		    SendMessage( SB_SETTIPTEXTW, part, StrToLPStr( text ) )
		  else
		    SendMessage( SB_SETTIPTEXTA, part, StrToLPStr( text ) )
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TipText(part as Integer) As String
		  Const SB_GETTIPTEXTA = &h412
		  Const SB_GETTIPTEXTW = &h413
		  
		  dim textBuf as new MemoryBlock( 2048 )
		  dim wParam as Integer = MAKELONG( part, textBuf.Size )
		  
		  if mUnicodeSavvy then
		    SendMessage( SB_GETTIPTEXTW, wParam, textBuf )
		    return textBuf.WString( 0 )
		  else
		    SendMessage( SB_GETTIPTEXTA, wParam, textBuf )
		    return textBuf.CString( 0 )
		  end if
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function WndProc(hWnd as Integer, msg as Integer, wParam as Integer, lParam as Integer, ByRef returnValue as Integer) As Boolean
		  // We want to capture the WM_SIZE and WM_SIZING messages
		  // so that we can resize without the window needing to do anything
		  // to cause us to take note.
		  Const WM_SIZE = &h5
		  Const WM_SIZING = &h214
		  Const WM_NCHITTEST = &h84
		  
		  select case msg
		  case WM_SIZE, WM_NCHITTEST
		    // We just forward these messages on to the status bar itself so that
		    // it can properly resize itself.
		    SendMessage( msg, wParam, lParam )
		  end select
		  
		  // Note that we return false always, that's because we're not
		  // changing anything -- just peeking at it
		  return false
		  
		  #pragma unused hWnd
		  #pragma unused returnValue // Was this an oversight? -KT
		End Function
	#tag EndMethod


	#tag Note, Name = Known Problems
		By Carlos Martinho (2006-SEP-10)
		
		Icon Masks
		The icon draws correctly with masks if:
		1. the bmp image has a black background and IF the mask is draw first:
		iconMask + iconWithBlackBg
		This works fine with XP and Win2K
		2. the bmp file has an Alpha Transparency and it's not necessary to apply any mask
		This only works fine with XP - on Win2K is wrong (there's a black bg)
		
		
		Resizing Problems
		When there's one part that is a percentage width, that part is not resized when
		the window is resized. The workaround for this is to use on windows' resizing
		and resized events:
		// We need to call the Resize in order to get the StatusBar parts (that uses a
		// percentage width) resized
		If mBar <> Nil Then mBar.Resize
	#tag EndNote

	#tag Note, Name = Modifications by Carlos Martinho
		Added on 2006-SEP-10
		
		- SimpleMode(simple As Boolean = True)
		  Specifies whether the status bar displays simple text or displays
		  all window parts. Setting True (default) the status bar displays
		  simple text. Setting False, it displays multiple parts.
		
		  The string that a status bar displays in simple mode is maintained
		  separately from the strings it displays when it is in multiple-part mode.
		
		  AFAIK it is not possible to display an icon on simple mode.
		
		
		- IsInSimpleMode() As Boolean
		  Returns if the status bar is in Simple Mode.
		  When it returns False it means is in multiple-part mode.
		
		
		- SetTextCenter( str as String, partNum as Integer = 0, raisedText as Boolean = false )
		  Aligns text on center
		
		- SetTextRight( str as String, partNum as Integer = 0, raisedText as Boolean = false )
		  Aligns text on right
		
		- SetTextOnSimple( str as String, centerText as Boolean = false, raisedText as Boolean = false )
		  To use ONLY when setting text on a simple mode status bar.
		  If centerText is true, the text is centered on the status bar.
		
		
		- Hide()
		  Hides the status bar
		
		- Show()
		  Shows the status bar
		
		
		- Handle() As Integer
		  Returns the status bar Handle.
		  Useful if you want to attach a control, like a progress bar ;-)
	#tag EndNote

	#tag Note, Name = Special note about MDIWindows
		If you are going to use the StatusBar class in an MDIWindow, you must
		manually call the destructor method from the App.Close event.  Failing to
		do so can cause crashes due to the window subclassing of the MDICLIENT
		window.
	#tag EndNote


	#tag Property, Flags = &h1
		Protected hWnd As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIconCache As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mID As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mOwner As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mParts(-1) As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		#tag Note
			// ADDED BY Carlos Martinho (2006-SEP-10)
			Set by the SimpleMode sub
		#tag EndNote
		Private mSimpleMode As Boolean
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mUnicodeSavvy As Boolean
	#tag EndProperty


	#tag Constant, Name = kRemainder, Type = Integer, Dynamic = False, Default = \"0", Scope = Public
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
