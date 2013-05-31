#tag Module
Protected Module SystemMetrics
	#tag Method, Flags = &h1
		Protected Function CaptionBarButtonHeight() As Integer
		  Initialize
		  
		  return mCaptionBarButtonHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function CaptionBarButtonWidth() As Integer
		  Initialize
		  
		  return mCaptionBarButtonWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function CaptionHeight() As Integer
		  Initialize
		  
		  return mCaptionHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function CleanBoot() As Boolean
		  Initialize
		  
		  return mCleanBoot
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function CursorHeight() As Integer
		  Initialize
		  
		  return mCursorHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function CursorWidth() As Integer
		  Initialize
		  
		  return mCursorWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DebugUser() As Boolean
		  Initialize
		  
		  return mDebugUser
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DialogBorderHeight() As Integer
		  Initialize
		  
		  return mDialogBorderHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DialogBorderWidth() As Integer
		  Initialize
		  
		  return mDialogBorderWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DoubleByteEnabled() As Boolean
		  Initialize
		  
		  return mDoubleByteEnabled
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DoubleClickHeight() As Integer
		  Initialize
		  
		  return mDoubeClickHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DoubleClickWidth() As Integer
		  Initialize
		  
		  return mDoubeClickWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DragHeight() As Integer
		  Initialize
		  
		  return mDragHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DragWidth() As Integer
		  Initialize
		  
		  return mDragWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FocusBorderHeight() As Integer
		  Initialize
		  
		  return mFocusBorderHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FocusBorderWidth() As Integer
		  Initialize
		  
		  return mFocusBorderWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FullScreenHeight() As Integer
		  Initialize
		  
		  return mFullScreenHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FullScreenWidth() As Integer
		  Initialize
		  
		  return mFullScreenWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function HorizontalMouseWheelPresent() As Boolean
		  Initialize
		  
		  return mHorizMouseWheelPresent
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function HorizontalScrollbarHeight() As Integer
		  Initialize
		  
		  return mHorizontalScrollbarHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function HorizontalScrollThumbWidth() As Integer
		  Initialize
		  
		  return mHorizontalScrollThumbWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IconHeight() As Integer
		  Initialize
		  return mIconHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IconSpacingHeight() As Integer
		  Initialize
		  
		  return mIconSpacingHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IconSpacingWidth() As Integer
		  Initialize
		  
		  return mIconSpacingWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IconWidth() As Integer
		  Initialize
		  
		  return mIconWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IMEEnabled() As Boolean
		  Initialize
		  
		  return mIMEEnabled
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub Initialize()
		  if mInitialized then return
		  mInitialized = true
		  
		  #if TargetWin32
		    Declare Function GetSystemMetrics Lib "user32"  (nIndex As integer) As integer
		    
		    Const SM_CXSCREEN = 0
		    Const SM_CYSCREEN = 1
		    mScreenCX = GetSystemMetrics( SM_CXSCREEN )
		    mScreenCY = GetSystemMetrics( SM_CYSCREEN )
		    
		    Const SM_CXVSCROLL = 2
		    mVerticalScrollbarWidth = GetSystemMetrics( SM_CXVSCROLL )
		    
		    Const SM_CYHSCROLL = 3
		    mHorizontalScrollbarHeight = GetSystemMetrics( SM_CYHSCROLL )
		    
		    Const SM_CYCAPTION = 4
		    mCaptionHeight = GetSystemMetrics( SM_CYCAPTION )
		    
		    Const SM_CXBORDER = 5
		    Const SM_CYBORDER = 6
		    mWindowBorderWidth = GetSystemMetrics( SM_CXBORDER )
		    mWindowBorderHeight = GetSystemMetrics( SM_CYBORDER )
		    
		    Const SM_CXDLGFRAME = 7
		    Const SM_CYDLGFRAME = 8
		    mDialogBorderWidth = GetSystemMetrics( SM_CXDLGFRAME )
		    mDialogBorderHeight = GetSystemMetrics( SM_CYDLGFRAME )
		    
		    Const SM_CYVTHUMB = 9
		    mVerticalScrollThumbHeight = GetSystemMetrics( SM_CYVTHUMB )
		    
		    Const SM_CXHTHUMB = 10
		    mHorizontalScrollThumbWidth = GetSystemMetrics( SM_CXHTHUMB )
		    
		    Const SM_CXICON = 11
		    Const SM_CYICON = 12
		    mIconWidth = GetSystemMetrics( SM_CXICON )
		    mIconHeight = GetSystemMetrics( SM_CYICON )
		    
		    Const SM_CXCURSOR = 13
		    Const SM_CYCURSOR = 14
		    mCursorWidth = GetSystemMetrics( SM_CXCURSOR )
		    mCursorHeight = GetSystemMetrics(SM_CYCURSOR )
		    
		    Const SM_CYMENU = 15
		    mMenuBarHeight = GetSystemMetrics( SM_CYMENU )
		    
		    Const SM_CXFULLSCREEN = 16
		    Const SM_CYFULLSCREEN = 17
		    mFullScreenWidth = GetSystemMetrics( SM_CXFULLSCREEN )
		    mFullScreenHeight = GetSystemMetrics( SM_CYFULLSCREEN )
		    
		    Const SM_CYKANJIWINDOW = 18
		    mKanjiWindowHeight = GetSystemMetrics( SM_CYKANJIWINDOW )
		    
		    Const SM_MOUSEPRESENT = 19
		    mMousePresent = (GetSystemMetrics( SM_MOUSEPRESENT ) <> 0)
		    
		    Const SM_CYVSCROLL = 20
		    mScrollbarArrowHeight = GetSystemMetrics( SM_CYVSCROLL )
		    
		    Const SM_CXHSCROLL = 21
		    mScrollbarArrowWidth = GetSystemMetrics( SM_CXHSCROLL )
		    
		    Const SM_DEBUG = 22
		    mDebugUser = (GetSystemMetrics( SM_DEBUG ) <> 0)
		    
		    Const SM_SWAPBUTTON = 23
		    mSwapButtons = (GetSystemMetrics( SM_SWAPBUTTON ) <> 0)
		    
		    Const SM_CXMIN = 28
		    Const SM_CYMIN = 29
		    mMinWindowWidth = GetSystemMetrics( SM_CXMIN )
		    mMinWindowHeight = GetSystemMetrics( SM_CYMIN )
		    
		    Const SM_CXSIZE = 30
		    Const SM_CYSIZE = 31
		    mCaptionBarButtonWidth = GetSystemMetrics( SM_CXSIZE )
		    mCaptionBarButtonHeight = GetSystemMetrics( SM_CYSIZE )
		    
		    Const SM_CXFRAME = 32
		    Const SM_CYFRAME = 33
		    mWindowSizingBorderWidth = GetSystemMetrics( SM_CXFRAME )
		    mWindowSizingBorderHeight = GetSystemMetrics( SM_CYFRAME )
		    
		    Const SM_CXMINTRACK = 34
		    Const SM_CYMINTRACK = 35
		    mMinWindowTrackingWidth = GetSystemMetrics( SM_CXMINTRACK )
		    mMinWindowTrackingHeight = GetSystemMetrics( SM_CYMINTRACK )
		    
		    Const SM_CXDOUBLECLK = 36
		    Const SM_CYDOUBLECLK = 37
		    mDoubeClickWidth = GetSystemMetrics( SM_CXDOUBLECLK )
		    mDoubeClickHeight = GetSystemMetrics( SM_CYDOUBLECLK )
		    
		    Const SM_CXICONSPACING = 38
		    Const SM_CYICONSPACING = 39
		    mIconSpacingWidth = GetSystemMetrics( SM_CXICONSPACING )
		    mIconSpacingHeight = GetSystemMetrics( SM_CYICONSPACING )
		    
		    Const SM_MENUDROPALIGNMENT = 40
		    mRightAlignMenuDrops = (GetSystemMetrics( SM_MENUDROPALIGNMENT ) <> 0)
		    
		    Const SM_PENWINDOWS = 41
		    mPenWindowsInstalled = (GetSystemMetrics( SM_PENWINDOWS ) <> 0)
		    
		    Const SM_DBCSENABLED = 42
		    mDoubleByteEnabled = (GetSystemMetrics( SM_DBCSENABLED ) <> 0)
		    
		    Const SM_CMOUSEBUTTONS = 43
		    mNumMouseButtons = GetSystemMetrics( SM_CMOUSEBUTTONS )
		    
		    Const SM_SECURE = 44
		    mSecurityPresent = (GetSystemMetrics( SM_SECURE ) <> 0)
		    
		    Const SM_CXEDGE = 45
		    Const SM_CYEDGE = 46
		    m3DBorderWidth = GetSystemMetrics( SM_CXEDGE )
		    m3DBorderHeight = GetSystemMetrics( SM_CYEDGE )
		    
		    Const SM_CXMINSPACING = 47
		    Const SM_CYMINSPACING = 48
		    mMinimizedWindowSpacingWidth = GetSystemMetrics( SM_CXMINSPACING )
		    mMinimizedWindowSpacingHeight = GetSystemMetrics( SM_CYMINSPACING )
		    
		    Const SM_CXSMICON = 49
		    Const SM_CYSMICON = 50
		    mSmallIconWidth = GetSystemMetrics( SM_CXSMICON )
		    mSmallIconHeight = GetSystemMetrics( SM_CYSMICON )
		    
		    Const SM_CYSMCAPTION = 51
		    mSmallCaptionHeight = GetSystemMetrics( SM_CYSMCAPTION )
		    
		    Const SM_CXSMSIZE = 52
		    Const SM_CYSMSIZE = 53
		    mSmallCaptionButtonWidth = GetSystemMetrics( SM_CXSMSIZE )
		    mSmallCaptionButtonHeight = GetSystemMetrics( SM_CYSMSIZE )
		    
		    Const SM_CXMENUSIZE = 54
		    Const SM_CYMENUSIZE = 55
		    mMenuBarButtonWidth = GetSystemMetrics(SM_CXMENUSIZE )
		    mMenuBarButtonHeight = GetSystemMetrics( SM_CYMENUSIZE )
		    
		    Const SM_CXMINIMIZED = 57
		    Const SM_CYMINIMIZED = 58
		    mMinimizedWindowWidth = GetSystemMetrics( SM_CXMINIMIZED )
		    mMinimizedWindowHeight = GetSystemMetrics( SM_CYMINIMIZED )
		    
		    Const SM_CXMAXTRACK = 59
		    Const SM_CYMAXTRACK = 60
		    mMaxTrackingWidth = GetSystemMetrics( SM_CXMAXTRACK )
		    mMaxTrackingHeight = GetSystemMetrics( SM_CYMAXTRACK )
		    
		    Const SM_CXMAXIMIZED = 61
		    Const SM_CYMAXIMIZED = 62
		    mMaxWindowWidth = GetSystemMetrics( SM_CXMAXIMIZED )
		    mMaxWindowHeight = GetSystemMetrics( SM_CYMAXIMIZED )
		    
		    Const SM_NETWORK = 63
		    mNetworkPresent = (BitwiseAnd( GetSystemMetrics( SM_NETWORK ), &h1 ) = 1)
		    
		    Const SM_CLEANBOOT = 67
		    mCleanBoot = (GetSystemMetrics( SM_CLEANBOOT ) = 0 )
		    
		    Const SM_CXDRAG = 68
		    Const SM_CYDRAG = 69
		    mDragWidth = GetSystemMetrics( SM_CXDRAG )
		    mDragHeight = GetSystemMetrics( SM_CYDRAG )
		    
		    Const SM_SHOWSOUNDS = 70
		    mShowSounds = (GetSystemMetrics( SM_SHOWSOUNDS ) <> 0)
		    
		    Const SM_CXMENUCHECK = 71
		    Const SM_CYMENUCHECK = 72
		    mMenuBitmapWidth = GetSystemMetrics( SM_CXMENUCHECK )
		    mMenuBitmapHeight = GetSystemMetrics( SM_CYMENUCHECK )
		    
		    Const SM_SLOWMACHINE = 73
		    mSlowMachine = (GetSystemMetrics( SM_SLOWMACHINE ) <> 0)
		    
		    Const SM_MIDEASTENABLED = 74
		    mMidEastEnabled = (GetSystemMetrics( SM_MIDEASTENABLED ) <> 0)
		    
		    Const SM_MOUSEWHEELPRESENT = 75
		    mMouseWheelPresent = (GetSystemMetrics( SM_MOUSEWHEELPRESENT ) <> 0)
		    
		    Const SM_XVIRTUALSCREEN = 76
		    Const SM_YVIRTUALSCREEN = 77
		    mVirtualScreenLeft = GetSystemMetrics( SM_XVIRTUALSCREEN )
		    mVirtualScreenTop = GetSystemMetrics( SM_YVIRTUALSCREEN )
		    
		    Const SM_CXVIRTUALSCREEN = 78
		    Const SM_CYVIRTUALSCREEN = 79
		    mVirtualScreenWidth = GetSystemMetrics( SM_CXVIRTUALSCREEN )
		    mVirtualScreenHeight = GetSystemMetrics( SM_CYVIRTUALSCREEN )
		    
		    Const SM_CMONITORS = 80
		    mNumMonitors = GetSystemMetrics( SM_CMONITORS )
		    
		    Const SM_SAMEDISPLAYFORMAT = 81
		    mSameDisplayFormat = (GetSystemMetrics( SM_SAMEDISPLAYFORMAT ) <> 0)
		    
		    Const SM_IMMENABLED = 82
		    mIMEEnabled = (GetSystemMetrics( SM_IMMENABLED ) <> 0)
		    
		    Const SM_CXFOCUSBORDER = 83
		    Const SM_CYFOCUSBORDER = 84
		    mFocusBorderWidth = GetSystemMetrics( SM_CXFOCUSBORDER )
		    mFocusBorderHeight = GetSystemMetrics( SM_CYFOCUSBORDER )
		    
		    Const SM_TABLETPC = 86
		    Const SM_MEDIACENTER = 87
		    Const SM_STARTER = 88
		    Const SM_SERVERR2 = 89
		    mTabletPC = (GetSystemMetrics( SM_TABLETPC ) <> 0)
		    mMediaCenterPC = (GetSystemMetrics( SM_MEDIACENTER ) <> 0)
		    mStarterPC = (GetSystemMetrics( SM_STARTER ) <> 0)
		    mServerR2PC = (GetSystemMetrics( SM_SERVERR2 ) <> 0)
		    
		    Const SM_MOUSEHORIZONTALWHEELPRESENT = 91
		    mHorizMouseWheelPresent = (GetSystemMetrics( SM_MOUSEHORIZONTALWHEELPRESENT ) <> 0)
		    
		    Const SM_CXPADDEDBORDER = 92
		    mCXPaddedBorder = GetSystemMetrics( SM_CXPADDEDBORDER )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function KanjiWindowHeight() As Integer
		  Initialize
		  
		  return mKanjiWindowHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MaxWindowHeight() As Integer
		  Initialize
		  
		  return mMaxWindowHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MaxWindowTrackingHeight() As Integer
		  Initialize
		  
		  return mMaxTrackingHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MaxWindowTrackingWidth() As Integer
		  Initialize
		  
		  return mMaxTrackingWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MaxWindowWidth() As Integer
		  Initialize
		  
		  return mMaxWindowWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MediaCenterPC() As Boolean
		  Initialize
		  return mMediaCenterPC
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MenuBarButtonHeight() As Integer
		  Initialize
		  
		  return mMenuBarButtonHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MenuBarButtonWidth() As Integer
		  Initialize
		  
		  return mMenuBarButtonWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MenuBarHeight() As Integer
		  Initialize
		  
		  return mMEnuBarHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MenuBitmapHeight() As Integer
		  Initialize
		  
		  return mMenuBitmapHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MenuBitmapWidth() As Integer
		  Initialize
		  
		  return mMenuBitmapWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MidEastEnabled() As Boolean
		  Initialize
		  
		  return mMidEastEnabled
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MinimizedWindowHeight() As Integer
		  Initialize
		  
		  return mMinimizedWindowHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MinimizedWindowSpacingHeight() As Integer
		  Initialize
		  
		  return mMinimizedWindowSpacingHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MinimizedWindowSpacingWidth() As Integer
		  Initialize
		  
		  return mMinimizedWindowSpacingWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MinimizedWindowWidth() As Integer
		  Initialize
		  
		  return mMinimizedWindowWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MinWindowHeight() As Integer
		  Initialize
		  
		  return mMinWindowHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MinWindowTrackingHeight() As Integer
		  Initialize
		  
		  return mMinWindowTrackingHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MinWindowTrackingWidth() As Integer
		  Initialize
		  
		  return mMinWindowTrackingWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MinWindowWidth() As Integer
		  Initialize
		  
		  return mMinWindowWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MousePresent() As Boolean
		  Initialize
		  
		  return mMousePresent
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MouseWheelPresent() As Boolean
		  Initialize
		  
		  return mMouseWheelPresent
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function NetworkPresent() As Boolean
		  Initialize
		  
		  return mNEtworkPresent
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function NumMonitors() As Integer
		  Initialize
		  
		  return mNumMonitors
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function NumMouseButtons() As Integer
		  Initialize
		  
		  return mNumMouseButtons
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function PaddedBorderWidth() As Integer
		  Initialize
		  return mCXPaddedBorder
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function PenWindowInstalled() As Boolean
		  Initialize
		  
		  return mPEnWindowsInstalled
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function RightAlignMenuDrops() As Boolean
		  Initialize
		  
		  return mRightAlignMenuDrops
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SameDisplayFormat() As Boolean
		  Initialize
		  
		  return mSameDisplayFormat
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ScreenHeight() As Integer
		  Initialize
		  
		  return mScreenCY
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ScreenWidth() As Integer
		  Initialize
		  
		  return mScreenCX
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ScrollbarArrowHeight() As Integer
		  Initialize
		  
		  return mScrollbarArrowHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ScrollbarArrowWidth() As Integer
		  Initialize
		  
		  return mScrollbarArrowWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SecurityPresent() As Boolean
		  Initialize
		  
		  return mSecurityPresent
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ServerR2PC() As Boolean
		  Initialize
		  return mServerR2PC
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ShowSounds() As Boolean
		  Initialize
		  
		  return mShowSounds
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SlowMachine() As Boolean
		  Initialize
		  
		  return mSlowMachine
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SmallCaptionButtonHeight() As Integer
		  Initialize
		  
		  return mSmallCaptionButtonHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SmallCaptionButtonWidth() As Integer
		  Initialize
		  
		  return mSmallCaptionButtonWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SmallCaptionHeight() As Integer
		  Initialize
		  
		  return mSmallCaptionHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SmallIconHeight() As Integer
		  Initialize
		  
		  return mSmallIconHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SmallIconWidth() As Integer
		  Initialize
		  
		  return mSmallIconWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function StarterPC() As Boolean
		  Initialize
		  return mStarterPC
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SwapButtons() As Boolean
		  Initialize
		  
		  return mSwapButtons
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function TabletPC() As Boolean
		  Initialize
		  
		  return mTabletPC
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function VerticalScrollbarWidth() As Integer
		  Initialize
		  
		  return mVerticalScrollbarWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function VerticalScrollThumbHeight() As Integer
		  Initialize
		  
		  return mVerticalScrollThumbHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function VirtualScreenHeight() As Integer
		  Initialize
		  
		  return mVirtualScreenHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function VirtualScreenLeft() As Integer
		  Initialize
		  
		  return mVirtualScreenLeft
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function VirtualScreenTop() As Integer
		  Initialize
		  
		  return mVirtualScreenTop
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function VirtualScreenWidth() As Integer
		  Initialize
		  
		  return mVirtualScreenWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Window3DBorderHeight() As Integer
		  Initialize
		  
		  return m3DBorderHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Window3DBorderWidth() As Integer
		  Initialize
		  
		  return m3DBorderWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function WindowBorderHeight() As Integer
		  Initialize
		  
		  return mWindowBorderHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function WindowBorderWidth() As Integer
		  Initialize
		  
		  return mWindowBorderWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function WindowSizingBorderHeight() As Integer
		  Initialize
		  
		  return mWindowSizingBorderHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function WindowSizingBorderWidth() As Integer
		  Initialize
		  
		  return mWindowSizingBorderWidth
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21
		Private m3DBorderHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private m3DBorderWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mCaptionBarButtonHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mCaptionBarButtonWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mCaptionHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mCleanBoot As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mCursorHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mCursorWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mCXPaddedBorder As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDebugUser As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDialogBorderHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDialogBorderWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDoubeClickHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDoubeClickWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDoubleByteEnabled As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDragHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDragWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mFocusBorderHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mFocusBorderWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mFullScreenHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mFullScreenWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mHorizMouseWheelPresent As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mHorizontalScrollbarHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mHorizontalScrollThumbWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIconHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIconSpacingHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIconSpacingWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIconWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIMEEnabled As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mInitialized As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mKanjiWindowHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMaxTrackingHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMaxTrackingWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMaxWindowHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMaxWindowWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMediaCenterPC As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMenuBarButtonHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMenuBarButtonWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMenuBarHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMenuBitmapHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMenuBitmapWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMidEastEnabled As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMinimizedWindowHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMinimizedWindowSpacingHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMinimizedWindowSpacingWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMinimizedWindowWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMinWindowHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMinWindowTrackingHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMinWindowTrackingWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMinWindowWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMousePresent As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMouseWheelPresent As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mNetworkPresent As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mNumMonitors As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mNumMouseButtons As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mPenWindowsInstalled As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mRightAlignMenuDrops As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mSameDisplayFormat As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mScreenCX As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mScreenCY As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mScrollbarArrowHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mScrollbarArrowWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mSecurityPresent As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mServerR2PC As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mShowSounds As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mSlowMachine As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mSmallCaptionButtonHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mSmallCaptionButtonWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mSmallCaptionHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mSmallIconHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mSmallIconWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mStarterPC As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mSwapButtons As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mTabletPC As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mVerticalScrollbarWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mVerticalScrollThumbHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mVirtualScreenHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mVirtualScreenLeft As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mVirtualScreenTop As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mVirtualScreenWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mWindowBorderHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mWindowBorderWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mWindowSizingBorderHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mWindowSizingBorderWidth As Integer
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
