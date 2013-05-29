#tag Module
Protected Module SystemColors
	#tag Method, Flags = &h1
		Protected Function ActiveBorder() As Color
		  Initialize
		  
		  return mActiveBorder
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ActiveCaption() As Color
		  Initialize
		  
		  return mActiveCaption
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AppWorkspace() As Color
		  Initialize
		  
		  return mAppWorkspace
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Background() As Color
		  Initialize
		  
		  return mBackgroundColor
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ButtonDarkShadow() As Color
		  Initialize
		  
		  return mButtonDkShadow
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ButtonFace() As Color
		  Initialize
		  
		  return mButtonFace
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ButtonHighlight() As Color
		  Initialize
		  
		  return mButtonHighlight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ButtonShadow() As Color
		  Initialize
		  
		  return mButtonShadow
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ButtonText() As Color
		  Initialize
		  
		  return mButtonText
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function CaptionText() As Color
		  Initialize
		  
		  return mCaptionText
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GreyText() As Color
		  Initialize
		  
		  return mGreyText
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Highlight() As Color
		  Initialize
		  
		  return mHighlight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function HighlightText() As Color
		  Initialize
		  
		  return mHighlightText
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function InactiveBorder() As Color
		  Initialize
		  
		  return mInactiveBorder
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function InactiveCaption() As Color
		  Initialize
		  
		  return mInactiveCaption
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function InactiveCaptionText() As Color
		  Initialize
		  
		  return mInactiveCaptionText
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub Initialize()
		  if mInitialized then return
		  mInitialized = true
		  
		  #if TargetWin32
		    // SysCol class wrapper for GetSysColor
		    dim res as memoryblock
		    dim c as color
		    
		    Const COLOR_ACTIVEBORDER = 10
		    Const COLOR_ACTIVECAPTION = 2
		    Const COLOR_ADJ_MAX = 100
		    Const COLOR_ADJ_MIN = -100
		    Const COLOR_APPWORKSPACE = 12
		    Const COLOR_BACKGROUND = 1
		    Const COLOR_BTNFACE = 15
		    Const COLOR_BTNHIGHLIGHT = 20
		    Const COLOR_BTNSHADOW = 16
		    Const COLOR_BTNTEXT = 18
		    Const COLOR_CAPTIONTEXT = 9
		    Const COLOR_GRAYTEXT = 17
		    Const COLOR_HIGHLIGHT = 13
		    Const COLOR_HIGHLIGHTTEXT = 14
		    Const COLOR_INACTIVEBORDER = 11
		    Const COLOR_INACTIVECAPTION = 3
		    Const COLOR_INACTIVECAPTIONTEXT = 19
		    Const COLOR_MENU = 4
		    Const COLOR_MENUTEXT = 7
		    Const COLOR_SCROLLBAR = 0
		    Const COLOR_WINDOW = 5
		    Const COLOR_WINDOWFRAME = 6
		    Const COLOR_WINDOWTEXT = 8
		    Const COLOR_BUTTONDKSHADOW = 21
		    
		    Declare Function GetSysColor Lib "user32" (nIndex As integer) As integer
		    res = new MemoryBlock(4)
		    
		    
		    res.long(0) = getsyscolor(COLOR_ACTIVEBORDER)
		    if res <> nil then
		      mActiveBorder = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_ACTIVECAPTION)
		    if res <> nil then
		      mActiveCaption = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_APPWORKSPACE)
		    if res <> nil then
		      mappworkspace = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_BACkGROUND)
		    if res <> nil then
		      mbackgroundColor = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_BTNFACE)
		    if res <> nil then
		      mButtonFace= rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_BTNHIGHLIGHT)
		    if res <> nil then
		      mButtonHighlight= rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_BTNSHADOW)
		    if res <> nil then
		      mButtonShadow= rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_BTNTEXT)
		    if res <> nil then
		      mButtonText= rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_CAPTIONTEXT)
		    if res <> nil then
		      mCaptionText = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_GRAYTEXT)
		    if res <> nil then
		      mGreyText = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_HIGHLIGHT)
		    if res <> nil then
		      mHighlight = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_HIGHLIGHTTEXT)
		    if res <> nil then
		      mHighlighttext = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_INACTIVEBORDER)
		    if res <> nil then
		      mInactiveBorder = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_INACTIVECAPTION)
		    if res <> nil then
		      mInactiveCaption = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_INACTIVECAPTIONTEXT)
		    if res <> nil then
		      mInactiveCaptionText = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_MENU)
		    if res <> nil then
		      mMenuColor = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_MENUTEXT)
		    if res <> nil then
		      mMenuText = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_SCROLLBAR)
		    if res <> nil then
		      mScrollBarColor= rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_WINDOW)
		    if res <> nil then
		      mWindowColor = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_WINDOWFRAME)
		    if res <> nil then
		      mWindowFrame = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_WINDOWTEXT)
		    if res <> nil then
		      mWindowText = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		    
		    res.long(0) = getsyscolor(COLOR_BUTTONDKSHADOW)
		    if res <> nil then
		      mButtonDkShadow = rgb(res.byte(0),res.byte(1),res.byte(2))
		    end
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Menu() As Color
		  Initialize
		  
		  return mMenuColor
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MenuText() As Color
		  Initialize
		  
		  return mMenuText
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ScrollBar() As Color
		  Initialize
		  
		  return mScrollBarColor
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Window() As Color
		  Initialize
		  
		  return mWindowColor
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function WindowFrame() As Color
		  Initialize
		  
		  return mWindowFrame
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function WindowText() As Color
		  Initialize
		  
		  return mWindowText
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mActiveBorder As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mActiveCaption As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mAppWorkspace As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mBackgroundColor As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mButtonDkShadow As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mButtonFace As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mButtonHighlight As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mButtonShadow As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mButtonText As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mCaptionText As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mGreyText As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mHighlight As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mHighlightText As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mInactiveBorder As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mInactiveCaption As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mInactiveCaptionText As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mInitialized As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMenuColor As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMenuText As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mScrollBarColor As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mWindowColor As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mWindowFrame As color
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mWindowText As color
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
