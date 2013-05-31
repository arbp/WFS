#tag Class
Protected Class Font
	#tag Method, Flags = &h0
		Sub Constructor()
		  mSize = 12
		  Weight = kWeightNormal
		  Quality = kQualityDefault
		  Pitch = kPitchDefault
		  FontFamily = kFontFamilyDontCare
		  Charset = kCharsetDefault
		  OutputPrecision = kOutputPrecisionDefault
		  ClippingPrecision = kClippingPrecisionDefault
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Destructor()
		  // Destroy anything we have left to destroy
		  Unlock
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Lock() As Integer
		  #if TargetWin32
		    Soft Declare Function CreateFontA Lib "Gdi32" ( height as Integer, width as Integer, escapement as Integer, _
		    orientation as Integer, weight as Integer, italic as Boolean, underline as Boolean, strikeout as Boolean, _
		    charSet as Integer, outputPrecision as Integer, clippingPrecision as Integer, quality as Integer, _
		    pitchAndFamily as Integer, typefaceName as CString ) as Integer
		    
		    Soft Declare Function CreateFontW Lib "Gdi32" ( height as Integer, width as Integer, escapement as Integer, _
		    orientation as Integer, weight as Integer, italic as Boolean, underline as Boolean, strikeout as Boolean, _
		    charSet as Integer, outputPrecision as Integer, clippingPrecision as Integer, quality as Integer, _
		    pitchAndFamily as Integer, typefaceName as WString ) as Integer
		    
		    Dim hFont as Integer
		    
		    if System.IsFunctionAvailable( "CreateFontW", "Gdi32" ) then
		      hFont = CreateFontW( -mSize, 0, Escapement, Orientation, Weight, Italic, Underline, Strikethrough, _
		      Charset, OutputPrecision, ClippingPrecision, _
		      Quality, Pitch + FontFamily, TypefaceName )
		    else
		      hFont = CreateFontA( -mSize, 0, Escapement, Orientation, Weight, Italic, Underline, Strikethrough, _
		      Charset, OutputPrecision, ClippingPrecision, _
		      Quality, Pitch + FontFamily, TypefaceName )
		    end if
		    
		    mStoredFont = hFont
		    
		    return hFont
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub RealizeFont(handle as Integer, forceRefresh as Boolean = true)
		  #if TargetWin32
		    Soft Declare Sub SendMessageA Lib "User32" ( hwnd as Integer, msg as Integer, wParam as Integer, lParam as Integer )
		    Soft Declare Sub SendMessageW Lib "User32" ( hwnd as Integer, msg as Integer, wParam as Integer, lParam as Integer )
		    
		    Const WM_SETFONT = &h30
		    
		    dim refresh as Integer
		    if forceRefresh then refresh = 1
		    
		    if System.IsFunctionAvailable( "SendMessageW", "User32" ) then
		      SendMessageW( handle, WM_SETFONT, me.Lock, refresh )
		    else
		      SendMessageA( handle, WM_SETFONT, me.Lock, refresh )
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Size(type as Integer = 0) As Integer
		  select case type
		  case kSizeTypePoint
		    return mSize
		    
		  else
		    raise new TypeMismatchException
		    return -1
		  end select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Size(type as Integer = 0, assigns set as Integer)
		  #if TargetWin32
		    select case type
		    case kSizeTypePoint
		      // We need to use MulDiv to convert this into the proper metric
		      Declare Function GetDeviceCaps Lib "Gdi32" ( hdc as Integer, index as Integer ) as Integer
		      Declare Function MulDiv Lib "Kernel32" ( number as Integer, numerator as Integer, divisor as Integer ) as Integer
		      Declare Function GetDC Lib "User32" ( hwnd as Integer ) as Integer
		      Declare Sub ReleaseDC Lib "User32" ( hwnd as Integer, dc as Integer )
		      
		      dim hdc as Integer = GetDC( 0 )
		      Const LOGPIXELSY = 90
		      mSize = MulDiv( set, GetDeviceCaps( hdc, LOGPIXELSY ), 72 )
		      ReleaseDC( 0, hdc )
		    else
		      raise new TypeMismatchException
		    end select
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Unlock()
		  #if TargetWin32
		    Declare Sub DeleteObject Lib "Gdi32" ( obj as Integer )
		    
		    if mStoredFont <> 0 then
		      DeleteObject( mStoredFont )
		    end if
		    
		    mStoredFont = 0
		  #endif
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		Charset As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ClippingPrecision As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Escapement As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		FontFamily As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Italic As Boolean
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mSize As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mStoredFont As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Orientation As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		OutputPrecision As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Pitch As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Quality As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Strikethrough As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		TypefaceName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Underline As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		Weight As Integer
	#tag EndProperty


	#tag Constant, Name = kCharsetANSI, Type = Double, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetArabic, Type = Double, Dynamic = False, Default = \"178", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetBaltic, Type = Double, Dynamic = False, Default = \"186", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetChineseBig5, Type = Double, Dynamic = False, Default = \"136", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetDefault, Type = Double, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetEastEurope, Type = Double, Dynamic = False, Default = \"238", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetGB2312, Type = Double, Dynamic = False, Default = \"134", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetGreek, Type = Double, Dynamic = False, Default = \"161", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetHangul, Type = Double, Dynamic = False, Default = \"129", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetHebrew, Type = Double, Dynamic = False, Default = \"177", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetJohab, Type = Double, Dynamic = False, Default = \"130", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetMac, Type = Double, Dynamic = False, Default = \"77", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetOEM, Type = Double, Dynamic = False, Default = \"255", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetRussian, Type = Double, Dynamic = False, Default = \"204", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetShiftJIS, Type = Double, Dynamic = False, Default = \"128", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetSymbol, Type = Double, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetThai, Type = Double, Dynamic = False, Default = \"222", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetTurkish, Type = Double, Dynamic = False, Default = \"162", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharsetVietnamese, Type = Double, Dynamic = False, Default = \"163", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kClippingPrecisionCharacter, Type = Double, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kClippingPrecisionDefault, Type = Double, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kClippingPrecisionDisableFontAssociation, Type = Double, Dynamic = False, Default = \"64", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kClippingPrecisionEmbedded, Type = Double, Dynamic = False, Default = \"128", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kClippingPrecisionLeftHandAngles, Type = Double, Dynamic = False, Default = \"16", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kClippingPrecisionStroke, Type = Double, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kClippingPrecisionTrueTypeAlways, Type = Double, Dynamic = False, Default = \"32", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kFontFamilyDecorative, Type = Double, Dynamic = False, Default = \"80", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kFontFamilyDontCare, Type = Double, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kFontFamilyModern, Type = Double, Dynamic = False, Default = \"48", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kFontFamilyRoman, Type = Double, Dynamic = False, Default = \"16", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kFontFamilyScript, Type = Double, Dynamic = False, Default = \"64", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kFontFamilySwiss, Type = Double, Dynamic = False, Default = \"32", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionCharacter, Type = Double, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionDefault, Type = Double, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionDevice, Type = Double, Dynamic = False, Default = \"5", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionOutline, Type = Double, Dynamic = False, Default = \"8", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionPostScriptOnly, Type = Double, Dynamic = False, Default = \"10", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionRaster, Type = Double, Dynamic = False, Default = \"6", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionScreenOutline, Type = Double, Dynamic = False, Default = \"9", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionString, Type = Double, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionStroke, Type = Double, Dynamic = False, Default = \"3", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionTrueType, Type = Double, Dynamic = False, Default = \"4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionTrueTypeOnly, Type = Double, Dynamic = False, Default = \"7", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kPitchDefault, Type = Double, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kPitchFixed, Type = Double, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kPitchVariable, Type = Double, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kQualityAntiAliased, Type = Double, Dynamic = False, Default = \"4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kQualityClearType, Type = Double, Dynamic = False, Default = \"5", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kQualityClearTypeNatural, Type = Double, Dynamic = False, Default = \"6", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kQualityDefault, Type = Double, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kQualityDraft, Type = Double, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kQualityNonAntiAliased, Type = Double, Dynamic = False, Default = \"3", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kQualityProof, Type = Double, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kSizeTypePoint, Type = Double, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightBlack, Type = Double, Dynamic = False, Default = \"900", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightBold, Type = Double, Dynamic = False, Default = \"700", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightDontCare, Type = Double, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightExtraBold, Type = Double, Dynamic = False, Default = \"800", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightExtraLight, Type = Double, Dynamic = False, Default = \"200", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightLight, Type = Double, Dynamic = False, Default = \"300", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightMedium, Type = Double, Dynamic = False, Default = \"500", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightNormal, Type = Double, Dynamic = False, Default = \"400", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightSemiBold, Type = Double, Dynamic = False, Default = \"600", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightThin, Type = Double, Dynamic = False, Default = \"100", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="Charset"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ClippingPrecision"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Escapement"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="FontFamily"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Italic"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
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
			Name="Orientation"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="OutputPrecision"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Pitch"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Quality"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Strikethrough"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
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
		#tag ViewProperty
			Name="TypefaceName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Underline"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Weight"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
