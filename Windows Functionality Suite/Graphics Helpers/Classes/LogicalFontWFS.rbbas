#tag Class
Protected Class LogicalFontWFS
	#tag Method, Flags = &h0
		Sub Constructor()
		  // Do nothing
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(mb as MemoryBlock, bUnicodeSavvy as Boolean)
		  // Convert from a memory block
		  
		  Height = mb.Long( 0 )
		  Width = mb.Long( 4 )
		  Escapement = mb.Long( 8 )
		  Orientation = mb.Long( 12 )
		  Weight = mb.Long( 16 )
		  Italic = (mb.Byte( 20 ) <> 0 )
		  Underline = (mb.Byte( 21 ) <> 0)
		  Strikeout = (mb.Byte( 22 ) <> 0)
		  CharSet = mb.Byte( 23 )
		  OutputPrecision = mb.Byte( 24 )
		  ClippingPrecision = mb.Byte( 25 )
		  Quality = mb.Byte( 26 )
		  
		  dim pitchAndFamily as Integer
		  PitchAndFamily = mb.Byte( 27 )
		  
		  Pitch = BitwiseAnd( pitchAndFamily, &h3 )
		  Family = Bitwise.ShiftRight( BitwiseAnd( pitchAndFamily, &hF0 ), 2 )
		  
		  if bUnicodeSavvy then
		    FaceName = mb.WString( 28 )
		  else
		    FaceName = mb.CString( 28 )
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ToMemoryBlock(bUnicodeSavvy as Boolean) As MemoryBlock
		  // Convert to a memory block
		  
		  // We make a 60 byte structure
		  dim ret as MemoryBlock
		  if bUnicodeSavvy then
		    ret = new MemoryBlock( kWideCharSize  )
		  else
		    ret = new MemoryBlock( kANSISize )
		  end if
		  
		  ret.Long( 0 ) = Height
		  ret.Long( 4 ) = Width
		  ret.Long( 8 ) = Escapement
		  ret.Long( 12 ) = Orientation
		  ret.Long( 16 ) = Weight
		  ret.BooleanValue( 20 ) = Italic
		  ret.BooleanValue( 21 ) = Underline
		  ret.BooleanValue( 22 ) = Strikeout
		  ret.Byte( 23 ) = CharSet
		  ret.Byte( 24 ) = OutputPrecision
		  ret.Byte( 25 ) = ClippingPrecision
		  ret.Byte( 26 ) = Quality
		  ret.Byte( 27 ) = Bitwise.ShiftRight( Family, 2 ) + Pitch
		  if bUnicodeSavvy then
		    ret.WString( 28 ) = FaceName
		  else
		    ret.CString( 28 ) = FaceName
		  end if
		  
		  Return ret
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		CharSet As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ClippingPrecision As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Escapement As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		FaceName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Family As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Height As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Italic As Boolean
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
		Strikeout As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		Underline As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		Weight As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Width As Integer
	#tag EndProperty


	#tag Constant, Name = kANSISize, Type = Double, Dynamic = False, Default = \"60", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetANSI, Type = Integer, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetArabic, Type = Integer, Dynamic = False, Default = \"178", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetBaltic, Type = Integer, Dynamic = False, Default = \"186", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetChineseBig5, Type = Integer, Dynamic = False, Default = \"136", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetDefault, Type = Integer, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetEastEurope, Type = Integer, Dynamic = False, Default = \"238", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetGB2312, Type = Integer, Dynamic = False, Default = \"134", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetGreek, Type = Integer, Dynamic = False, Default = \"161", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetHangeul, Type = Integer, Dynamic = False, Default = \"129", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetHebrew, Type = Integer, Dynamic = False, Default = \"177", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetJohab, Type = Integer, Dynamic = False, Default = \"130", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetMac, Type = Integer, Dynamic = False, Default = \"77", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetOEM, Type = Integer, Dynamic = False, Default = \"255", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetRussian, Type = Integer, Dynamic = False, Default = \"204", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetShiftJIS, Type = Integer, Dynamic = False, Default = \"128", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetSymbol, Type = Integer, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetThai, Type = Integer, Dynamic = False, Default = \"222", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetTurkish, Type = Integer, Dynamic = False, Default = \"162", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCharSetVietnemese, Type = Integer, Dynamic = False, Default = \"163", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kClipPrecisionCharacter, Type = Integer, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kClipPrecisionDefault, Type = Integer, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kClipPrecisionEmbedded, Type = Integer, Dynamic = False, Default = \"128", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kClipPrecisionLeftHandAngles, Type = Integer, Dynamic = False, Default = \"4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kClipPrecisionStroke, Type = Integer, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kClipPrecisionTrueTypeAlways, Type = Integer, Dynamic = False, Default = \"32", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kFamilyDecorative, Type = Integer, Dynamic = False, Default = \"80", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kFamilyDontCare, Type = Integer, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kFamilyModern, Type = Integer, Dynamic = False, Default = \"48", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kFamilyRoman, Type = Integer, Dynamic = False, Default = \"16", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kFamilyScript, Type = Integer, Dynamic = False, Default = \"64", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kFamilySwiss, Type = Integer, Dynamic = False, Default = \"32", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionCharacter, Type = Integer, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionDefault, Type = Integer, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionDevice, Type = Integer, Dynamic = False, Default = \"5", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionOutline, Type = Integer, Dynamic = False, Default = \"8", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionPostScriptOnly, Type = Integer, Dynamic = False, Default = \"10", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionRaster, Type = Integer, Dynamic = False, Default = \"6", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionScreenOutline, Type = Integer, Dynamic = False, Default = \"9", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionString, Type = Integer, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionStroke, Type = Integer, Dynamic = False, Default = \"3", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionTrueType, Type = Integer, Dynamic = False, Default = \"4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kOutputPrecisionTrueTypeOnly, Type = Integer, Dynamic = False, Default = \"7", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kPitchDefault, Type = Integer, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kPitchFixed, Type = Integer, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kPitchMonoFont, Type = Integer, Dynamic = False, Default = \"8", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kPitchVariable, Type = Integer, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kQualityAntiAliased, Type = Integer, Dynamic = False, Default = \"4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kQualityClearType, Type = Integer, Dynamic = False, Default = \"5", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kQualityClearTypeNatural, Type = Integer, Dynamic = False, Default = \"6", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kQualityDefault, Type = Integer, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kQualityDraft, Type = Integer, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kQualityNonAntiAliased, Type = Integer, Dynamic = False, Default = \"3", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kQualityProof, Type = Integer, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightBold, Type = Integer, Dynamic = False, Default = \"700", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightDontCare, Type = Integer, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightExtraBold, Type = Integer, Dynamic = False, Default = \"800", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightExtraLight, Type = Integer, Dynamic = False, Default = \"200", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightHeavy, Type = Integer, Dynamic = False, Default = \"900", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightLight, Type = Integer, Dynamic = False, Default = \"300", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightMedium, Type = Integer, Dynamic = False, Default = \"500", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightNormal, Type = Integer, Dynamic = False, Default = \"400", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightSemiBold, Type = Integer, Dynamic = False, Default = \"600", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWeightThin, Type = Integer, Dynamic = False, Default = \"100", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWideCharSize, Type = Double, Dynamic = False, Default = \"92", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="CharSet"
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
			Name="FaceName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Family"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Height"
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
			Name="Strikeout"
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
		#tag ViewProperty
			Name="Width"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
