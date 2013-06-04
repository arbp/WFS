#tag Class
Protected Class IconMetricsWFS
	#tag Method, Flags = &h0
		Sub Constructor()
		  // Do nothing
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(mb as MemoryBlock, bUnicodeSavvy as Boolean)
		  // Convert from a memory block
		  StructureSize = mb.Long( 0 )
		  HorizontalSpacing = mb.Long( 4 )
		  VerticalSpacing = mb.Long( 8 )
		  TitleWrap = (mb.Long( 12 ) <> 0)
		  
		  dim theFont as MemoryBlock = mb.StringValue( 16, mb.Size - 16 )
		  Font = new LogicalFontWFS( theFont, bUnicodeSavvy )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ToMemoryBlock(bUnicodeSavvy as Boolean) As MemoryBlock
		  // Convert to a MemoryBlock
		  
		  // We make a 16 byte structure, plus 60 bytes for the font
		  dim us as new MemoryBlock( 16 )
		  if bUnicodeSavvy then
		    us.Long( 0 ) = 16 + LogicalFontWFS.kWideCharSize
		  else
		    us.Long( 0 ) = 16 + LogicalFontWFS.kANSISize
		  end if
		  
		  us.Long( 4 ) = HorizontalSpacing
		  us.Long( 8 ) = VerticalSpacing
		  if TitleWrap then
		    us.Long( 12 ) = 1
		  else
		    us.Long( 12 ) = 0
		  end if
		  
		  return us + Font.ToMemoryBlock( bUnicodeSavvy )
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		Font As LogicalFontWFS
	#tag EndProperty

	#tag Property, Flags = &h0
		HorizontalSpacing As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		StructureSize As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		TitleWrap As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		VerticalSpacing As Integer
	#tag EndProperty


	#tag Constant, Name = kANSISize, Type = Double, Dynamic = False, Default = \"76", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kWideCharSize, Type = Double, Dynamic = False, Default = \"108", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="HorizontalSpacing"
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
			Name="StructureSize"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TitleWrap"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="VerticalSpacing"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
