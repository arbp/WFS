#tag Class
Protected Class FilterKeysWFS
	#tag Method, Flags = &h0
		Sub Constructor()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(mb as MemoryBlock)
		  Flags = mb.Long( 4 )
		  WaitMilliseconds = mb.Long( 8 )
		  DelayMilliseconds = mb.Long( 12 )
		  RepeatMilliseconds = mb.Long( 16 )
		  BoundMilliseconds = mb.Long( 20 )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ToMemoryBlock() As MemoryBlock
		  dim ret as new MemoryBlock( 24 )
		  ret.Long( 0 ) = ret.Size
		  ret.Long( 4 ) = Flags
		  ret.Long( 8 ) = WaitMilliseconds
		  ret.Long( 12 ) = DelayMilliseconds
		  ret.Long( 16 ) = RepeatMilliseconds
		  ret.Long( 20 ) = BoundMilliseconds
		  
		  return ret
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		BoundMilliseconds As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		DelayMilliseconds As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Flags As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		RepeatMilliseconds As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		WaitMilliseconds As Integer
	#tag EndProperty


	#tag Constant, Name = kAvailable, Type = Integer, Dynamic = False, Default = \"&h2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kClickOn, Type = Integer, Dynamic = False, Default = \"&h40", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kConfirmHotKeys, Type = Integer, Dynamic = False, Default = \"&h8", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kFilterKeysOn, Type = Integer, Dynamic = False, Default = \"&h1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kHotKeyActive, Type = Integer, Dynamic = False, Default = \"&h4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kHotKeySound, Type = Integer, Dynamic = False, Default = \"&h10", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kIndicator, Type = Integer, Dynamic = False, Default = \"&h20", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="BoundMilliseconds"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DelayMilliseconds"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Flags"
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
			Name="RepeatMilliseconds"
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
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="WaitMilliseconds"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
