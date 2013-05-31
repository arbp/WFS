#tag Class
Protected Class ToggleKeys
	#tag Method, Flags = &h0
		Sub Constructor()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(mb as MemoryBlock)
		  Flags = mb.Long( 4 )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ToMemoryBlock() As MemoryBlock
		  dim ret as new MemoryBlock( 8 )
		  ret.Long( 0 ) = ret.Size
		  ret.Long( 4 ) = Flags
		  
		  return ret
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		Flags As Integer
	#tag EndProperty


	#tag Constant, Name = kAvailable, Type = Integer, Dynamic = False, Default = \"&h2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kConfirmHotKey, Type = Integer, Dynamic = False, Default = \"&h8", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kHotKeyActive, Type = Integer, Dynamic = False, Default = \"&h4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kHotKeySound, Type = Integer, Dynamic = False, Default = \"&h10", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kIndicator, Type = Integer, Dynamic = False, Default = \"&h20", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kToggleKeysOn, Type = Integer, Dynamic = False, Default = \"&h1", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
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
