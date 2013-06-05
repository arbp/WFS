#tag Class
Protected Class AccessTimeoutWFS
	#tag Method, Flags = &h0
		Sub Constructor()
		  // Do nothing
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(mb as MemoryBlock)
		  Flags = mb.Long( 4 )
		  MillisecondsTimeout = mb.Long( 8 )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ToMemoryBlock() As MemoryBlock
		  dim ret as new MemoryBlock( 12 )
		  ret.Long( 0 ) = ret.Size
		  ret.Long( 4 ) = Flags
		  ret.Long( 8 ) = MillisecondsTimeout
		  
		  return ret
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		Flags As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		MillisecondsTimeout As Integer
	#tag EndProperty


	#tag Constant, Name = kOnOffFeedback, Type = Integer, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kTimeoutOn, Type = Integer, Dynamic = False, Default = \"1", Scope = Public
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
			Name="MillisecondsTimeout"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
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
