#tag Class
Protected Class ThreadInformation
	#tag Method, Flags = &h0
		Sub Constructor()
		  // Default constructor, do nothing
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(mb as MemoryBlock)
		  // Make sure our data is sane
		  if mb = nil or mb.Long( 0 ) <> mb.Size then return
		  
		  // Grab the usage count
		  ReferenceCount = mb.Long( 4 )
		  
		  // Then the IDs
		  ThreadID = mb.Long( 8 )
		  OwnerProcessID = mb.Long( 12 )
		  
		  // Then the priorities
		  BasePriority = mb.Long( 16 )
		  PriorityDelta = mb.Long( 20 )
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		BasePriority As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		OwnerProcessID As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		PriorityDelta As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ReferenceCount As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ThreadID As Integer
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="BasePriority"
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
			Name="OwnerProcessID"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="PriorityDelta"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ReferenceCount"
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
			Name="ThreadID"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
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
