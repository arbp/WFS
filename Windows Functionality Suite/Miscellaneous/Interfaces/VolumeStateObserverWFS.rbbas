#tag Interface
Protected Interface VolumeStateObserverWFS
	#tag Method, Flags = &h0
		Function CancelVolumeRemove(drives() as String, isMedia as Boolean, isNetwork as Boolean) As Boolean
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub VolumeAdded(drives() as String, isMedia as Boolean, isNetwork as Boolean)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub VolumeRemoved(drives() as String, isMedia as Boolean, isNetwork as Boolean)
		  
		End Sub
	#tag EndMethod


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
End Interface
#tag EndInterface
