#tag Module
Protected Module BinaryStreamExtensions
	#tag Method, Flags = &h0
		Function ReadCString(extends bs as BinaryStream, enc as TextEncoding = nil) As String
		  // Read one bytes at a time until we come to a 0-byte.
		  dim ret as String
		  dim bytes as UInt8
		  
		  do
		    bytes = bs.ReadUInt8
		    
		    if enc <> nil then
		      ret = ret + enc.Chr( bytes )
		    else
		      ret = ret + Encodings.ASCII.Chr( bytes )
		    end if
		  loop until bytes = 0
		  
		  return ret
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ReadWString(extends bs as BinaryStream) As String
		  // Read one word at a time until we come to a 0-word.
		  dim ret as String
		  dim word as UInt16
		  
		  do
		    word = bs.ReadUInt16
		    ret = ret + Encodings.UTF16.Chr( word )
		  loop until word = 0
		  
		  return ret
		End Function
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
End Module
#tag EndModule
