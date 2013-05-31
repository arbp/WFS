#tag Class
Protected Class HeapEntryInformation
	#tag Method, Flags = &h0
		Sub Constructor()
		  // Default constructor, do nothing
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(mb as MemoryBlock)
		  // Ensure that things are sane
		  if mb = nil or mb.Long( 0 ) <> mb.Size then return
		  
		  // Get the handle to the heap
		  Handle = mb.Long( 4 )
		  
		  // As well as the address and size
		  Address = mb.Long( 8 )
		  BlockSize = mb.Long( 12 )
		  
		  Const LF32_FIXED = &h1
		  Const LF32_FREE = &h2
		  Const LF32_MOVEABLE = &h4
		  
		  // Get the flags
		  dim flags as Integer = mb.Long( 16 )
		  Fixed = Bitwise.BitAnd( flags, LF32_FIXED ) <> 0
		  Free = Bitwise.BitAnd( flags, LF32_FREE ) <> 0
		  Moveable = Bitwise.BitAnd( flags, LF32_MOVEABLE ) <> 0
		  
		  // Then comes the lock count
		  LockCount = mb.Long( 20 )
		  
		  // Skip the reserved stuff and get the IDs
		  OwnerProcessID = mb.Long( 28 )
		  OwnerHeapListID = mb.Long( 32 )
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadData()
		  #if TargetWin32
		    Soft Declare Function Toolhelp32ReadProcessMemory Lib "Kernel32" ( processID as Integer, baseAddress as Integer, _
		    buffer as Ptr, bytesToRead as Integer, ByRef numBytesRead as Integer ) as Boolean
		    
		    // Allocate a block large enough to hold all our data
		    Data = new MemoryBlock( BlockSize )
		    
		    if Data = nil then return
		    
		    // And try to read the data in
		    dim bytesRead as Integer = 0
		    if not Toolhelp32ReadProcessMemory( OwnerProcessID, Address, Data, Data.size, bytesRead ) then
		      Data = nil
		    else
		      Data.Size = bytesRead
		    end if
		  #endif
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		Address As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		BlockSize As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Data As MemoryBlock
	#tag EndProperty

	#tag Property, Flags = &h0
		Fixed As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		Free As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		Handle As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		LockCount As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Moveable As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		OwnerHeapListID As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		OwnerProcessID As Integer
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Address"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="BlockSize"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Fixed"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Free"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Handle"
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
			Name="LockCount"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Moveable"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="OwnerHeapListID"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="OwnerProcessID"
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
	#tag EndViewBehavior
End Class
#tag EndClass
