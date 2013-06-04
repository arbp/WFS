#tag Module
Protected Module PerformanceCounterWFS
	#tag Method, Flags = &h1
		Protected Function Frequency() As Double
		  #if TargetWin32
		    Declare Function QueryPerformanceFrequency Lib "Kernel32" ( freq as Ptr ) as Boolean
		    
		    dim freq as new MemoryBlock( 8 )
		    
		    if not QueryPerformanceFrequency( freq ) then
		      return -1
		    end if
		    
		    #if RBVersion < 2006.01 then
		      return LongLongToDouble( freq )
		    #else
		      return freq.UInt64Value( 0 )
		    #endif
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetCount() As Double
		  #if TargetWin32
		    Declare Function QueryPerformanceCounter Lib "Kernel32" ( perfCount as Ptr ) as Boolean
		    
		    dim cnt1 as new MemoryBlock( 8 )
		    if not QueryPerformanceCounter( cnt1 ) then return -1
		    
		    #if RBVersion < 2006.01 then
		      return LongLongToDouble( cnt1 )
		    #else
		      return cnt1.UInt64Value( 0 )
		    #endif
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function LongLongToDouble(mb as MemoryBlock) As Double
		  dim ret as Double
		  
		  // Take the high 4 bytes and shift them
		  ret = mb.Long( 4 ) * (2 ^32)
		  
		  // Then add in the low 4 bytes
		  ret = ret + mb.Long( 0 )
		  
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
