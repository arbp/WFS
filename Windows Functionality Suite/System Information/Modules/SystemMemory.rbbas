#tag Module
Protected Module SystemMemory
	#tag Method, Flags = &h1
		Protected Function AvailablePageFile() As Double
		  Dim mb as MemoryBlock = Query
		  
		  if mb <> nil and mb.Long( 0 ) >= 24 then return mb.Long( 20 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AvailablePhysical() As Double
		  Dim mb as MemoryBlock = Query
		  
		  if mb <> nil and mb.Long( 0 ) >= 16 then return mb.Long( 12 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AvailableVirtual() As Double
		  Dim mb as MemoryBlock = Query
		  
		  if mb <> nil and mb.Long( 0 ) >= 32 then return mb.Long( 28 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Load() As Double
		  Dim mb as MemoryBlock = Query
		  
		  if mb <> nil then return mb.Long( 4 ) / 100
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function Query() As MemoryBlock
		  Dim ret as new MemoryBlock( 32 )
		  
		  Soft Declare Sub GlobalMemoryStatus Lib "Kernel32" ( data as Ptr )
		  
		  if System.IsFunctionAvailable( "GlobalMemoryStatus", "Kernel32" ) then
		    GlobalMemoryStatus( ret )
		    
		    return ret
		  end if
		  
		  return nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function TotalPageFile() As Double
		  Dim mb as MemoryBlock = Query
		  
		  if mb <> nil and mb.Long( 0 ) >= 20 then return mb.Long( 16 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function TotalPhysical() As Double
		  Dim mb as MemoryBlock = Query
		  
		  if mb <> nil and mb.Long( 0 ) >= 12 then return mb.Long( 8 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function TotalVirtual() As Double
		  Dim mb as MemoryBlock = Query
		  
		  if mb <> nil and mb.Long( 0 ) >= 28 then return mb.Long( 24 )
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
