#tag Class
Protected Class HeapListInformationWFS
	#tag Method, Flags = &h0
		Sub Constructor()
		  // Default constructor, do nothing
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(mb as MemoryBlock)
		  // Make sure our data is sane
		  if mb = nil or mb.Long( 0 ) <> mb.Size then return
		  
		  OwnerProcessID = mb.Long( 4 )
		  HeapListID = mb.Long( 8 )
		  
		  dim flags as Integer
		  flags = mb.Long( 12 )
		  
		  Const HF32_DEFAULT = 1
		  Const HF32_SHARED = 2
		  
		  DefaultHeap = Bitwise.BitAnd( flags, HF32_DEFAULT ) <> 0
		  SharedHeap = Bitwise.BitAnd( flags, HF32_SHARED ) <> 0
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadHeapEntries(entryCallback as HeapEntryLoadedCallbackWFS = nil)
		  #if TargetWin32
		    
		    Soft Declare Function Heap32First Lib "Kernel32" ( heapEntry as Ptr, processID as Integer, heapID as Integer ) as Boolean
		    Soft Declare Function Heap32Next Lib "Kernel32" ( entry as Ptr ) as Integer
		    
		    dim mb as new MemoryBlock( 36 )
		    
		    dim entry as HeapEntryInformation
		    
		    mb.Long( 0 ) = mb.Size
		    if not Heap32First( mb, OwnerProcessID, HeapListID ) then return
		    
		    dim good as Integer
		    
		    do
		      entry = new HeapEntryInformation( mb )
		      
		      HeapEntries.Append( entry )
		      
		      if entryCallback <> nil then
		        if entryCallback.HeapEntryLoaded( entry ) then return
		      end if
		      
		      good = Heap32Next( mb )
		    loop until good = 0
		    
		  #else
		    
		    #pragma unused entryCallback
		    
		  #endif
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		DefaultHeap As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		HeapEntries(-1) As HeapEntryInformationWFS
	#tag EndProperty

	#tag Property, Flags = &h0
		HeapListID As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		OwnerProcessID As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		SharedHeap As Boolean
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="DefaultHeap"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="HeapListID"
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
			Name="SharedHeap"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
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
