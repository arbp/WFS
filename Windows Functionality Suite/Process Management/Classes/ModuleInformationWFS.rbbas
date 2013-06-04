#tag Class
Protected Class ModuleInformationWFS
	#tag Method, Flags = &h0
		Sub Constructor()
		  // Do nothing -- default constructor
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(mb as MemoryBlock, unicodeSavvy as Boolean)
		  // Make sure the memory block is sane
		  if mb = nil or mb.Long( 0 ) <> mb.Size then return
		  
		  // Construct the object from a MemoryBlock automatically
		  
		  // Grab the module id, but it's only used by the toolhelp functions -- not
		  // much use outside of them
		  ModuleID = mb.Long( 4 )
		  
		  // Now grab the process ID as well.
		  ProcessID = mb.Long( 8 )
		  
		  // Then comes the global reference count
		  GlobalReferenceCount = mb.Long( 12 )
		  
		  // We're also told the process reference count for this module to see how
		  // many times the process hooks into the module
		  ProcessReferenceCount = mb.Long( 16 )
		  
		  // Then the base address and module size
		  BaseAddress = mb.Long( 20 )
		  ModuleBaseSize = mb.Long( 24 )
		  
		  // The module's handle
		  ModuleHandle = mb.Long( 28 )
		  
		  // And then comes the name information
		  Const MAX_MODULE_NAME32 = 255
		  if unicodeSavvy then
		    ModuleName = mb.WString( 32 )
		  else
		    ModuleName = mb.CString( 32 )
		  end if
		  
		  // Finally, there's the module's file path
		  if unicodeSavvy then
		    ModuleItem = GetFolderItem( mb.WString( 32 + (MAX_MODULE_NAME32 + 1 ) * 2 ) )
		  else
		    ModuleItem = GetFolderItem( mb.CString( 32 + MAX_MODULE_NAME32 + 1 ) )
		  end if
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		BaseAddress As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		GlobalReferenceCount As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ModuleBaseSize As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ModuleHandle As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ModuleID As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ModuleItem As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h0
		ModuleName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		ProcessID As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ProcessReferenceCount As Integer
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="BaseAddress"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="GlobalReferenceCount"
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
			Name="ModuleBaseSize"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ModuleHandle"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ModuleID"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ModuleName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ProcessID"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ProcessReferenceCount"
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
