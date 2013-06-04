#tag Class
Protected Class FindFileWFS
	#tag Method, Flags = &h0
		Sub Constructor(mb as MemoryBlock)
		  // We need to convert the data in the
		  // MemoryBlock into something human-readable.
		  
		  // Are we unicode or not?
		  Dim isUnicode as Boolean = mb.Size = kUnicodeSize
		  
		  FileAttributes = mb.Long( 0 )
		  CreationDate = FileTimeToDate( mb.Long( 4 ), mb.Long( 8 ) )
		  AccessDate = FileTimeToDate( mb.Long( 12 ), mb.Long( 16 ) )
		  WriteDate = FileTimeToDate( mb.Long( 20 ), mb.Long( 24 ) )
		  
		  // Now we have the file size
		  dim lo, hi as Integer
		  hi = mb.Long( 28 )
		  lo = mb.Long( 32 )
		  
		  #if RBVersion < 2006.01 then
		    // NOTE!!  This is wrong, we need UInt32 support
		    // before this will work properly.
		    Length = LongLongToDouble( lo, hi )
		  #else
		    // 2006r1 and higher has direct support
		    // for UInt64s so we can just convert to a
		    // double directly
		    Length = mb.UInt64Value( 28 )
		  #endif
		  
		  // There are now 8 bytes of reserved data
		  
		  if isUnicode then
		    Name = mb.WString( 44 )
		    AlternateName = mb.WString( 234 )
		  else
		    Name = mb.CString( 44 )
		    AlternateName = mb.CString( 130 )
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function FileTimeToDate(lo as Integer, hi as Integer) As Date
		  // Now we have a FILETIME structure, which is 8
		  // bytes of lo followed by hi.  It's in 100-nanosecond
		  // intervals since Jan 01, 1601 (UTC)
		  
		  // And instead of doing very hard math conversions, we
		  // will just get a SYSTEMTIME structure from it, and fill out
		  // our RB Date object from that.  Yay!
		  #if TargetWin32
		    Declare Function FileTimeToSystemTime Lib "Kernel32" ( ft as Ptr, sysTime as Ptr ) as Boolean
		    
		    dim ft as new MemoryBlock( 8 )
		    ft.Long( 0 ) = lo
		    ft.Long( 4 ) = hi
		    
		    dim st as new MemoryBlock( 16 )
		    
		    if FileTimeToSystemTime( ft, st ) then
		      dim d as new Date
		      
		      d.Year = st.Short( 0 )
		      d.Month = st.Short( 2 )
		      d.Day = st.Short( 6 )
		      d.Hour = st.Short( 8 )
		      d.Minute = st.Short( 10 )
		      d.Second = st.Short( 12 )
		      
		      return d
		    else
		      return nil
		    end if
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function LongLongToDouble(lo as Integer, hi as Integer) As Double
		  dim ret as Double
		  
		  // Take the high 4 bytes and shift them
		  ret = hi * (2 ^32)
		  
		  // Then add in the low 4 bytes
		  ret = ret + lo
		  
		  return ret
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TestAttribute(attrib as Integer) As Boolean
		  return Bitwise.BitAnd( FileAttributes, attrib ) = attrib
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		AccessDate As Date
	#tag EndProperty

	#tag Property, Flags = &h0
		AlternateName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		CreationDate As Date
	#tag EndProperty

	#tag Property, Flags = &h0
		FileAttributes As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Length As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		Name As String
	#tag EndProperty

	#tag Property, Flags = &h0
		WriteDate As Date
	#tag EndProperty


	#tag Constant, Name = kANSISize, Type = Double, Dynamic = False, Default = \"140", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kArchive, Type = Double, Dynamic = False, Default = \"&h20", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCompressed, Type = Double, Dynamic = False, Default = \"&h800", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kDirectory, Type = Double, Dynamic = False, Default = \"&h10", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kEncrypted, Type = Double, Dynamic = False, Default = \"&h40", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kHidden, Type = Double, Dynamic = False, Default = \"&h2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kNormal, Type = Double, Dynamic = False, Default = \"&h80", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReadOnly, Type = Double, Dynamic = False, Default = \"&h1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kSystem, Type = Double, Dynamic = False, Default = \"&h4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kTemporary, Type = Double, Dynamic = False, Default = \"&h100", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kUnicodeSize, Type = Double, Dynamic = False, Default = \"250", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="AlternateName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="FileAttributes"
			Group="Behavior"
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
			Name="Length"
			Group="Behavior"
			InitialValue="0"
			Type="Double"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
			EditorType="MultiLineEditor"
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
