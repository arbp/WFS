#tag Module
Protected Module IniFileWFS
	#tag Method, Flags = &h1
		Protected Function ReadEntry(iniFile as folderItem, Key As string, Name As string) As string
		  #pragma unused iniFile
		  #pragma unused Key
		  #pragma unused Name
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub WriteEntry(iniFile as folderItem, Key As String, Name As String, Text As String)
		  #if TargetWin32
		    
		    Soft Declare Function WritePrivateProfileStringA Lib "kernel32"  ( lpApplicationName As CString, lpKeyName As CString, lpString As CString, lpFileName As CString ) As integer
		    Soft Declare Function WritePrivateProfileStringW Lib "kernel32"  ( lpApplicationName As WString, lpKeyName As WString, lpString As WString, lpFileName As WString ) As integer
		    
		    Dim intLen As Integer
		    if System.IsFunctionAvailable ( "GetPrivateProfileStringW", "kernel32" ) then
		      intLen = WritePrivateProfileStringW( Key, Name, Text, iniFile.AbsolutePath )
		    ElseIf  System.IsFunctionAvailable ( "GetPrivateProfileStringA", "kernel32" ) then
		      intLen = WritePrivateProfileStringA( Key, Name, Text, iniFile.AbsolutePath )
		    end if
		    
		  #else
		    
		    #pragma unused iniFile
		    #pragma unused Key
		    #pragma unused Name
		    #pragma unused Text
		    
		  #endif
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
End Module
#tag EndModule
