#tag Module
Protected Module InternetConnection
	#tag Method, Flags = &h1
		Protected Function Configured() As Boolean
		  return Bitwise.BitAnd( Query, &h40 ) <> 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function LAN() As Boolean
		  return Bitwise.BitAnd( Query, &h2 ) <> 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Modem() As Boolean
		  return Bitwise.BitAnd( Query, &h1 ) <> 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ModemBusy() As Boolean
		  return Bitwise.BitAnd( Query, &h8 ) <> 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function NoConnection() As Boolean
		  if Query = 0 then return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Offline() As Boolean
		  return Bitwise.BitAnd( Query, &h20 ) <> 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Proxy() As Boolean
		  return Bitwise.BitAnd( Query, &h4 ) <> 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function Query() As Integer
		  Soft Declare Function InternetGetConnectedState Lib "WinInet" ( ByRef state as Integer, reserved as Integer ) as Boolean
		  
		  if System.IsFunctionAvailable( "InternetGetConnectedState", "WinInet" ) then
		    Dim n as Integer
		    if not InternetGetConnectedState( n, 0 ) then return 0 else return n
		  end if
		  
		  return 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function RASInstalled() As Boolean
		  return Bitwise.BitAnd( Query, &h10 ) <> 0
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
