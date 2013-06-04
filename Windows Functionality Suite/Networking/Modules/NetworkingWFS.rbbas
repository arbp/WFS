#tag Module
Protected Module NetworkingWFS
	#tag Method, Flags = &h1
		Protected Function MapNetworkDrive(remotePath as String, localPath as String, userName as String = "", password as String = "", interactive as Boolean = true) As FolderItem
		  // We want to map the network drive the user gave us, which is in UNC format (like //10.10.10.116/foobar)
		  // and map it to the local drive they gave us (like f:).
		  
		  Soft Declare Function WNetAddConnection2A Lib "Mpr" ( netRes as Ptr, password as CString, userName as CString, flags as Integer ) as Integer
		  Soft Declare Function WNetAddConnection2W Lib "Mpr" ( netRes as Ptr, password as WString, userName as WString, flags as Integer ) as Integer
		  
		  dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "WNetAddConnection2W", "Mpr" )
		  
		  Const CONNECT_INTERACTIVE = &h8
		  Const RESOURCETYPE_DISK = &h1
		  
		  // Create and set up our network resource structure
		  dim netRes as new MemoryBlock( 30 )
		  netRes.Long( 4 ) = RESOURCETYPE_DISK
		  
		  dim localName as new MemoryBlock( 1024 )
		  dim remoteName as new MemoryBlock( 1024 )
		  if unicodeSavvy then
		    localName.WString( 0 ) = localPath
		    remoteName.WString( 0 ) = remotePath
		  else
		    locaLName.CString( 0 ) = localPath
		    remoteName.CString( 0 ) = remotePath
		  end if
		  netRes.Ptr( 16 ) = localName
		  netRes.Ptr( 20 ) = remoteName
		  
		  dim flags As Integer
		  if interactive then flags = flags + CONNECT_INTERACTIVE
		  
		  // Now make the call
		  dim ret as Integer
		  
		  if unicodeSavvy then
		    ret = WNetAddConnection2W( netRes, password, userName, flags )
		  else
		    ret = WNetAddConnection2A( netRes, password, userName, flags )
		  end if
		  
		  Const NO_ERROR = 0
		  if ret = NO_ERROR then
		    return new FolderItem( localPath )
		  else
		    return nil
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MapNetworkDriveDialog(owner as Window) As FolderItem
		  Soft Declare Function WNetConnectionDialog1W Lib "Mpr" ( dlgstruct as Ptr ) as Integer
		  Soft Declare Function WNetConnectionDialog1A Lib "Mpr" ( dlgstruct as Ptr ) as Integer
		  
		  dim dlgstruct as new MemoryBlock( 20 )
		  dim netRsrc as new MemoryBlock( 30 )
		  
		  Const RESOURCETYPE_DISK = &h1
		  
		  netRsrc.Long( 4 ) = RESOURCETYPE_DISK
		  
		  dlgStruct.Long( 0 ) = dlgstruct.Size
		  'dlgStruct.Long( 4 ) = owner.WinHWND
		  dlgStruct.Long( 4 ) = owner.Handle
		  dlgstruct.Ptr( 8 ) = netRsrc
		  
		  // Now make the call
		  dim ret as Integer
		  if System.IsFunctionAvailable( "WNetConnectionDialog1W", "Mpr" ) then
		    ret = WNetConnectionDialog1W( dlgstruct )
		  else
		    ret = WNetConnectionDialog1A( dlgstruct )
		  end if
		  
		  if ret = 0 then
		    // The drive letter is stored in the dlgstruct as an integer.  1 = a, 2 = b, etc
		    dim drive as String = Chr( 65 + dlgstruct.Long( 16 ) - 1 ) + ":"
		    return new FolderItem( drive )
		  else
		    return nil
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Ping(addy as String) As double
		  #if TargetWin32
		    
		    Declare Function IcmpCreateFile Lib "ICMP" ( ) as Integer
		    Declare Sub IcmpCloseHandle Lib "ICMP" ( handle as Integer )
		    Declare Function IcmpSendEcho Lib "ICMP" ( handle as Integer, address as Integer, data as Integer, _
		    size as Integer, options as Ptr, reply as Ptr, replySize as Integer, timeout as Integer ) as Integer
		    Declare Function inet_addr Lib "ws2_32" ( addr as CString ) as Integer
		    Declare Function gethostbyname Lib "ws2_32" ( addr as CString ) as Ptr
		    Declare Sub WSAStartup Lib "ws2_32" ( versRequest as Integer, data as Ptr )
		    Declare Sub WSACleanup Lib "ws2_32" ()
		    
		    ' Initialize WinSock
		    dim mb as new MemoryBlock( 256 + 128 + 8 + 4 )
		    
		    WSAStartup( &h0101, mb )
		    
		    if mb.Short( 0 ) <> &h0101 then
		      WSACleanup
		      return -1.0
		    end
		    
		    dim address as Integer
		    address = inet_addr( addy )
		    dim addrMemBlock as MemoryBlock
		    dim addrTemp as new MemoryBlock( 16 )
		    
		    if address = -1 then
		      ' we couldn't resolve the address that way, so we need to try a
		      ' different approach
		      addrMemBlock = gethostbyname( addy )
		      address = addrMemBlock.Ptr( 12 ).Ptr( 0 ).Long( 0 )
		      'addrTemp.StringValue( 0, addrTemp.Size ) = addrMemBlock.StringValue( 0, 16 )
		    end
		    
		    ' Create the ICMP handle
		    dim icmpFile as Integer = IcmpCreateFile
		    
		    ' Set up the ping structure
		    dim pingInfo as new MemoryBlock( 8 )
		    pingInfo.Byte( 0 ) = 255
		    
		    ' Send the ping
		    dim reply as new MemoryBlock( 28 )
		    dim ret as Integer
		    ret = IcmpSendEcho( icmpFile, address, 0, 0,pingInfo, reply, 28, 5000 )
		    
		    ' Clean everything up
		    WSACleanup
		    IcmpCloseHandle( icmpFile )
		    
		    ' return the proper value
		    if ret <> 0 then
		      return reply.Long( 8 )
		    else
		      return -1.0
		    end
		    
		  #else
		    
		    #pragma unused addy
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub TransmitFile(extends tcp as TCPSocket, bs as BinaryStream)
		  Soft Declare Sub TransmitFile Lib "mswsock" ( socket as Integer, file as Integer, _
		  bytesToWrite as Integer, bytesPerSend as Integer, _
		  overlapped as Integer, transmitBuffers as Integer, flags as Integer )
		  
		  if System.IsFunctionAvailable( "TransmitFile", "mswsock" ) then
		    Const TF_WRITE_BEHIND = &h4
		    Const TF_USE_SYSTEM_THREAD = &h10
		    Const TF_DISCONNECT = &h1
		    
		    if bs = nil then return
		    
		    TransmitFile( tcp.Handle, _
		    bs.Handle( BinaryStream.HandleTypeWin32Handle ), 0, 0, 0, 0, _
		    TF_WRITE_BEHIND + TF_USE_SYSTEM_THREAD )
		    
		    // Note that we have to keep bs open because the thread is going
		    // to send the file asynchronously.  This makes things kind of tough
		    // since there's also no notification of when the file can safely be closed
		    // either.
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function UnmapNetworkDrive(drive as String, force as Boolean = false) As Boolean
		  Soft Declare Function WNetCancelConnection2W Lib "Mpr" ( name as WString, flags as Integer, force as Boolean ) as Integer
		  Soft Declare Function WNetCancelConnection2A Lib "Mpr" ( name as CString, flags as Integer, force as Boolean ) as Integer
		  
		  if System.IsFunctionAvailable( "WNetCancelConnection2W", "Mpr" ) then
		    return WNetCancelConnection2W( drive, 0, force ) = 0
		  else
		    return WNetCancelConnection2A( drive, 0, force ) = 0
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function UnmapNetworkDriveDialog(owner as Window, localPath as String) As Boolean
		  Soft Declare Function WNetDisconnectDialog1W Lib "Mpr" ( dlgstruct as Ptr ) as Integer
		  Soft Declare Function WNetDisconnectDialog1A Lib "Mpr" ( dlgstruct as Ptr ) as Integer
		  
		  dim dlgstruct as new MemoryBlock( 20 )
		  dlgstruct.Long( 0 ) = dlgstruct.Size
		  'dlgstruct.Long( 4 ) = owner.WinHWND
		  dlgstruct.Long( 4 ) = owner.Handle
		  
		  if Right( localPath, 1 ) = "\" then localPath = Left( localPath, 2 )
		  
		  dim localName as new MemoryBlock( 1024 )
		  if System.IsFunctionAvailable( "WNetDisconnectDialog1W", "Mpr" ) then
		    localName.WString( 0 ) = localPath
		  else
		    localName.CString( 0 ) = localPath
		  end if
		  
		  dlgstruct.Ptr( 8 ) = localName
		  
		  dim ret as Integer
		  if System.IsFunctionAvailable( "WNetDisconnectDialog1W", "Mpr" ) then
		    ret = WNetDisconnectDialog1W( dlgstruct )
		  else
		    ret = WNetDisconnectDialog1A( dlgstruct )
		  end if
		  
		  return ret = 0
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
