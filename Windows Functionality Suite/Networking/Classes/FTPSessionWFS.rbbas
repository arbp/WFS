#tag Class
Protected Class FTPSessionWFS
Inherits InternetSessionWFS
	#tag Method, Flags = &h0
		Sub Connect(url as String, username as String, password as String)
		  #if TargetWin32
		    
		    Soft Declare Function InternetConnectA Lib "WinInet" ( handle as Integer, server as CString, _
		    port as Integer, username as CString, password as CString, servic as Integer, flags as Integer, _
		    context as Integer ) as Integer
		    Soft Declare Function InternetConnectW Lib "WinInet" ( handle as Integer, server as WString, _
		    port as Integer, username as WString, password as WString, servic as Integer, flags as Integer, _
		    context as Integer ) as Integer
		    
		    // Setup our connection flags
		    dim flags as Integer
		    Const INTERNET_FLAG_PASSIVE = &h8000000
		    if Passive then flags = INTERNET_FLAG_PASSIVE
		    
		    Const INTERNET_SERVICE_FTP = 1
		    // Try to do the connection
		    if System.IsFunctionAvailable( "InternetConnectW", "WinInet" ) then
		      mFTPHandle = InternetConnectW( mInetHandle, url, Port, username, password, _
		      INTERNET_SERVICE_FTP, flags, 0 )
		    else
		      mFTPHandle = InternetConnectA( mInetHandle, url, Port, username, password, _
		      INTERNET_SERVICE_FTP, flags, 0 )
		    end if
		    
		    // Now check to make sure we were able to open the connection
		    if mFTPHandle = 0 then
		      FireException( "Could not open the FTP connection" )
		      return
		    end if
		    
		  #else
		    
		    #pragma unused url
		    #pragma unused username
		    #pragma unused password
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  // Make sure we get our inet handle
		  super.Constructor( "" )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CreateDirectory(name as String)
		  // Sanity checks
		  if mFTPHandle = 0 then
		    FireException( "Trying to create a directory while not connected" )
		    return
		  end if
		  
		  #if TargetWin32
		    
		    Soft Declare Function FtpCreateDirectoryW Lib "WinInet" ( handle as Integer, name as WString ) as Boolean
		    Soft Declare Function FtpCreateDirectoryA Lib "WinInet" ( handle as Integer, name as CString ) as Boolean
		    
		    dim success as Boolean
		    if System.IsFunctionAvailable( "FtpCreateDirectoryW", "WinInet" ) then
		      success = FtpCreateDirectoryW( mFTPHandle, name )
		    else
		      success = FtpCreateDirectoryA( mFTPHandle, name )
		    end if
		    
		    if not success then
		      FireException( "Could not create the directory" )
		      return
		    end if
		    
		  #else
		    
		    #pragma unused name
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CurrentDirectory() As String
		  // Sanity checks
		  if mFTPHandle = 0 then
		    FireException( "Trying to get the current directory while not connected" )
		    return ""
		  end if
		  
		  #if TargetWin32
		    Soft Declare Function FtpGetCurrentDirectoryW Lib "WinInet" ( handle as Integer, buf as Ptr, ByRef size as Integer ) as Boolean
		    Soft Declare Function FtpGetCurrentDirectoryA Lib "WinInet" ( handle as Integer, buf as Ptr, ByRef size as Integer ) as Boolean
		    
		    dim success as Boolean
		    
		    dim buf as new MemoryBlock( 260 * 2 )
		    dim size as Integer = buf.Size
		    if System.IsFunctionAvailable( "FtpGetCurrentDirectoryW", "WinInet" ) then
		      success = FtpGetCurrentDirectoryW( mFTPHandle, buf, size )
		      
		      if success then return buf.WString( 0 )
		    else
		      success = FtpGetCurrentDirectoryW( mFTPHandle, buf, size )
		      
		      if success then return buf.CString( 0 )
		    end if
		    
		    if not success then
		      FireException( "Could not get the current remote directory" )
		      return ""
		    end if
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CurrentDirectory(assigns name as String)
		  // Sanity checks
		  if mFTPHandle = 0 then
		    FireException( "Trying to set the current directory while not connected" )
		    return
		  end if
		  
		  #if TargetWin32
		    
		    Soft Declare Function FtpSetCurrentDirectoryW Lib "WinInet" ( handle as Integer, dir as WString ) as Boolean
		    Soft Declare Function FtpSetCurrentDirectoryA Lib "WinInet" ( handle as Integer, dir as CString ) as Boolean
		    
		    dim success as Boolean
		    if System.IsFunctionAvailable( "FtpSetCurrentDirectoryW", "WinInet" ) then
		      success = FtpSetCurrentDirectoryW( mFTPHandle, name )
		    else
		      success = FtpSetCurrentDirectoryA( mFTPHandle, name )
		    end if
		    
		    if not success then
		      FireException( "Could not set the current remote directory" )
		      return
		    end if
		    
		  #else
		    
		    #pragma unused name
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DeleteDirectory(name as String)
		  // Sanity checks
		  if mFTPHandle = 0 then
		    FireException( "Trying to delete a directory while not connected" )
		    return
		  end if
		  
		  #if TargetWin32
		    
		    Soft Declare Function FtpRemoveDirectoryW Lib "WinInet" ( handle as Integer, name as WString ) as Boolean
		    Soft Declare Function FtpRemoveDirectoryA Lib "WinInet" ( handle as Integer, name as CString ) as Boolean
		    
		    dim success as Boolean
		    if System.IsFunctionAvailable( "FtpRemoveDirectoryW", "WinInet" ) then
		      success = FtpRemoveDirectoryW( mFTPHandle, name )
		    else
		      success = FtpRemoveDirectoryA( mFTPHandle, name )
		    end if
		    
		    if not success then
		      FireException( "Could not delete the directory" )
		      return
		    end if
		    
		  #else
		    
		    #pragma unused name
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DeleteFile(name as String)
		  // Sanity checks
		  if mFTPHandle = 0 then
		    FireException( "Trying to delete a file while not connected" )
		    return
		  end if
		  
		  #if TargetWin32
		    
		    Soft Declare Function FtpDeleteFileW Lib "WinInet" ( handle as Integer, name as WString ) as Boolean
		    Soft Declare Function FtpDeleteFileA Lib "WinInet" ( handle as Integer, name as CString ) as Boolean
		    
		    dim success as Boolean
		    if System.IsFunctionAvailable( "FtpDeleteFileW", "WinInet" ) then
		      success = FtpDeleteFileW( mFTPHandle, name )
		    else
		      success = FtpDeleteFileA( mFTPHandle, name )
		    end if
		    
		    if not success then
		      FireException( "Could not delete the file" )
		      return
		    end if
		    
		  #else
		    
		    #pragma unused name
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Destructor()
		  CloseHandle( mFTPHandle )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub FindFinish()
		  // The user wants to stop this find operation
		  // early.
		  if mInternalFindHandle <> 0 then
		    CloseHandle( mInternalFindHandle )
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FindFirstFile(search as String = "") As FindFileWFS
		  // Sanity checks
		  if mFTPHandle = 0 then
		    FireException( "Trying to find the first file while not connected" )
		    return nil
		  end if
		  
		  // If we're already in the process of doing a find, then bail out
		  if mInternalFindHandle <> 0 then
		    FireException( "Can only do one find at a time." )
		    return nil
		  end if
		  
		  #if TargetWin32
		    Soft Declare Function FtpFindFirstFileW Lib "WinInet" ( handle as Integer, search as WString, data as Ptr, _
		    flags as Integer, context as Integer ) as Integer
		    Soft Declare Function FtpFindFirstFileA Lib "WinInet" ( handle as Integer, search as CString, data as Ptr, _
		    flags as Integer, context as Integer ) as Integer
		    
		    Const INTERNET_FLAG_RELOAD = &h80000000
		    dim mb as MemoryBlock
		    if System.IsFunctionAvailable( "FtpFindFirstFileW", "WinInet" ) then
		      mb = new MemoryBlock( FindFileWFS.kUnicodeSize )
		      mInternalFindHandle = FtpFindFirstFileW( mFTPHandle, search, mb, INTERNET_FLAG_RELOAD, 0 )
		    else
		      mb = new MemoryBlock( FindFileWFS.kANSISize )
		      mInternalFindHandle = FtpFindFirstFileA( mFTPHandle, search, mb, INTERNET_FLAG_RELOAD, 0 )
		    end if
		    
		    Const ERROR_NO_MORE_FILES = 18
		    if mInternalFindHandle = 0 then
		      // It could be that there just aren't any files in the directory
		      if GetLastError = ERROR_NO_MORE_FILES  then return nil
		      
		      FireException( "Could not find the first file" )
		      return nil
		    end if
		    
		    return new FindFileWFS( mb )
		    
		  #else
		    
		    #pragma unused search
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub GetFile(remoteName as String, local as FolderItem, bFailIfExists as Boolean = false)
		  // Sanity checks
		  if mFTPHandle = 0 then
		    FireException( "Trying to set the current directory while not connected" )
		    return
		  end if
		  
		  // If the local item is a directory, turn it into
		  // a file
		  if local.Directory then
		    local = local.Child( remoteName )
		  end if
		  
		  #if TargetWin32
		    
		    Soft Declare Function FtpGetFileW Lib "WinInet" ( handle as Integer, remote as WString, local as WString, _
		    fail as Boolean, attribs as Integer, flags as Integer, context as Integer ) as Boolean
		    Soft Declare Function FtpGetFileA Lib "WinInet" ( handle as Integer, remote as WString, local as WString, _
		    fail as Boolean, attribs as Integer, flags as Integer, context as Integer ) as Boolean
		    
		    dim flags as Integer = TransferType
		    
		    Const FILE_ATTRIBUTE_NORMAL = &h80
		    
		    dim success as Boolean
		    if System.IsFunctionAvailable( "FtpGetFileW", "WinInet" ) then
		      success = FtpGetFileW( mFTPHandle, remoteName, local.AbsolutePath, bFailIfExists, _
		      FILE_ATTRIBUTE_NORMAL, flags, 0 )
		    else
		      success = FtpGetFileW( mFTPHandle, remoteName, local.AbsolutePath, bFailIfExists, _
		      FILE_ATTRIBUTE_NORMAL, flags, 0 )
		    end if
		    
		    if not success then
		      FireException( "Could not get the remote file" )
		      return
		    end if
		    
		  #else
		    
		    #pragma unused bFailIfExists
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub PutFile(f as FolderItem, remoteName as String = "")
		  #if TargetWin32
		    
		    // Sanity checks
		    if f.Directory then
		      FireException( "Trying to put a directory by calling PutFile" )
		      return
		    end if
		    
		    if mFTPHandle = 0 then
		      FireException( "Trying to put a file while not connected" )
		      return
		    end if
		    
		    // The first thing we need to do is make
		    // sure that our local directory is correct.
		    SetLocalDirectory( f.Parent )
		    
		    // Setup our support code
		    if remoteName = "" then remoteName = f.Name
		    
		    dim flags as Integer = TransferType
		    
		    Soft Declare Function FtpPutFileA Lib "WinInet" ( handle as Integer, localFile as CString, remoteFile as CString, _
		    flags as Integer, context as Integer ) as Boolean
		    Soft Declare Function FtpPutFileW Lib "WinInet" ( handle as Integer, localFile as WString, remoteFile as WString, _
		    flags as Integer, context as Integer ) as Boolean
		    
		    // Now that we've done that, we can upload the file
		    dim success as Boolean
		    if System.IsFunctionAvailable( "FtpPutFileW", "WinInet" ) then
		      success = FtpPutFileW( mFTPHandle, f.Name, remoteName, flags, 0 )
		    else
		      success = FtpPutFileA( mFTPHandle, f.Name, remoteName, flags, 0 )
		    end if
		    
		    if not success then
		      FireException( "Could not upload the file" )
		      return
		    end if
		    
		  #else
		    
		    #pragma unused f
		    #pragma unused remoteName
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub PutFolder(f as FolderItem)
		  // This is a helper function that puts an entire
		  // folder of data (recursively) onto the server
		  if f = nil then return
		  
		  if f.Directory then
		    // We want to create a directory
		    CreateDirectory( f.Name )
		    
		    // Advance into the directory
		    CurrentDirectory = f.Name
		    
		    // Loop over all of our children and put
		    // them into the new directory
		    dim count as Integer = f.Count
		    for i as Integer = 1 to count
		      PutFolder( f.TrueItem( i ) )
		    next i
		    
		    // Then go back down a directory
		    CurrentDirectory = ".."
		  else
		    // Just put the file into the current folder
		    PutFile( f )
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub RenameFile(existing as String, newName as String)
		  // Sanity checks
		  if mFTPHandle = 0 then
		    FireException( "Trying to get the rename a file while not connected" )
		    return
		  end if
		  
		  #if TargetWin32
		    Soft Declare Function FtpRenameFileW Lib "WinInet" ( handle as Integer, old as WString, newName as WString ) as Boolean
		    Soft Declare Function FtpRenameFileA Lib "WinInet" ( handle as Integer, old as CString, newName as CString ) as Boolean
		    
		    dim success as Boolean
		    if System.IsFunctionAvailable( "FtpRenameFileW", "WinInet" ) then
		      success = FtpRenameFileW( mFTPHandle, existing, newName )
		    else
		      success = FtpRenameFileA( mFTPHandle, existing, newName )
		    end if
		    
		    if not success then
		      FireException( "Could not rename the file" )
		      return
		    end if
		    
		  #else
		    
		    #pragma unused existing
		    #pragma unused newName
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetLocalDirectory(dir as FolderItem)
		  #if TargetWin32
		    
		    Soft Declare Function SetCurrentDirectoryA Lib "Kernel32" ( dir as CString ) as Boolean
		    Soft Declare Function SetCurrentDirectoryW Lib "Kernel32" ( dir as WString ) as Boolean
		    
		    // Sanity check
		    if not dir.Directory then
		      FireException( "Trying to set the local directory to a non-directory FolderItem" )
		      return
		    end if
		    
		    // Set the directory
		    dim success as Boolean
		    if System.IsFunctionAvailable( "SetCurrentDirectoryW", "Kernel32" ) then
		      success = SetCurrentDirectoryW( dir.AbsolutePath )
		    else
		      success = SetCurrentDirectoryA( dir.AbsolutePath )
		    end if
		    
		  #else
		    
		    #pragma unused dir
		    
		  #endif
		End Sub
	#tag EndMethod


	#tag Note, Name = About this class
		This class derives from the InternetSession class so that
		we can don't have to reinvent the wheel.  So be certain
		to look on that class for some functionality as well (such as
		FindNextFile).
		
		This class should work on all versions of Windows.
		
		Note that the current implementation is synchronous in
		nature and does all error handling via Exceptions.  A future
		implementation could be implemented asynchronously (with
		modifications to the InternetSession class) as well.
	#tag EndNote


	#tag Property, Flags = &h1
		Protected mFTPHandle As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Passive As Boolean = true
	#tag EndProperty

	#tag Property, Flags = &h0
		Port As Integer = 21
	#tag EndProperty

	#tag Property, Flags = &h0
		TransferType As Integer = &h2
	#tag EndProperty


	#tag Constant, Name = kTransferTypeASCII, Type = Double, Dynamic = False, Default = \"&h1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kTransferTypeBinary, Type = Double, Dynamic = False, Default = \"&h2", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			InheritedFrom="InternetSession"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="InternetSession"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InheritedFrom="InternetSession"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Passive"
			Group="Behavior"
			InitialValue="true"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Port"
			Group="Behavior"
			InitialValue="21"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InheritedFrom="InternetSession"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="InternetSession"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TransferType"
			Group="Behavior"
			InitialValue="kTransferTypeBinary"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
