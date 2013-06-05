#tag Module
Protected Module FolderItemExtensionsWFS
	#tag Method, Flags = &h0
		Sub AddToRecentItemsWFS(extends f as FolderItem)
		  // We want to add this folder item to the recent docs menu
		  #if TargetWin32
		    
		    Soft Declare Sub SHAddToRecentDocs Lib "Shell32" ( type as Integer, path as Ptr )
		    
		    Const SHARED_PATHA = 2
		    Const SHARED_PATHW = 3
		    
		    if System.IsFunctionAvailable( "SHAddToRecentDocs", "Shell32" ) then
		      dim path as MemoryBlock
		      dim type as Integer
		      // This is just a cheap hack to tell whether we're on NT or not without
		      // having to rely on the Win32DeclareLibrary module
		      if System.IsFunctionAvailable( "CreateWindowExW", "User32" ) then
		        path = DefineEncoding( f.AbsolutePath + Chr( 0 ), Encodings.UTF16 )
		        type = SHARED_PATHW
		      else
		        path = DefineEncoding( f.AbsolutePath + Chr( 0 ), Encodings.ASCII )
		        type = SHARED_PATHA
		      end if
		      
		      SHAddToRecentDocs( type, path )
		    end if
		    
		  #else
		    
		    #pragma unused f
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AssociateExtensionWFS(extends f as FolderItem, set as Boolean)
		  ' We wanna make two different registry items in CLASSES_ROOT
		  ' to deal with this.
		  
		  dim theExtension, theAppName as String
		  dim numDots as Integer
		  
		  numDots = CountFields( f.Name, "." )
		  theExtension = "." + nthField( f.Name, ".", numDots )
		  theAppName = nthField( App.ExecutableFile.Name, ".", 1 )
		  
		  dim exten as new RegistryItem( "HKEY_CLASSES_ROOT" )
		  
		  if set then
		    try
		      // Try to get a folder with the extension
		      exten = exten.Child( theExtension )
		    catch
		      // The extension isn't listed in the folder, so we
		      // need to make it
		      exten = exten.AddFolder( theExtension )
		    end
		    
		    // Now we need to set the externsion to the application
		    // name
		    exten.DefaultValue = theAppName
		    
		    // Now we need to make a folder with our application
		    // name
		    exten = new RegistryItem( "HKEY_CLASSES_ROOT" )
		    
		    try
		      // Get the folder with the application name
		      exten = exten.Child( theAppName )
		    catch
		      // The app name doesn't exist, so we need
		      // to make it
		      exten = exten.AddFolder( theAppName )
		    end
		    
		    // Now we need to try opening the Shell folder from
		    // the exten folder
		    try
		      exten = exten.Child( "Shell" )
		    catch
		      // It doesn't exist, so make it
		      exten = exten.AddFolder( "Shell" )
		    end
		    
		    // And we need the Open folder as well
		    try
		      exten = exten.Child( "open" )
		    catch
		      // Make that one too
		      exten = exten.AddFolder( "open" )
		    end
		    
		    // Finally, we need to set the command property
		    // to our application
		    try
		      exten = exten.Child( "command" )
		    catch
		      // Make it as well
		      exten = exten.AddFolder( "command" )
		    end
		    
		    exten.DefaultValue = """" + App.ExecutableFile.AbsolutePath + """" + " ""%1"""
		  else
		    // We want to delete the extension folder
		    exten.Delete( theExtension )
		  end
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CopyFileToWithProgressWFS(extends fromFile as FolderItem, toFile as FolderItem)
		  // Delegate to the heavy-lifting function
		  CopyMoveOp( fromFile, toFile, false )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub CopyMoveOp(fromFile as FolderItem, toFile as FolderItem, move as Boolean)
		  #if TargetWin32
		    
		    Soft Declare Sub SHFileOperationW Lib "Shell32" ( op as Ptr )
		    Soft Declare Sub SHFileOperationA Lib "Shell32" ( op as Ptr )
		    
		    dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "SHFileOperationW", "Shell32" )
		    
		    Const FO_MOVE = &h1
		    Const FO_COPY = &h2
		    
		    dim mb as new MemoryBlock( 8 * 4 )
		    
		    #if TargetHasGUI
		      dim wnd as Window = Window( 0 )
		      if wnd <> nil then
		        mb.Long( 0 ) = wnd.Handle
		      end if
		    #endif
		    
		    if move then
		      mb.Long( 4 ) = FO_MOVE
		    else
		      mb.Long( 4 ) = FO_COPY
		    end if
		    
		    dim fromFilePath as String
		    dim fromStr as MemoryBlock
		    
		    if unicodeSavvy then
		      fromFilePath = ConvertEncoding( fromFile.AbsolutePath, Encodings.UTF16 )
		      fromStr = new MemoryBlock( LenB( fromFilePath ) + 4 )
		      fromStr.WString( 0 ) = fromFilePath
		    else
		      fromFilePath = ConvertEncoding( fromFile.AbsolutePath, Encodings.SystemDefault )
		      fromStr = new MemoryBlock( LenB( fromFilePath ) + 2 )
		      fromStr.CString( 0 ) = fromFilePath
		    end if
		    
		    mb.Ptr( 8 ) = fromStr
		    
		    dim toFilePath as String
		    dim toStr as MemoryBlock
		    if unicodeSavvy then
		      toFilePath = ConvertEncoding( toFile.AbsolutePath, Encodings.UTF16 )
		      toStr = new MemoryBlock( LenB( toFilePath ) + 4 )
		      toStr.WString( 0 ) = toFilePath
		    else
		      toFilePath = ConvertEncoding( toFile.AbsolutePath, Encodings.SystemDefault )
		      toStr = new MemoryBlock( LenB( toFilePath ) + 2 )
		      toStr.CString( 0 ) = toFilePath
		    end if
		    
		    mb.Ptr( 12 ) = toStr
		    
		    mb.Long( 16 ) = 0  // Flags
		    mb.Long( 20 ) = 0  // Byref "bool" to let you know if operation was aborted
		    mb.Long( 24 ) = 0  // Don't care about name mappings
		    mb.Ptr( 28 ) = nil  // Don't care about the progress title caption
		    
		    if unicodeSavvy then
		      SHFileOperationW( mb )
		    else
		      SHFileOperationA( mb )
		    end if
		    
		  #else
		    
		    #pragma unused fromFile
		    #pragma unused toFile
		    #pragma unused move
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DecryptNTFSFileWFS(extends item as FolderItem)
		  if item = nil then return
		  
		  #if TargetWin32
		    
		    Soft Declare Sub DecryptFileA Lib "AdvApi32" ( file as CString, zero as Integer )
		    Soft Declare Sub DecryptFileW Lib "AdvApi32" ( file as WString, zero as Integer )
		    
		    if System.IsFunctionAvailable( "DecryptFileW", "AdvApi32" ) then
		      DecryptFileW( item.AbsolutePath, 0 )
		    else
		      DecryptFileA( item.AbsolutePath, 0 )
		    end if
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DeleteOnRebootWFS(extends f as FolderItem)
		  if f is nil then return
		  
		  #if TargetWin32 then
		    
		    Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4
		    Soft Declare Sub MoveFileExW Lib "Kernel32"( OldFilename As WString, NewFileName As Integer, nWord As Integer )
		    Soft Declare Sub MoveFileExA Lib "Kernel32"( OldFilename As CString, NewFileName As Integer, nWord As Integer )
		    
		    if System.IsFunctionAvailable( "MoveFileExW", "Kernel32" ) then
		      MoveFileExW( f.AbsolutePath, 0, MOVEFILE_DELAY_UNTIL_REBOOT )
		    else
		      MoveFileExA( ConvertEncoding( f.AbsolutePath, Encodings.SystemDefault ), 0, MOVEFILE_DELAY_UNTIL_REBOOT )
		    end if
		    
		  #else
		    
		    #pragma unused f
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DoVersionJiggery(data as MemoryBlock) As String
		  // We want to hack our way thru the data given.  As evil as it
		  // seems, we're going to search for "StringFileInfo" as a WString
		  // and skip to the part where we can get the language and code
		  // page IDs.
		  
		  dim dataStr as String = data.StringValue( 0, data.Size )
		  dim strFileInfo as String = ConvertEncoding( "StringFileInfo", Encodings.UTF16 )
		  dim pos as Integer = InStrB( dataStr, strFileInfo )
		  
		  if pos > 0 then
		    // We found it in the data, so let's dive on ahead by the proper amount.
		    dim skipLen as Integer = LenB( strFileInfo ) + 8  // for the null byte
		    
		    // Get the textual information
		    dim infoStr as String = Mid( dataStr, pos + skipLen, 16 )
		    
		    // Return it
		    return DefineEncoding( infoStr, Encodings.UTF16 )
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub EmptyTrashWFS(extends f as folderItem)
		  '// This method will empty the RecycleBins on the given drive
		  '// By Anthony G. Cyphers
		  '// 05/17/2007
		  
		  #if TargetWin32 then
		    
		    while f.Parent <> nil
		      f = f.Parent
		    wend
		    
		    Soft Declare Function SHEmptyRecycleBinA Lib "shell32" ( hwnd As Integer, pszRootPath As CString, dwFlags As Integer) As Integer
		    Soft Declare Function SHEmptyRecycleBinW Lib "shell32" ( hwnd As Integer, pszRootPath As WString, dwFlags As Integer) As Integer
		    Soft Declare Function SHUpdateRecycleBinIcon Lib "shell32" () As Integer
		    
		    dim updateIcon as Boolean
		    if System.IsFunctionAvailable( "SHEmptyRecycleBinW", "shell32" ) then
		      updateIcon = SHEmptyRecycleBinW( 0, f.AbsolutePath, 0 ) = 0
		    elseif System.IsFunctionAvailable( "SHEmptyRecycleBinA", "shell32" ) then
		      updateIcon = SHEmptyRecycleBinA( 0, f.AbsolutePath, 0 ) = 0
		    end if
		    
		    if updateIcon and System.IsFunctionAvailable( "SHUpdateRecycleBinIcon", "shell32" ) then
		      Call SHUpdateRecycleBinIcon
		    end if
		    
		  #else
		    
		    #pragma unused f
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub EncryptNTFSFileWFS(extends item as FolderItem)
		  if item = nil then return
		  
		  #if TargetWin32
		    Soft Declare Sub EncryptFileA Lib "AdvApi32" ( file as CString )
		    Soft Declare Sub EncryptFileW Lib "AdvApi32" ( file as CString )
		    
		    if System.IsFunctionAvailable( "EncryptFileW", "AdvApi32" ) then
		      EncryptFileW( item.AbsolutePath )
		    else
		      EncryptFileA( item.AbsolutePath )
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetShortPathWFS(extends f as FolderItem) As string
		  #if TargetWin32
		    
		    /// This takes a long path and returns the truncated path
		    /// for example GetShortPath("c:\program files")  will return C:\progra~1
		    /// useful in shells and other places where spaces in paths are hard to deal with
		    Dim Res As integer
		    Dim TruncatedPath as new MemoryBlock( 1024 )
		    
		    Soft Declare Function GetShortPathNameA Lib "kernel32" (lpszLongPath As cString, lpszShortPath As ptr, lBuffer As integer) As integer
		    Soft Declare Function GetShortPathNameW Lib "kernel32" (lpszLongPath As WString, lpszShortPath As ptr, lBuffer As integer) As integer
		    
		    if System.IsFunctionAvailable( "GetShortPathNameW", "Kernel32" ) then
		      // We do 512 here because it's the number of characters, not bytes
		      res = GetShortPathNameW( f.AbsolutePath, TruncatedPath, 512 )
		      if res > 0 then return TruncatedPath.WString( 0 )
		    else
		      res = GetShortPathNameA( f.AbsolutePath, TruncatedPath, 1024 )
		      
		      if res > 0 then return TruncatedPath.CString( 0 )
		    end if
		    
		  #else
		    
		    #pragma unused f
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetTrashCountWFS(extends f as folderItem) As Int64
		  '// This function return the number of items for the given drive's RecycleBin
		  '// By Anthony G. Cyphers
		  '// 05/17/2007
		  
		  #if TargetWin32 then
		    
		    while f.Parent <> nil
		      f = f.Parent
		    wend
		    
		    Soft Declare Function SHQueryRecycleBinW Lib "shell32" ( pszRootPath As WString, pSHQueryRBInfo As Ptr ) As Integer
		    Soft Declare Function SHQueryRecycleBinA Lib "shell32" ( pszRootPath As CString, pSHQueryRBInfo As Ptr ) As Integer
		    
		    dim newInfo as new MemoryBlock( 20 )
		    newInfo.Long( 0 ) = newInfo.Size
		    dim x as Integer
		    if System.IsFunctionAvailable( "SHQueryRecycleBinW", "shell32" ) then
		      x = shQueryrecyclebinW( f.AbsolutePath, newInfo )
		    elseif System.IsFunctionAvailable( "SHQueryRecycleBinA", "shell32" ) then
		      x = shQueryrecyclebinA( f.AbsolutePath, newInfo )
		    end if
		    
		    if x = 0 then
		      return newInfo.Int64Value( 12 )
		    end if
		    
		    return -1
		    
		  #else
		    
		    #pragma unused f
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetTrashSizeWFS(extends f as FolderItem) As Int64
		  '// This function will return the size of the given RecycleBin in bytes.
		  '// By Anthony G. Cyphers
		  '// 05/17/2007
		  
		  #if TargetWin32 then
		    
		    while f.Parent <> nil
		      f = f.Parent
		    wend
		    
		    Soft Declare Function SHQueryRecycleBinW Lib "shell32" ( pszRootPath As WString, pSHQueryRBInfo As Ptr ) As Integer
		    Soft Declare Function SHQueryRecycleBinA Lib "shell32" ( pszRootPath As CString, pSHQueryRBInfo As Ptr ) As Integer
		    
		    dim newInfo as new MemoryBlock( 20 )
		    newInfo.Long( 0 ) = newInfo.Size
		    dim x as Integer
		    if System.IsFunctionAvailable( "SHQueryRecycleBinW", "shell32" ) then
		      x = shQueryrecyclebinW( f.AbsolutePath, newInfo )
		    elseif System.IsFunctionAvailable( "SHQueryRecycleBinA", "shell32" ) then
		      x = shQueryrecyclebinA( f.AbsolutePath, newInfo )
		    end if
		    
		    if x = 0 then
		      return newInfo.Int64Value( 4 )
		    end if
		    
		  #else
		    
		    #pragma unused f
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetVersionInformationWFS(extends f as FolderItem) As VersionInformationWFS
		  #if TargetWin32
		    
		    Soft Declare Function GetFileVersionInfoA lib "Version" ( fileName as CString, ignored as Integer, len as Integer, buffer as Ptr) as Boolean
		    Soft Declare Function GetFileVersionInfoW lib "Version" ( fileName as WString, ignored as Integer, len as Integer, buffer as Ptr) as Boolean
		    Soft Declare Function GetFileVersionInfoSizeA lib "Version" ( fileName as CString, ByRef ignored as Integer ) as Integer
		    Soft Declare Function GetFileVersionInfoSizeW lib "Version" ( fileName as WString, ByRef ignored as Integer ) as Integer
		    Soft Declare Function VerQueryValueA lib "Version" ( info as Ptr, subBlock as CString, value as Ptr, ByRef len as Integer ) as Boolean
		    Soft Declare Function VerQueryValueW lib "Version" ( info as Ptr, subBlock as WString, value as Ptr, ByRef len as Integer ) as Boolean
		    
		    dim isUnicode as Boolean = System.IsFunctionAvailable( "GetFileVersionInfoW", "Version" )
		    
		    // First, get the size of the version information
		    dim size as Integer
		    if isUnicode then
		      dim ignored as Integer
		      size = GetFileVersionInfoSizeW( f.AbsolutePath, ignored )
		    else
		      dim ignored as Integer
		      size = GetFileVersionInfoSizeA( ConvertEncoding( f.AbsolutePath, Encodings.SystemDefault ), ignored )
		    end if
		    
		    // If our size is legit, then we know how much data to allocate
		    if size = 0 then return nil
		    
		    // Allocate a buffer big enough to fit all of our data
		    dim buffer as new MemoryBlock( size )
		    
		    // Now obtain the data itself
		    dim success as Boolean
		    if isUnicode then
		      success = GetFileVersionInfoW( f.AbsolutePath, 0, size, buffer )
		    else
		      success = GetFileVersionInfoA( ConvertEncoding( f.AbsolutePath, Encodings.SystemDefault ), 0, size, buffer )
		    end if
		    
		    // If we couldn't get the data, then bail out
		    if not success then return nil
		    
		    // Now we want to find the language and code page information
		    dim langInfoPtr as new MemoryBlock( 4 )
		    dim langInfoLen as Integer
		    if isUnicode then
		      success = VerQueryValueW( buffer, "\VarFileInfo\Translation", langInfoPtr, langInfoLen )
		    else
		      success = VerQueryValueA( buffer, "\VarFileInfo\Translation", langInfoPtr, langInfoLen )
		    end if
		    
		    dim langCodePage as String
		    if not success then
		      // It's possible that there's no translation table, in which case we will just
		      // fake it by using the system's language and code page ID's.  According to
		      // MSDN, the Translation table should always be present and always provide
		      // you with correct information.  However, I've found that some linkers decide
		      // to exclude this information, such as CodeWarrior.  Yee haw.  So if we can't
		      // find the translation table, then we want to do some jiggery to find the first
		      // language and code page that's in the structure.
		      langCodePage = DoVersionJiggery( buffer )
		    else
		      // Now that we have language and codepage information, we want to
		      // find the language and codepage that match the user's
		      langCodePage = PadHexToFourDigits( langInfoPtr.Ptr( 0 ).Short( 0 ) ) + _
		      PadHexToFourDigits( langInfoPtr.Ptr( 0 ).Short( 2 ) )
		    end if
		    
		    dim data as new MemoryBlock( 4 )
		    dim ret as new VersionInformationWFS
		    dim len as Integer
		    
		    // Grab all the info we can
		    if isUnicode then
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\Comments", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.Comments = mb.WString( 0 )
		      end if
		    else
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\Comments", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.Comments = mb.CString( 0 )
		      end if
		      
		    end if
		    
		    if isUnicode then
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\CompanyName", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.CompanyName = mb.WString( 0 )
		      end if
		    else
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\CompanyName", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.CompanyName = mb.CString( 0 )
		      end if
		      
		    end if
		    
		    if isUnicode then
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\FileDescription", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.FileDescription = mb.WString( 0 )
		      end if
		    else
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\FileDescription", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.FileDescription = mb.CString( 0 )
		      end if
		      
		    end if
		    
		    if isUnicode then
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\FileVersion", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.FileVersion = mb.WString( 0 )
		      end if
		    else
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\FileVersion", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.FileVersion = mb.CString( 0 )
		      end if
		      
		    end if
		    
		    if isUnicode then
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\InternalName", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.InternalName = mb.WString( 0 )
		      end if
		    else
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\InternalName", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.InternalName = mb.CString( 0 )
		      end if
		      
		    end if
		    
		    if isUnicode then
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\LegalCopyright", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.LegalCopyright = mb.WString( 0 )
		      end if
		    else
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\LegalCopyright", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.LegalCopyright = mb.CString( 0 )
		      end if
		    end if
		    
		    if isUnicode then
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\LegalTrademarks", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.LegalTrademarks = mb.WString( 0 )
		      end if
		    else
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\LegalTrademarks", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.LegalTrademarks = mb.CString( 0 )
		      end if
		      
		    end if
		    
		    if isUnicode then
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\OriginalFilename", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.OriginalFilename = mb.WString( 0 )
		      end if
		    else
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\OriginalFilename", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.OriginalFilename = mb.CString( 0 )
		      end if
		      
		    end if
		    
		    if isUnicode then
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\ProductName", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.ProductName = mb.WString( 0 )
		      end if
		    else
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\ProductName", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.ProductName = mb.CString( 0 )
		      end if
		    end if
		    
		    if isUnicode then
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\ProductVersion", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.ProductVersion = mb.WString( 0 )
		      end if
		    else
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\ProductVersion", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.ProductVersion = mb.WString( 0 )
		      end if
		    end if
		    
		    if isUnicode then
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\PrivateBuild", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.PrivateBuild = mb.WString( 0 )
		      end if
		    else
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\PrivateBuild", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.PrivateBuild = mb.CString( 0 )
		      end if
		      
		    end if
		    
		    if isUnicode then
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\SpecialBuild", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.SpecialBuild = mb.WString( 0 )
		      end if
		    else
		      if VerQueryValueW( buffer, "\StringFileInfo\" + langCodePage + "\SpecialBuild", data, len ) then
		        dim mb as MemoryBlock = data.Ptr( 0 ).StringValue( 0, len )
		        ret.SpecialBuild = mb.CString( 0 )
		      end if
		    end if
		    
		    Break
		Exception err as NilObjectException
		  return nil
		  
		  #else
		    
		    #pragma unused f
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function HasChangedWFS(extends f as FolderItem) As Boolean
		  #if TargetWin32
		    
		    // Check to see if this folder item has a handle in the map
		    dim handle as Integer
		    try
		      handle = mCHangeHandles.Value( f.AbsolutePath )
		    catch
		      // If this didn't work, then something's wrong!  Either there's
		      // no map, or the folder item isn't in it.  Either way, bail out
		      return false
		    end
		    
		    // Now that we have a handle, let's try to see if anything's changed
		    // with it
		    Declare Function WaitForSingleObject Lib "Kernel32" ( handle as Integer, waitTime as Integer ) as Integer
		    dim status as Integer
		    status = WaitForSingleObject( handle, 0 )
		    
		    // If we've had a change, the status will be WAIT_OBJECT_0
		    Const WAIT_OBJECT_0 = &h0
		    if status = WAIT_OBJECT_0 then
		      // We've had a change!  Reset the handle so we can use it again
		      Declare Sub FindNextChangeNotification Lib "Kernel32" ( handle as Integer )
		      FindNextChangeNotification( handle )
		      
		      // And return the change
		      return true
		    end
		    
		    // If we got here, no change has taken place!
		    return false
		    
		  #else
		    
		    #pragma unused f
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsExtensionAssociatedWFS(extends f as FolderItem) As Boolean
		  ' We wanna make two different registry items in CLASSES_ROOT
		  ' to deal with this.
		  
		  dim theExtension, theAppName as String
		  dim numDots as Integer
		  
		  numDots = CountFields( f.Name, "." )
		  theExtension = "." + nthField( f.Name, ".", numDots )
		  theAppName = nthField( App.ExecutableFile.Name, ".", 1 )
		  
		  dim exten as new RegistryItem( "HKEY_CLASSES_ROOT" )
		  
		  try
		    // Try to get a folder with the extension
		    exten = exten.Child( theExtension )
		  catch
		    // The folder doesn't exist, so we're SO not associated
		    return false
		  end
		  
		  // Now we need to set the externsion to the application
		  // name
		  return (exten.DefaultValue = theAppName)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsStartupItemWFS(extends f as FolderItem, machineWide as Boolean) As Boolean
		  #if TargetWin32
		    
		    dim runItem as RegistryItem
		    
		    if machineWide then
		      runItem = new RegistryItem( "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\" )
		    else
		      runItem = new RegistryItem( "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\" )
		    end
		    
		    
		    dim theAppName as String
		    theAppName =  nthField( f.Name, ".", 1 )
		    
		    return (runItem.Value( theAppName ) <> "")
		    
		  #else
		    
		    #pragma unused f
		    #pragma unused machineWide
		    
		  #endif
		  
		exception
		  return false
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LaunchAndWait(extends f as FolderItem, params as String = "", deskName as String = "")
		  // We want to launch the application and wait for
		  // it to finish executing before we return.
		  #if TargetWin32
		    
		    Soft Declare Function CreateProcessW Lib "Kernel32" ( appName as WString, params as WString, _
		    procAttribs as Integer, threadAttribs as Integer, inheritHandles as Boolean, flags as Integer, _
		    env as Integer, curDir as Integer, startupInfo as Ptr, procInfo as Ptr ) as Boolean
		    
		    Soft Declare Function CreateProcessA Lib "Kernel32" ( appName as CString, params as CString, _
		    procAttribs as Integer, threadAttribs as Integer, inheritHandles as Boolean, flags as Integer, _
		    env as Integer, curDir as Integer, startupInfo as Ptr, procInfo as Ptr ) as Boolean
		    
		    dim startupInfo, procInfo as MemoryBlock
		    
		    startupInfo = new MemoryBlock( 17 * 4 )
		    procInfo = new MemoryBlock( 16 )
		    
		    dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "CreateProcessW", "Kernel32" )
		    
		    startupInfo.Long( 0 ) = startupInfo.Size
		    
		    dim deskStr, deskPtr as MemoryBlock
		    if deskName <> "" then
		      if unicodeSavvy then
		        deskStr = ConvertEncoding( deskName + Chr( 0 ), Encodings.UTF16 )
		      else
		        deskStr = deskName + Chr( 0 )
		      end if
		      
		      startupInfo.Ptr( 8 ) = deskStr
		    end if
		    
		    dim ret as Boolean
		    if unicodeSavvy then
		      ret = CreateProcessW( f.AbsolutePath, params, 0, 0, false, 0, 0, 0, startupInfo, procInfo )
		    else
		      ret = CreateProcessA( f.AbsolutePath, params, 0, 0, false, 0, 0, 0, startupInfo, procInfo )
		    end if
		    
		    if not ret then return
		    
		    Declare Function WaitForSingleObject Lib "Kernel32" ( handle as Integer, howLong as Integer ) as Integer
		    Const INFINITE = -1
		    Const WAIT_TIMEOUT = &h00000102
		    Const WAIT_OBJECT_0 = &h0
		    
		    // We want to loop here so that we can yield time back
		    // to other threads.  This means threaded applications
		    // will "just work", but non-threaded ones will still appear hung.
		    while WaitForSingleObject( procInfo.Long( 0 ), 1 ) = WAIT_TIMEOUT
		      App.SleepCurrentThread( 10 )
		    wend
		    
		    Declare Sub CloseHandle Lib "Kernel32" ( handle as Integer )
		    CloseHandle( procInfo.Long( 0 ) )
		    CloseHandle( procInfo.Long( 4 ) )
		    
		  #else
		    
		    #pragma unused f
		    #pragma unused params
		    #pragma unused deskName
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LaunchAsAdministrator(extends f as FolderItem, ParamArray args as String)
		  #if TargetWin32
		    
		    Soft Declare Sub ShellExecuteA Lib "Shell32" ( hwnd as Integer, operation as CString, _
		    file as CString, params as CString, directory as CString, show as Integer )
		    Soft Declare Sub ShellExecuteW Lib "Shell32" ( hwnd as Integer, operation as WString, _
		    file as WString, params as WString, directory as WString, show as Integer )
		    
		    dim params as String
		    params = Join( args, " " )
		    
		    if System.IsFunctionAvailable( "ShellExecuteW", "Shell32" ) then
		      ShellExecuteW( 0, "runas", f.AbsolutePath, params, "", 1 )
		    else
		      ShellExecuteA( 0, "runas", f.AbsolutePath, params, "", 1 )
		    end if
		    
		  #else
		    
		    #pragma unused f
		    #pragma unused args
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LaunchWFS(extends f as FolderItem, ParamArray args as String)
		  #if TargetWin32
		    
		    Soft Declare Sub ShellExecuteA Lib "Shell32" ( hwnd as Integer, operation as CString, _
		    file as CString, params as CString, directory as CString, show as Integer )
		    Soft Declare Sub ShellExecuteW Lib "Shell32" ( hwnd as Integer, operation as WString, _
		    file as WString, params as WString, directory as WString, show as Integer )
		    
		    dim params as String
		    params = Join( args, " " )
		    
		    if System.IsFunctionAvailable( "ShellExecuteW", "Shell32" ) then
		      ShellExecuteW( 0, "open", f.AbsolutePath, params, "", 1 )
		    else
		      ShellExecuteA( 0, "open", f.AbsolutePath, params, "", 1 )
		    end if
		    
		  #else
		    
		    #pragma unused f
		    #pragma unused args
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub MoveFileToWithProgressWFS(extends fromFile as FolderItem, toFile as FolderItem)
		  // Delegate to the heavy-lifting function
		  CopyMoveOp( fromFile, toFile, true )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function PadHexToFourDigits(i as Integer) As String
		  // Get our hex string
		  dim hexStr as String = Hex( i )
		  
		  // Now pad whatever we need on the left to get
		  // to four bytes of string
		  return Left("0000", 4 - Len( hexStr ) ) + hexStr
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub RevealWFS(extends f as FolderItem)
		  #if TargetWin32
		    
		    dim param as String = "/select, """ + f.AbsolutePath + """"
		    
		    Soft Declare Sub ShellExecuteA Lib "Shell32" ( hwnd as Integer, op as CString, file as CString, _
		    params as CString, directory as Integer, cmd as Integer )
		    Soft Declare Sub ShellExecuteW Lib "Shell32" ( hwnd as Integer, op as WString, file as WString, _
		    params as WString, directory as Integer, cmd as Integer )
		    
		    Const SW_SHOW = 5
		    
		    if System.IsFunctionAvailable( "ShellExecuteW", "Shell32" ) then
		      ShellExecuteW( 0, "open", "explorer", param, 0, SW_SHOW )
		    else
		      ShellExecuteA( 0, "open", "explorer", param, 0, SW_SHOW )
		    end if
		    
		  #else
		    
		    #pragma unused f
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub StartupItemWFS(extends f as FolderItem, machineWide as Boolean, assigns set as Boolean)
		  #if TargetWin32
		    
		    dim runItem as RegistryItem
		    
		    if machineWide then
		      runItem = new RegistryItem( "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\" )
		    else
		      runItem = new RegistryItem( "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\" )
		    end
		    
		    
		    dim theAppName, theAppPath as String
		    theAppName =  nthField( f.Name, ".", 1 )
		    theAppPath = """" + f.AbsolutePath + """"
		    if set then
		      runItem.Value( theAppName ) = theAppPath
		    else
		      runItem.Delete( theAppName )
		    end
		    
		  #else
		    
		    #pragma unused f
		    #pragma unused machineWide
		    #pragma unused set
		    
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub StartWatchingForChangesWFS(extends f as FolderItem, watchSubDirs as Boolean)
		  #if TargetWin32
		    
		    Soft Declare Function FindFirstChangeNotificationA Lib "Kernel32" ( path as CString, watchSubDirs as Boolean, flags as Integer ) as Integer
		    Soft Declare Function FindFirstChangeNotificationW Lib "Kernel32" ( path as WString, watchSubDirs as Boolean, flags as Integer ) as Integer
		    
		    Const FILE_NOTIFY_CHANGE_FILE_NAME = &h00000001
		    Const FILE_NOTIFY_CHANGE_DIR_NAME = &h00000002
		    
		    // Try to start watching for changes
		    dim handle as Integer
		    if System.IsFunctionAvailable( "FindFirstChangeNotificationW", "Kernel32" ) then
		      handle = FindFirstChangeNotificationA( f.AbsolutePath, watchSubDirs, FILE_NOTIFY_CHANGE_FILE_NAME + FILE_NOTIFY_CHANGE_DIR_NAME )
		    else
		      handle = FindFirstChangeNotificationW( f.AbsolutePath, watchSubDirs, FILE_NOTIFY_CHANGE_FILE_NAME + FILE_NOTIFY_CHANGE_DIR_NAME )
		    end if
		    
		    if handle <> -1 then
		      // If we don't have a handle map, then make one
		      if mChangeHandles = nil then mChangeHandles = new Dictionary
		      
		      // And store this handle since we'll need it
		      mChangeHandles.Value( f.AbsolutePath ) = handle
		    end
		    
		  #else
		    
		    #pragma unused f
		    #pragma unused watchSubDirs
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub StopWatchingForChangesWFS(extends f as FolderItem)
		  #if TargetWin32
		    
		    // Check to see if this folder item has a handle in the map
		    dim handle as Integer
		    try
		      handle = mCHangeHandles.Value( f.AbsolutePath )
		    catch
		      // If this didn't work, then something's wrong!  Either there's
		      // no map, or the folder item isn't in it.  Either way, bail out
		      return
		    end
		    
		    // Now that we have a handle, close it
		    Declare Sub FindCloseChangeNotification Lib "Kernel32" ( handle as Integer )
		    FindCloseChangeNotification( handle )
		    
		    // And remove the item from the map
		    mChangeHandles.Remove( f.AbsolutePath )
		    
		  #else
		    
		    #pragma unused f
		    
		  #endif
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mChangeHandles As Dictionary
	#tag EndProperty


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
