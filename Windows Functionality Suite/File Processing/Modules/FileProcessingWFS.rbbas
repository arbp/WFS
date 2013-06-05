#tag Module
Protected Module FileProcessingWFS
	#tag Method, Flags = &h1
		Protected Sub EmptyTrashes()
		  '// This method will empty all RecycleBins in the system
		  '// By Anthony G. Cyphers
		  '// 05/17/2007
		  
		  #if TargetWin32 then
		    Soft Declare Function SHEmptyRecycleBinA Lib "shell32" ( hwnd As Integer, pszRootPath As Integer, dwFlags As Integer) As Integer
		    Soft Declare Function SHUpdateRecycleBinIcon Lib "shell32" () As Integer
		    
		    if System.IsFunctionAvailable( "SHEmptyRecycleBinA", "shell32" ) then
		      if SHEmptyRecycleBinA( 0, 0, 0 ) = 0 then
		        if System.IsFunctionAvailable( "SHUpdateRecycleBinIcon", "shell32" ) then
		          Call SHUpdateRecycleBinIcon
		        end if
		      end if
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetDriveStrings() As String()
		  #if TargetWin32
		    Soft Declare Function GetLogicalDriveStringsA Lib "Kernel32" ( size as Integer, buffer as Ptr ) as Integer
		    Soft Declare Function GetLogicalDriveStringsW Lib "Kernel32" ( size as Integer, buffer as Ptr ) as Integer
		    
		    dim numChars as Integer
		    dim mb as new MemoryBlock( 1024 )
		    dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "GetLogicalDriveStringsW", "Kernel32" )
		    
		    if unicodeSavvy then
		      numChars = GetLogicalDriveStringsW( mb.Size, mb )
		    else
		      numChars = GetLogicalDriveStringsA( mb.Size, mb )
		    end if
		    
		    dim ret(), theStr as String
		    dim i as Integer
		    while i < numChars
		      if unicodeSavvy then
		        // We multiply by two because there are two bytes
		        // per character.  i counts characters, but the MemoryBlock
		        // position is in bytes.
		        theStr = mb.WString( i * 2 )
		      else
		        theStr = mb.CString( i )
		      end if
		      
		      ret.Append( theStr )
		      i = i + Len( theStr ) + 1
		    wend
		    
		    return ret
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetDriveType(root as FolderItem) As String
		  #if TargetWin32
		    
		    Soft Declare Function GetDriveTypeA Lib "Kernel32" ( drive as CString ) as Integer
		    Soft Declare Function GetDriveTypeW Lib "Kernel32" ( drive as WString ) as Integer
		    
		    dim rootStr as String
		    if root <> nil then
		      rootStr = root.AbsolutePath
		    end
		    
		    dim type as Integer
		    if System.IsFunctionAvailable( "GetDriveTypeW", "Kernel32" ) then
		      type = GetDriveTypeW( rootStr )
		    else
		      type = GetDriveTypeA( rootStr )
		    end if
		    
		    select case type
		    case 0
		      return "Unknown"
		    case 1
		      return "No root directory"
		    case 2
		      return "Removable drive"
		    case 3
		      return "Fixed  drive"
		    case 4
		      return "Remote drive"
		    case 5
		      return "CD-ROM drive"
		    case 6
		      return "RAM disk"
		    end
		    
		    return "Unknown"
		    
		  #else
		    
		    #pragma unused root
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetFreeDiskSpaceForCaller(root as FolderItem) As Double
		  #if TargetWin32
		    
		    Soft Declare Sub GetDiskFreeSpaceExA Lib "Kernel32" ( directory as CString, freeBytesForCaller as Ptr, _
		    totalBytes as Ptr, totalFreeBytes as Ptr )
		    Soft Declare Sub GetDiskFreeSpaceExW Lib "Kernel32" ( directory as WString, freeBytesForCaller as Ptr, _
		    totalBytes as Ptr, totalFreeBytes as Ptr )
		    
		    if root = nil then return 0.0
		    
		    Dim free, total, totalFree as MemoryBlock
		    free = new MemoryBlock( 8 )
		    total = new MemoryBlock( 8 )
		    totalFree = new MemoryBlock( 8 )
		    
		    if System.IsFunctionAvailable( "GetFreeDiskSpaceExW", "Kernel32" ) then
		      GetDiskFreeSpaceExW( root.AbsolutePath, free, total, totalFree )
		    else
		      GetDiskFreeSpaceExA( root.AbsolutePath, free, total, totalFree )
		    end if
		    
		    dim ret as Double
		    dim high as double = free.Long( 4 )
		    dim low as double = free.Long( 0 )
		    
		    if high < 0 then high = high + pow( 2, 32 )
		    if low < 0 then low= low + pow( 2, 32 )
		    
		    ret = high * Pow( 2, 32 )
		    ret = ret + low
		    
		    return ret
		    
		  #else
		    
		    #pragma unused root
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetSpecialFolder(CSIDL as integer) As folderItem
		  Dim f as FolderItem
		  
		  #if targetWin32
		    
		    Dim myPidl as integer
		    Dim myErr as integer
		    Dim sPath as string
		    Dim mb as new MemoryBlock(256)
		    
		    declare Function SHGetSpecialFolderLocation Lib "shell32"(hwnd as integer, nFolder as integer, byref pidl as integer) as integer
		    
		    myErr = SHGetSpecialFolderLocation(0, CSIDL,myPidl)
		    
		    
		    Soft Declare Function SHGetPathFromIDListA Lib "shell32" (pidl as integer, path as ptr) as integer
		    Soft Declare Function SHGetPathFromIDListW Lib "shell32" (pidl as integer, path as ptr) as integer
		    
		    if System.IsFunctionAvailable( "SHGetPathFromIDListW", "Shell32" ) then
		      myErr = SHGetPathFromIDListW(myPidl,mb)
		      f = GetFolderItem(mb.WString(0))
		    else
		      myErr = SHGetPathFromIDListA(myPidl,mb)
		      f = GetFolderItem(mb.CString(0))
		    end if
		    
		  #else
		    
		    #pragma unused CSIDL
		    
		  #endif
		  
		  return f
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetTotalBytes(root as FolderItem) As double
		  #if TargetWin32
		    
		    Soft Declare Sub GetDiskFreeSpaceExA Lib "Kernel32" ( directory as CString, freeBytesForCaller as Ptr, _
		    totalBytes as Ptr, totalFreeBytes as Ptr )
		    Soft Declare Sub GetDiskFreeSpaceExW Lib "Kernel32" ( directory as WString, freeBytesForCaller as Ptr, _
		    totalBytes as Ptr, totalFreeBytes as Ptr )
		    
		    if root = nil then return 0.0
		    
		    Dim free, total, totalFree as MemoryBlock
		    free = new MemoryBlock( 8 )
		    total = new MemoryBlock( 8 )
		    totalFree = new MemoryBlock( 8 )
		    
		    if System.IsFunctionAvailable( "GetDiskFreeSpaceExW", "Kernel32" ) then
		      GetDiskFreeSpaceExW( root.AbsolutePath, free, total, totalFree )
		    else
		      GetDiskFreeSpaceExA( root.AbsolutePath, free, total, totalFree )
		    end if
		    
		    dim ret as Double
		    dim high as double = total.Long( 4 )
		    dim low as double = total.Long( 0 )
		    
		    if high < 0 then high = high + pow( 2, 32 )
		    if low < 0 then low= low + pow( 2, 32 )
		    
		    ret = high * Pow( 2, 32 )
		    ret = ret + low
		    
		    return ret
		    
		  #else
		    
		    #pragma unused root
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetTotalFreeSpace(root as FolderItem) As double
		  #if TargetWin32
		    
		    Soft Declare Sub GetDiskFreeSpaceExA Lib "Kernel32" ( directory as CString, freeBytesForCaller as Ptr, _
		    totalBytes as Ptr, totalFreeBytes as Ptr )
		    Soft Declare Sub GetDiskFreeSpaceExW Lib "Kernel32" ( directory as WString, freeBytesForCaller as Ptr, _
		    totalBytes as Ptr, totalFreeBytes as Ptr )
		    
		    if root = nil then return 0.0
		    
		    Dim free, total, totalFree as MemoryBlock
		    free = new MemoryBlock( 8 )
		    total = new MemoryBlock( 8 )
		    totalFree = new MemoryBlock( 8 )
		    
		    if System.IsFunctionAvailable( "GetDiskFreeSpaceExW", "Kernel32" ) then
		      GetDiskFreeSpaceExW( root.AbsolutePath, free, total, totalFree )
		    else
		      GetDiskFreeSpaceExA( root.AbsolutePath, free, total, totalFree )
		    end if
		    
		    dim ret as Double
		    dim high as double = totalFree.Long( 4 )
		    dim low as double = totalFree.Long( 0 )
		    
		    if high < 0 then high = high + pow( 2, 32 )
		    if low < 0 then low= low + pow( 2, 32 )
		    
		    ret = high * Pow( 2, 32 )
		    ret = ret + low
		    
		    return ret
		    
		  #else
		    
		    #pragma unused root
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetTrashesCount() As Int64
		  '// This function return the number of items for all RecycleBins on the system
		  '// By Anthony G. Cyphers
		  '// 05/17/2007
		  
		  #if TargetWin32 then
		    Soft Declare Function SHQueryRecycleBinA Lib "shell32" ( pszRootPath As Integer, pSHQueryRBInfo As Ptr) As Integer
		    
		    dim newInfo as new MemoryBlock( 20 )
		    newInfo.Long( 0 ) = newInfo.Size
		    dim x as Integer
		    if System.IsFunctionAvailable( "SHQueryRecycleBinA", "shell32" ) then
		      x = shQueryrecyclebinA( 0, newInfo )
		    end if
		    
		    if x = 0 then
		      return newInfo.Int64Value( 12 )
		    end if
		    
		    return -1
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetTrashesSize() As Int64
		  '// This function return the size in bytes for all RecycleBins on the system
		  '// By Anthony G. Cyphers
		  '// 05/17/2007
		  
		  #if TargetWin32 then
		    Soft Declare Function SHQueryRecycleBinA Lib "shell32" ( pszRootPath As Integer, pSHQueryRBInfo As Ptr) As Integer
		    
		    dim newInfo as new MemoryBlock( 20 )
		    newInfo.Long( 0 ) = newInfo.Size
		    dim x as Integer
		    if System.IsFunctionAvailable( "SHQueryRecycleBinA", "shell32" ) then
		      x = shQueryrecyclebinA( 0, newInfo )
		    end if
		    
		    if x = 0 then
		      return newInfo.Int64Value( 4 )
		    end if
		    
		    return -1
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetVolumeName(root as FolderItem) As String
		  #if TargetWin32
		    
		    Soft Declare Function GetVolumeInformationA Lib "Kernel32" ( root as CString, _
		    volName as Ptr, volNameSize as Integer, ByRef volSer as Integer, ByRef _
		    maxCompLength as Integer, ByRef sysFlags as Integer, sysName as Ptr, _
		    sysNameSize as Integer ) as Boolean
		    Soft Declare Function GetVolumeInformationW Lib "Kernel32" ( root as WString, _
		    volName as Ptr, volNameSize as Integer, ByRef volSer as Integer, ByRef _
		    maxCompLength as Integer, ByRef sysFlags as Integer, sysName as Ptr, _
		    sysNameSize as Integer ) as Boolean
		    
		    
		    dim volName as new MemoryBlock( 256 )
		    dim sysName as new MemoryBlock( 256 )
		    dim volSerial, maxCompLength, sysFlags as Integer
		    
		    if System.IsFunctionAvailable( "GetVolumeInformationW", "Kernel32" ) then
		      Call GetVolumeInformationW( left( root.AbsolutePath, 3 ), volName, 256, volSerial, maxCompLength, _
		      sysFlags, sysName, 256 )
		      
		      return volName.WString( 0 )
		    else
		      Call GetVolumeInformationA( left( root.AbsolutePath, 3 ), volName, 256, volSerial, maxCompLength, _
		      sysFlags, sysName, 256 )
		      
		      return volName.CString( 0 )
		    end if
		    
		  #else
		    
		    #pragma unused root
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetVolumeSerial(root as FolderItem) As String
		  #if TargetWin32
		    
		    Soft Declare Function GetVolumeInformationA Lib "Kernel32" ( root as CString, _
		    volName as Ptr, volNameSize as Integer, ByRef volSer as Integer, ByRef _
		    maxCompLength as Integer, ByRef sysFlags as Integer, sysName as Ptr, _
		    sysNameSize as Integer ) as Boolean
		    Soft Declare Function GetVolumeInformationW Lib "Kernel32" ( root as WString, _
		    volName as Ptr, volNameSize as Integer, ByRef volSer as Integer, ByRef _
		    maxCompLength as Integer, ByRef sysFlags as Integer, sysName as Ptr, _
		    sysNameSize as Integer ) as Boolean
		    
		    
		    dim volName as new MemoryBlock( 256 )
		    dim sysName as new MemoryBlock( 256 )
		    dim volSerial, maxCompLength, sysFlags as Integer
		    
		    if System.IsFunctionAvailable( "GetVolumeInformationW", "Kernel32" ) then
		      Call GetVolumeInformationW( left( root.AbsolutePath, 3 ), volName, 256, volSerial, maxCompLength, _
		      sysFlags, sysName, 256 )
		    else
		      Call GetVolumeInformationA( left( root.AbsolutePath, 3 ), volName, 256, volSerial, maxCompLength, _
		      sysFlags, sysName, 256 )
		    end if
		    
		    dim hexStr as String = Hex( volSerial )
		    return Left( hexStr, 4 ) + "-" + Right( hexStr, 4 )
		    
		  #else
		    
		    #pragma unused root
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetVolumeSerialNumber(root as FolderItem) As Integer
		  #if TargetWin32
		    
		    Soft Declare Function GetVolumeInformationA Lib "Kernel32" ( root as CString, _
		    volName as Ptr, volNameSize as Integer, ByRef volSer as Integer, ByRef _
		    maxCompLength as Integer, ByRef sysFlags as Integer, sysName as Ptr, _
		    sysNameSize as Integer ) as Boolean
		    Soft Declare Function GetVolumeInformationW Lib "Kernel32" ( root as CString, _
		    volName as Ptr, volNameSize as Integer, ByRef volSer as Integer, ByRef _
		    maxCompLength as Integer, ByRef sysFlags as Integer, sysName as Ptr, _
		    sysNameSize as Integer ) as Boolean
		    
		    
		    dim volName as new MemoryBlock( 256 )
		    dim sysName as new MemoryBlock( 256 )
		    dim volSerial, maxCompLength, sysFlags as Integer
		    
		    if System.IsFunctionAvailable( "GetVolumeInformationW", "Kernel32" ) then
		      Call GetVolumeInformationW( left( root.AbsolutePath, 3 ), volName, 256, volSerial, maxCompLength, _
		      sysFlags, sysName, 256 )
		    else
		      Call GetVolumeInformationA( left( root.AbsolutePath, 3 ), volName, 256, volSerial, maxCompLength, _
		      sysFlags, sysName, 256 )
		    end if
		    return volSerial
		    
		  #else
		    
		    #pragma unused root
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MapNetworkDrive(remotePath as String, localPath as String, userName as String = "", password as String = "", interactive as Boolean = true) As FolderItem
		  // We want to map the network drive the user gave us, which is in UNC format (like //10.10.10.116/foobar)
		  // and map it to the local drive they gave us (like f:).
		  
		  #if TargetWin32
		    
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
		    
		  #else
		    
		    #pragma unused remotePath
		    #pragma unused localPath
		    #pragma unused userName
		    #pragma unused password
		    #pragma unused interactive
		    
		  #endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MapNetworkDriveDialog(owner as Window) As FolderItem
		  #if TargetWin32
		    
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
		    
		  #else
		    
		    #pragma unused owner
		    
		  #endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SelectMultipleFiles(parentWindow as Window, filterTypes() as String) As FolderItem()
		  #if TargetWin32
		    dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "GetOpenFileNameW", "CommDlg32" )
		    
		    dim ofn as new MemoryBlock( 76 )
		    
		    ofn.Long( 0 ) = ofn.Size
		    if parentWindow <> nil then
		      ofn.Long( 4 ) = parentWindow.WinHWND
		    end
		    
		    dim filter as String
		    dim catFilters as String
		    dim filterPtr as MemoryBlock
		    for each filter in filterTypes
		      if unicodeSavvy then
		        catFilters = catFilters + ConvertEncoding( filter + Chr( 0 ), Encodings.UTF16 )
		      else
		        catFilters = catFilters + filter + Chr( 0 )
		      end if
		    next
		    if unicodeSavvy then
		      catFilters = catFilters + ConvertEncoding( Chr( 0 ), Encodings.UTF16 )
		    else
		      catFilters = catFilters + Chr( 0 )
		    end if
		    
		    filterPtr = catFilters
		    ofn.Ptr( 12 ) = filterPtr
		    ofn.Long( 24 ) = 1
		    
		    dim filePaths as new MemoryBlock( 4098 )
		    ofn.Ptr( 28 ) = filePaths
		    ofn.Long( 32 ) = filePaths.Size
		    
		    'dim titlePtr as new MemoryBlock( LenB( title ) + 1 )
		    'titlePtr.CString( 0 ) = title
		    'ofn.Ptr( 48 ) = titlePtr
		    
		    Const OFN_EXPLORER = &H80000
		    Const OFN_LONGNAMES = &H200000
		    Const OFN_ALLOWMULTISELECT = &H200
		    ofn.Long( 52 ) = BitWiseOr(OFN_ALLOWMULTISELECT,OFN_EXPLORER)
		    
		    Soft Declare Function GetOpenFileNameA Lib "comdlg32" (pOpenfilename As ptr) As Boolean
		    Soft Declare Function GetOpenFileNameW Lib "comdlg32" (pOpenfilename As ptr) As Boolean
		    
		    dim res as Boolean
		    dim ret( -1 ) as FolderItem
		    dim strs( -1 ) as String
		    dim i as Integer
		    dim s, filePath as String
		    dim sPtr as MemoryBlock
		    
		    if unicodeSavvy then
		      res = GetOpenFileNameW( ofn )
		    else
		      res = GetOpenFileNameA( ofn )
		    end if
		    
		    if res then
		      sPtr = ofn.Ptr( 28 )
		      s = sPtr.StringValue( 0, filePaths.Size )
		      if unicodeSavvy then
		        s = DefineEncoding( s, Encodings.UTF16 )
		      end if
		      
		      strs = Split( s, Chr( 0 ) )
		      
		      filePath = strs( 0 )
		      
		      for i = 0 to UBound( strs )
		        if strs( i ) <> "" then
		          try
		            ret.Append( new FolderItem( filePath + "\" + strs( i ) ) )
		          catch err as UnsupportedFormatException
		            // If we couldn't make a file from it, then chances are we only
		            // got one file and not multiple ones.  Try something else
		            
		            try
		              dim test as FolderItem = new FolderItem( filePath )
		              
		              // We have to make sure this isn't a directory because that's
		              // the first entry in the list if the list contains multiple items.  Goofy
		              // enough, if the list only contains one item, then the first entry is
		              // that item.
		              if not test.Directory then
		                ret.Append( test )
		              end if
		            catch err2 as UnsupportedFormatException
		              // We're really hosed now
		            end try
		          end try
		        end
		      next
		    end
		    
		    return ret
		    
		  #else
		    
		    #pragma unused parentWindow
		    #pragma unused filterTypes
		    
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function UnmapNetworkDrive(drive as String, force as Boolean = false) As Boolean
		  #if TargetWin32
		    
		    Soft Declare Function WNetCancelConnection2W Lib "Mpr" ( name as WString, flags as Integer, force as Boolean ) as Integer
		    Soft Declare Function WNetCancelConnection2A Lib "Mpr" ( name as CString, flags as Integer, force as Boolean ) as Integer
		    
		    if System.IsFunctionAvailable( "WNetCancelConnection2W", "Mpr" ) then
		      return WNetCancelConnection2W( drive, 0, force ) = 0
		    else
		      return WNetCancelConnection2A( drive, 0, force ) = 0
		    end if
		    
		  #else
		    
		    #pragma unused drive
		    #pragma unused force
		    
		  #endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function UnmapNetworkDriveDialog(owner as Window, localPath as String) As Boolean
		  #if TargetWin32
		    
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
		    
		  #else
		    
		    #pragma unused owner
		    #pragma unused localPath
		    
		  #endif
		  
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
