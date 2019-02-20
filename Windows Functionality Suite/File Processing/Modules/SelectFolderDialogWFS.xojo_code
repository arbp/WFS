#tag Module
Protected Module SelectFolderDialogWFS
	#tag Method, Flags = &h0
		Function SelectFolderWFS(title as string="", root as SelectFolderDialogWFS.RootFolders=SelectFolderDialogWFS.RootFolders.CSIDL_DESKTOP, flags as Integer=0) As FolderItem
		  #if TargetWin32
		    declare Function SHBrowseForFolderA lib "Shell32"(byref info as BROWSEINFOA) as ptr
		    declare Function SHGetPathFromIDListA lib "Shell32"(list as ptr,path as ptr) as Boolean
		    
		    dim info as BROWSEINFOA
		    info.title=title
		    info.root=uint32(root)
		    info.Flags=Flags
		    dim displayNameMB as new MemoryBlock(260)
		    dim pathMB as new MemoryBlock(260)
		    info.displayName=displayNameMB
		    dim idList as ptr=SHBrowseForFolderA(info)
		    if idList<>nil then 
		      if SHGetPathFromIDListA(idList,pathMB) then
		        dim path as CString=pathMB.CString(0)
		        dim f as new FolderItem(path,FolderItem.PathTypeNative)
		        Return f
		      end if
		    end if
		  #endif
		End Function
	#tag EndMethod


	#tag Note, Name = About
		
		SelectFolderWFS(rootFolder,flags) Uses the minimalistic Folder Seletion Dialog
		RootFolders Enumeration values allow setting a staring point for navigation
		Flags are a combination of constants listed in module OR'd together
	#tag EndNote


	#tag Constant, Name = kBIF_BROWSEFILEJUNCTIONS, Type = Double, Dynamic = False, Default = \"&h00010000", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kBIF_BROWSEFORCOMPUTER, Type = Double, Dynamic = False, Default = \"&h00001000", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kBIF_BROWSEFORPRINTER, Type = Double, Dynamic = False, Default = \"&h00002000", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kBIF_BROWSEINCLUDEFILES, Type = Double, Dynamic = False, Default = \"&h00004000", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kBIF_BROWSEINCLUDEURLS, Type = Double, Dynamic = False, Default = \"&h00000080", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kBIF_DONTGOBELOWDOMAIN, Type = Double, Dynamic = False, Default = \"&h00000002", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kBIF_EDITBOX, Type = Double, Dynamic = False, Default = \"&h00000010", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kBIF_NEWDIALOGSTYLE, Type = Double, Dynamic = False, Default = \"&h00000040", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kBIF_NONEWFOLDERBUTTON, Type = Double, Dynamic = False, Default = \"&h00000200", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kBIF_NOTRANSLATETARGETS, Type = Double, Dynamic = False, Default = \"&h00000400", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kBIF_RETURNFSANCESTORS, Type = Double, Dynamic = False, Default = \"&h00000008", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kBIF_RETURNONLYFSDIRS, Type = Double, Dynamic = False, Default = \"&h00000001", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kBIF_SHAREABLE, Type = Double, Dynamic = False, Default = \"&h00008000", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kBIF_STATUSTEXT, Type = Double, Dynamic = False, Default = \"&h00000004", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kBIF_UAHINT, Type = Double, Dynamic = False, Default = \"&h00000100", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kBIF_USENEWUI, Type = Double, Dynamic = False, Default = \"&h00000050", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kBIF_VALIDATE, Type = Double, Dynamic = False, Default = \"&h00000020", Scope = Public
	#tag EndConstant


	#tag Structure, Name = BROWSEINFOA, Flags = &h21, Attributes = \"StructureAlignment \x3D 1"
		Owner as ptr
		  root as int32
		  displayName as ptr
		  Title as cstring
		  Flags as uint32
		  fn as ptr
		image as int32
	#tag EndStructure


	#tag Enum, Name = RootFolders, Type = Integer, Flags = &h1
		CSIDL_DESKTOP                   = &h00000        // <desktop>
		  CSIDL_INTERNET                  = &h00001        // Internet Explorer (icon on desktop)
		  CSIDL_PROGRAMS                  = &h00002        // Start Menu\Programs
		  CSIDL_CONTROLS                  = &h00003        // My Computer\Control Panel
		  CSIDL_PRINTERS                  = &h00004        // My Computer\Printers
		  CSIDL_PERSONAL                  = &h00005        // My Documents
		  CSIDL_FAVORITES                 = &h00006        // <user name>\Favorites
		  CSIDL_STARTUP                   = &h00007        // Start Menu\Programs\Startup
		  CSIDL_RECENT                    = &h00008        // <user name>\Recent
		  CSIDL_SENDTO                    = &h00009        // <user name>\SendTo
		  CSIDL_BITBUCKET                 = &h0000a        // <desktop>\Recycle Bin
		  CSIDL_STARTMENU                 = &h0000b        // <user name>\Start Menu
		  CSIDL_MYDOCUMENTS               = CSIDL_PERSONAL //  Personal was just a silly name for My Documents
		  CSIDL_MYMUSIC                   = &h0000d        // "My Music" folder
		  CSIDL_MYVIDEO                   = &h0000e        // "My Videos" folder
		  CSIDL_DESKTOPDIRECTORY          = &h00010        // <user name>\Desktop
		  CSIDL_DRIVES                    = &h00011        // My Computer
		  CSIDL_NETWORK                   = &h00012        // Network Neighborhood (My Network Places)
		  CSIDL_NETHOOD                   = &h00013        // <user name>\nethood
		  CSIDL_FONTS                     = &h00014        // windows\fonts
		  CSIDL_TEMPLATES                 = &h00015
		  CSIDL_COMMON_STARTMENU          = &h00016        // All Users\Start Menu
		  CSIDL_COMMON_PROGRAMS           = &h00017        // All Users\Start Menu\Programs
		  CSIDL_COMMON_STARTUP            = &h00018        // All Users\Startup
		  CSIDL_COMMON_DESKTOPDIRECTORY   = &h00019        // All Users\Desktop
		  CSIDL_APPDATA                   = &h0001a        // <user name>\Application Data
		  CSIDL_PRINTHOOD                 = &h0001b        // <user name>\PrintHood
		  CSIDL_LOCAL_APPDATA             = &h0001c        // <user name>\Local Settings\Applicaiton Data (non roaming)
		  CSIDL_ALTSTARTUP                = &h0001d        // non localized startup
		  CSIDL_COMMON_ALTSTARTUP         = &h0001e        // non localized common startup
		  CSIDL_COMMON_FAVORITES          = &h0001f
		  CSIDL_INTERNET_CACHE            = &h00020
		  CSIDL_COOKIES                   = &h00021
		  CSIDL_HISTORY                   = &h00022
		  CSIDL_COMMON_APPDATA            = &h00023        // All Users\Application Data
		  CSIDL_WINDOWS                   = &h00024        // GetWindowsDirectory()
		  CSIDL_SYSTEM                    = &h00025        // GetSystemDirectory()
		  CSIDL_PROGRAM_FILES             = &h00026        // C:\Program Files
		  CSIDL_MYPICTURES                = &h00027        // C:\Program Files\My Pictures
		  CSIDL_PROFILE                   = &h00028        // USERPROFILE
		  CSIDL_SYSTEMX86                 = &h00029        // x86 system directory on RISC
		  CSIDL_PROGRAM_FILESX86          = &h0002a        // x86 C:\Program Files on RISC
		  CSIDL_PROGRAM_FILES_COMMON      = &h0002b        // C:\Program Files\Common
		  CSIDL_PROGRAM_FILES_COMMONX86   = &h0002c        // x86 Program Files\Common on RISC
		  CSIDL_COMMON_TEMPLATES          = &h0002d        // All Users\Templates
		  CSIDL_COMMON_DOCUMENTS          = &h0002e        // All Users\Documents
		  CSIDL_COMMON_ADMINTOOLS         = &h0002f        // All Users\Start Menu\Programs\Administrative Tools
		  CSIDL_ADMINTOOLS                = &h00030        // <user name>\Start Menu\Programs\Administrative Tools
		  CSIDL_CONNECTIONS               = &h00031        // Network and Dial-up Connections
		  CSIDL_COMMON_MUSIC              = &h00035        // All Users\My Music
		  CSIDL_COMMON_PICTURES           = &h00036        // All Users\My Pictures
		  CSIDL_COMMON_VIDEO              = &h00037        // All Users\My Video
		  CSIDL_RESOURCES                 = &h00038        // Resource Direcotry
		  CSIDL_RESOURCES_LOCALIZED       = &h00039        // Localized Resource Direcotry
		  CSIDL_COMMON_OEM_LINKS          = &h0003a        // Links to All Users OEM specific apps
		  CSIDL_CDBURN_AREA               = &h0003b        // USERPROFILE\Local Settings\Application Data\Microsoft\CD Burning
		CSIDL_COMPUTERSNEARME           = &h0003d        // Computers Near Me (computered from Workgroup membership)
	#tag EndEnum


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
