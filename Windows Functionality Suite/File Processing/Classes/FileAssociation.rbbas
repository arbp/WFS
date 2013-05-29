#tag Class
Protected Class FileAssociation
	#tag Method, Flags = &h0
		Sub AddExtension(extension as String, longExtension as String, perceivedType as String, mimeType as String)
		  Dim d as new Dictionary
		  d.Value( "Name" ) = extension
		  d.Value( "Long Name" ) = longExtension
		  d.Value( "Perceived Type" ) = perceivedType
		  d.Value( "MIME Type" ) = mimeType
		  
		  mExt.Append( d )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddVerb(verb as String, command as String)
		  // We want to add a new verb, such as open or play
		  mVerbs.Append( verb )
		  mVerbCommands.Append( command )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CompanyName() As String
		  return mCompanyName
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CompanyName(assigns s as String)
		  mCompanyName = s
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FriendlyName() As String
		  return mFriendlyName
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub FriendlyName(assigns s as String)
		  mFriendlyName = s
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ProductName() As String
		  return mProductName
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ProductName(assigns s as String)
		  mProductName = s
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Register() As Boolean
		  // If we don't have enough information, then we
		  // can't complete the regitration.
		  
		  if CompanyName = "" or ProductName = "" or Version = 0 or _
		    UBound( mExt ) = -1 or UBound( mVerbs ) = -1 or _
		    mFriendlyName = "" then
		    
		    return false
		  end if
		  
		  // We want to register the prog ID first
		  if not RegisterProgID then return false
		  
		  // Then we want to register each file extension
		  for i as Integer = 0 to UBound( mExt )
		    dim d as Dictionary = mExt( i )
		    
		    if not RegisterExtension( d.Value( "Name" ), d.Value( "Long Name" ), _
		      d.Value( "Perceived Type" ), d.Value( "MIME Type" ) ) then
		      return false
		    end if
		  next i
		  
		  #if TargetWin32
		    // Now we want to let the shell know that a change has
		    // happened.
		    Declare Sub SHChangeNotify Lib "Shell32" ( eventId as Integer, flags as Integer, _
		    item1 as Integer, item2 as Integer )
		    
		    Const SHCNE_ASSOCCHANGED = &h08000000
		    Const SHCNF_FLUSH = &h1000
		    
		    SHChangeNotify( SHCNE_ASSOCCHANGED, SHCNF_FLUSH, 0, 0 )
		    
		    return true
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RegisterExtension(ext as String, longExt as String = "", perceivedType as String = "", mimeType as String = "") As Boolean
		  // We want to make a key under HKEY_CLASSES_ROOT with
		  // the following format:
		  //
		  // .ext
		  //      (Default) = ProgID
		  //      PerceivedType = (Optional) perceivedType
		  //      Content Type = (Optional) mimeType
		  
		  // First, get a handle to the HKEY_CLASSES_ROOT key
		  Dim hkcr as RegistryItem
		  
		  try
		    hkcr = new RegistryItem( "HKEY_CLASSES_ROOT" )
		  catch err as RegistryAccessErrorException
		    return false
		  end try
		  
		  // Now get a handle to the extension we care about
		  dim extKey as RegistryItem
		  try
		    // We need to make sure that the extension
		    // is prefixed with a dot.
		    if Left( ext, 1 ) <> "." then ext = "." + ext
		    
		    // Make the extension
		    extKey = hkcr.AddFolder( ext )
		    
		    // Set it's default value
		    extKey.DefaultValue = mCompanyName + "." + mProductName + "." + Str( mVersion )
		  catch err as RegistryAccessErrorException
		    return false
		  end try
		  
		  // If the user has a perceivedType or a mimeType,
		  // then we want to set those now as well
		  if perceivedType <> "" then
		    try
		      extKey.Value( "PerceivedType" ) = perceivedType
		    catch err as RegistryAccessErrorException
		      return false
		    end try
		  end if
		  
		  if mimeType <> "" then
		    try
		      extKey.Value( "ContentType" ) = mimeType
		    catch err as RegistryAccessErrorException
		      return false
		    end try
		  end if
		  
		  // Now we want to do the same dance for the
		  // long extension.
		  dim longExtKey as RegistryItem
		  try
		    // We need to make sure that the extension
		    // is prefixed with a dot.
		    if Left( longExt, 1 ) <> "." then longExt = "." + longExt
		    
		    // Make the extension
		    longExtKey = hkcr.AddFolder( longExt )
		    
		    // Set it's default value
		    longExtKey.DefaultValue = mCompanyName + "." + mProductName + "." + Str( mVersion )
		  catch err as RegistryAccessErrorException
		    return false
		  end try
		  
		  // If the user has a perceivedType or a mimeType,
		  // then we want to set those now as well
		  if perceivedType <> "" then
		    try
		      longExtKey.Value( "PerceivedType" ) = perceivedType
		    catch err as RegistryAccessErrorException
		      return false
		    end try
		  end if
		  
		  if mimeType <> "" then
		    try
		      longExtKey.Value( "ContentType" ) = mimeType
		    catch err as RegistryAccessErrorException
		      return false
		    end try
		  end if
		  
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RegisterProgID() As Boolean
		  // We want to make a key under HKEY_CLASSES_ROOT with
		  // the following format:
		  //
		  // CompanyName.ProductName.Version
		  //      (Default) = Friendly Product Name
		  //      FriendlyTypeName = Same as (Default), but it must be an indirect string (so not available to RB applications)
		  //      InfoTip = help tag for this prog id, however, must be an indirect string again
		  //      EditFlags = bitmask for FILETYPEATTRIBUTES
		  //      CurVer
		  //          (Default) = CompanyName.ProductName.Version
		  //      DefaultIcon
		  //          (Default) = Path to icon
		  //      shell
		  //          verb (play, open, edit, etc)
		  //              command
		  //                  (Default) = Path to executable %1
		  //
		  
		  // First, get a handle to the HKEY_CLASSES_ROOT key
		  Dim hkcr as RegistryItem
		  
		  try
		    hkcr = new RegistryItem( "HKEY_CLASSES_ROOT" )
		  catch err as RegistryAccessErrorException
		    return false
		  end try
		  
		  // Now we want to get a key that points to our
		  // ProgID
		  dim progIDStr as String = mCompanyName + "." + mProductName + "." + Str( mVersion )
		  dim progID as RegistryItem
		  
		  try
		    progID = hkcr.AddFolder( progIDStr )
		    
		    // Now that we've got the main key we're
		    // interested in, let's set up the values for
		    // the key.
		    progID.DefaultValue = mFriendlyName
		  catch err as RegistryAccessErrorException
		    return false
		  end try
		  
		  // If we wanted to set things like InfoTip, FriendlyTypeName or
		  // EditFlags, this would be the place to do it.  But we'll save
		  // that for some other time.
		  
		  // Add the CurVer subkey
		  dim curVer as RegistryItem
		  try
		    curVer = progID.AddFolder( "CurVer" )
		    
		    // Now set the current version default
		    // value.  We want to set this to the progID
		    // string
		    curVer.DefaultValue = progIDStr
		  catch err as RegistryAccessErrorException
		    return false
		  end try
		  
		  // Now we want to add the default icon
		  // subkey
		  if mDefaultIcon <> "" then
		    dim defIcon as RegistryItem
		    try
		      defIcon = progID.AddFolder( "DefaultIcon" )
		      
		      // Now we can set the default value for this
		      // key to the location of the icon.  This is
		      // usually an indirect string.
		      defIcon.DefaultValue = mDefaultIcon
		    catch err as RegistryAccessErrorException
		      return false
		    end try
		  end if
		  
		  // Now we need to add the shell subkey so that
		  // we can add the verbs.
		  dim shellKey as RegistryItem
		  try
		    shellKey = progID.AddFolder( "shell" )
		  catch err as RegistryAccessErrorException
		    return false
		  end try
		  
		  // Now we can add our verbs and their commands
		  for i as Integer = 0 to UBound( mVerbs )
		    dim verb as String = mVerbs( i )
		    dim command as String = mVerbCommands( i )
		    
		    dim verbKey, commandKey as RegistryItem
		    try
		      verbKey = shellKey.AddFolder( verb )
		      commandKey = verbKey.AddFolder( "command" )
		      
		      // Once we've got the command subkey created,
		      // set its default value to be the actual command
		      // we want to fire for the given verb
		      commandKey.DefaultValue = command
		    catch err as RegistryAccessErrorException
		      return false
		    end try
		  next i
		  
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Version() As Integer
		  return mVersion
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Version(assigns i as Integer)
		  mVersion = i
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mCompanyName As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDefaultIcon As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mExt(-1) As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mFriendlyName As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mProductName As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mVerbCommands(-1) As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mVerbs(-1) As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mVersion As Integer
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
End Class
#tag EndClass
