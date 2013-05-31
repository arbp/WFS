#tag Module
Protected Module VB
	#tag Method, Flags = &h1
		Protected Sub AppActivate(processID as Integer, unused as Boolean = false)
		  // We have a process ID, so let's load all the processes
		  // that are running and see if any have a matching ID.
		  dim processes( -1 ) as ProcessInformation = ProcessManagement.GetActiveProcesses
		  
		  for each process as ProcessInformation in processes
		    // Check to see if its ID matches
		    if process.ProcessID = processID then
		      // Bring it to the front
		      ProcessManagement.BringToFront( process )
		      // And we're done
		      return
		    end if
		  next process
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AppActivate(title as String, unused as Boolean = false)
		  // We want to activate an "application" based on the
		  // window title, or partial window title.  From the VB6 docs:
		  //
		  // In determining which application to activate, title is compared to the title
		  // string of each running application. If there is no exact match, any application
		  // whose title string begins with title is activated. If there is more than one
		  // instance of the application named by title, one instance is arbitrarily activated.
		  //
		  // The way I interpret this is that we should loop over all the processes
		  // and look at the name.  If the name is an exact match, then we use
		  // that one.  If no matches are found, we go back and look at the
		  // beginning of the each process name to see if we can find a match.  If
		  // we still can't find one, then I think we're supposed search window titles
		  // in the same fashion.  It's ambiguous though.  The parameter specifier
		  // from the same docs says: "...the title in the title bar of the application window you
		  // want to activate..."  So which is it -- process name or window title?
		  //
		  // Turns out that the answer is "exact window title", as learned from
		  // http://support.microsoft.com/kb/q147659/
		  //
		  // "The Visual Basic AppActivate command can only activate a window if
		  // you know the exact window title. Similarly, the Windows FindWindow
		  // function can only find a window handle if you know the exact window title. "
		  
		  #if TargetWin32
		    Soft Declare Function FindWindowA Lib "User32" ( className as Integer, title as CString ) as Integer
		    Soft Declare Function FindWindowW Lib "User32" ( className as Integer, title as WString ) as Integer
		    
		    // Find the window via an exact match using FindWindow
		    dim handle as Integer
		    if System.IsFunctionAvailable( "FindWindowW", "User32" ) then
		      handle = FindWindowW( 0, title )
		    else
		      handle = FindWindowA( 0, title )
		    end if
		    
		    // If we found a handle, then we want to bring it to the front
		    if handle <> 0 then
		      Declare Sub SetForegroundWindow Lib "User32" ( hwnd as Integer )
		      SetForegroundWindow( handle )
		    end if
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ASCIIToScanKey(char as String) As Integer
		  #if TargetWin32
		    Declare Function VkKeyScanA Lib "User32" ( ch as Integer ) as Short
		    
		    dim theAscVal as Integer = Asc( char )
		    return VkKeyScanA( theAscVal )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ATn(d as Double) As Double
		  // Just a straight name change
		  return ATan( d )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ChDir(path as String)
		  // We want to change the active directory.  Easy enough!
		  #if TargetWin32
		    Soft Declare Function SetCurrentDirectoryA Lib "Kernel32" ( dir as CString ) as Boolean
		    Soft Declare Function SetCurrentDirectoryW Lib "Kernel32" ( dir as WString ) as Boolean
		    
		    // Set the directory
		    dim success as Boolean
		    if System.IsFunctionAvailable( "SetCurrentDirectoryW", "Kernel32" ) then
		      success = SetCurrentDirectoryW( path )
		    else
		      success = SetCurrentDirectoryA( path )
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ChDrive(letter as String)
		  // We only want to look at the first character, so let's
		  // ensure that we have only one.
		  letter = Left( letter, 1 )
		  
		  // If the letter is empty, we can bail out
		  if letter = "" then return
		  
		  // Now we want to change the drive.  We can
		  // do this with ChDir.
		  ChDir( letter + ":\" )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Command() As String
		  // This should return just the command line part after the file
		  // name.  It returns it without parsing it.  In REALbasic, the
		  // command line comes with the filename, so we need to parse
		  // that out.
		  
		  // Get the entire command line
		  dim commandLine as String = System.CommandLine
		  
		  // Let the helper function deal with the hard stuff
		  dim temp as String
		  return GetParams( commandLine, temp )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function CurDir(unused as String = "") As String
		  // The user just wants the path to the current directory
		  return GetCurrentDirectory.AbsolutePath
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Date() As String
		  // This one's really easy -- it just gets the current date as a string
		  dim now as new Date
		  return now.LongDate
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Day() As Integer
		  // Returns the day of the month for the current date
		  return now.Day
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub DeleteSetting(appName as String, section as String, key as String = "")
		  // First, we want to get a registry key that points
		  // to the default location of all VB settings.
		  dim base as new RegistryItem( kSettingsLocation )
		  
		  // Now we want to delve into the appName folder
		  base = base.Child( appName )
		  
		  // If we don't have a key name, then we want to
		  // delete the entire section.  Otherwise, we want
		  // to delve into the section and delete the
		  // key specified.
		  if key = "" then
		    base.Delete( section )
		  else
		    // Dive into the section
		    base = base.Child( section )
		    // And delete the key
		    base.Delete( key )
		  end if
		  
		Exception err as RegistryAccessErrorException
		  // Something bad happened, so let's just bail out
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub FileCopy(source as String, dest as String)
		  // Boy, this can't get much easier.
		  dim sourceItem as new FolderItem( source, FolderItem.PathTypeAbsolute )
		  dim destItem as new FolderItem( dest, FolderItem.PathTypeAbsolute )
		  
		  // Copy the source to the dest
		  sourceItem.CopyFileTo( destItem )
		  
		Exception err as UnsupportedFormatException
		  // Bail out
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub FillKeyMap(ByRef map as Dictionary)
		  map.Value( "BACKSPACE" ) = &h8
		  map.Value( "BS" ) = &h8
		  map.Value( "BKSP" ) = &h8
		  map.Value( "BREAK" ) = &h3
		  map.Value( "CAPSLOCK" ) = &h14
		  map.Value( "DELETE" ) = &h2E
		  map.Value( "DEL" ) = &h2E
		  map.Value( "DOWN" ) = &h28
		  map.Value( "END" ) = &h23
		  map.Value( "ENTER" ) = &h0D
		  map.Value( "ESC" ) = &h1B
		  map.Value( "HELP" ) = &h2F
		  map.Value( "HOME" ) = &h24
		  map.Value( "INSERT" ) = &h2D
		  map.Value( "INS" ) = &h2D
		  map.Value( "LEFT" ) = &h25
		  map.Value( "NUMLOCK" ) = &h90
		  map.Value( "PGDN" ) = &h22
		  map.Value( "PGUP" ) = &h21
		  map.Value( "PRTSC" ) = &h2C
		  map.Value( "RIGHT" ) = &h27
		  map.Value( "SCROLLLOCK" ) = &h91
		  map.Value( "TAB" ) = &h09
		  map.Value( "UP" ) = &h26
		  
		  map.Value( "+" ) = ASCIIToScanKey( "+" )
		  map.Value( "^" ) = ASCIIToScanKey( "^" )
		  map.Value( "%" ) = ASCIIToScanKey( "%" )
		  map.Value( "~" ) = ASCIIToScanKey( "~" )
		  map.Value( "(" ) = ASCIIToScanKey( "(" )
		  map.Value( ")" ) = ASCIIToScanKey( ")" )
		  map.Value( "{" ) = ASCIIToScanKey( "{" )
		  map.Value( "}" ) = ASCIIToScanKey( "}" )
		  map.Value( "[" ) = ASCIIToScanKey( "[" )
		  map.Value( "]" ) = ASCIIToScanKey( "]" )
		  
		  for i as Integer = 1 to 16
		    map.Value( "F" + Str( i ) ) = &h70 + (i - 1)
		  next i
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function FillString(char as String, numChars as Integer) As String
		  #if TargetWin32
		    Declare Sub memset lib "msvcrt" ( dest as Ptr, char as Integer, count as Integer )
		    
		    dim mb as new MemoryBlock( LenB( char ) * numChars )
		    memset( mb, AscB( char ), numChars )
		    
		    return DefineEncoding( mb, Encoding( char ) )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Filter(source() as String, match as String, include as Boolean = true, compare as Integer = 1) As String()
		  // We want to filter the entries from source which match the match
		  // string.  The include flag says whether we want to return entries
		  // which do match, or which don't match.  The compare flag says
		  // what type of comparison to use.
		  
		  dim ret(-1) as String
		  for each s as String in source
		    dim add as Boolean
		    select case compare
		    case 0, 1  // Binary or text comparison
		      if StrComp( s, match, compare ) = 0 then add = true
		    else
		      if s = match then add = true
		    end select
		    
		    // If we're doing exclusion, then add means
		    // we don't want to add it
		    if not include and add then add = false
		    
		    // If we want to add it, then do it
		    if add then ret.Append( s )
		  next s
		  
		  // Return the results
		  return ret
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Fix(num as Double) As Integer
		  // This function returns the integer portion of the number passed.
		  // If the number is negative, Fix returns the first negative integer
		  // greater than or equal to the number.  For example, Fix converts -8.4
		  // to -8.  If you want -9, then you should be using Int instead.
		  
		  return Sgn( num ) * Int( Abs( num ) )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FV(rate as Double, nper as Integer, pmt as Double, pv as Double = 0, type as Integer = 0) As Double
		  // These equations come from gnucash
		  // http://www.gnucash.org/docs/v1.8/C/gnucash-guide/loans_calcs1.html
		  dim a as Double = (1 + rate) ^ nper - 1
		  dim b as Double = (1 + rate * type) / rate
		  dim c as Double = pmt * b
		  
		  return -(pv + a * (pv + c))
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetAllSettings(appName as String, section as String) As Dictionary
		  // First, we want to get to the default location for all VB apps
		  dim base as new RegistryItem( kSettingsLocation )
		  
		  // Then we want to dive into the app and section folders
		  base = base.Child( appName ).Child( section )
		  
		  // Loop over all the children and return their values
		  dim i, count as Integer
		  dim ret as new Dictionary
		  
		  // How many keys do we have?
		  count = base.KeyCount
		  
		  for i = 0 to count - 1
		    // Grab the key and value and add it to the dictionary
		    ret.Value( base.Name( i ) ) = base.Value( i )
		  next i
		  
		  // Return our list
		  return ret
		Exception err as RegistryAccessErrorException
		  // Just bail out
		  return nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetCurrentDirectory() As FolderItem
		  #if TargetWin32
		    Soft Declare Sub GetCurrentDirectoryA Lib "Kernel32" ( len as Integer, buf as Ptr )
		    Soft Declare Sub GetCurrentDirectoryW Lib "Kernel32" ( len as Integer, buf as Ptr )
		    
		    dim path as String
		    dim buf as new MemoryBlock( 520 )
		    if System.IsFunctionAvailable( "GetCurrentDirectoryW", "Kernel32" ) then
		      GetCurrentDirectoryW( buf.Size, buf )
		      path = buf.WString( 0 )
		    else
		      GetCurrentDirectoryA( buf.Size, buf )
		      path = buf.CString( 0 )
		    end if
		    
		    return new FolderItem( path, FolderItem.PathTypeAbsolute )
		    
		  #endif
		  
		Exception err as UnsupportedFormatException
		  return nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetParams(commandLine as String, ByRef file as String) As String
		  // Now parse the command line so that we can get just what
		  // is after the application name.  We have to do this one
		  // character at a time, unfortunately
		  dim length as Integer = Len( commandLine )
		  dim ignoreSpaces as Boolean = false
		  
		  for curPos as Integer = 1 to length
		    dim char as String = Mid( commandLine, curPos, 1 )
		    
		    if char = """" then
		      // We found a quote, so we can ignore any spaces.  If
		      // this is the second quote, then we can pay attention
		      // to spaces again.
		      ignoreSpaces = not ignoreSpaces
		    elseif not ignoreSpaces and (char = " " or char = Chr( 9 )) then
		      // We have a space.  If we're ignoring spaces, then
		      // it doesn't matter.  But if we're not, then we've found
		      // the end of the application name
		      file = Trim( Left( commandLine, curPos ) )
		      return Trim( Mid( commandLine, curPos ) )
		    end if
		  next curPos
		  
		  file = commandLine
		  return ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetSetting(appName as String, section as String, key as String, default as Variant = "") As Variant
		  // First, we want to get to the default location for all VB apps
		  dim base as new RegistryItem( kSettingsLocation )
		  
		  // Then we want to dive into the app and section folders
		  base = base.Child( appName ).Child( section )
		  
		  // Now we want to get the value from the key
		  return base.Value( key )
		  
		Exception err as RegistryAccessErrorException
		  // Just bail out
		  return default
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Hour() As Integer
		  // Return the hour for the current time
		  return now.Hour
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function InStrRev(source As String, substr As String, startPos As Integer = 1, compare as Integer = 1) As Integer
		  if source = "" then return 0
		  if substr.Len = 0 then return startPos
		  // Similar to InStr, but searches backwards from the given position
		  // (or if startPos = -1, then from the end of the string).
		  // If substr can't be found, returns 0.
		  
		  Dim srcLen As Integer
		  if compare = 0 then
		    srcLen = source.LenB
		  else
		    srcLen = source.Len
		  end if
		  
		  if startPos > srcLen then return 0
		  
		  // Here's an easy way...
		  // There may be a faster implementation, but then again, there may not -- it probably
		  // depends on what you're trying to do.
		  Dim reversedSource As String = StrReverse(source)
		  Dim reversedSubstr As String = StrReverse(substr)
		  Dim reversedPos As Integer
		  if compare = 0 then
		    reversedPos = InStrB( startPos, reversedSource, reversedSubstr )
		  else
		    reversedPos = InStr( startPos, reversedSource, reversedSubstr )
		  end if
		  if reversedPos < 1 then return 0
		  
		  if compare = 0 then
		    return srcLen - reversedPos - substr.LenB + 2
		  else
		    return srcLen - reversedPos - substr.Len + 2
		  end if
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Int(num as Double) As Integer
		  // This function returns the integer portion of the number passed.
		  // If the number is negative, Int returns the first negative integer
		  // less than or equal to the number.  For example, Int converts -8.4
		  // to -9.  If you want -8, then you should be using Fix instead.
		  
		  dim i as Integer = num
		  if num > 0 then
		    return i
		  else
		    return i - 1
		  end if
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IPmt(rate as Double, per as Integer, nper as Integer, pv as Double, fv as Double = 0, type as Integer = 0) As Double
		  // IPmt is the principle for the previous month times the interest rate
		  // http://www.gnome.org/projects/gnumeric/doc/gnumeric-IPMT.shtml
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsObject(v as Variant) As Boolean
		  // This is easy -- if the variant holds an object, this is true.  Also, if
		  // it holds nil, then it's true as well.
		  return v.Type = Variant.TypeObject or v.Type = Variant.TypeNil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub KeyDown(virtualKeyCode as Integer, extendedKey as Boolean = false)
		  #if TargetWin32
		    Declare Sub keybd_event Lib "User32" ( keyCode as Integer, scanCode as Integer, _
		    flags as Integer, extraData as Integer )
		    
		    dim flags as Integer
		    Const KEYEVENTF_EXTENDEDKEY = &h1
		    if extendedKey then
		      flags = KEYEVENTF_EXTENDEDKEY
		    end
		    
		    ' Press the key
		    keybd_event( virtualKeyCode, 0, flags, 0 )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub KeyUp(virtualKeyCode as Integer, extendedKey as Boolean = false)
		  #if TargetWin32
		    Declare Sub keybd_event Lib "User32" ( keyCode as Integer, scanCode as Integer, _
		    flags as Integer, extraData as Integer )
		    
		    dim flags as Integer
		    Const KEYEVENTF_EXTENDEDKEY = &h1
		    if extendedKey then
		      flags = KEYEVENTF_EXTENDEDKEY
		    end
		    
		    Const KEYEVENTF_KEYUP = &h2
		    flags = BitwiseOr( flags, KEYEVENTF_KEYUP )
		    keybd_event( virtualKeyCode, 0, flags, 0 )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Kill(path as String)
		  // Oddly enough, this deletes files from the disk.  Also,
		  // it supports wildcard characters such as * (for multiple
		  // characters) and ? (for single characters) as a way to
		  // specify multiple files.
		  
		  // We need to use FindFirstFile as a way to find all the
		  // files that we want to delete.  We will come up with a
		  // list of FolderItems, and then we can just use
		  // FolderItem.Delete on them.
		  
		  // Make sure the path points to our current directory as well
		  dim curDir as String = GetCurrentDirectory.AbsolutePath
		  path = curDir + path
		  
		  dim toBeDeleted( -1 ) as FolderItem
		  #if TargetWin32
		    Soft Declare Function FindFirstFileA Lib "Kernel32" ( name as CString, data as Ptr ) as Integer
		    Soft Declare Function FindFirstFileW Lib "Kernel32" ( name as WString, data as Ptr ) as Integer
		    Soft Declare Function FindNextFileA Lib "Kernel32" ( handle as Integer, data as Ptr ) as Boolean
		    Soft Declare Function FindNextFileW Lib "Kernel32" ( handle as Integer, data as Ptr ) as Boolean
		    Declare Sub FindClose Lib "Kernel32" ( handle as Integer )
		    
		    // Check to see whether we're doing unicode processing or not
		    dim isUnicode as Boolean = false
		    if System.IsFunctionAvailable( "FindNextFileW", "Kernel32" ) then isUnicode = true
		    
		    // Get the search handle
		    dim searchHandle as Integer
		    dim searchData as new MemoryBlock( 44 + 520 + 28 )
		    if isUnicode then
		      searchHandle = FindFirstFileW( path, searchData )
		    else
		      searchHandle = FindFirstFileA( path, searchData )
		    end if
		    
		    // If the search handle is 0, then we know that something's wrong and
		    // we can bail out
		    if searchHandle = 0 then return
		    
		    // Loop over all the files and add them to our kill list
		    dim done as Boolean
		    do
		      // Add the file to our delete list
		      try
		        if isUnicode then
		          toBeDeleted.Append( new FolderItem( curDir + searchData.WString( 44 ), FolderItem.PathTypeAbsolute ) )
		        else
		          toBeDeleted.Append( new FolderItem( curDir + searchData.CString( 44 ), FolderItem.PathTypeAbsolute ) )
		        end if
		      catch err as UnsupportedFormatException
		        // We had an error, but I think we should keep trying.
		      end try
		      
		      // Find the next file in our list
		      if isUnicode then
		        done = not FindNextFileW( searchHandle, searchData )
		      else
		        done = not FindNextFileA( searchHandle, searchData )
		      end if
		    loop until done
		    
		    // Close the search handle
		    FindClose( searchHandle )
		  #endif
		  
		  // Now we can loop over all the files to be deleted
		  // and delete them
		  for each item as FolderItem in toBeDeleted
		    item.Delete
		  next item
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Like(toSearch As String, matchingPattern As String) As Boolean
		  Static re As RegEx
		  
		  if re = nil then re = new RegEx
		  
		  // convert Like syntax to RegEx syntax
		  matchingPattern = matchingPattern.ReplaceAll(".", "\.")
		  matchingPattern = matchingPattern.ReplaceAll("*", ".*")
		  matchingPattern = matchingPattern.ReplaceAll("#","\d")
		  matchingPattern = matchingPattern.ReplaceAll("[!", "[^")
		  
		  // special replace for "[x]" syntax in Like
		  re.SearchPattern = "\[(.)\]" // match 1 char in brackets
		  re.ReplacementPattern = "\\\1"
		  re.Options.ReplaceAllMatches = True
		  matchingPattern = re.Replace(matchingPattern)
		  
		  // special replace for "?"
		  re.SearchPattern = "(?<!\\)\?"
		  re.ReplacementPattern = "."
		  matchingPattern = re.Replace(matchingPattern)
		  
		  // now set up RegEx
		  re.SearchPattern = "^" + matchingPattern + "$"
		  
		  // and see if it matches toSearch
		  if nil = re.Search(toSearch) then
		    // no match found?
		    return false
		  else
		    // it did match
		    return true
		  end if
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LSet(ByRef dest as String, assigns source as String)
		  // We want to take the source string and left-align it in the
		  // destination string.  What this essentially does is puts the
		  // source in the left-hand part of dest, and fills the rest of
		  // dest with spaces.  So:
		  //
		  // Dim MyString as String = "0123456789"
		  // Lset MyString = "<-Left"
		  //
		  // Means that MyString contains "<-Left    ".
		  
		  // First, calculate the end length of the destination
		  dim destLen as Integer = Len( dest )
		  dim sourceLen as Integer = Len( source )
		  
		  // If the source string is greater than the
		  // destination, we want to trim the source
		  // string and just be done
		  if sourceLen >= destLen then
		    dest = Left( source, destLen )
		    return
		  end if
		  
		  // Otherwise, we're stuck doing it the "hard" way.
		  // First, assign the source (this would make it left-aligned).
		  dest = source
		  
		  // Then calculate how many spaces we need to
		  // add to fill the rest of the length
		  dim numSpaces as Integer = destLen - sourceLen
		  
		  // Now add the spaces
		  dest = dest + Space( numSpaces )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Mid(ByRef text As String, startPos As Integer, length As Integer = - 1, assigns subStr As String)
		  // handle optional length parameter
		  dim max as Integer = Len(  text )
		  
		  // Assign the replacement to the original data
		  text = left( text, startPos ) + Left( subStr, length ) + _
		  mid( text, startPos + length + 1 )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Minute() As Integer
		  // Return the current minute
		  return now.Minute
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MkDir(name as String)
		  // First, get the current directory
		  dim curDir as FolderItem = GetCurrentDirectory
		  if curDir = nil then return
		  
		  // Now, we want to make a new directory as a
		  // child of the current one
		  dim newDir as FolderItem = curDir.Child( name )
		  newDir.CreateAsFolder
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Month() As Integer
		  // Return the current month
		  return now.Month
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MonthName(month as Integer, abbr as Boolean = false) As String
		  // We want to return the localized name for the month
		  // and abbreviate it if needed.
		  return LocaleInformation.DefaultLocale.MonthName( month, abbr )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Name(oldPathName as String, newPathName as String)
		  // This moves and/or renames a file.  Sound familiar?
		  // Check to see whether the old path or the new path
		  // really are paths.  If they're not, we need to use the
		  // current directory.
		  
		  dim oldPathIsAbsolute, newPathIsAbsolute as Boolean
		  if Mid( oldPathName, 2, 2 ) = ":\" or Left( oldPathName, 2 ) = "//" then
		    oldPathIsAbsolute = true
		  end if
		  
		  if Mid( newPathName, 2, 2 ) = ":\" or Left( newPathName, 2 ) = "//" then
		    newPathIsAbsolute = true
		  end if
		  
		  // Now we can get folder items for both the new and the
		  // old path.
		  dim oldPath as FolderItem
		  if oldPathIsAbsolute then
		    oldPath = new FolderItem( oldPathName, FolderItem.PathTypeAbsolute )
		  else
		    oldPath = GetCurrentDirectory.Child( oldPathName )
		  end if
		  
		  dim newPath as FolderItem
		  if newPathIsAbsolute then
		    newPath = new FolderItem( newPathName, FolderItem.PathTypeAbsolute )
		  else
		    newPath = GetCurrentDirectory.Child( newPathName )
		  end if
		  
		  // Now we can do a move operation.  This will also do a rename if
		  // the oldPath and the newPath reside in the same directory
		  oldPath.MoveFileTo( newPath )
		  
		exception err as UnsupportedFormatException
		  return
		exception err as NilObjectException
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Now() As Date
		  return new Date
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Pmt(rate as Double, nper as Integer, pv as Double, fv as Double = 0, type as Integer = 0) As Double
		  // These equations come from gnucash
		  // http://www.gnucash.org/docs/v1.8/C/gnucash-guide/loans_calcs1.html
		  dim a as Double = (1 + rate) ^ nper - 1
		  dim b as Double = (1 + rate * type) / rate
		  
		  return -(fv + pv * (a + 1)) / (a * b)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function PPmt(rate as Double, per as Integer, nper as Integer, pv as Double, fv as Double = 0, type as Integer = 0) As Double
		  // PPmt is just the Pmt - IPmt, according to
		  // http://www.gnome.org/projects/gnumeric/doc/gnumeric-PPMT.shtml
		  return Pmt( rate, nper, pv, fv, type ) - IPmt( rate, per, nper, pv, fv, type )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function PV(rate as Double, nper as Integer, pmt as Double, fv as Double = 0, type as Integer = 0) As Double
		  // These equations come from gnucash
		  // http://www.gnucash.org/docs/v1.8/C/gnucash-guide/loans_calcs1.html
		  dim a as Double = (1 + rate) ^ nper - 1
		  dim b as Double = (1 + rate * type) / rate
		  dim c as Double = pmt * b
		  
		  return -(fv + a * c) / (a + 1)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function QBColor(index as Integer) As Integer
		  select case index
		  case 0  // Black
		    return &h000000
		  case 1  // Blue
		    return &h800000
		  case 2  // Green
		    return &h8000
		  case 3  // Cyan
		    return &h808000
		  case 4  // Red
		    return &h80
		  case 5  // Magenta
		    return &h800080
		  case 6  // Yellow
		    return &h8080
		  case 7  // White
		    return &hC0C0C0
		  case 8  // Gray
		    return &h808080
		  case 9  // Light blue
		    return &hFF0000
		  case 10  // Light green
		    return &hFF00
		  case 11  // Light cyan
		    return &hFFFF00
		  case 12  // Light red
		    return &hFF
		  case 13  // Light magenta
		    return &hFF00FF
		  case 14  // Light yellow
		    return &hFFFF
		  case 15  // Bright white
		    return &hFFFFFF
		  else
		    return 0
		  end select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Randomize(seed as Integer = - 1)
		  if seed <> -1 then
		    mRnd.Seed = seed
		  else
		    mRnd.Seed = mRnd.Number * &hFFFFFFFF
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Replace(source as String, find as String, rep as String, start as Integer = 1, count as Integer = - 1, compare as Integer = 1) As String
		  // We want to replace the search string a certain number of times
		  // in the source string.  This is different than our Replace function which
		  // only replaces the first instance and ReplaceAll, which replaces all
		  // instances.
		  
		  // Do our santity checks
		  if source = "" then return ""
		  if find = "" then return source
		  if rep = "" then return source
		  if count = 0 then return source
		  
		  // If the user wants to start farther up the string than at
		  // the first character, we need to do some wiggling since
		  // REALbasic doesn't let you do specify a start position for
		  // the source string in Replace
		  dim searchStr as String = Mid( source, start )
		  dim curPos as Integer = 1
		  
		  if count = -1 then
		    // We just want to do a replace all
		    if compare = 0 then
		      searchStr = REALbasic.ReplaceAllB( searchStr, find, rep )
		    else
		      searchStr = REALbasic.ReplaceAll( searchStr, find, rep )
		    end if
		  else
		    // Now we want to do the replaces over and over again.
		    while count > 0
		      if compare = 0 then
		        searchStr = REALbasic.ReplaceB( searchStr, find, rep )
		      else
		        searchStr = REALbasic.Replace( searchStr, find, rep )
		      end if
		      
		      // We have one less replace to do
		      count = count - 1
		    wend
		  end if
		  
		  // Now we're set.  The only thing we might have to do
		  // is reconstitute the original part of the search string if
		  // start is greater than 1.
		  if start > 1 then
		    return Left( source, start - 1 ) + searchStr
		  else
		    return searchStr
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub RmDir(path as String)
		  // Check to see if the path is an absolute path, or
		  // just a local one
		  dim itemToDelete as FolderItem
		  if Mid( path, 2, 2 ) = ":\" or Left( path, 2 ) = "//" then
		    itemToDelete = new FolderItem( path, FolderItem.PathTypeAbsolute )
		  else
		    itemToDelete = GetCurrentDirectory.Child( path )
		  end if
		  
		  // Then delete the item
		  itemToDelete.Delete
		  
		Exception err as UnsupportedFormatException
		  return
		Exception err as NilObjectException
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Rnd() As Double
		  return mRnd.Number
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub RSet(ByRef dest as String, assigns source as String)
		  // This is the sibling to RSet.
		  //
		  // Dim MyString as String = "0123456789"
		  // Rset MyString = "Right->"
		  //
		  // MyString contains "   Right->".
		  
		  // First, calculate the end length of the destination
		  dim destLen as Integer = Len( dest )
		  dim sourceLen as Integer = Len( source )
		  
		  // If the source string is greater than the
		  // destination, we want to trim the source
		  // string and just be done
		  if sourceLen >= destLen then
		    dest = Right( source, destLen )
		    return
		  end if
		  
		  // Otherwise, we're stuck doing it the "hard" way.
		  
		  // First, calculate how many spaces we need to
		  // add to fill the rest of the length
		  dim numSpaces as Integer = destLen - sourceLen
		  
		  // Now add the spaces
		  dest = Space( numSpaces )
		  
		  // Then, assign the source (this would make it right-aligned).
		  dest = dest + source
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function SaveAsFormatFromName(name As String) As Integer
		  // Added by Kem Tekinay.
		  // Determines the Save As format from the extension of the filename given.
		  // Returns MostCompatible if it can't be identified.
		  
		  dim parts() as string = name.Split( "." )
		  dim ext as string = parts( parts.Ubound )
		  
		  select case ext
		  case "bmp"
		    return Picture.SaveAsWindowsBMP
		    
		  case "png"
		    return Picture.SaveAsPNG
		    
		  case "jpg", "jpeg"
		    return Picture.SaveAsJPEG
		    
		  case "gif"
		    return Picture.SaveAsGIF
		    
		  case "emf"
		    return Picture.SaveAsWindowsEMF
		    
		  case "wmf"
		    return Picture.SaveAsWindowsWMF
		    
		  case "tif", "tiff"
		    return Picture.SaveAsTIFF
		    
		  else
		    return Picture.SaveAsMostCompatible
		    
		  end select
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub SavePicture(p as Picture, name as String)
		  // We want to save the given picture in a file, which
		  // REALbasic pretty much already handles for you.
		  
		  // Check to see if the path is an absolute path, or
		  // just a local one
		  dim fileToSave as FolderItem
		  fileToSave = GetCurrentDirectory.Child( name )
		  
		  // Then save the picture out
		  'fileToSave.SaveAsPicture( p )
		  p.Save( fileToSave, SaveAsFormatFromName( name ) ) // Modified by Kem Tekinay,
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub SaveSetting(appName as String, section as String, key as String, setting as Variant)
		  // First, we want to get to the default location for all VB apps
		  dim base as new RegistryItem( kSettingsLocation )
		  
		  // Then we want to dive into the app and section folders
		  base = base.Child( appName ).Child( section )
		  
		  // Now we want to save the key and value
		  base.Value( key ) = setting
		  
		Exception err as RegistryAccessErrorException
		  // Just bail out
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Second() As Integer
		  // Return the current second
		  return now.Second
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub SendKeys(keys as String, unused as Boolean = false)
		  // We want to initialize all of our virtual keys.  Some of them
		  // are going to be constants, others will be figured out while
		  // we parse, and still others will reside in a map.
		  Const VK_SHIFT = &h10
		  Const VK_CONTROL = &h11
		  Const VK_MENU = &h12
		  Const VK_ENTER = &h0D
		  
		  Static sMap as Dictionary
		  
		  if sMap = nil then
		    // Create our map
		    sMap = new Dictionary
		    
		    // Fill the map out (note that this is ByRef)
		    FillKeyMap( sMap )
		  end if
		  
		  // We have to write a very simple parser to parse
		  // the keys string that's passed in.
		  
		  // Split the entire string into a set of tokens.  Each token
		  // is a single character.
		  dim chars( -1 ) as String = Split( keys, "" )
		  
		  // This holds the current virtual key (so we don't have to
		  // process the ASCII character multiple times).
		  dim virtKey as Integer
		  
		  // This holds the current set of depressed modifier keys.
		  // As we press the keys, we should be adding to this list,
		  // and when we need to release the keys, we know which
		  // ones to generate a key up for.
		  dim modifierList( -1 ) as Integer
		  
		  // Sometimes we want to hold the modifiers for a while, for
		  // instance, if the user enters +(EC), then we want to hold
		  // the shift key for E and C.  But other times, we only want the
		  // modifier held for one key press, such as +EC, in which case
		  // there would be a Shift+E, and then a C.
		  dim holdModifiers as Boolean
		  
		  // We may need to do some special processing for things inside
		  // of a {} tag, such as {TAB} or {LEFT 42}
		  dim specialProcessing as Boolean
		  dim specialProcessingStr as String
		  
		  // Loop over every token and do the appropriate
		  // action.
		  for each token as String in chars
		    // If we're doing special processing, then we
		    // should be doing that instead of the normal
		    // processing.
		    if specialProcessing then
		      // Check to see whether we have a { or
		      // not.  If we have one, and we've already
		      // processed at least one character, then we
		      // are done with our special processing.  The
		      // reason we check to see whether we have a
		      // character or not is because of the string {}},
		      // which should print a } character.
		      if token = "}" and Len( specialProcessingStr ) > 0 then
		        specialProcessing = false
		      else
		        // Add the character to our processing string
		        specialProcessingStr = specialProcessingStr + token
		      end if
		    end if
		    
		    // If we're still doing special processing, then we want
		    // to continue doing it.  But since the state may have changed
		    // we need to check again
		    if specialProcessing then continue
		    
		    select case token
		    case "("
		      // We're starting a token group.  The group
		      // should have the current modifiers applied
		      // to it.
		      holdModifiers = true
		      
		    case ")"
		      // If we're not holding modifies, then
		      // that means we've never gotten a ( and
		      // something is wrong
		      if not holdModifiers then return
		      
		      // We're ending a token group.  This means
		      // that we can clear the modifiers.
		      for each modifier as Integer in modifierList
		        // Release the key
		        KeyUp( modifier )
		      next modifier
		      
		      // Now clear our list
		      Redim modifierList( -1 )
		      
		      // We're no longer holding the modifiers
		      holdModifiers = false
		      
		    case "{"
		      // We have a special token to parse, such as
		      // {TAB} or {F1}.  We should search until we
		      // get to } and figure out what key to press
		      // from there.
		      //
		      // This could also be a repeat modifier that is
		      // in the form {key number}.  If we find a space
		      // while parsing, then we know we have this form.
		      
		      // So note that we need to parse until we hit the }.
		      // Once we have that, we can do the processing.
		      specialProcessing = true
		      
		    case "}"
		      // Our special processing is done now, so we should
		      // check the data we have in the string.  It could be
		      // in the form "key number", or it could just be "key".
		      
		      // First, get the key from our map.  If we don't have
		      // the key in the map, then we should try a regular
		      // ASCII key.
		      dim firstField as String = NthField( specialProcessingStr, " ", 1 )
		      dim keyToken as Integer = sMap.Lookup( firstField, -1 )
		      if keyToken = -1 then
		        keyToken = ASCIIToScanKey( firstField )
		      end if
		      
		      // If we're still in a bad state, then we should
		      // bail out.
		      if keyToken = 0 then return
		      
		      // We want our repeats to always be at least one, but
		      // possibly more, if that field exists.
		      dim numRepeats as Integer = Max( Val( NthField( specialProcessingStr, " ", 2 ) ), 1 )
		      for count as Integer = 1 to numRepeats
		        // Press the key and then release it
		        KeyDown( keyToken )
		        KeyUp( keyToken )
		      next count
		      
		      // Clear out our special processing data
		      specialProcessingStr = ""
		    case "~"
		      // We want to send an enter key
		      KeyDown( VK_ENTER )
		      KeyUp( VK_ENTER )
		      
		    case "+"
		      // The shift key modifier should be pressed
		      if modifierList.IndexOf( VK_SHIFT ) = -1 then
		        KeyDown( VK_SHIFT )
		        modifierList.Append( VK_SHIFT )
		      end if
		      
		    case "^"
		      // The control key modifier should be pressed
		      if modifierList.IndexOf( VK_CONTROL ) = -1 then
		        KeyDown( VK_CONTROL )
		        modifierList.Append( VK_CONTROL )
		      end if
		      
		    case "%"
		      // The Alt key modifier should be pressed
		      if modifierList.IndexOf( VK_MENU ) = -1 then
		        KeyDown( VK_MENU )
		        modifierList.Append( VK_MENU )
		      end if
		      
		    else
		      // We have a regular key press, such as A or 4.
		      virtKey = ASCIIToScanKey( token )
		      
		      // Check to see if the scan key we got back
		      // has any of the modifier keys pressed
		      // or not.
		      Dim releaseShift as Boolean
		      if Bitwise.BitAnd( virtKey, &h100 ) = &h100 then
		        KeyDown( VK_SHIFT )
		        releaseShift = true
		      end if
		      
		      Dim releaseControl as Boolean
		      if Bitwise.BitAnd( virtKey, &h200 ) = &h200 then
		        KeyDown( VK_CONTROL )
		        releaseControl = true
		      end if
		      
		      Dim releaseAlt as Boolean
		      if Bitwise.BitAnd( virtKey, &h400 ) = &h400 then
		        KeyDown( VK_MENU )
		        releaseAlt = true
		      end if
		      
		      // Press the key and then release it
		      KeyDown( virtKey )
		      KeyUp( virtKey )
		      
		      if releaseAlt then KeyUp( VK_MENU )
		      if releaseControl then KeyUp( VK_CONTROL )
		      if releaseShift then KeyUp( VK_SHIFT )
		      
		      // If we aren't holding modifiers, then we
		      // should release them all here.
		      if not holdModifiers then
		        for each modifier as Integer in modifierList
		          // Release the key
		          KeyUp( modifier )
		        next modifier
		        
		        // Now clear our list
		        Redim modifierList( -1 )
		      end if
		    end select
		  next token
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Sgn(number as Double) As Integer
		  if number = 0 then return 0
		  if number < 0 then return -1
		  if number > 0 then return 1
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Shell(pathname as String, style as Integer = 1) As Integer
		  // We want to launch the application given by the path name, and
		  // we need to return the application's PID if the launch was successful
		  #if TargetWin32
		    Soft Declare Function CreateProcessW Lib "Kernel32" ( appName as Integer, params as WString, _
		    procAttribs as Integer, threadAttribs as Integer, inheritHandles as Boolean, flags as Integer, _
		    env as Integer, curDir as Integer, startupInfo as Ptr, procInfo as Ptr ) as Boolean
		    
		    Soft Declare Function CreateProcessA Lib "Kernel32" ( appName as Integer, params as CString, _
		    procAttribs as Integer, threadAttribs as Integer, inheritHandles as Boolean, flags as Integer, _
		    env as Integer, curDir as Integer, startupInfo as Ptr, procInfo as Ptr ) as Boolean
		    
		    dim startupInfo, procInfo as MemoryBlock
		    
		    startupInfo = new MemoryBlock( 17 * 4 )
		    procInfo = new MemoryBlock( 16 )
		    
		    dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "CreateProcessW", "Kernel32" )
		    
		    startupInfo.Long( 0 ) = startupInfo.Size
		    
		    // Create the application
		    dim ret as Boolean
		    if unicodeSavvy then
		      ret = CreateProcessW( 0, pathname, 0, 0, false, 0, 0, 0, startupInfo, procInfo )
		    else
		      ret = CreateProcessA( 0, pathname, 0, 0, false, 0, 0, 0, startupInfo, procInfo )
		    end if
		    
		    // If we couldn't make it, then we're stuck
		    if not ret then return 0
		    
		    // We want to return the process identifier for the application
		    dim retVal as Integer = procInfo.Long( 8 )
		    
		    // We should wait for the input idle so that we can switch to the app
		    Declare Function WaitForInputIdle Lib "User32" ( handle as Integer, wait as Integer ) as Integer
		    dim wait as Integer = WaitForInputIdle( procInfo.Long( 0 ), 2500 )
		    
		    // Clean the application up
		    Declare Sub CloseHandle Lib "Kernel32" ( handle as Integer )
		    CloseHandle( procInfo.Long( 0 ) )
		    CloseHandle( procInfo.Long( 4 ) )
		    
		    return retVal
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Space(num as Integer) As String
		  // Return a string with the proper number
		  // of spaces.  I really wish memset was an
		  // option for MemoryBlocks.  Grr!
		  
		  return FillString( " ", num )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Sqr(d as Double) As Double
		  // This is the squareroot function, which RB has a different name for
		  return Sqrt( d )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Stop()
		  // This behaves just like Break, so do that.
		  Break
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function StrReverse(s As String) As String
		  // Return s with the characters in reverse order.
		  
		  if Len(s) < 2 then return s
		  
		  Dim m As MemoryBlock
		  Dim c As String
		  Dim pos, mpos, csize As Integer
		  
		  m = new MemoryBlock( s.LenB )
		  
		  pos = 1
		  mpos = m.Size
		  while mpos > 0
		    c = Mid( s, pos, 1 )
		    csize = c.LenB
		    mpos = mpos - csize
		    m.StringValue( mpos, csize ) = c
		    pos = pos + 1
		  wend
		  
		  return DefineEncoding( m.StringValue(0, m.Size), s.Encoding )
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Time() As Date
		  return new Date
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Timer() As Single
		  // This returns the number of seconds that have elapsed since
		  // midnight.  Weird...
		  
		  // Get the current time
		  dim d as Date = now
		  
		  // Calculate the number of seconds since
		  // midnight.  We do this the cheap way.
		  dim midnight as new Date
		  midnight.Hour = 0
		  
		  // Now that we have midnight and now, we
		  // can figure out the number of seconds by
		  // subtracting the totalseconds of each
		  return now.TotalSeconds - midnight.TotalSeconds
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Year() As Integer
		  // Return the current year
		  return now.Year
		End Function
	#tag EndMethod


	#tag ComputedProperty, Flags = &h21
		#tag Getter
			Get
			  static r as new Random
			  return r
			End Get
		#tag EndGetter
		Private mRnd As Random
	#tag EndComputedProperty


	#tag Constant, Name = kSettingsLocation, Type = String, Dynamic = False, Default = \"HKEY_CURRENT_USER\\Software\\VB and VBA Program Settings\\", Scope = Private
	#tag EndConstant


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
