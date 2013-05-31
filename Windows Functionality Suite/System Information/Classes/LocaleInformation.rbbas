#tag Class
Protected Class LocaleInformation
	#tag Method, Flags = &h21
		Private Shared Function CommonLocaleInfoProc(idStr as String) As Boolean
		  dim localID as Integer
		  if Left( idStr, 1 ) = "0" then
		    localID = Val( "&h" + idStr )
		  else
		    localID = Val( idStr )
		  end if
		  
		  sLocales.Append( new LocaleInformation( localID ) )
		  
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(localeID as Integer = - 1)
		  #if TargetWin32
		    
		    Dim mb as new MemoryBlock( 2048 )
		    Dim ret as Integer
		    Dim retVal as String
		    
		    Const LOCALE_USER_DEFAULT = &H400
		    if localeID = -1 then localeID = LOCALE_USER_DEFAULT
		    
		    LCID = localeID
		    
		    Const LOCALE_ICENTURY = &h00000024
		    ret = GetLocaleInfo( LOCALE_ICENTURY, mb, retVal )
		    UseFourDigitCentury = retVal = "1"
		    
		    Const LOCALE_SENGCOUNTRY = &H1002 ' English name of country
		    ret = GetLocaleInfo( LOCALE_SENGCOUNTRY, mb, EnglishCountryName )
		    
		    Const LOCALE_SENGLANGUAGE = &H1001  ' English name of language
		    ret = GetLocaleInfo( LOCALE_SENGLANGUAGE, mb, EnglishLanguageName )
		    
		    Const LOCALE_SNATIVELANGNAME = &H4  ' native name of language
		    ret = GetLocaleInfo( LOCALE_SNATIVELANGNAME, mb, NativeLanguageName )
		    
		    Const LOCALE_SNATIVECTRYNAME = &H8  ' native name of country
		    ret = GetLocaleInfo( LOCALE_SNATIVECTRYNAME, mb, NativeCountryName )
		    
		    Const LOCALE_IDIGITS = &h0000011
		    ret = GetLocaleInfo( LOCALE_IDIGITS, mb, retVal )
		    NumberOfFractionalDigits = Val( retVal )
		    
		    Const LOCALE_ILZERO = &h0000012
		    ret = GetLocaleInfo( LOCALE_ILZERO, mb, retVal )
		    UseLeadingZeros = retVal = "1"
		    
		    Const LOCALE_IMEASURE = &h000000D
		    ret = GetLocaleInfo( LOCALE_IMEASURE, mb, retVal )
		    UseMetricMeasurements = retVal = "0"
		    
		    Const LOCALE_ITIME = &h0000023
		    ret = GetLocaleInfo( LOCALE_ITIME, mb, retVal )
		    Use24HourClock = retVal = "1"
		    
		    Const LOCALE_S1159 = &h0000028
		    ret = GetLocaleInfo( LOCALE_S1159, mb, Am )
		    
		    Const LOCALE_S2359 = &h0000029
		    ret = GetLocaleInfo( LOCALE_S2359, mb, Pm )
		    
		    Const LOCALE_SDAYNAME1 = 42'&h0000002A
		    Const LOCALE_SDAYNAME2 = &h0000002B
		    Const LOCALE_SDAYNAME3 = &h0000002C
		    Const LOCALE_SDAYNAME4 = &h0000002D
		    Const LOCALE_SDAYNAME5 = &h0000002E
		    Const LOCALE_SDAYNAME6 = &h0000002F
		    Const LOCALE_SDAYNAME7 = &h00000030
		    dim dayConst( 7 ) as Integer
		    dayConst = Array( LOCALE_SDAYNAME1, LOCALE_SDAYNAME2, LOCALE_SDAYNAME3, _
		    LOCALE_SDAYNAME4, LOCALE_SDAYNAME5, LOCALE_SDAYNAME6, LOCALE_SDAYNAME7 )
		    
		    dim i as Integer
		    for i = 0 to 6
		      ret = GetLocaleInfo( dayConst( i ), mb, retVal )
		      DayNames.Append( retVal )
		    next
		    
		    Const LOCALE_SABBREVDAYNAME1 = 49'&h00000031
		    Const LOCALE_SABBREVDAYNAME2 = &h00000032
		    Const LOCALE_SABBREVDAYNAME3 = &h00000033
		    Const LOCALE_SABBREVDAYNAME4 = &h00000034
		    Const LOCALE_SABBREVDAYNAME5 = &h00000035
		    Const LOCALE_SABBREVDAYNAME6 = &h00000036
		    Const LOCALE_SABBREVDAYNAME7 = &h00000037
		    dim dayAbbrConst( 7 ) as Integer
		    dayAbbrConst = Array( LOCALE_SABBREVDAYNAME1, LOCALE_SABBREVDAYNAME2, LOCALE_SABBREVDAYNAME3, _
		    LOCALE_SABBREVDAYNAME4, LOCALE_SABBREVDAYNAME5, LOCALE_SABBREVDAYNAME6, LOCALE_SABBREVDAYNAME7 )
		    
		    for i = 0 to 6
		      ret = GetLocaleInfo( dayAbbrConst( i ), mb, retVal )
		      AbbreviatedDayNames.Append( retVal )
		    next
		    
		    Const LOCALE_SMONTHNAME1 = 56'&h00000038
		    Const LOCALE_SMONTHNAME2 = &h00000039
		    Const LOCALE_SMONTHNAME3 = &h0000003A
		    Const LOCALE_SMONTHNAME4 = &h0000003B
		    Const LOCALE_SMONTHNAME5 = &h0000003C
		    Const LOCALE_SMONTHNAME6 = &h0000003D
		    Const LOCALE_SMONTHNAME7 = &h0000003E
		    Const LOCALE_SMONTHNAME8 = &h0000003F
		    Const LOCALE_SMONTHNAME9 = &h00000040
		    Const LOCALE_SMONTHNAME10 = &h00000041
		    Const LOCALE_SMONTHNAME11 = &h00000042
		    Const LOCALE_SMONTHNAME12 = &h00000043
		    Const LOCALE_SMONTHNAME13 = &h0000100E
		    dim monthConst( 13 ) as Integer
		    monthConst = Array( LOCALE_SMONTHNAME1, LOCALE_SMONTHNAME2, LOCALE_SMONTHNAME3, _
		    LOCALE_SMONTHNAME4, LOCALE_SMONTHNAME5, LOCALE_SMONTHNAME6, LOCALE_SMONTHNAME7, _
		    LOCALE_SMONTHNAME8, LOCALE_SMONTHNAME9, LOCALE_SMONTHNAME10, LOCALE_SMONTHNAME11, _
		    LOCALE_SMONTHNAME12, LOCALE_SMONTHNAME13 )
		    for i = 0 to 12
		      ret = GetLocaleInfo( monthConst( i ), mb, retVal )
		      MonthNames.Append( retVal )
		    next
		    
		    Const LOCALE_SABBREVMONTHNAME1 = 68'&h00000044
		    Const LOCALE_SABBREVMONTHNAME2 = &h00000045
		    Const LOCALE_SABBREVMONTHNAME3 = &h00000046
		    Const LOCALE_SABBREVMONTHNAME4 = &h00000047
		    Const LOCALE_SABBREVMONTHNAME5 = &h00000048
		    Const LOCALE_SABBREVMONTHNAME6 = &h00000049
		    Const LOCALE_SABBREVMONTHNAME7 = &h0000004A
		    Const LOCALE_SABBREVMONTHNAME8 = &h0000004B
		    Const LOCALE_SABBREVMONTHNAME9 = &h0000004C
		    Const LOCALE_SABBREVMONTHNAME10 = &h0000004D
		    Const LOCALE_SABBREVMONTHNAME11 = &h0000004E
		    Const LOCALE_SABBREVMONTHNAME12 = &h0000004F
		    Const LOCALE_SABBREVMONTHNAME13 = &h0000100F
		    dim monthConstAbbr( 13 ) as Integer
		    monthConstAbbr = Array( LOCALE_SABBREVMONTHNAME1, LOCALE_SABBREVMONTHNAME2, LOCALE_SABBREVMONTHNAME3, _
		    LOCALE_SABBREVMONTHNAME4, LOCALE_SABBREVMONTHNAME5, LOCALE_SABBREVMONTHNAME6, LOCALE_SABBREVMONTHNAME7, _
		    LOCALE_SABBREVMONTHNAME8, LOCALE_SABBREVMONTHNAME9, LOCALE_SABBREVMONTHNAME10, LOCALE_SABBREVMONTHNAME11, _
		    LOCALE_SABBREVMONTHNAME12, LOCALE_SABBREVMONTHNAME13 )
		    for i = 0 to 12
		      ret = GetLocaleInfo( monthConstAbbr( i ), mb, retVal )
		      AbbreviatedMonthNames.Append( retVal )
		    next
		    
		    Const LOCALE_SCURRENCY = &h00000014
		    ret = GetLocaleInfo( LOCALE_SCURRENCY, mb, CurrencySymbol )
		    
		    Const LOCALE_SDATE = &h0000001D
		    ret = GetLocaleInfo( LOCALE_SDATE, mb, DateSeparator )
		    
		    Const LOCALE_STIME = &h0000001E
		    ret = GetLocaleInfo( LOCALE_STIME, mb, TimeSeparator )
		    
		    Const LOCALE_ICALENDARTYPE = &h1009
		    ret = GetLocaleInfo( LOCALE_ICALENDARTYPE, mb, retVal )
		    CalendarType = Val( retVal )
		    
		    Const LOCALE_ICOUNTRY = &h5
		    ret = GetLocaleInfo( LOCALE_ICOUNTRY, mb, CountryCode )
		    
		    Const LOCALE_ICURRDIGITS = &h19
		    ret = GetLocaleInfo( LOCALE_ICURRDIGITS, mb, retVal )
		    NumberOfCurrencyDigits = Val( retVal )
		    
		    Const LOCALE_ICURRENCY = &h1B
		    ret = GetLocaleInfo( LOCALE_ICURRENCY, mb, retVal )
		    select case retVal
		    case "0"
		      PositiveCurrencyDisplayFormat = "$X.Y"
		    case "1"
		      PositiveCurrencyDisplayFormat = "X.Y$"
		    case "2"
		      PositiveCurrencyDisplayFormat = "$ X.Y"
		    case "3"
		      PositiveCurrencyDisplayFormat = "X.Y $"
		    end select
		    
		    Const LOCALE_INEGCURR = &h1C
		    ret = GetLocaleInfo( LOCALE_INEGCURR, mb, retVal )
		    select case retVal
		    case "0"
		      NegativeCurrencyDisplayFormat = "($X.Y)"
		    case "1"
		      NegativeCurrencyDisplayFormat = "-$X.Y"
		    case "2"
		      NegativeCurrencyDisplayFormat = "$-X.Y"
		    case "3"
		      NegativeCurrencyDisplayFormat = "$X.Y-"
		    case "4"
		      NegativeCurrencyDisplayFormat = "(X.Y$)"
		    case "5"
		      NegativeCurrencyDisplayFormat = "-X.Y$"
		    case "6"
		      NegativeCurrencyDisplayFormat = "X.Y-$"
		    case "7"
		      NegativeCurrencyDisplayFormat = "X.Y$-"
		    case "8"
		      NegativeCurrencyDisplayFormat = "-X.Y $"
		    case "9"
		      NegativeCurrencyDisplayFormat = "-$ X.Y"
		    case "10"
		      NegativeCurrencyDisplayFormat = "X.Y $-"
		    case "11"
		      NegativeCurrencyDisplayFormat = "$ X.Y-"
		    case "12"
		      NegativeCurrencyDisplayFormat = "$ -X.Y"
		    case "13"
		      NegativeCurrencyDisplayFormat = "X.Y- $"
		    case "14"
		      NegativeCurrencyDisplayFormat = "($ X.Y)"
		    case "15"
		      NegativeCurrencyDisplayFormat = "(X.Y $)"
		      
		    end select
		    
		    Const LOCALE_SDECIMAL = &hE
		    ret = GetLocaleInfo( LOCALE_SDECIMAL, mb, DecimalSeparator )
		    
		    Const LOCALE_IDATE = &h21
		    ret = GetLocaleInfo( LOCALE_IDATE, mb, retVal )
		    select case retVal
		    case "0"
		      ShortDateFormat = "Month-Day-Year"
		    case "1"
		      ShortDateFormat = "Day-Month-Year"
		    case "2"
		      ShortDateFormat = "Year-Month-Day"
		    end select
		    
		    Const LOCALE_IDAYLZERO = &h26
		    ret = GetLocaleInfo( LOCALE_IDAYLZERO, mb, retVal )
		    UseLeadingZerosForDay = retVal = "1"
		    
		    Const LOCALE_IMONTHLZERO = &h27
		    ret = GetLocaleInfo( LOCALE_IMONTHLZERO, mb, retVal )
		    UseLeadingZerosForMonth = retVal = "1"
		    
		    Const LOCALE_ITLZERO = &h25
		    ret = GetLocaleInfo( LOCALE_ITLZERO, mb, retVal )
		    
		    Const LOCALE_IDEFAULTCOUNTRY = &hA
		    ret = GetLocaleInfo( LOCALE_IDEFAULTCOUNTRY, mb, DefaultCountryCode )
		    UseLeadingZerosForHour = retVal = "1"
		    
		    Const LOCALE_IDEFAULTLANGUAGE = &h9
		    ret = GetLocaleInfo( LOCALE_IDEFAULTLANGUAGE, mb, DefaultLanguage )
		    
		    Const LOCALE_IFIRSTDAYOFWEEK  = &h100C
		    ret = GetLocaleInfo( LOCALE_IFIRSTDAYOFWEEK, mb, retVal )
		    FirstDayOfWeek = Val( retVal )
		    
		    Const LOCALE_IFIRSTWEEKOFYEAR = &h100D
		    ret = GetLocaleInfo( LOCALE_IFIRSTWEEKOFYEAR, mb, retVal )
		    FirstWeekOfYear = Val( retVal )
		    
		    Const LOCALE_IINTLCURRDIGITS = &h1A
		    ret = GetLocaleInfo( LOCALE_IINTLCURRDIGITS, mb, retVal )
		    NumberOfInternationalCurrencyDigits = Val( retVal )
		    
		    Const LOCALE_ILDATE = &h22
		    ret = GetLocaleInfo( LOCALE_ILDATE, mb, retVal )
		    select case retVal
		    case "0"
		      LongDateFormat = "Month-Day-Year"
		    case "1"
		      LongDateFormat = "Day-Month-Year"
		    case "2"
		      LongDateFormat = "Year-Month-Day"
		    end select
		    
		    Const LOCALE_SNEGATIVESIGN = &h51
		    ret = GetLocaleInfo( LOCALE_SNEGATIVESIGN, mb, NegativeSymbol )
		    
		    Const LOCALE_SPOSITIVESIGN = &h50
		    ret = GetLocaleInfo( LOCALE_SPOSITIVESIGN, mb, PositiveSymbol )
		    
		    Const LOCALE_ITIMEMARKPOSN = &h1005
		    ret = GetLocaleInfo( LOCALE_ITIMEMARKPOSN, mb, retVal )
		    TimeMarkerIsPrefix = retVal = "1"
		    
		    Const LOCALE_SABBREVCTRYNAME = &h7
		    ret = GetLocaleInfo( LOCALE_SABBREVCTRYNAME, mb, AbbreviatedCountry )
		    
		    Const LOCALE_SABBREVLANGNAME = &h3
		    ret = GetLocaleInfo( LOCALE_SABBREVLANGNAME, mb, AbbreviatedLanguage )
		    
		    Const LOCALE_SENGCURRNAME  = &h1007
		    ret = GetLocaleInfo( LOCALE_SENGCURRNAME, mb, EnglishCurrencyName )
		    
		    Const LOCALE_SNATIVECURRNAME = &h1008
		    ret = GetLocaleInfo( LOCALE_SNATIVECURRNAME, mb, NativeCurrencyName )
		    
		    Const LOCALE_SINTLSYMBOL = &h15
		    ret = GetLocaleInfo( LOCALE_SINTLSYMBOL, mb, InternationalMonetarySymbol )
		    
		    Const LOCALE_SLIST = &hC
		    ret = GetLocaleInfo( LOCALE_SLIST, mb, ListSymbol )
		    
		    Const LOCALE_SMONDECIMALSEP = &h16
		    ret = GetLocaleInfo( LOCALE_SMONDECIMALSEP, mb, MonetaryDecimalSeparator )
		    
		    Const LOCALE_SNATIVEDIGITS = &h13
		    ret = GetLocaleInfo( LOCALE_SNATIVEDIGITS, mb, retVal )
		    for i = 0 to 9
		      NativeDigits.Append( Mid( retVal, i + 1, 1 ) )
		    next i
		    
		    Const LOCALE_STHOUSAND = &hF
		    ret = GetLocaleInfo( LOCALE_STHOUSAND, mb, ThousandsSeparator )
		    
		    Const LOCALE_SYEARMONTH = &h1006
		    ret = GetLocaleInfo( LOCALE_SYEARMONTH, mb, YearMonthFormat )
		    
		    Const LOCALE_SMONTHOUSANDSEP = &h17
		    ret = GetLocaleInfo( LOCALE_SMONTHOUSANDSEP, mb, MonetaryThousandsSeparator )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CurrencyDisplayString(value as Double) As String
		  Soft Declare Sub GetCurrencyFormatW Lib "Kernel32" ( local as Integer, flags as Integer, num as WString, formatting as Integer, out as Ptr, size as Integer )
		  Soft Declare Sub GetCurrencyFormatA Lib "Kernel32" ( local as Integer, flags as Integer, num as CString, formatting as Integer, out as Ptr, size as Integer )
		  
		  if System.IsFunctionAvailable( "GetCurrencyFormatW", "Kernel32" ) then
		    dim mb as new MemoryBlock( 1024 )
		    GetCurrencyFormatW( LCID, 0, Str( value ), 0, mb, mb.Size )
		    
		    return mb.WString( 0 )
		  else
		    dim mb as new MemoryBlock( 1024 )
		    GetCurrencyFormatA( LCID, 0, Str( value ), 0, mb, mb.Size )
		    
		    return mb.CString( 0 )
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CurrencyDisplayStringByHand(currencyValue as Double) As String
		  '// We're given a monetary value that we want to display to the user
		  '// in the proper format.
		  '
		  'dim form as String
		  'if currencyValue >= 0 then
		  '// We're going to use the positive currency format
		  'form = PositiveCurrencyDisplayFormat
		  'else
		  '// We have a negative currency, so use that format instead
		  'form = NegativeCurrencyDisplayFormat
		  'end if
		  '
		  '// We use the Str function because we don't want it to be
		  '// i8n savvy yet.  We'll take care of making it savvy ourselves
		  'dim dataStr as String = Str( currencyValue )
		  '
		  '// Split our string into two parts -- prior to the . and after the .
		  'dim whole, partial as String
		  'whole = NthField( dataStr, ".", 1 )
		  'partial = NthField( dataStr, ".", 2 )
		  '
		  '// Now we want to truncate the partial down to the proper
		  '// number of digits.  Or, it could grow it up to the proper number
		  '// of digits as well.
		  'partial = Left( partial, NumCurrencyDigits )
		  'if Len( partial ) < NumCurrencyDigits then partial = partial + Mid( "0000000000", 1, NumCurrencyDigits - Len( partial ) )
		  '
		  '// Now, based on our form string, let's do some replacements.
		  '// First, replace "X" with the whole value
		  'form = Replace( form, "X", Replace( whole, "-", "" ) )
		  '
		  '// Next, replace the . with the separator value, but
		  '// only if we've got a partial to deal with
		  'if Len( partial ) > 0 then
		  'form = Replace( form, ".", MonetaryDecimalSeparator )
		  'else
		  'form = Replace( form, ".", "" )
		  'end if
		  '
		  '// Then replace "Y" with the partial value
		  'form = Replace( form, "Y", partial )
		  '
		  '// And then, do the currency symbol
		  'form = Replace( form, "$", CurrencySymbol )
		  '
		  '// If we have a negative symbol, then we need to replace
		  '// that with the locale's negative symbol.
		  'if currencyValue < 0 then
		  'form = Replace( form, "-", NegativeSymbol )
		  'end if
		  '
		  'return form
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DayName(i as Integer, abbr as Boolean = false) As String
		  // Days are one-based, but the array is 0-based
		  i = i - 1
		  
		  if abbr then
		    return AbbreviatedDayNames( i )
		  else
		    return DayNames( i )
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		 Shared Function DefaultLocale() As LocaleInformation
		  static s as new LocaleInformation
		  return s
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		 Shared Function GetAllLocales() As LocaleInformation()
		  Soft Declare Sub EnumSystemLocalesW Lib "Kernel32" ( proc as Ptr, flags as Integer )
		  Soft Declare Sub EnumSystemLocalesA Lib "Kernel32" ( proc as Ptr, flags as Integer )
		  
		  // Clear out all our information
		  Redim sLocales( -1 )
		  
		  // Enumerate over the system information
		  // Due to a bug with RB2006r3 and earlier, we can't use the WString version
		  // of the callback.  So we will fall back on the CString version, which should
		  // work just as well given the fact that the string passed to the callback only
		  // contains numeric information.
		  Const LCID_SUPPORTED = &h2
		  if RBVersion > 2006.3 and System.IsFunctionAvailable( "EnumSystemLocalesW", "Kernel32" ) then
		    EnumSystemLocalesW( AddressOf LocaleInfoProcW, LCID_SUPPORTED )
		  else
		    EnumSystemLocalesA( AddressOf LocaleInfoProcA, LCID_SUPPORTED )
		  end if
		  
		  // Return the data we found
		  return sLocales
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetLocaleInfo(type as Integer, mb as MemoryBlock, ByRef retVal as String) As Integer
		  Soft Declare Function GetLocaleInfoA Lib "kernel32" (Locale As integer, LCType As integer, lpLCData As ptr, cchData As integer) As Integer
		  Soft Declare Function GetLocaleInfoW Lib "kernel32" (Locale As integer, LCType As integer, lpLCData As ptr, cchData As integer) As Integer
		  
		  dim returnValue as Integer
		  dim size as Integer
		  
		  if mb <> nil then size = mb.Size
		  
		  if System.IsFunctionAvailable( "GetLocaleInfoW", "Kernel32" ) then
		    if mb <> nil then
		      returnValue = GetLocaleInfoW( LCID, type, mb, size ) * 2
		      retVal = ReplaceAll( DefineEncoding( mb.StringValue( 0, returnValue ), Encodings.UTF16 ), Chr( 0 ), "" )
		    else
		      returnValue = GetLocaleInfoW( LCID, type, nil, size ) * 2
		    end if
		  else
		    if mb <> nil then
		      returnValue = GetLocaleInfoA( LCID, type, mb, size ) * 2
		      retVal = ReplaceAll( DefineEncoding( mb.StringValue( 0, returnValue ), Encodings.ASCII ), Chr( 0 ), "" )
		    else
		      returnValue = GetLocaleInfoA( LCID, type, nil, size ) * 2
		    end if
		  end if
		  
		  return returnValue
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Shared Function LocaleInfoProcA(idStr as CString) As Boolean
		  #pragma X86CallingConvention StdCall
		  
		  // Append the data to our shared property
		  dim data as String = idStr
		  
		  return CommonLocaleInfoProc( data )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Shared Function LocaleInfoProcW(idStr as WString) As Boolean
		  #pragma X86CallingConvention StdCall
		  
		  // Append the data to our shared property
		  dim data as String
		  
		  #if RBVersion >= 2006.3
		    data = idStr
		  #endif
		  
		  return CommonLocaleInfoProc( data )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function MonthName(i as Integer, abbr as Boolean = false) As String
		  // Months are one-based, but the array is 0-based
		  i = i - 1
		  
		  if abbr then
		    return AbbreviatedMonthNames( i )
		  else
		    return MonthNames( i )
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function NumberDisplayString(value as Double) As String
		  Soft Declare Sub GetNumberFormatW Lib "Kernel32" ( local as Integer, flags as Integer, num as WString, formatting as Integer, out as Ptr, size as Integer )
		  Soft Declare Sub GetNumberFormatA Lib "Kernel32" ( local as Integer, flags as Integer, num as CString, formatting as Integer, out as Ptr, size as Integer )
		  
		  if System.IsFunctionAvailable( "GetNumberFormatW", "Kernel32" ) then
		    dim mb as new MemoryBlock( 1024 )
		    GetNumberFormatW( LCID, 0, Str( value ), 0, mb, mb.Size )
		    
		    return mb.WString( 0 )
		  else
		    dim mb as new MemoryBlock( 1024 )
		    GetNumberFormatA( LCID, 0, Str( value ), 0, mb, mb.Size )
		    
		    return mb.CString( 0 )
		  end if
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		AbbreviatedCountry As String
	#tag EndProperty

	#tag Property, Flags = &h0
		AbbreviatedDayNames() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		AbbreviatedLanguage As String
	#tag EndProperty

	#tag Property, Flags = &h0
		AbbreviatedMonthNames() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		AM As String
	#tag EndProperty

	#tag Property, Flags = &h0
		CalendarType As Integer
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  select case CalendarType
			  case kCalendarTypeTangunEra
			    return "Tangun Era (Korea)"
			    
			  case kCalendarTypeArabicGregorian
			    return "Gregorian Arabic calendar"
			    
			  case kCalendarTypeEnglishGregorian
			    return "Gregorian (English strings always)"
			    
			  case kCalendarTypeHebrew
			    return "Hebrew (Lunar)"
			    
			  case kCalendarTypeHijri
			    return "Hijri (Arabic lunar)"
			    
			  case kCalendarTypeLocalizedGregorian
			    return "Gregorian (localized)"
			    
			  case kCalendarTypeMiddleEastFrenchGregorian
			    return "Gregorian Middle East French calendar"
			    
			  case kCalendarTypeTaiwan
			    return "Taiwan Calendar"
			    
			  case kCalendarTypeThai
			    return "Thai"
			    
			  case kCalendarTypeTransliteratedEnglishGregorian
			    return "Gregorian Transliterated English calendar"
			    
			  case kCalendarTypeTransliteratedFrenchGregorian
			    return "Gregorian Transliterated French calendar"
			    
			  case kCalendarTypeYearOfTheEmporer
			    return "Year of the Emperor (Japan)"
			    
			  end select
			End Get
		#tag EndGetter
		CalendarTypeString As String
	#tag EndComputedProperty

	#tag Property, Flags = &h0
		CountryCode As String
	#tag EndProperty

	#tag Property, Flags = &h0
		CurrencySymbol As String
	#tag EndProperty

	#tag Property, Flags = &h0
		DateSeparator As String
	#tag EndProperty

	#tag Property, Flags = &h0
		DayNames() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		DecimalSeparator As String
	#tag EndProperty

	#tag Property, Flags = &h0
		DefaultCountryCode As String
	#tag EndProperty

	#tag Property, Flags = &h0
		DefaultLanguage As String
	#tag EndProperty

	#tag Property, Flags = &h0
		EnglishCountryName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		EnglishCurrencyName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		EnglishLanguageName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		FirstDayOfWeek As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		FirstWeekOfYear As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		InternationalMonetarySymbol As String
	#tag EndProperty

	#tag Property, Flags = &h0
		LCID As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ListSymbol As String
	#tag EndProperty

	#tag Property, Flags = &h0
		LongDateFormat As String
	#tag EndProperty

	#tag Property, Flags = &h0
		MonetaryDecimalSeparator As String
	#tag EndProperty

	#tag Property, Flags = &h0
		MonetaryThousandsSeparator As String
	#tag EndProperty

	#tag Property, Flags = &h0
		MonthNames() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		NativeCountryName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		NativeCurrencyName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		NativeDigits() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		NativeLanguageName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		NegativeCurrencyDisplayFormat As String
	#tag EndProperty

	#tag Property, Flags = &h0
		NegativeSymbol As String
	#tag EndProperty

	#tag Property, Flags = &h0
		NumberOfCurrencyDigits As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		NumberOfFractionalDigits As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		NumberOfInternationalCurrencyDigits As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		PM As String
	#tag EndProperty

	#tag Property, Flags = &h0
		PositiveCurrencyDisplayFormat As String
	#tag EndProperty

	#tag Property, Flags = &h0
		PositiveSymbol As String
	#tag EndProperty

	#tag Property, Flags = &h0
		ShortDateFormat As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private Shared sLocales(-1) As LocaleInformation
	#tag EndProperty

	#tag Property, Flags = &h0
		ThousandsSeparator As String
	#tag EndProperty

	#tag Property, Flags = &h0
		TimeMarkerIsPrefix As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		TimeSeparator As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Use24HourClock As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		UseFourDigitCentury As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		UseLeadingZeros As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		UseLeadingZerosForDay As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		UseLeadingZerosForHour As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		UseLeadingZerosForMonth As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		UseMetricMeasurements As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		YearMonthFormat As String
	#tag EndProperty


	#tag Constant, Name = kCalendarTypeArabicGregorian, Type = Double, Dynamic = False, Default = \"10", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCalendarTypeEnglishGregorian, Type = Double, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCalendarTypeHebrew, Type = Double, Dynamic = False, Default = \"8", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCalendarTypeHijri, Type = Double, Dynamic = False, Default = \"6", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCalendarTypeLocalizedGregorian, Type = Double, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCalendarTypeMiddleEastFrenchGregorian, Type = Double, Dynamic = False, Default = \"9", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCalendarTypeTaiwan, Type = Double, Dynamic = False, Default = \"4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCalendarTypeTangunEra, Type = Double, Dynamic = False, Default = \"5", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCalendarTypeThai, Type = Double, Dynamic = False, Default = \"7", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCalendarTypeTransliteratedEnglishGregorian, Type = Double, Dynamic = False, Default = \"11", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCalendarTypeTransliteratedFrenchGregorian, Type = Double, Dynamic = False, Default = \"12", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCalendarTypeYearOfTheEmporer, Type = Double, Dynamic = False, Default = \"3", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kFirstWeekContaining11, Type = Double, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kFirstWeekFollowing11, Type = Double, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kFirstWeekWithFourDays, Type = Double, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="AbbreviatedCountry"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="AbbreviatedLanguage"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="AM"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CalendarType"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CalendarTypeString"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CountryCode"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CurrencySymbol"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DateSeparator"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DecimalSeparator"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DefaultCountryCode"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DefaultLanguage"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="EnglishCountryName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="EnglishCurrencyName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="EnglishLanguageName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="FirstDayOfWeek"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="FirstWeekOfYear"
			Group="Behavior"
			InitialValue="0"
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
			Name="InternationalMonetarySymbol"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LCID"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ListSymbol"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LongDateFormat"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="MonetaryDecimalSeparator"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="MonetaryThousandsSeparator"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="NativeCountryName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="NativeCurrencyName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="NativeLanguageName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="NegativeCurrencyDisplayFormat"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="NegativeSymbol"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="NumberOfCurrencyDigits"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="NumberOfFractionalDigits"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="NumberOfInternationalCurrencyDigits"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="PM"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="PositiveCurrencyDisplayFormat"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="PositiveSymbol"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ShortDateFormat"
			Group="Behavior"
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
			Name="ThousandsSeparator"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TimeMarkerIsPrefix"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TimeSeparator"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Use24HourClock"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="UseFourDigitCentury"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="UseLeadingZeros"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="UseLeadingZerosForDay"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="UseLeadingZerosForHour"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="UseLeadingZerosForMonth"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="UseMetricMeasurements"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="YearMonthFormat"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
