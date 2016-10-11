'********************************************************************
'*
'* VBScript for enumerating computer, group and user objects from Active Directory
'* with attributes. The results saved to a textfile.
'*
'* Parameters:	TBC
'*
'*
'********************************************************************
Option Explicit

'********************************************************************
'* Declare global constants and variables - Scripts default parameters
'********************************************************************
Const defaultAgeAccountSuspend = "60"		'To be confirmed
Const defaultAgeAccountRemoval = "90"		'To be confirmed

'Const defaultType = "1"		'Computer
'Const defaultType = "2"		'Group
Const defaultType = "3"			'User

'Const defaultScope = "1"		'Target OU only
Const defaultScope = "2"		'Target OU and all child OUs

Dim delimiter:	delimiter = chr(9)	' Tab
'Dim delimiter:	delimiter = chr(44)	' Comma
'********************************************************************


'********************************************************************
'* Returns the FQDN of the currently logged in user !!!! NEED TO TEST AGAINST LOCAL ACCOUNT
'********************************************************************
Dim objSysInfo
Set objSysInfo = Createobject("ADSystemInfo")
Dim strDefaultOU
strDefaultOU = split(objSysInfo.UserName,",",2)(1)

'Wscript.Echo "The variable xxx is a : " & VarType(xxx)


'********************************************************************
'* Declare Global Constants
'********************************************************************
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Const ADS_OBJ_TYPE_COMPUTER = "computer"
Const ADS_OBJ_TYPE_GROUP = "group"
Const ADS_OBJ_TYPE_USER = "user"

'*************** User Account Control Flags ****************
Const ADS_UF_SCRIPT = &H1
Const ADS_UF_ACCOUNTDISABLE = &H2
Const ADS_UF_HOMEDIR_REQUIRED = &H8
Const ADS_UF_LOCKOUT = &H10
Const ADS_UF_PASSWD_NOTREQD = &H20
Const ADS_UF_PASSWD_CANT_CHANGE = &H40
Const ADS_UF_ENCRYPTED_TEXT_PWD_ALLOWED = &H80
Const ADS_UF_TEMP_DUPLICATE_ACCOUNT = &H100
Const ADS_UF_NORMAL_ACCOUNT = &H200
Const ADS_UF_INTERDOMAIN_TRUST_ACCOUNT = &H800
Const ADS_UF_WORKSTATION_TRUST_ACCOUNT = &H1000
Const ADS_UF_SERVER_TRUST_ACCOUNT = &H2000
Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000
Const ADS_UF_MNS_LOGON_ACCOUNT = &H20000
Const ADS_UF_SMARTCARD_REQUIRED = &H40000
Const ADS_UF_TRUSTED_FOR_DELEGATION = &H80000
Const ADS_UF_NOT_DELEGATED = &H100000
Const ADS_UF_USE_DES_KEY_ONLY = &H200000
Const ADS_UF_DONT_REQ_PREAUTH = &H400000
Const ADS_UF_PASSWORD_EXPIRED = &H800000
Const ADS_UF_TRUSTED_TO_AUTH_FOR_DELEGATION = &H1000000
Const ADS_UF_PARTIAL_SECRETS_ACCOUNT = &H4000000

'******************* Group Type / Scope ********************
Const ADS_UF_SYSTEM = &H1
Const ADS_UF_SCOPE_GLOBAL = &H2
Const ADS_UF_SCOPE_LOCAL = &H4
Const ADS_UF_SCOPE_UNIVERSAL = &H8
Const ADS_UF_SECURITY_GROUP = &H80000000

'********************************************************************
'* Declare Global Variables - WMI Object Management
'********************************************************************
Dim strDN
Dim strScope
Dim strType
Dim oRootDSE

Dim objLogFile

'Dim dicUserAccountControls
'Set dicUserAccountControls = CreateObject("Scripting.Dictionary")
'dicUserAccountControls.Add "ADS_UF_SCRIPT", &H1
'dicUserAccountControls.Add "ADS_UF_ACCOUNTDISABLE", &H2
'dicUserAccountControls.Add "ADS_UF_HOMEDIR_REQUIRED", &H8
'dicUserAccountControls.Add "ADS_UF_LOCKOUT", &H10
'dicUserAccountControls.Add "ADS_UF_PASSWD_NOTREQD", &H20
'dicUserAccountControls.Add "ADS_UF_PASSWD_CANT_CHANGE", &H40
'dicUserAccountControls.Add "ADS_UF_ENCRYPTED_TEXT_PWD_ALLOWED", &H80
'dicUserAccountControls.Add "ADS_UF_TEMP_DUPLICATE_ACCOUNT", &H100
'dicUserAccountControls.Add "ADS_UF_NORMAL_ACCOUNT", &H200
'dicUserAccountControls.Add "ADS_UF_INTERDOMAIN_TRUST_ACCOUNT", &H800





'********************************************************************
'* TBC
dim listAttributes
Set listAttributes = CreateObject("System.Collections.ArrayList")

'********* Attribute of: Computer, Group, and User *********
listAttributes.Add "distinguishedName"
listAttributes.Add "ou"
listAttributes.Add "cn"
listAttributes.Add "whenCreated"
listAttributes.Add "whenChanged"
listAttributes.Add "managedBy"
listAttributes.Add "description"

'************* Attribute of: Computer and User *************
listAttributes.Add "displayName"

'************** Attribute of: Group and User ***************
listAttributes.Add "mail"

'************** Attribute of: User ***************
listAttributes.Add "UserPrincipalName"
listAttributes.Add "givenName"
listAttributes.Add "sn"
listAttributes.Add "department"
listAttributes.Add "telephoneNumber"
listAttributes.Add "mobile"
listAttributes.Add "streetAddress"						'Street address
listAttributes.Add "l"									'Location
listAttributes.Add "st"									'State
listAttributes.Add "co"									'Country
listAttributes.Add "postalCode"							'Postal code
listAttributes.Add "homeMDB"							'Exchange Mailstore
listAttributes.Add "mDBUseDefaults"						'Exchange Default settings enforced
listAttributes.Add "mDBStorageQuota"					'Exchange Storage Quota
listAttributes.Add "mDBOverQuotaLimit"					'Exchange Storage Quota Limit
listAttributes.Add "msExchHideFromAddressLists"			'Exchange Address Hide From GAL
listAttributes.Add "msRTCSIP-FederationEnabled"			'Microsoft Communicator Federation enabled
listAttributes.Add "msRTCSIP-InternetAccessEnabled"		'Microsoft Communicator Internet access enabled
listAttributes.Add "manager"
listAttributes.Add "userWorkstations"
listAttributes.Add "msExchHomeServerName"
listAttributes.Add "msExchMobileMailboxFlags"
listAttributes.Add "msExchRecipientDisplayType"
listAttributes.Add "msExchRecipientTypeDetails"
listAttributes.Add "legacyExchangeDN"
listAttributes.Add "HomeMTA"


'dim listGroupAttributes
'Set listGroupAttributes = CreateObject("System.Collections.ArrayList")
'listGroupAttributes.Add "GroupType"
'listGroupAttributes.Add "GroupTypeValue"
'listGroupAttributes.Add "CreatedBySystem"
'listGroupAttributes.Add "GroupScope"



'********************************************************************
'* Function: TBC   !!!!!!!!!!!!!!!!
'* Purpose: 
'* Input:	
'* Output:  
'* Notes:  
'*
Sub writeLog(objParent, strDescription)
	Dim objAttributeValue
	Dim tempString

	On Error Resume Next
	objAttributeValue = objParent.get(strDescription)

	If Err.Number = 0 Then
		tempString = Replace(objAttributeValue, chr(10), "")
		tempString = Replace(tempString , chr(13), "")
		objLogFile.Write chr(34) & tempString & chr(34) & delimiter
		
	Else
		objLogFile.Write delimiter
		Err.Clear
	End if
	On Error Goto 0
End Sub




'********************************************************************
'* Function: ExtractCommon_OpenLDAP(strDN, strFilter)
'* Purpose: 
'* Input:   strDN		:	Distinguished Name as a string
'*			strFilter	:	Object type to be filtered on as a string (e.g. Computer, Group, or User)
'* Output:  Object
'* Notes:  
'*
Function ExtractCommon_OpenLDAP(strDN, strFilter)
	Set ExtractCommon_OpenLDAP = GetObject("LDAP://" & strDN)
	ExtractCommon_OpenLDAP.Filter = Array(strFilter)
End Function
'********************************************************************


'********************************************************************
'* Function: hexCompare(objAttribute, ObjComparator)
'* Purpose: 
'* Input:
'* Output:
'* Notes:  Try using Case to improve speed
'*
Function hexCompare(objAttribute, ObjComparator)
	If IsEmpty(objAttribute) = TRUE Then
		Set hexCompare = Nothing
	Else
		If objAttribute AND ObjComparator Then
			hexCompare = "TRUE"
		Else 
			hexCompare = "FALSE"
		End If			
	End if
End Function
'********************************************************************


'********************************************************************
'* Function:	ExtractCommon_IntegerToDate(ByVal intDateEpoch)
'* Purpose:		Function to convert Integer value to a date, not adjustment is made for local time zone bias
'* Input:		intDateEpoch	:	Integer which represents a date as seconds from epoch
'* Output:		Date
'* Notes:  
'*
Function ExtractCommon_IntegerToDate(ByVal intDateEpoch)
	ExtractCommon_IntegerToDate = CDate(intDateEpoch) + #1/1/1601#
End Function
'********************************************************************


'********************************************************************
'* Function:	ExtractCommon_Integer8ToInteger(ByVal objInteger8)
'* Purpose:	Function to convert Integer8 (64-bit) value from an object to an Integer
'* Input:   intDateEpoch	:	Integer which represents a date as seconds from epoch
'* Output:  Date
'* Notes:  
'*
Function ExtractCommon_Integer8ToInteger(ByVal objInteger8)
	Dim intInteger8
	Dim intHighPart:intHighPart = objInteger8.HighPart
	Dim intLowPart:intLowPart = objInteger8.LowPart
	
	If (intLowPart < 0) Then
		intHighPart = intHighPart + 1
	End If
	
	intInteger8 = intHighPart * (2^32) + intLowPart 
	intInteger8 = intInteger8 / (60 * 10000000)
	intInteger8 = intInteger8 / 1440
	ExtractCommon_Integer8ToInteger = intInteger8
End Function
'********************************************************************


'********************************************************************
'* Sub ExtractObject_ExportRecursive
'* Purpose:
'* Input:   strZoneOU
'* Output:  None
'* Notes:  
'*
Sub ExtractObject_ExportRecursive (ByVal strZoneOU)
	Dim objZoneOU
	Dim objZoneChildOU
	
	Set objZoneOU = ExtractCommon_OpenLDAP(strZoneOU, "organizationalUnit")
	ExtractObject_ExportSite(objZoneOU.distinguishedName)
	
	For Each objZoneChildOU In objZoneOU
		ExtractObject_ExportRecursive (objZoneChildOU.distinguishedName)
	Next
End Sub
'********************************************************************


'********************************************************************
'* Sub ExtractObject_ExportSite
'* Purpose: Connects to site OU via LDAP and extracts all objects of specified type.
'* The Sub then populates Excel with a set of attributes for each object found
'* Input:   strSiteOU
'* Output:  
'* Notes:  
'*
Sub ExtractObject_ExportSite (ByVal strSiteOU)
	Dim objSiteOU
	Dim objSiteChild

	Dim strAccountType
	
	Dim dblLastLogonTimestamp
	Dim dblLockoutTime
	Dim dblPwdLastSet
	Dim lngDateDiffCheck
	Dim dtmPwdLastSet

	Dim strMyElements
	Dim intLinesWritten
	
	intLinesWritten = 0
	
'	Dim intWhenCreated
'	Dim memberCounter
'	Dim memberOfCounter
'	Dim objMember
'	Dim objMemberOf
'	Dim colMembers
'	Dim sGroupMember
'	Dim colMembersOf
'	Dim sGroupMemberOf


	Set objSiteOU = ExtractCommon_OpenLDAP(strSiteOU, strType)
	For Each objSiteChild In objSiteOU
		intLinesWritten = intLinesWritten + 1
'		objLogFile.Write chr(34) & intLinesWritten & chr(34) & delimiter
		strAccountType = objSiteChild.class
		If strType <> strAccountType Then
			Exit For
		End If

		For Each strMyElements In listAttributes
			Call writeLog(objSiteChild, strMyElements)
		Next

		'********************************************************************
		'* Doing stuff with passwords

		If isEmpty(objSiteChild.pwdLastSet) = FALSE Then
			dblPwdLastSet = ExtractCommon_Integer8ToInteger(objSiteChild.pwdLastSet)
			dtmPwdLastSet = ExtractCommon_IntegerToDate(dblPwdLastSet)
			lngDateDiffCheck = DateDiff("d",dtmPwdLastSet,Now)


			
			If dblPwdLastSet = 0 Then
				objLogFile.Write chr(34) & "0" & chr(34) & delimiter
				objLogFile.Write chr(34) & "TRUE" & chr(34) & delimiter
			Else
				objLogFile.Write chr(34) & dtmPwdLastSet & chr(34) & delimiter
				objLogFile.Write chr(34) & "FALSE" & chr(34) & delimiter
			End if
		Else
			objLogFile.Write chr(34) & "isEmpty," & chr(34) & "isEmpty,"
		End if


		'********************************************************************
		'*lastLogonTimestamp
		If IsEmpty(objSiteChild.lastLogonTimestamp) = FALSE Then
			dblLastLogonTimestamp = ExtractCommon_Integer8ToInteger(objSiteChild.lastLogonTimestamp)
			If dblLastLogonTimestamp = 0 Then
				objLogFile.Write chr(34) & "0" & chr(34) & delimiter
			Else
				objLogFile.Write chr(34) & ExtractCommon_IntegerToDate(dblLastLogonTimestamp) & chr(34) & delimiter
			End if
		Else
			objLogFile.Write chr(34) & "isEmpty" & chr(34) & delimiter
		End if




		'********************************************************************
		'*lockoutTime
		If IsEmpty(objSiteChild.lockoutTime) = FALSE Then
			dblLockoutTime = ExtractCommon_Integer8ToInteger(objSiteChild.lockoutTime)
			If dblLockoutTime = 0 Then
				objLogFile.Write chr(34) & "0" & chr(34) & delimiter
			Else
				objLogFile.Write chr(34) & ExtractCommon_IntegerToDate(dblLockoutTime) & chr(34) & delimiter
			End if
		Else
			objLogFile.Write chr(34) & "isEmpty" & chr(34) & delimiter
		End if



		If isEmpty(objSiteChild.userAccountControl) = FALSE Then
			objLogFile.Write chr(34) & objSiteChild.UserAccountControl & chr(34) & delimiter
		
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_SCRIPT) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_ACCOUNTDISABLE) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_HOMEDIR_REQUIRED) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_PASSWD_NOTREQD) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_PASSWD_CANT_CHANGE) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_ENCRYPTED_TEXT_PWD_ALLOWED) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_TEMP_DUPLICATE_ACCOUNT) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_NORMAL_ACCOUNT) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_INTERDOMAIN_TRUST_ACCOUNT) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_WORKSTATION_TRUST_ACCOUNT) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_SERVER_TRUST_ACCOUNT) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_DONT_EXPIRE_PASSWD) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_MNS_LOGON_ACCOUNT) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_SMARTCARD_REQUIRED) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_TRUSTED_FOR_DELEGATION) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_NOT_DELEGATED) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_USE_DES_KEY_ONLY) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_DONT_REQ_PREAUTH) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_PASSWORD_EXPIRED) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_TRUSTED_TO_AUTH_FOR_DELEGATION) & chr(34) & delimiter
			objLogFile.Write chr(34) & hexCompare(objSiteChild.userAccountControl, ADS_UF_PARTIAL_SECRETS_ACCOUNT) & chr(34) & delimiter
			'********************************************************************
		End if


		objLogFile.Writeline

	Next
End Sub



'********************************************************************
'* Sub Main()
'* Purpose: Main component of Active Directory Extraction script.
'* Input:   
'* Output:  
'*
'********************************************************************
Sub Main()

'* Variables used to calculate execute time of script
Dim startTime:	startTime = Now
Dim stopTime
Dim elapsedTime


'********************************************************************
Dim objArgs
Set objArgs = WScript.Arguments


'********************************************************************
'* InputBox to select object type to be extracted from AD. Includes input validation
'*
If WScript.Arguments.Count = 3 Then
	strType = objArgs(0)
	strScope = objArgs(1)
	strDN = objArgs(2)

	Select Case strType
	Case "1"
		strType = ADS_OBJ_TYPE_COMPUTER
	Case "2"
		strType = ADS_OBJ_TYPE_GROUP
	Case "3"
		strType = ADS_OBJ_TYPE_USER
	End Select
	
Else
	strType = InputBox("Enter the type object to extract from Active Directory Computer[1], Group[2], or User[3] ","Input Object Type",defaultType)
		Select Case strType
		Case "1"
			strType = ADS_OBJ_TYPE_COMPUTER
		Case "2"
			strType = ADS_OBJ_TYPE_GROUP
		Case "3"
			strType = ADS_OBJ_TYPE_USER
		End Select

	'* InputBox to select mode of operation. Extract data from: Single OU or Recursively 
	strScope = InputBox("Enter the mode of operation. Single OU Only  [1] or additionally query child OU [2]","Input Scope",defaultScope)

	'* InputBox to select the DN of the target OU
	strDN = InputBox("Enter the distinguished name of a Site container","Input Site OU",strDefaultOU)

End If
'********************************************************************




'********************************************************************
'* Setup Filesystem access
Dim OutputFile
Dim objFSO 
Dim element


OutputFile = "AD_" & strType & "_" & Day(Now) & MonthName(Month(Now),True) & Year(Now) & ".csv"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.CreateTextFile(OutputFile , ForWriting, True)

For Each element In listAttributes
	objLogFile.Write chr(34) & element & chr(34) & delimiter
Next

objLogFile.Write chr(34) & "PasswordLastSet" & chr(34) & delimiter
objLogFile.Write chr(34) & "MustChangePasswordNextLogin" & chr(34) & delimiter
objLogFile.Write chr(34) & "LastLogonTimestampGMT" & chr(34) & delimiter
objLogFile.Write chr(34) & "LockoutTime" & chr(34) & delimiter
objLogFile.Write chr(34) & "UserAccountControl" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_SCRIPT" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_ACCOUNTDISABLE" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_HOMEDIR_REQUIRED" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_PASSWD_NOTREQD" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_PASSWD_CANT_CHANGE" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_ENCRYPTED_TEXT_PWD_ALLOWED" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_TEMP_DUPLICATE_ACCOUNT" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_NORMAL_ACCOUNT" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_INTERDOMAIN_TRUST_ACCOUNT" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_WORKSTATION_TRUST_ACCOUNT" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_SERVER_TRUST_ACCOUNT" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_DONT_EXPIRE_PASSWD" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_MNS_LOGON_ACCOUNT" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_SMARTCARD_REQUIRED" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_TRUSTED_FOR_DELEGATION" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_NOT_DELEGATED" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_USE_DES_KEY_ONLY" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_DONT_REQ_PREAUTH" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_PASSWORD_EXPIRED" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_TRUSTED_TO_AUTH_FOR_DELEGATION" & chr(34) & delimiter
objLogFile.Write chr(34) & "ADS_UF_PARTIAL_SECRETS_ACCOUNT" & chr(34) & delimiter
objLogFile.Writeline
'********************************************************************


'********************************************************************
'* Execute either Active Directory extract in either single or recursive mode
'*
Select Case strScope
	Case "1"
		ExtractObject_ExportSite(strDN)
	Case "2"
		ExtractObject_ExportRecursive(strDN)
	End Select
'********************************************************************


'********************************************************************
'* Setup variables to calculate time taken to execute script
'*
stopTime = Now
elapsedTime = DateDiff("s",startTime,stopTime)
'********************************************************************









'********************************************************************
objLogFile.Close
Wscript.Echo "Script Completed in : " & elapsedTime & " seconds"



End Sub

Main

'********************************************************************
'*                                                                  *
'*                           End of File                            *
'*                                                                  *
'********************************************************************








