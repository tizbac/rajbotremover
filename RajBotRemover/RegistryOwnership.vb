' RegistryOwnership v1.01
' v1.0 -- June 12, 2012
' v1.01 -- July 10, 2015 -- corrected HKLM access; bug report by 'Jonas'
' This work by Eric Siron is licensed under a Creative Commons Attribution-ShareAlike 3.0 Unported License.
' http://creativecommons.org/licenses/by-sa/3.0/
Imports Microsoft.Win32
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Security.AccessControl
Imports System.Security.Principal

Module RegistryOwnership
    Private Const WIN_NT_ANYSIZE_ARRAY As Integer = 1
    Private Const WIN32_NO_ERROR = 0
    Private Const SE_PRIVILEGE_ENABLED As Integer = &H2 ' WinNT.h line 7660; there are other possibilities, but none useful in this context

    ' most of the Win32 objects are directly from WinNT.h or AccCtrl.h
#Region "Win32 Enums"

    ''' <remarks></remarks>
    Private Enum SE_OBJECT_TYPE ' the only type used in this assembly is SE_REGISTRY_KEY; list included for completeness and expandability
        SE_UNKNOWN_OBJECT_TYPE = 0
        SE_FILE_OBJECT
        SE_SERVICE
        SE_PRINTER
        SE_REGISTRY_KEY
        SE_LMSHARE
        SE_KERNEL_OBJECT
        SE_WINDOW_OBJECT
        SE_DS_OBJECT
        SE_DS_OBJECT_ALL
        SE_PROVIDER_DEFINED_OBJECT
        SE_WMIGUID_OBJECT
        SE_REGISTRY_WOW64_32KEY
    End Enum


    ''' <remarks></remarks>
    Private Enum SECURITY_INFORMATION As UInteger
        OWNER_SECURITY_INFORMATION = &H1L
        GROUP_SECURITY_INFORMATION = &H2L
        DACL_SECURITY_INFORMATION = &H4L
        SACL_SECURITY_INFORMATION = &H8L
        LABEL_SECURITY_INFORMATION = &H10L
        PROTECTED_DACL_SECURITY_INFORMATION = &H80000000L
        PROTECTED_SACL_SECURITY_INFORMATION = &H40000000L
        UNPROTECTED_DACL_SECURITY_INFORMATION = &H20000000L
        UNPROTECTED_SACL_SECURITY_INFORMATION = &H10000000L
    End Enum
#End Region

#Region "Win32 Structures"

    ''' <remarks></remarks>
    <StructLayout(LayoutKind.Sequential)> _
    Private Structure LUID
        Dim LowPart As UInt32
        Dim HighPart As UInt32
    End Structure


    ''' <remarks></remarks>
    <StructLayout(LayoutKind.Sequential)> _
    Private Structure LUID_AND_ATTRIBUTES
        Dim Luid As LUID
        Dim Attributes As UInt32
    End Structure

    ''' <remarks>Adjust SizeConst for uses beyond original design</remarks>
    <StructLayout(LayoutKind.Sequential)> _
    Private Structure TOKEN_PRIVILEGES
        Dim PrivilegeCount As UInt32
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=WIN_NT_ANYSIZE_ARRAY)> Dim Privileges() As LUID_AND_ATTRIBUTES
    End Structure
#End Region
#Region "Win32 Functions"



    ''' <param name="lpSystemName">The system to find the privilege for. Use Nothing for the local computer.</param>
    ''' <param name="lpName">The name of the privilege to lookup</param>
    ''' <param name="lpLuid">OUT: LUID structure</param>
    ''' <returns>True for success; false for failure. Use Marshal.GetLastWin32Error to determine error code.</returns>
    ''' <remarks></remarks>
    <DllImport("advapi32.dll", SetLastError:=True)> _
    Private Function LookupPrivilegeValue(ByVal lpSystemName As String, ByVal lpName As String, ByRef lpLuid As LUID) As Boolean
    End Function



    ''' <param name="TokenHandle">Handle to the user token</param>
    ''' <param name="DisableAllPrivileges">True to disable all privileges.</param>
    ''' <param name="NewState">TOKEN_PRIVILEGES structure with desired privilege set.</param>
    ''' <param name="BufferLength">Size of the structure that will be passed in the PreviousState parameter.</param>
    ''' <param name="PreviousState">TOKEN_PRIVILEGES structure to hold the user's privileges as they were prior to
    ''' this function call.</param>
    ''' <param name="ReturnLength">OUT: The size of the PreviousState variable after the function runs, or in the event
    ''' of an error, the minimum size that the PreviousState variable should have been.</param>
    ''' <returns>0 for success, a non-zero error code that can initialize a new Win32Exception for errorchecking.</returns>
    ''' <remarks></remarks>
    <DllImport("advapi32.dll", SetLastError:=True)> _
    Private Function AdjustTokenPrivileges(ByVal TokenHandle As IntPtr, ByVal DisableAllPrivileges As Boolean, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Integer, <Out()> ByRef PreviousState As TOKEN_PRIVILEGES, <Out()> ByRef ReturnLength As IntPtr) As Boolean
    End Function



    ''' <param name="pObjectName">Name of the object.</param>
    ''' <param name="ObjectType">Type of the object</param>
    ''' <param name="SecurityInfo">Flag indicating which security item(s) will be modified.</param>
    ''' <param name="psidOwner">SID of the new owner, if ownership is changing.</param>
    ''' <param name="psidGroup">SID of the object's primary group, if it is changing.</param>
    ''' <param name="pDacl">Discretionary access list to be set on the object, if it is changing.</param>
    ''' <param name="pSacl">Security access list to be set on the object, if it is changing.</param>
    ''' <returns>0 for success, a non-zero error code that can initialize a new Win32Exception for errorchecking.</returns>
    ''' <remarks></remarks>
    Private Declare Auto Function SetNamedSecurityInfo Lib "advapi32.dll" (ByVal pObjectName As String, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As SECURITY_INFORMATION, ByVal psidOwner As Byte(), ByVal psidGroup As IntPtr, ByVal pDacl As IntPtr, ByVal pSacl As IntPtr) As Integer
#End Region


    ''' <remarks>.Net functions typically use the HKEY_ prefix while the SetNamedSecurityInfo used here does not, so this
    ''' can help avoid confusion.
    ''' </remarks>
    Friend Enum BaseKeySelector
        ClassesRoot
        CurrentConfig
        CurrentUser
        LocalMachine
        Users
        ' allow for DynData as well?
    End Enum



    ''' <param name="RegistryBaseKeyName">An enumerator that indicates which registry base key to work with.</param>
    ''' <param name="RegistrySubKeyName">Subkey to use. Format example: "Software\MyApplication\MyKey"</param>
    ''' <param name="ErrorMessage">This string will be populated with any error conditions that the function traps.</param>
    ''' <returns>True if the function succeeed. If False, check the value of ErrorMessage.</returns>
    ''' <remarks></remarks>
    Friend Function SeizeRegistryKey(ByVal RegistryBaseKeyName As BaseKeySelector, ByRef RegistrySubKeyName As String, ByRef ErrorMessage As String) As Boolean
        '' Variable Declarations ''
        Dim Win32ReturnValue As Integer = 0 ' Windows APIs return numeric values from most functions.
        Dim Win32ErrorCode As Integer = 0 ' Used to hold the last recorded Win32 error code
        Dim CurrentUserToken As IntPtr = IntPtr.Zero ' Security token of the user this process is running as
        Dim CurrentUserIdentity As SecurityIdentifier ' Identity of "" ""
        Dim CurrentUserSDDL As Byte() = Nothing ' SDDL of "" ""
        Dim CurrentUsername As String = "" ' Full name of current user in DOMAIN\username format
        Dim TakeOwnershipLUIDandAttr As LUID_AND_ATTRIBUTES ' container for the SeTakeOwnership LUID and its attributes
        Dim NewState As TOKEN_PRIVILEGES ' Desired privilege set for the token
        Dim PreviousState As TOKEN_PRIVILEGES ' Previous privilege set for the token
        Dim ReturnLength As IntPtr = IntPtr.Zero ' Some Windows APIs set a number indicated how much data they returned, or tried to return
        Dim FullKeyName As String ' Combination of the registry base key and the subkey
        Dim BaseKey As RegistryKey ' The Microsoft.Win32.Registry representation of the base key
        Dim Subkey As RegistryKey ' The Microsoft.Win32.Registry representation of the key to be manipulated
        Dim KeyAccessSettings As RegistrySecurity ' Encapsulation of security settings on the registry key
        Dim KeyRuleList As AuthorizationRuleCollection ' Encapsulation of authorization rules on the registry key (essentially ACEs)
        Dim KeyFullControlRule As RegistryAccessRule ' Registry access rule that will give the current user full control

        '' Initializations ''
        ErrorMessage = "No error"
        Subkey = Nothing

        '' Step 1
        '' To see if the current user has the SeTakeOwnershipPrivilege, it is first necessary to determine
        '' what this computer calls that privilege (indicated by its LUID).
        TakeOwnershipLUIDandAttr = New LUID_AND_ATTRIBUTES
        If Not LookupPrivilegeValue(Nothing, "SeTakeOwnershipPrivilege", TakeOwnershipLUIDandAttr.Luid) Then
            ErrorMessage = String.Format("Cannot determine identifier for SeTakeOwnershipPrivilege: {0}", (New Win32Exception(Marshal.GetLastWin32Error)))
            Return False
        End If

        '' Step 2
        '' With the LUID of SeTakeOwnershipPrivilege in hand, the next step is to enable the privilege
        ' for the user within this process.
        ' Get the user's token first. Tokens taken this way do not need to be manually released. Note that GetCurrent()
        ' takes binary flags, so the "Or" is combining the two indicated access levels.
        CurrentUserToken = WindowsIdentity.GetCurrent(TokenAccessLevels.AdjustPrivileges Or TokenAccessLevels.Query).Token
        NewState = New TOKEN_PRIVILEGES ' Create an empty TOKEN_PRIVILEGES
        NewState.PrivilegeCount = 1 ' This will always be 1 in this usage, but structures can't have initializers without using Shared, which may have unexpected side effects in this context
        NewState.Privileges = New LUID_AND_ATTRIBUTES() {TakeOwnershipLUIDandAttr} ' Create the privilege set directly from the retrieved LUID; the only thing of meaning in here is SeTakeOwnershipPrivilege's LUID
        NewState.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED ' this indicates to AdjustTokenPrivileges what is to change
        PreviousState = New TOKEN_PRIVILEGES ' Will hold the privilege state of the token prior to modifications

        ' Documentation for the Win32 API indicates that passing a zero for BufferLength allows you to send in a
        ' NULL (Nothing in VB) for PreviousState and ReturnLength. However, attempting this always causes this function
        ' to return an error that insufficient space was provided for PreviousState. So, even though this assembly
        ' completely ignores the value of PreviousState, it must be captured.
        If Not AdjustTokenPrivileges(CurrentUserToken, False, NewState, Marshal.SizeOf(PreviousState), PreviousState, ReturnLength) Then
            Win32ErrorCode = Marshal.GetLastWin32Error
            If Win32ErrorCode = 122 Then ' PreviousState variable isn't large enough for the amount of data that AdjustTokenPrivileges is returning
                ErrorMessage = String.Format("The ""PreviousState"" variable passed to AdjustTokenPrivileges is not large enough. Its size was {0}. The required size was {1}", Marshal.SizeOf(PreviousState), ReturnLength)
                ' TODO: Else/Select Case: Set up traps for common returns, like security problems
            Else
                ErrorMessage = String.Format("An error occurred while attempting to adjust privileges for {0}: {1}", WindowsIdentity.GetCurrent.User, (New Win32Exception(Win32ErrorCode).Message))
            End If
            Return False
        End If

        '' Step 3
        '' The user token is now in a state where it can take ownership of objects in the system. Seize the registry key.
        ' Convert the passed-in registry key parts to a single string and get the base key
        Select Case RegistryBaseKeyName ' TODO: add an entry for DynData? edge-case usage probably not worth the effort
            Case BaseKeySelector.ClassesRoot
                FullKeyName = "CLASSES_ROOT"
                BaseKey = Registry.ClassesRoot
            Case BaseKeySelector.CurrentConfig
                FullKeyName = "CURRENT_CONFIG"
                BaseKey = Registry.CurrentConfig
            Case BaseKeySelector.CurrentUser
                FullKeyName = "CURRENT_USER"
                BaseKey = Registry.CurrentUser
            Case BaseKeySelector.LocalMachine
                FullKeyName = "MACHINE"
                BaseKey = Registry.LocalMachine
            Case BaseKeySelector.Users
                FullKeyName = "USERS"
                BaseKey = Registry.Users
            Case Else
                ErrorMessage = "Invalid registry base key selected"
                Return False
        End Select

        ' TODO: the "RegistrySubKeyName" variable is the most fragile part of this function; consider adding error-checking,
        ' convert signature to ByVal (makes a copy, potentially wasteful of memory) for string manipulation operations
        FullKeyName &= "\" & RegistrySubKeyName
        ' need the binary form of the current user's SID for SetNamedSecurityInfo(S-1-5-XX-XXXXXXXXXX...)
        CurrentUserIdentity = New SecurityIdentifier(WindowsIdentity.GetCurrent.User.Value) ' get the identity object first
        ReDim CurrentUserSDDL(CurrentUserIdentity.BinaryLength) ' prepare the SDDL binary array to hold it
        CurrentUserIdentity.GetBinaryForm(CurrentUserSDDL, 0) ' retrieve the binary SDDL and place it in the array

        ' take ownership of the registry key
        Win32ReturnValue = SetNamedSecurityInfo(FullKeyName, SE_OBJECT_TYPE.SE_REGISTRY_KEY,
        SECURITY_INFORMATION.OWNER_SECURITY_INFORMATION, CurrentUserSDDL, Nothing, Nothing, Nothing)
        If Win32ReturnValue <> WIN32_NO_ERROR Then
            ErrorMessage = String.Format("Error taking ownership of {0}: {1}", FullKeyName, (New Win32Exception(Win32ReturnValue).Message))
            Return False
        End If

        '' Step 4
        '' Having ownership is great, but all that does on its own is allow the current user to change permissions on the object.
        '' Without an explicitly granted permission, the current user will still be unable to change any values contained in
        '' the key. However, it is now possible to get a handle to the key to manipulate permissions.
        Try
            Subkey = BaseKey.OpenSubKey(RegistrySubKeyName, RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.ChangePermissions) ' get a handle to the key, explicitly indicate the need for ChangePermisssions rights
            If Subkey Is Nothing Then
                ' included for completeness as this should never happen; the only possibility is that another process deleted the key
                Throw New Exception("The specified registry key could not be found.")
            End If
            KeyAccessSettings = Subkey.GetAccessControl ' copy the current security settings into KeyAccessSettings
            KeyAccessSettings.SetAccessRuleProtection(True, False) ' True in the first parameter means that in case of an inheritance conflict, inheritance loses (except in the case of a deny). False in the second parameter means to remove inherited rules. This is the only combination that will strip an inherited Deny.
            KeyRuleList = KeyAccessSettings.GetAccessRules(True, True, GetType(SecurityIdentifier)) ' this is effectively an abstraction of the ACL on the key

            ' clean all rules from the list
            For Each RuleEntry As RegistryAccessRule In KeyRuleList
                KeyAccessSettings.RemoveAccessRuleSpecific(RuleEntry)
            Next RuleEntry
            KeyFullControlRule = New RegistryAccessRule(CurrentUserIdentity, RegistryRights.FullControl, InheritanceFlags.ContainerInherit, PropagationFlags.None, AccessControlType.Allow) ' create a new rule that sets the current user's permissions as Full Control: Allow with downward inheritance enabled
            KeyAccessSettings.SetAccessRule(KeyFullControlRule) ' place the Full Control rule into the list (now the sole entry)
            Subkey.SetAccessControl(KeyAccessSettings) ' replace the key's security settings with the new one (no inheritance, the current user has full control, no one else has any permissions)
        Catch ex As Exception
            ErrorMessage = String.Format("An error occurred while attempting to access and set permissions for {0}\{1}: {2}", BaseKey.Name, RegistrySubKeyName, ex.Message)
            Return False
        End Try
        ' All done
        Return True
    End Function
End Module
