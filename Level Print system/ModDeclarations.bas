Attribute VB_Name = "ModDeclarations"
'This Module contains all the Global variables used in this Software

'List of data type declaration characters that might be used in this software
'
'1. x$ -> x is a string data type
'2. x& -> x is a long data type
'3. x! -> x is a single data type
'4. x@ -> x is a currency data type
'5. x# -> x is a double data type
'6. x% -> x is an integer data type

Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam%, ByVal lParam As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public iIndex(2) As Long
Public iBuffer(5) As String
Public vOpenForms(100) As Form
Public iSetNo&, UserGroupLevel&
Public iOpenForms&, vCancelOperation%
Public Displaying, iFrmRefreshing, mCanClearPassword As Boolean
Public iSelTbl$, EditRecordID$, vMultiSelData$, vMultiData$, iSystemDateFormat$

Public iFrm(0 To 5) As Form

'Holds Picture in bytes for storage
Public iDataBytes() As Byte

Public iArrayList() As String
Public iTmpArrayList() As String

'DATABASE
'-----------------------------------------------------

'Represents the entire set of records from
'a base table or the results of an executed command
Public iRs As New ADODB.Recordset
Public iTempRs As New ADODB.Recordset

'Represents an open connection to a data source.
Public iAdoCN As New ADODB.Connection

'FILE ACCESS
'-----------------------------------------------------
'Provides access to the computer's file system.
'Public iFso As New FileSystemObject

Public Type Groups
   Group_ID As Integer
   Group_Name As String
   Hierarchy As Integer
End Type

Public Type Users
   UserID As Integer
   UserName As String
   LoginID As String
   LoginName As String
   SubHierarchy As Integer
   Group As Groups
End Type

Public Type Licences
    User_Name As String
    Computer_Name As String
    Users As Long
End Type

Public Type Configurations

    iForeignGuestAddedCharges As String 'Default = 10%
    iCommissionCharges As String        'Default = 5%
    iRoomServiceCharges As String       'Default = 2%
    
    ManyRoomsOneGuest As Boolean
    ManyGuestsOneRoom As Boolean
    ReAssignFacilityToOneGuest As Boolean
    
End Type

Public iSettings As Configurations

Public User As Users 'Set variable to handle all the Software Groups
Public Licence As Licences




