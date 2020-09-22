VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit '- Force declaration of all variables. Much safer!

' These are the two interfaces that must be implemented for Explorer to be
' able to transfer data to and from the dll. They are declared in the
' type libraries.
Implements IContextMenu
Implements IShellExtInit

Private Selection() As String

Private Sub Class_Initialize()
' Initialize the selection string
ReDim Selection(0 To 0) As String

' Because of the troubles with IContextMenu_QueryContextMenu,
' it has to be reworked, and thus, rereferenced. The code below places
' the function stored within the module instead of IContextMenu_QueryContextMenu,
' changing the VTable. Don't ask me how anybody knew to do that... but it works.

Dim pImp As Long
Dim oImp As IContextMenu
Set oImp = Me
Call CopyMemory(pImp, oImp, 4)
pOldFunction = ReplaceVtableEntry(pImp, 4, AddressOf sc_QueryContextMenu)
End Sub

Private Sub Class_Terminate()
ReDim Selection(0 To 0) As String

' Must undo the damage we have caused...
Dim pImp As Long
Dim oImp As IContextMenu
Set oImp = Me
Call CopyMemory(pImp, oImp, 4)
pOldFunction = ReplaceVtableEntry(pImp, 4, pOldFunction)

End Sub

Private Sub IContextMenu_QueryContextMenu(ByVal hMenu As Long, ByVal indexMenu As Long, ByVal idCmdFirst As Long, ByVal idCmdLast As Long, ByVal uFlags As Long)
' There's a problem in the Type Library here - as you can see, this procedure
' has been defined as void - or a subroutine, returning no value
' But Explorer expects to get a value back. What can we do?
' See the Class_Initialize section.

' Remapped to modContext.sc_QueryContextMenu
End Sub

' This is called when the shell initialises the class for action!
Private Sub IShellExtInit_Initialize(ByVal pidlFolder As Long, _
    ByVal lpdobj As Long, _
    ByVal hKeyProgID As Long)

Dim strVal      As String
Dim pIDataObj   As IUnknown
Dim oIDataObj   As IDataObject
Dim FmtEtc      As FORMATETC
Dim pMedium     As STGMEDIUM
Dim szTemp      As String
Dim iFile       As Long
Dim iFiles      As Long
Dim Rc          As Long
Dim pos As Long
Dim pos2 As Long
' lpdobj stores a pointer to the object that holds all of the filenames that
' are selected. We have to be a little naughty to get them out, but I don't
' think that there's any other way.


' Copy dataobject pointer to IUnknown variable
' WARNING: This does not addref the actual object- it doesn't know that we've
' referenced it
CopyMemory pIDataObj, lpdobj, 4

' Make IDataObject Variable QueryInterface from IUnknown
Set oIDataObj = pIDataObj
' oIDataObj now stores the object required

' Set IUnknown Variable to nothing with calling release
CopyMemory pIDataObj, 0&, 4
' Setting PIDataObj = Nothing will screw Explorer's references...

' Set the format to get the data in...
With FmtEtc
  .cfFormat = CF_HDROP
  .ptd = 0
  .dwAspect = DVASPECT_CONTENT
  .lindex = -1
  .tymed = TYMED_HGLOBAL
End With

' ... and store it in our medium
oIDataObj.GetData ByVal VarPtr(FmtEtc), ByVal VarPtr(pMedium)
Rc = Err.LastDllError 'Check for errors

' Now comes the fun - we tally up all of the filenames
' The format that the filenames are stored in is exactly the same as
' that of the drag operations - so we use DragQueryFile
' &hFFFFFFFF returns the number of files in the drag
iFiles = DragQueryFile(pMedium.hGlobal, &HFFFFFFFF, vbNullString, 0)
' Notice that there are 2 different variables - iFile and iFiles
' Iterate through the number of files
For iFile = 0 To (iFiles - 1)
  szTemp = String(500, 0)                     ' Set the buffer size
  Rc = DragQueryFile(pMedium.hGlobal, iFile, szTemp, Len(szTemp))      ' Get the data
  If (Rc > 0) Then              ' If the returned number of characters are > 0
    szTemp = Left$(szTemp, Rc)      ' Truncate to the full string lenth (- the buffer)
    ' Make the array a tad bigger...
    ReDim Preserve Selection(0 To UBound(Selection) + 1) As String
    ' ... andwhack the string into it
    If iFiles = 1 Then
        m_FileName = Mid$(szTemp, InStrRev(szTemp, "\") + 1)
    Else
        pos = InStrRev(szTemp, "\") + 1
        pos2 = InStrRev(szTemp, "\", pos - 2)
        m_FileName = Mid$(szTemp, pos2 + 1, pos - pos2 - 2)
    End If
    Selection(UBound(Selection)) = szTemp
  End If
Next

Rc = ReleaseStgMedium(pMedium)
Set oIDataObj = Nothing

' These next two lines don't do much - they just get the ProgID from
' the hKeyProgID in the registry.
' If you want to look at them in greater detail, you'll have to sort out
' the data thats returned, because it will still be 512 chars, and with null chars
' to terminate the string. Since it is a variable passed, I place it here
' for you to use if necessary.
' This should work, but I'm not sure that the hKeyProgID is an open key.
' I'm assuming it is. I can't really be bothered to test it, because I
' don't know what it is anyway. If this doesn't work, just open the key and
' then take a peek at it.
strVal = String$(512, 0)
Rc = RegQueryValueEx(hKeyProgID, vbNullString, 0&, REG_SZ, strVal, Len(strVal))

End Sub

' Nothing to do with the shell interfacing, but a useful routine
' to get the short path name out of a complete one
Private Function GetShortName(ByVal FileAndPath As String) As String
Dim Rc As Long
Dim sp As String
sp = String(500, 0)
Rc = GetShortPathName(FileAndPath, sp, Len(sp))
If (Rc > 0) Then
  sp = Left$(sp, Rc)
End If
GetShortName = sp
End Function

' This simply asks for information, after the menu has been created
' - for the status bar, etc.
Private Sub IContextMenu_GetCommandString(ByVal idCmd As Long, _
    ByVal uFlags As Long, _
    ByVal pwReserved As Long, _
    ByVal pszName As Long, _
    ByVal cchMax As Long)
'Referencing pwReserved will cause a GPF, it is a null pointer.
Dim bszName()   As Byte
Dim cChars      As Long
Dim szName      As String
Dim temp        As String

' Select the command that we have to respond to.
' Have a look at where the different messages pop up when you actually click.
Select Case uFlags
  Case GCS_VERBA, GCS_VERBW
    Select Case idCmd
      Case 0
        szName = "(GCS_VERB)MENUITEM1"
    End Select
  Case GCS_HELPTEXTA, GCS_HELPTEXTW
    Select Case idCmd
      Case 0
        szName = mSTATUS_TEXT
    End Select
  Case GCS_VALIDATEA, GCS_VALIDATEW
    Select Case idCmd
      Case 0
        szName = "(GCS_VALIDATE)New menu item number 1"
    End Select
End Select

' Sign off the string,
szName = Left$(szName, cchMax - 1) & vbNullChar

' then convert to Unicode the ones that have to be converted to Unicode...
Select Case uFlags
  Case GCS_VERBA, GCS_HELPTEXTA, GCS_VALIDATEA
    If (szName <> "") Then
      bszName = StrConv(szName, vbFromUnicode)
      Call CopyMemory(ByVal pszName, bszName(0), UBound(bszName) + 1)
    End If
  Case GCS_VERBW, GCS_HELPTEXTW, GCS_VALIDATEW
    If (szName <> "") Then
      bszName = szName
      Call CopyMemory(ByVal pszName, bszName(0), UBound(bszName) + 1)
    End If
End Select

End Sub

' Now the fun starts - our code has been invoked, and now we ACT!
Private Sub IContextMenu_InvokeCommand(ByVal lpcmi As Long)

'
' N.B. There is no support for Unicode in here, although it is a very real
' possibility that it might be needed. I, for one, HATE Unicode, because its
' a pain in the neck to convert to ansi, and back again, and blah blah blah...
' If you need to be able to work with unicode, look in the MSDN library under
' "IContextMenu::InvokeCommand".
'

Dim CMI         As CMINVOKECOMMANDINFO
Dim Directory   As String
Dim Parameters  As String
Dim Verb        As String
Dim sLen        As Long

' Get the information into a structure
CopyMemory CMI, ByVal lpcmi, Len(CMI)

' Get the directory from the pointer
Directory = String(lstrlen(ByVal CMI.lpDirectory), Chr$(0))
lstrcpy ByVal Directory, ByVal CMI.lpDirectory

' Parameters - get them
Parameters = String(lstrlen(ByVal CMI.lpParameters), Chr$(0))
sLen = lstrlen(ByVal CMI.lpParameters)
If (sLen > 0) Then
  Dim bData() As Byte
  ReDim bData(sLen) As Byte
  CopyMemory bData(0), ByVal CMI.lpParameters, sLen
End If
lstrcpy ByVal Parameters, ByVal CMI.lpParameters

' The command can be sent one of two ways.
' The first is by the identifier, and the second by the verb.
' Both will have to be parsed differently.
If ((CMI.lpVerb And &HFFFF0000) = 0) Then
  idCmd = CMI.lpVerb
  '
  ' The message has been sent by the command identifier-
  ' use the numbers that the commands were given.
  '
  'Select Case idCmd
    'Case 0
      ' MenuItem 0 selected...
      Call IGotClicked1(CMI.hWnd, idCmd, CMI)
    'Case 1
      ' MenuItem 1 selected...
      ' etc.
  'End Select
Else
  Verb = String(lstrlen(ByVal CMI.lpVerb), Chr$(0))
  lstrcpy ByVal Verb, ByVal CMI.lpVerb
  ' The message was sent by verb.
  ' I'm not sure if this responds to the text that is in the menu, but I
  ' assume so. Otherwise it might be the text that it was was passed
  ' in the getcommandstring sub
  If Verb = modContext.mMENU_ITEM_TEXT Then
    Call IGotClicked1(CMI.hWnd, idCmd, CMI)
  End If
End If

End Sub

Private Sub IGotClicked1(ByVal hWnd As Long, id As Long, cm As CMINVOKECOMMANDINFO)
'This is where you would put code to do actions based upon what is returned.
Dim Display As String
Dim i As Integer
Dim submenu As String

    Select Case id
        Case 0
            submenu = "Submenu 1"
        Case 1
            submenu = "Submenu 2"
        Case Else
            submenu = id
    End Select
    If UBound(Selection) = 0 Then
      Display = "No files or folders were selected."
    Else
      Display = Trim(str(UBound(Selection))) & " file(s) and/or folder(s) were selected:"
      For i = 1 To UBound(Selection)
        Display = Display & vbCrLf & Selection(i)
      Next i
    End If
    
    MessageBox hWnd, Display & " from " & submenu, "Shell Extension ", vbInformation
End Sub
