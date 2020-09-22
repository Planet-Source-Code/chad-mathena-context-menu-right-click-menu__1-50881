Attribute VB_Name = "modContext"
Option Explicit
Public m_FileName As String

' Hello everybody who will ready this!
' This program is a DLL that adds to the context menu when right-clicking
' in explorer, but can also be used as a base to program onto the more
' demanding area of namepsace extention.
' These areas are very poorly documented, and mostly in C, making it very hard
' to do anything in VB. I did manage to find a bit of code written in VB to almost
' do exactly the same thing as this, but I have rewritten it for ease of use
' and understanding. This piece of code was very poorly commented, too.
' However, it did include two type library's that make interfacing to Explorer
' possible in this circumstance. I certainly would struggle a lot before
' actually writing a type library that worked!
' Well, here's the source code in its entirety. There should also be a .REG
' file that imports the registry settings to make this program work. The DLL
' file will need to be registered, too. Then right-click on a file, and
' see the new menu item thats added. Then try it with a group of files,
' even folders ;)
' I didn't include a bitmap as an icon in the menu, 'coz that's asking for trouble.
' I won't say that this DLL will actually be useful for anybody, but it shows
' how to do some really neat stuff that is otherwise untouched by the hands
' of VB Programmers.
'
' I've worked hard to write this for you, and all of the other people out there,
' so I will invoke a [Greek, I think] proverb:
' "Give credit where credit is due!" - and I WANT CREDIT!
'
' Likewise, I must stick to my own principles - Thanks to Andy Stotzer for the
' type libraries that he created/found that this project revolves upon!
'
' So, Enjoy the code!
'
' Jolyon Bloomfield   ICQ: 11084041    E-mail Jolyon_B@Hotmail.Com
'
' P.S., Remember VERSION COMPATIBILITY for upgrades and stuff, because the GUID will
' change with every compile you do if you don't put on binary compatibility!
'
' P.P.S., BTW, the resource file stores the icon for this dll ;)

'Addition********
' to get the GUID compile the dll and then do a registry search for ContextMenu.cMenu.
'The registry entry you find will have the clsid, this is the number you want.
' Place this number in the registry file ContextMenu.reg located with this project.
'To test after you compile and after you reg it and after you place correct registry entires
'then right click on a file.  To test you changes, you will have to close all instances of the explorer
'before you can recompile, if you have VERSION COMPATIBILITY set to Binary COMPATIBILITY then you can just
'open explorer again and right click the file.


'
' These are the types necessary for transmitting data to the shell
'
Public Type STGMEDIUM
  tymed               As Long
  hGlobal             As Long
  pUnkForRelease      As IUnknown
End Type

Public Type FORMATETC
  cfFormat            As Long
  ptd                 As Long
  dwAspect            As Long
  lindex              As Long
  tymed               As Long
End Type

Public Type CMINVOKECOMMANDINFO
    cbSize              As Long    ' sizeof(CMINVOKECOMMANDINFO)
    fMask               As Long    ' any combination of CMIC_MASK_*
    hWnd                As Long    ' might be NULL (indicating no owner window)
    lpVerb              As Long    ' either a string or MAKEINTRESOURCE(idOffset)
    lpParameters        As Long    ' might be NULL (indicating no parameter)
    lpDirectory         As Long    ' might be NULL (indicating no specific directory)
    nShow               As Long    ' one of SW_ values for ShowWindow() API
    dwHotKey            As Long
    hIcon               As Long
End Type


'
' The API calls...
'
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As Long, ByVal pUINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function ReleaseStgMedium Lib "ole32.dll" (pMedium As STGMEDIUM) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (lpstring As Any) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
   (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   LParam As Any) _
   As Long
   
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal imageType As Long, ByVal newWidth As Long, ByVal newHeight As Long, ByVal lFlags As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function CreateIC Lib "gdi32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As String) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

'
' And the constants
'
Public Const CF_HDROP = 15              ' For gettings the files
Public Const DVASPECT_CONTENT = 1       '  "
Public Const TYMED_HGLOBAL = 1          '  "
Public Const REG_SZ = 1&                ' Registry access
Public Const PAGE_EXECUTE_READWRITE = &H40&      ' Memory functioning

Private Const IMAGE_BITMAP = &O0
Private Const IMAGE_ICON = 1
Private Const WM_SETICON As Long = &H80
Private Const ICON_SMALL As Long = 0
Private Const SM_CXSMICON    As Long = 49
Private Const SM_CYSMICON    As Long = 50
Private Const LR_COPYFROMRESOURCE = 3
Private Const DI_NORMAL = &H3

Private Const LR_LOADFROMFILE = &H10
Private Const LR_CREATEDIBSECTION = 8192 ' Menu flags for Add/Check/EnableMenuItem/etc

Public Const MF_INSERT = &H0&
Public Const MF_CHANGE = &H80&
Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_REMOVE = &H1000&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_SEPARATOR = &H800&
Public Const MF_ENABLED = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_DISABLED = &H2&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_USECHECKBITMAPS = &H200&
Public Const MF_STRING = &H0&
Public Const MF_BITMAP = &H4&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_POPUP = &H10&
Public Const MF_MENUBARBREAK = &H20&
Public Const MF_MENUBREAK = &H40&
Public Const MF_UNHILITE = &H0&
Public Const MF_HILITE = &H80&
Public Const MF_DEFAULT = &H1000&
Public Const MF_SYSMENU = &H2000&
Public Const MF_HELP = &H4000&
Public Const MF_RIGHTJUSTIFY = &H4000&
Public Const MF_MOUSESELECT = &H8000&
Public Const MF_END = &H80&

Public Const MFT_STRING = MF_STRING
Public Const MFT_BITMAP = MF_BITMAP
Public Const MFT_MENUBARBREAK = MF_MENUBARBREAK
Public Const MFT_MENUBREAK = MF_MENUBREAK
Public Const MFT_OWNERDRAW = MF_OWNERDRAW
Public Const MFT_RADIOCHECK = &H200&
Public Const MFT_SEPARATOR = MF_SEPARATOR
Public Const MFT_RIGHTORDER = &H2000&
Public Const MFT_RIGHTJUSTIFY = MF_RIGHTJUSTIFY

' Menu flags for Add/Check/EnableMenuItem/etc
Public Const MFS_GRAYED = &H3&
Public Const MFS_DISABLED = MFS_GRAYED
Public Const MFS_CHECKED = MF_CHECKED
Public Const MFS_HILITE = MF_HILITE
Public Const MFS_ENABLED = MF_ENABLED
Public Const MFS_UNCHECKED = MF_UNCHECKED
Public Const MFS_UNHILITE = MF_UNHILITE
Public Const MFS_DEFAULT = MF_DEFAULT

' QueryContextMenu uFlags
Public Const CMF_NORMAL = &H0&
Public Const CMF_DEFAULTONLY = &H1&
Public Const CMF_VERBSONLY = &H2&
Public Const CMF_EXPLORE = &H4&
Public Const CMF_NOVERBS = &H8&
Public Const CMF_CANRENAME = &H10&
Public Const CMF_NODEFAULT = &H20&
Public Const CMF_INCLUDESTATIC = &H40&
Public Const CMF_RESERVED = &HFFFF0000

' GetCommandString uFlags
Public Const GCS_VERBA = &H0&                   ' canonical verb
Public Const GCS_HELPTEXTA = &H1&               ' help text (for status bar)
Public Const GCS_VALIDATEA = &H2&               ' validate command exists
Public Const GCS_VERBW = &H4&                   ' canonical verb (Unicode)
Public Const GCS_HELPTEXTW = &H5&               ' help text (Unicode version)
Public Const GCS_VALIDATEW = &H6&               ' validate command exists (Unicode)

Public Const CMDSTR_NEWFOLDER = "NewFolder"
Public Const CMDSTR_VIEWLIST = "ViewList"
Public Const CMDSTR_VIEWDETAILS = "ViewDetails"


'--------------------------------------------------------------------------
'-      And now for the program's code, not the API calls and stuff       -
'--------------------------------------------------------------------------

' For storing the old address of a remapped procedure
Public pOldFunction As Long
' Two constants that our DLL recognises the menu systems by.
' They have to be in here so that sc_QueryContextMenu can access them
Public Const mPROGRAM_CLASS_NAME = "ContextMenu.cMenu"
Public Const mMENU_ITEM_TEXT = "My NIFTY little &Extension Chad"
Public Const mSTATUS_TEXT = "Check out what my NIFTY little extension can do!"
' The Ampersand ("&") says to add the little line under the letter
' like in a normal VB menu.
Public idCmd As Long

 
    


 

Public Function ReplaceVtableEntry(pObj As Long, _
    EntryNumber As Integer, _
    ByVal lpfn As Long) As Long
'
' Don't even ask about this procedure... I've basically ripped it out of the
' original source code; I certainly couldn't write it myself, but I do understand it.
' Basically, it rips out the reference to the class' function that needs to
' be replaced, and replaces it with the address of another function.
' It actually alters memory inside the VB workings of the VTable, and I really
' suggest that you stay away from it as far as possible...
'

Dim lOldAddr        As Long
Dim lpVtableHead    As Long
Dim lpfnAddr        As Long
Dim lOldProtect     As Long

CopyMemory lpVtableHead, ByVal pObj, 4
lpfnAddr = lpVtableHead + (EntryNumber - 1) * 4
CopyMemory lOldAddr, ByVal lpfnAddr, 4

Call VirtualProtect(lpfnAddr, 4, PAGE_EXECUTE_READWRITE, lOldProtect)
CopyMemory ByVal lpfnAddr, lpfn, 4
Call VirtualProtect(lpfnAddr, 4, lOldProtect, lOldProtect)

ReplaceVtableEntry = lOldAddr
End Function

' This is the function that modifies the menu to add our own items to it.
' Do anything that needs to be done to the menu here!
Public Function sc_QueryContextMenu(ByVal This As IContextMenu, _
    ByVal hMenu As Long, _
    ByVal indexMenu As Long, _
    ByVal idCmdFirst As Long, _
    ByVal idCmdLast As Long, _
    ByVal uFlags As Long) As Long

Dim Rc              As Long
Dim idCmd           As Long
Dim szMenu          As String
Dim szMenuText      As String
Dim bAppendItems    As Boolean
Dim szTemp          As String
Dim phMenu As Long
Dim bmp As Long
Dim str As String

On Error GoTo er
    'str = "C:\car.ico" ' Use if you want to load from a specific file

    'bmp = LoadImage(0, str, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
    
    ReDim SubMenuHandle(1)

    Dim p As stdole.IPictureDisp
    Set p = LoadResPicture(101, vbResBitmap)
    Dim l As Long
    bmp = CopyImage(p.handle, IMAGE_BITMAP, 13, 13, LR_COPYFROMRESOURCE)
    Set p = Nothing
    
    idCmd = idCmdFirst
    bAppendItems = True

    ' Check to see if the items need to be added
    If ((uFlags And &HF&) = CMF_NORMAL) Then
      szMenuText = mMENU_ITEM_TEXT
    ElseIf ((uFlags And CMF_VERBSONLY) = CMF_VERBSONLY) Then
      szMenuText = mMENU_ITEM_TEXT
    ElseIf ((uFlags And CMF_EXPLORE) = CMF_EXPLORE) Then
      szMenuText = mMENU_ITEM_TEXT
    ElseIf ((uFlags And CMF_DEFAULTONLY) = CMF_DEFAULTONLY) Then
      bAppendItems = False
    Else
      bAppendItems = False
    End If
    
    If bAppendItems Then
        ' Insert our menu item
        ' Copy this a few times for multiple items
        ' If you really know your way around the API and menus, you can get all of
        ' the information currently in the menu. Bare in mind, that not all of
        ' the context handlers might have yet been initialized, so the menu will
        ' not be complete.
        
        ' Note: I've included a few variations in the menu system for twofold:
        ' 1 - so that you might learn something
        ' 2 - so that I can have some fun B)
        ' Only use on at a time, unless you write more handler code.
      
        phMenu = CreatePopupMenu ' Creates the popup menu to be used for submenus
        
        ' Here's a straight menu item
        'InsertMenu phMenu, indexMenu, MF_BYPOSITION, idCmd, "Chad 3"
        InsertMenu hMenu, indexMenu, MF_POPUP Or MF_BYPOSITION, phMenu, m_FileName
        SetMenuItemBitmaps hMenu, indexMenu, MF_BITMAP Or MF_BYPOSITION, bmp, bmp
        
        indexMenu = indexMenu + 1
        InsertMenu phMenu, indexMenu, MF_BYPOSITION, idCmd, "SubMenu 1 "
        SetMenuItemBitmaps phMenu, indexMenu - 1, MF_BITMAP Or MF_BYPOSITION, bmp, bmp
        
        indexMenu = indexMenu + 1
        idCmd = idCmd + 1
        InsertMenu phMenu, indexMenu, MF_BYPOSITION, idCmd, "SubMenu 2 "
        SetMenuItemBitmaps phMenu, indexMenu - 1, MF_BITMAP Or MF_BYPOSITION, bmp, bmp
        
        'To add more submenus add
        
        'indexMenu = indexMenu + 1
        'idCmd = idCmd + 1
        'InsertMenu phMenu, indexMenu, MF_BYPOSITION, idCmd, "SubMenu 3 "
        'SetMenuItemBitmaps phMenu, indexMenu - 1, MF_BITMAP Or MF_BYPOSITION, bmp, bmp
        
        
          ' This one should hopefully be tagged onto the end
          'InsertMenu hMenu, &HFFFFFFFF, MF_BYPOSITION, idCmd, szMenuText
            
          ' Here's a go at making one checked
          'InsertMenu hMenu, indexMenu, MF_BYPOSITION Or MF_CHECKED, idCmd, szMenuText
          
          ' This one will go onto a new column
          'InsertMenu hMenu, indexMenu, MF_BYPOSITION Or MFT_MENUBARBREAK, idCmd, szMenuText
          
          ' This one has a radiobutton type check next to it
          'InsertMenu hMenu, indexMenu, MF_BYPOSITION Or MF_CHECKED Or MFT_RADIOCHECK, idCmd, szMenuText
           
           
          '
          ' If anyone can figure out how to make your item the default item,
          ' or how to control the position that your menu items are stored in,
          ' can you please tell me? Jolyon_B@Hotmail.Com
          ' It might be next to impossible, or it might just be me being tired at 1 in
          ' the morning. You find out. I dare ya ;)
          '
           
           
          ' With a it of work, you should be able to place anything anywhere,
          ' using a few API calls to get the structure of the menu as it stands
          ' You can even edit the menu, chaning around items, and deleting them,
          ' but think about it. Would you want somebody deleting your menu item?
              
          ' After ever Menu insertion, we must have this:
          ' Increment Index and menu count...
          indexMenu = indexMenu + 1
          idCmd = idCmd + 1
        
          ' Here's a menu separator, just for the sake of putting one in
          Call InsertMenu(hMenu, indexMenu, MF_SEPARATOR Or MF_BYPOSITION, 0, vbNullString)
            
          ' Must increase the number of the index, but not the command, only for separator bars
          indexMenu = indexMenu + 1
        
          ' Must return number of menu items inserted
        sc_QueryContextMenu = (idCmd - idCmdFirst)
    Else
        ' Must return number of menu items inserted
        sc_QueryContextMenu = 0
    End If
    
    Exit Function
er:
    MsgBox Err.Description
End Function

