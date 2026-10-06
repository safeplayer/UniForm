VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UniForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' UniForm — One class to dress up VB6 forms with true Unicode and native Win32 controls — zero external dependencies.
'  A wrapper class for creating all win32 controls on a VB Form, with VB-like events and methods, naturally support Unicode and all Windows Messages
'
' Advantages: Unicode, Owner-Drawn controls, automatic subclassing that fires MouseWheel, MouseLeave events and user events, additional convenient control wrappers
'
' THE CONCEPT:
'    1. In a VB Form, create Labels as placeholders for controls' position and size
'    2. Give the labels names that fit the controls, like: lstFiles, txtAddress, cmdPlay, etc
'    3. Set labels' font, backcolor, forecolor, backstyle, alignment, border style, appearance - those will be copied to the actual controls
'    4. For each label, enter the Windows Control Class in LinkItem property.
'       Any valid Windows Class is possible, but internal supported classes are:
'       Edit(TextBox), ListBox, ComboBox, Frame, Static(For labels and images),
'       Button(CommandButton), CheckBox, 3State(a checkbox with 3 states), RadioButton(OptionButton), RadioButtonGroup(to start a new radio buttons group)
'    5. In Form Code, add a class object member: Public WithEvents Uni As UniForm
'    6. At Form_Load: Set Uni = New UniForm: Uni.Init Me ' <- This will create the controls and hide the labels
'    7. The class name can contain some special style flags after a hyphen. e.g. For multiline and locked Edit control: Edit-lm
'       The complete list of all supported flags per control - below.
'       For a complete control of styles and some rare styles, you can catch Create event. For a ListBox catch CreateListBox event, for a TextBox catch CreateEdit event etc. For other controls use generic Create event
'       Every Create_xxx event uses a special styles enum for the specific control-type. Add (or exclude) the desired styles and ex_styles to the parameters
'    8. The class provides VB-style wrappers for useful control methods. The general syntax is Uni.Method(control). e.g. Uni.ToolTipText(lstFiles) = "Drag files to the list"
'       A list of all wrapper methods and properties - below.
'    9. Events: Standard VB events like MouseMove, KeyDown etc, In addition: MouseWheel, MouseEnter, MouseLeave, MouseHover, OwnerDraw(for owner-drawn controls), Message(all messages, including user messages)
'       Unlike VB-events, a specific event handler is for ALL CONTROLS: Uni_Event(ControlName...). Use ControlName to handle an individual control inside the event handler
'       The Compare events is used when calling SortList for ListBox or ComboBox with eSortCustom flag. The event expects Cmp parameter to become -1,1 or 0 according to items comparison.
'   10. Control-Arrays: It is possible to create arrays of Labels and refer to each element by Index: Uni.Text(txtName(i)) = asName(i). All UniForm events pass iIndex parameter (which is -1 if not an array)
'       It is also possible to dynamically add controls to the form: Load txtName(n) and then add to UniForm: Uni.AddControl txtName(n)
'   11. Bonus: For tooltips, there is HideToolTip and ShowToolTip to show the tooltip with arbitrary text and position
'
'   List of all supported style flags for controls creation: (The flags appear after the class name and a hyphen e.g. ListBox-sm)
'           Edit:lmn = l(locked)/m(multiline)/n(number only)
'           Button/Checkbox/Radiobutton/3State:og = o(owner-draw)/g(graphical-pushlike)
'           Listbox:smceov = s(sorted)/m(multiselection)/c(multicolumn)/e(extended selection)/o(owner-draw)/v(owner-draw-variable)
'           ComboBox:scdpaov = s(sorted)/c(combo+edit)/d(droplist)/p(simple-no-drop)/a(auto-hscroll)/o(owner-draw)/v(owner-draw-variable)
'           Static:ibo = i(Icon)/b(Bitmap)/o(Owner-draw)
'
'   List of all supported wrapper methods: (methods with (rw) are properties with Get(r) and/or Let(w))
'       ListBox and ComboBox: AddItem, RemoveItem, List(rw), ListCount(r), ListIndex(rw), ItemData(rw), ClearList, SortList, MouseRow(r), TopIndex(rw),
'                             FindItem, ListWidth(w), ItemHeight(rw), VisibleItems(r), IsItemVisible(r)
'       TextBox (and ComboBox's edit part): SetPasswordChar, HideSelection(rw), Locked(rw), MaxLength(w), NumbersOnly(w), RTL(w), ScrollBars(w), CueBanner(rw)
'                                           SelStart(r), SelLen(r), SelSet, SelText(rw), CurLine(r), LineStart, LineLen, GetLine, CharPoint, CharAtPoint,
'       ComboBox: DropState(rw), CueBanner(rw) (This is a custom CueBanner, not Windows Controls CueBanner)
'       CheckBox/RadioButton: Checked(rw)
'       For all controls: Tag(Variant)(rw), Text(rw), Visible(rw), Enable(rw), ToolTipText(rw), ToolTipFlags(alignment of tooltiptext)(w), TextWidth(text width with control's font)(r), TextHeight(r)
'       Left, Top, Width, Height, ClientWidth, ClientHeight, Move -> All values r/w are in pixels!
'       MouseInWindow, AcceptFiles(w), SetMouseCursor, BackColor(rw), ForeColor(rw), BackStyle(rw), SetFocus, ZOrder, Alignment(w), BorderStyle(w), Appearance(w), SetFont, Refresh, GetWindow
'       SetMouseCursor accepts an hCursor or enum with system cursor index. SetFont accepts VB IFont or hFont.

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Event Create(ByVal ControlName As String, ByVal iIndex As Long, hwndParent As Long, ControlClass As String, Style As Long, exStyle As EX_STYLES)
Event CreateListBox(ByVal ControlName As String, ByVal iIndex As Long, hwndParent As Long, Style As LISTBOX_STYLES, exStyle As EX_STYLES)
Event CreateComboBox(ByVal ControlName As String, ByVal iIndex As Long, hwndParent As Long, Style As COMBO_STYLES, exStyle As EX_STYLES)
Event CreateStatic(ByVal ControlName As String, ByVal iIndex As Long, hwndParent As Long, StaticType As STATIC_TYPE, StaticFlags As STATIC_FLAGS, exStyle As EX_STYLES)
Event CreateButton(ByVal ControlName As String, ByVal iIndex As Long, hwndParent As Long, Style As BUTTON_STYLES, exStyle As EX_STYLES)
Event CreateCheckBox(ByVal ControlName As String, ByVal iIndex As Long, hwndParent As Long, Style As BUTTON_STYLES, exStyle As EX_STYLES)
Event CreateRadioButton(ByVal ControlName As String, ByVal iIndex As Long, hwndParent As Long, Style As BUTTON_STYLES, exStyle As EX_STYLES)
Event Create3State(ByVal ControlName As String, ByVal iIndex As Long, hwndParent As Long, Style As BUTTON_STYLES, exStyle As EX_STYLES)
Event CreateFrame(ByVal ControlName As String, ByVal iIndex As Long, hwndParent As Long, Style As BUTTON_STYLES, exStyle As EX_STYLES)
Event CreateEdit(ByVal ControlName As String, ByVal iIndex As Long, hwndParent As Long, Style As EDIT_STYLES, exStyle As EX_STYLES)
Event GotFocus(ByVal ControlName As String, ByVal iIndex As Long)
Event LostFocus(ByVal ControlName As String, ByVal iIndex As Long)
Event Change(ByVal ControlName As String, ByVal iIndex As Long)
Event Scroll(ByVal ControlName As String, ByVal iIndex As Long)
Event BeforePaste(ByVal ControlName As String, ByVal iIndex As Long)
Event AfterPaste(ByVal ControlName As String, ByVal iIndex As Long)
Event Click(ByVal ControlName As String, ByVal iIndex As Long)
Event DblClick(ByVal ControlName As String, ByVal iIndex As Long)
Event SelChange(ByVal ControlName As String, ByVal iIndex As Long)
Event DropDown(ByVal ControlName As String, ByVal iIndex As Long)
Event DropClose(ByVal ControlName As String, ByVal iIndex As Long)
Event KeyDown(ByVal ControlName As String, ByVal iIndex As Long, KeyCode As Integer, Shift As Integer)
Event KeyUp(ByVal ControlName As String, ByVal iIndex As Long, KeyCode As Integer, Shift As Integer)
Event KeyPress(ByVal ControlName As String, ByVal iIndex As Long, WChar As Integer)
Event MouseDown(ByVal ControlName As String, ByVal iIndex As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Event MouseUp(ByVal ControlName As String, ByVal iIndex As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Event MouseMove(ByVal ControlName As String, ByVal iIndex As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Event MouseWheel(ByVal ControlName As String, ByVal iIndex As Long, ByVal Delta As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Event DropFiles(ByVal ControlName As String, ByVal iIndex As Long, oFiles As Collection)
Event MouseEnter(ByVal ControlName As String, ByVal iIndex As Long)
Event MouseLeave(ByVal ControlName As String, ByVal iIndex As Long)
Event MouseHover(ByVal ControlName As String, ByVal iIndex As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Event OwnerDraw(ByVal ControlName As String, ByVal iIndex As Long, ByVal hdc As Long, ByVal Item As Long, ByVal iLeft As Long, ByVal iTop As Long, ByVal iWidth As Long, ByVal iHeight As Long, ByVal Action As eCtlAction, ByVal State As eDrawState)
Event Compare(ControlName As String, ByVal iIndex As Long, ByVal Item1 As Long, ByVal Item2 As Long, Cmp As Long, Param)
Event Message(ByVal ControlName As String, ByVal iIndex As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, RetValue As Long, Handled As Boolean)

Private Type tControl
    m_Subclass As IUnknown
    m_SubclassItem As IUnknown
    m_SubclassList As IUnknown
    hWnd As Long
    hEdit As Long
    hList As Long
    sName As String
    sClass As String
    iIndex As Long
    BackColor As Long
    ForeColor As Long
    BackStyle As Long
    hCursor As Long
    ToolTipText As String
    ToolTipFlag As Long
    Tag As Variant
    CueBanner As String
    bCue As Boolean
End Type

Private m_pSubclassForm As IUnknown
'Private m_pSubclassTTT As IUnknown
Private m_Controls() As tControl
Private m_CtlCount&
Private m_hTTT&
Private m_idTTT&
Private m_hwndParent&
Private WM_CUEBANNER As Long

Private Type RECTL
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt(0 To 1) As Long
End Type

Private Type LOGFONTW
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To 32) As Integer
End Type

Private Type PAINTSTRUCT
    hdc As Long
    fErase As Long
    rcPaint As RECTL
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(0 To 31) As Byte
End Type

Private Type tCOMBOBOXINFO
    bSize As Long
    rcItem As RECTL
    rcButton As RECTL
    stateButton As Long
    hwndCombo As Long
    hwndItem As Long
    hwndList As Long
End Type

Private Type tTRACKMOUSEEVENT
  cbSize As Long
  dwFlags As Long
  hwndTrack As Long
  dwHoverTime As Long
End Type

Private Type tDRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECTL
    ItemData As Long
End Type

Private Type WNDCLASSEX
    cbSize As Long
    Style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    hIconSm As Long
End Type

Private Type tMONITORINFO
    cbSize As Long
    rcMonitor As RECTL
    rcWork As RECTL
    dwFlags As Long
End Type

Private Declare Function GetModuleHandle& Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String)
Private Declare Function GetProcAddress& Lib "kernel32" (ByVal hModule&, ByVal lpProcName As String)
Private Declare Function CallWindowProcW& Lib "user32" (ByVal lpPrevWndFunc&, ByVal hWnd&, ByVal iMsg&, ByVal wParam&, ByVal lParam&)
Private Declare Function DefSubclassProc413 Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetProcByOrdinal& Lib "kernel32" Alias "GetProcAddress" (ByVal hModule&, ByVal lpProcOrdinal&)
'Private Declare Function CallWindowProc& Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, ByVal hwnd&, ByVal iMsg&, ByVal wParam&, ByVal lParam&)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination&, ByVal Source&, ByVal Length&)
Private Declare Function FindWindowExA& Lib "user32" (ByVal hwndParent&, ByVal hWndChildAfter&, ByVal lpszClass As String, ByVal lpszWindow As String)
Private Declare Function GetWindowThreadProcessId& Lib "user32.dll" (ByVal hWnd&, ByVal lpdwProcessId&)
Private Declare Function GetCurrentProcessId& Lib "kernel32.dll" ()
Private Declare Function VirtualAlloc& Lib "kernel32" (ByVal lpAddress&, ByVal dwSize&, ByVal flAllocationType&, ByVal flProtect&)
Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
Private Declare Function WindowFromPoint& Lib "user32" (ByVal X&, ByVal Y&)
Private Declare Function LoadCursorV& Lib "user32" Alias "LoadCursorA" (ByVal hInstance&, ByVal lpCursorName&)
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32.dll" (ByVal lpPoint As Long) As Long
Private Declare Function TrackMouseEvent& Lib "user32.dll" (ByVal lpEventTrack&)
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function PostMessageW& Lib "user32.dll" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
Private Declare Function GetDlgItem Lib "user32" (ByVal hWndDlg As Long, ByVal id As Long) As Long
Private Declare Function GetDlgCtrlID Lib "user32" (ByVal hWndCtl As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function MonitorFromWindow& Lib "user32" (ByVal hWnd&, ByVal dwFlags&)
Private Declare Function GetMonitorInfoA& Lib "user32" (ByVal hMonitor&, ByVal lpmi&)
Private Declare Function CompareString& Lib "kernel32.dll" Alias "CompareStringW" (ByVal Locale&, ByVal dwCmpFlags&, ByVal lpString1&, ByVal cchCount1&, ByVal lpString2&, ByVal cchCount2&)
Private Declare Function DragQueryFileW Lib "Shell32.dll" (ByVal drop_handle As Long, ByVal UINT As Long, ByVal lpStr As Long, ByVal ch As Long) As Long
Private Declare Sub DragFinish Lib "Shell32" (ByVal hDrop&)
Private Declare Sub DragAcceptFiles Lib "Shell32" (ByVal hWnd&, ByVal fAccept&)

Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthW" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowTextW Lib "user32.dll" (ByVal hWnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
Private Declare Function SetWindowTextW Lib "user32.dll" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function EnableWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd&, ByVal hWndInsertAfter&, ByVal X&, ByVal Y&, ByVal cx&, ByVal cy&, ByVal wFlags&) As Long
Private Declare Function IsWindowVisible Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetComboBoxInfo& Lib "user32" (ByVal hwndCombo&, ByVal pcbi&)
Private Declare Function GetClientRect& Lib "user32" (ByVal hWnd&, lpRect As RECTL)
Private Declare Function GetWindowRect& Lib "user32" (ByVal hWnd&, lpRect As RECTL)
Private Declare Function ClientToScreen& Lib "user32" (ByVal hWnd&, ByVal lpPoint&)
Private Declare Function ScreenToClient& Lib "user32" (ByVal hWnd&, ByVal lpPoint&)
Private Declare Function GetParent& Lib "user32" (ByVal hWnd&)
Private Declare Function InvalidateRectV Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function GetWindowLongW Lib "user32" (ByVal hWnd&, ByVal nIndex&) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd&, ByVal nIndex&, ByVal dwNewLong&) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExW" (ByVal pcWndClassEx As Long) As Integer
Private Declare Function RegisterWindowMessageA& Lib "user32" (ByVal lpString$)
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetProp& Lib "user32.dll" Alias "GetPropA" (ByVal hWnd&, ByVal lpString As String)
Private Declare Function SetProp& Lib "user32.dll" Alias "SetPropA" (ByVal hWnd&, ByVal lpString As String, ByVal hData&)
Private Declare Function RemoveProp& Lib "user32.dll" Alias "RemovePropA" (ByVal hWnd&, ByVal lpString As String)
Private Declare Function IsDialogMessageW& Lib "user32" (ByVal hDlg&, ByVal lpMsg&)
Private Declare Function IsDialogMessageA& Lib "user32" (ByVal hDlg&, ByVal lpMsg&)
Private Declare Function GetFocus Lib "user32.dll" () As Long
Private Declare Function SetFocusApi Lib "user32.dll" Alias "SetFocus" (ByVal hWnd As Long) As Long

Private Declare Function SelectObject Lib "Gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "Gdi32" Alias "GetObjectW" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetStockObject Lib "Gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "Gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "Gdi32" (ByVal hObject As Long) As Long
Private Declare Function PatBlt Lib "Gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetTextColor Lib "Gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SetDCBrushColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetDCPenColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetBkMode Lib "Gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetBkMode Lib "Gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetBkColor Lib "Gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetBkColor Lib "Gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, ByVal lpRect As Long, ByVal wFormat As Long) As Long
Private Declare Function PtInRect Lib "user32.dll" (ByVal lpRect As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (ByVal lpRect As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetDeviceCaps Lib "Gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFontA& Lib "Gdi32" (ByVal cHeight&, ByVal cWidth&, ByVal cEscapement&, ByVal cOrientation&, ByVal cWeight&, ByVal bItalic&, ByVal bUnderline&, ByVal bStrikeOut&, ByVal iCharSet&, ByVal iOutPrecision&, ByVal iClipPrecision&, ByVal iQuality&, ByVal iPitchAndFamily&, ByVal pszFaceName$)
Private Declare Function CreateFontIndirectW Lib "Gdi32" (ByVal pLogFont As Long) As Long
Private Declare Function GdiFlush& Lib "Gdi32" ()
Private Declare Function IntersectClipRect& Lib "Gdi32" (ByVal hdc&, ByVal left&, ByVal top&, ByVal right&, ByVal bottom&)

Private Declare Function SysAllocString& Lib "oleaut32" (ByVal ptr&)
Private Declare Function SysAllocStringLen& Lib "oleaut32" (ByVal ptr&, ByVal Length&)
Private Declare Function SysAllocStringByteLen& Lib "oleaut32" (ByVal ptr&, ByVal Length&)
Private Declare Sub SysFreeString Lib "oleaut32" (ByVal ptr&)
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr&, RetVal As Byte)
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr&, RetVal As Integer)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr&, RetVal&)
Private Declare Function GetMem4R Lib "msvbvm60" Alias "GetMem4" (ByRef src As Any, ByRef dst As Any) As Long
Private Declare Sub GetMem8 Lib "msvbvm60" (ByVal Addr&, RetVal As Currency)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal Addr&, ByVal NewVal As Byte)
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal Addr&, ByVal NewVal As Integer)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr&, ByVal NewVal&)
Private Declare Sub PutMem8 Lib "msvbvm60" (ByVal Addr&, ByVal NewVal As Currency)

Private Const SIGN_BIT& = &H80000000
Private Const PAGE_EXECUTE_READWRITE& = &H40&
Private Const MEM_COMMIT& = &H1000&
Private Const MONITOR_DEFAULTTONEAREST& = &O2&

Enum eCursorsCtl
    CUR_DEFAULT = -1
    CUR_NONE = 0
    CUR_ARROW = 32512
    CUR_IBEAM = 32513
    CUR_WAIT = 32514
    CUR_CROSS = 32515
    CUR_UPARROW = 32516
    'CUR_SIZE = 32640 ' OBSOLETE: use IDC_SIZEALL
    'CUR_ICON = 32641 ' OBSOLETE: use IDC_ARROW
    CUR_SIZENWSE = 32642
    CUR_SIZENESW = 32643
    CUR_SIZEWE = 32644
    CUR_SIZENS = 32645
    CUR_SIZEALL = 32646
    CUR_NO = 32648
    CUR_HAND = 32649
    CUR_APPSTARTING = 32650
    CUR_HELP = 32651
    CUR_PIN = 32671
    CUR_PERSON = 32672
End Enum

Enum eCtlAction
    ODA_DRAWENTIRE = 1
    ODA_SELECT = 2
    ODA_FOCUS = 4
End Enum

Enum eDrawState
    ODS_SELECTED = 1
    ODS_GRAYED = 2
    ODS_DISABLED = 4
    ODS_CHECKED = 8
    ODS_FOCUS = 16
    ODS_DEFAULT = 32
    ODS_HOTLIGHT = 64
    ODS_INACTIVE = 128
    ODS_NOACCEL = 256
    ODS_NOFOCUSRECT = 512
    ODS_COMBOBOXEDIT = 4096
End Enum

Enum eToolTipFlags
    eTTFLeft = 0
    eTTFCenter = 1
    eTTFRight = 2
    eTTFRTL = 4
End Enum

Enum eSortList
    eSortByStringText
    eSortByStringBinary
    eSortAsNumbers
    eSortByItemData
    eSortCustom
    eSortDescending = 16
End Enum

Private Const WM_PAINT& = &HF&
Private Const WM_COMMAND As Long = &H111
Private Const WM_SETFONT As Long = &H30&
Private Const WM_GETFONT As Long = &H31&
Private Const WM_SETCURSOR As Long = &H20&
Private Const WM_SETFOCUS As Long = &H7&
Private Const WM_KILLFOCUS As Long = &H8&
Private Const WM_KEYDOWN As Long = &H100&
Private Const WM_KEYUP As Long = &H101&
Private Const WM_CHAR As Long = &H102&
Private Const WM_MOUSEMOVE As Long = &H200&
Private Const WM_LBUTTONDOWN As Long = &H201&
Private Const WM_LBUTTONUP As Long = &H202&
Private Const WM_LBUTTONDBLCLK As Long = &H203&
Private Const WM_RBUTTONDOWN As Long = &H204&
Private Const WM_RBUTTONUP As Long = &H205&
Private Const WM_MOUSEWHEEL As Long = &H20A&
Private Const WM_CTLCOLOREDIT As Long = &H133&
Private Const WM_CTLCOLORSTATIC& = &H138&
Private Const WM_DROPFILES As Long = &H233&
Private Const WM_MOUSELEAVE& = 675&
Private Const WM_MOUSEHOVER& = &H2A1&
Private Const WM_DRAWITEM& = &H2B&
Private Const WM_SETREDRAW& = &HB&

Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_VSCROLL As Long = &H200000
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_BORDER As Long = &H800000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_GROUP As Long = &H20000
Private Const WS_POPUP As Long = &H80000000
Private Const WS_TABSTOP& = &H10000

Enum EX_STYLES
    WS_EX_DLGMODALFRAME = &H1&
    WS_EX_NOPARENTNOTIFY = &H4&
    WS_EX_TOPMOST = &H8&
    WS_EX_ACCEPTFILES = &H10&
    WS_EX_TRANSPARENT = &H20&
    WS_EX_MDICHILD = &H40&
    WS_EX_TOOLWINDOW = &H80&
    WS_EX_WINDOWEDGE = &H100&
    WS_EX_CLIENTEDGE = &H200&
    WS_EX_CONTEXTHELP = &H400&
    WS_EX_RIGHT = &H1000&
    WS_EX_LEFT = &H0&
    WS_EX_RTLREADING = &H2000&
    WS_EX_LTRREADING = &H0&
    WS_EX_LEFTSCROLLBAR = &H4000&
    WS_EX_RIGHTSCROLLBAR = &H0&
    WS_EX_CONTROLPARENT = &H10000
    WS_EX_STATICEDGE = &H20000
    WS_EX_APPWINDOW = &H40000
    WS_EX_OVERLAPPEDWINDOW = WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE
    WS_EX_PALETTEWINDOW = WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST
    WS_EX_LAYERED = &H80000
    WS_EX_NOINHERITLAYOUT = &H100000    ' Disable inheritence of mirroring by children
    WS_EX_NOREDIRECTIONBITMAP = &H200000
    WS_EX_LAYOUTRTL = &H400000          ' Right to left mirroring
    WS_EX_COMPOSITED = &H2000000
    WS_EX_NOACTIVATE = &H8000000
End Enum

Enum LISTBOX_STYLES
    eLB_SORT& = &H2&
    eLB_MULTIPLESEL& = &H8&
    eLB_OWNERDRAWFIXED& = &H10&
    eLB_OWNERDRAWVARIABLE& = &H20&
    eLB_NOINTEGRALHEIGHT& = &H100&
    eLB_MULTICOLUMN& = &H200&
    eLB_EXTENDEDSEL& = &H800&
End Enum

Private Const LBS_NOTIFY& = &H1&
Private Const LBS_SORT& = &H2&
Private Const LBS_NOREDRAW& = &H4&
Private Const LBS_MULTIPLESEL& = &H8&
Private Const LBS_OWNERDRAWFIXED& = &H10&
Private Const LBS_OWNERDRAWVARIABLE& = &H20&
Private Const LBS_HASSTRINGS& = &H40&
Private Const LBS_USETABSTOPS& = &H80&
Private Const LBS_NOINTEGRALHEIGHT& = &H100&
Private Const LBS_MULTICOLUMN& = &H200&
Private Const LBS_WANTKEYBOARDINPUT& = &H400&
Private Const LBS_EXTENDEDSEL& = &H800&
Private Const LBS_DISABLENOSCROLL& = &H1000&
Private Const LBS_NODATA& = &H2000&

Enum COMBO_STYLES
    eCB_SIMPLE& = &H1&
    eCB_DROPEDIT& = &H2&
    eCB_DROPSTATIC& = &H3&
    eCB_OWNERDRAWFIXED& = &H10&
    eCB_OWNERDRAWVARIABLE& = &H20&
    eCB_AUTOHSCROLL& = &H40&
    eCB_SORT& = &H100&
    eCB_NOINTEGRALHEIGHT& = &H400&
    eCB_DISABLENOSCROLL& = &H800&
End Enum

Private Const CBS_SIMPLE& = &H1&
Private Const CBS_DROPDOWN& = &H2&
Private Const CBS_DROPDOWNLIST& = &H3&
Private Const CBS_OWNERDRAWFIXED& = &H10&
Private Const CBS_OWNERDRAWVARIABLE& = &H20&
Private Const CBS_AUTOHSCROLL& = &H40&
Private Const CBS_OEMCONVERT& = &H80&
Private Const CBS_SORT& = &H100&
Private Const CBS_HASSTRINGS& = &H200&
Private Const CBS_NOINTEGRALHEIGHT& = &H400&
Private Const CBS_DISABLENOSCROLL& = &H800&

Enum STATIC_TYPE
    eTEXT_LEFT& = &H0&
    eTEXT_CENTER& = &H1&
    eTEXT_RIGHT& = &H2&
    eSS_ICON& = &H3&
    eCOLOR_BLACKRECT& = &H4&
    eCOLOR_GRAYRECT& = &H5&
    eCOLOR_WHITERECT& = &H6&
    eCOLOR_BLACKFRAME& = &H7&
    eCOLOR_GRAYFRAME& = &H8&
    eCOLOR_WHITEFRAME& = &H9&
    eTEXT_SIMPLE& = &HB&
    eTEXT_LEFTNOWORDWRAP& = &HC&
    eSS_OWNERDRAW& = &HD&
    eSS_BITMAP& = &HE&
    eSS_ENHMETAFILE& = &HF&
    eColor_ETCHEDHORZ& = &H10&
    eColor_ETCHEDVERT& = &H11&
    eColor_ETCHEDFRAME& = &H12&
End Enum

Enum STATIC_FLAGS
    eSS_RESIZEIMAGE& = &H40&
    eTEXT_NOPREFIX& = &H80&              ' Don't do "&" character translation */
    eSS_CENTERIMAGE& = &H200&
    eSS_RIGHTJUST& = &H400&
    eSS_SUNKEN& = &H1000&
    eTEXT_EDITCONTROL& = &H2000&
    eTEXT_ENDELLIPSIS& = &H4000&
    eTEXT_PATHELLIPSIS& = &H8000&
    eTEXT_WORDELLIPSIS& = &HC000&
End Enum

Private Const SS_LEFT& = &H0&
Private Const SS_CENTER& = &H1&
Private Const SS_RIGHT& = &H2&
Private Const SS_ICON& = &H3&
Private Const SS_BLACKRECT& = &H4&
Private Const SS_GRAYRECT& = &H5&
Private Const SS_WHITERECT& = &H6&
Private Const SS_BLACKFRAME& = &H7&
Private Const SS_GRAYFRAME& = &H8&
Private Const SS_WHITEFRAME& = &H9&
Private Const SS_USERITEM& = &HA&
Private Const SS_SIMPLE& = &HB&
Private Const SS_LEFTNOWORDWRAP& = &HC&
Private Const SS_OWNERDRAW& = &HD&
Private Const SS_BITMAP& = &HE&
Private Const SS_ENHMETAFILE& = &HF&
Private Const SS_ETCHEDHORZ& = &H10&
Private Const SS_ETCHEDVERT& = &H11&
Private Const SS_ETCHEDFRAME& = &H12&
Private Const SS_TYPEMASK& = &H1F&
Private Const SS_REALSIZECONTROL& = &H40&
Private Const SS_NOPREFIX& = &H80&              ' Don't do "&" character translation */
Private Const SS_NOTIFY& = &H100&
Private Const SS_CENTERIMAGE& = &H200&
Private Const SS_RIGHTJUST& = &H400&
Private Const SS_REALSIZEIMAGE& = &H800&
Private Const SS_SUNKEN& = &H1000&
Private Const SS_EDITCONTROL& = &H2000&
Private Const SS_ENDELLIPSIS& = &H4000&
Private Const SS_PATHELLIPSIS& = &H8000&
Private Const SS_WORDELLIPSIS& = &HC000&
Private Const SS_ELLIPSISMASK& = &HC000&

Enum BUTTON_STYLES
    eDEFPUSHBUTTON& = &H1&
    eBUTTON_RTL& = &H20&
    eBUTTON_TEXT& = &H0&
    eBUTTON_ICON& = &H40&
    eBUTTON_BITMAP& = &H80&
    eBUTTON_LEFT& = &H100&
    eBUTTON_RIGHT& = &H200&
    eBUTTON_CENTER& = &H300&
    eBUTTON_TOP& = &H400&
    eBUTTON_BOTTOM& = &H800&
    eBUTTON_VCENTER& = &HC00&
    eBUTTON_GRAPHICAL& = &H1000&
    eBUTTON_MULTILINE& = &H2000&
    eBUTTON_FLAT& = &H8000&
End Enum

Private Const BS_PUSHBUTTON& = &H0&
Private Const BS_DEFPUSHBUTTON& = &H1&
Private Const BS_CHECKBOX& = &H2&
Private Const BS_AUTOCHECKBOX& = &H3&
Private Const BS_RADIOBUTTON& = &H4&
Private Const BS_3STATE& = &H5&
Private Const BS_AUTO3STATE& = &H6&
Private Const BS_GROUPBOX& = &H7&
Private Const BS_USERBUTTON& = &H8&
Private Const BS_AUTORADIOBUTTON& = &H9&
Private Const BS_PUSHBOX& = &HA&
Private Const BS_OWNERDRAW& = &HB&
Private Const BS_TYPEMASK& = &HF&
Private Const BS_LEFTTEXT& = &H20&
Private Const BS_TEXT& = &H0&
Private Const BS_ICON& = &H40&
Private Const BS_BITMAP& = &H80&
Private Const BS_LEFT& = &H100&
Private Const BS_RIGHT& = &H200&
Private Const BS_CENTER& = &H300&
Private Const BS_TOP& = &H400&
Private Const BS_BOTTOM& = &H800&
Private Const BS_VCENTER& = &HC00&
Private Const BS_PUSHLIKE& = &H1000&
Private Const BS_MULTILINE& = &H2000&
Private Const BS_NOTIFY& = &H4000&
Private Const BS_FLAT& = &H8000&
Private Const BS_RIGHTBUTTON& = BS_LEFTTEXT

Enum EDIT_STYLES
    eEDIT_MULTILINE& = &H4&
    eEDIT_UPPERCASE& = &H8&
    eEDIT_LOWERCASE& = &H10&
    eEDIT_NOHIDESEL& = &H100&
    eEDIT_LOCKED& = &H800&
    eEDIT_WANTRETURN& = &H1000&
    eEDIT_NUMBER& = &H2000&
End Enum

Private Const ES_LEFT& = &H0&
Private Const ES_CENTER& = &H1&
Private Const ES_RIGHT& = &H2&
Private Const ES_MULTILINE& = &H4&
Private Const ES_UPPERCASE& = &H8&
Private Const ES_LOWERCASE& = &H10&
Private Const ES_PASSWORD& = &H20&
Private Const ES_AUTOVSCROLL& = &H40&
Private Const ES_AUTOHSCROLL& = &H80&
Private Const ES_NOHIDESEL& = &H100&
Private Const ES_OEMCONVERT& = &H400&
Private Const ES_READONLY& = &H800&
Private Const ES_WANTRETURN& = &H1000&
Private Const ES_NUMBER& = &H2000&

Private Const LB_ADDSTRING& = &H180&
Private Const LB_INSERTSTRING& = &H181&
Private Const LB_DELETESTRING& = &H182&
Private Const LB_SELITEMRANGEEX& = &H183&
Private Const LB_RESETCONTENT& = &H184&
Private Const LB_SETSEL& = &H185&
Private Const LB_SETCURSEL& = &H186&
Private Const LB_GETSEL& = &H187&
Private Const LB_GETCURSEL& = &H188&
Private Const LB_GETTEXT& = &H189&
Private Const LB_GETTEXTLEN& = &H18A&
Private Const LB_GETCOUNT& = &H18B&
Private Const LB_SELECTSTRING& = &H18C&
Private Const LB_DIR& = &H18D&
Private Const LB_GETTOPINDEX& = &H18E&
Private Const LB_FINDSTRING& = &H18F&
Private Const LB_GETSELCOUNT& = &H190&
Private Const LB_GETSELITEMS& = &H191&
Private Const LB_SETTABSTOPS& = &H192&
Private Const LB_GETHORIZONTALEXTENT& = &H193&
Private Const LB_SETHORIZONTALEXTENT& = &H194&
Private Const LB_SETCOLUMNWIDTH& = &H195&
Private Const LB_ADDFILE& = &H196&
Private Const LB_SETTOPINDEX& = &H197&
Private Const LB_GETITEMRECT& = &H198&
Private Const LB_GETITEMDATA& = &H199&
Private Const LB_SETITEMDATA& = &H19A&
Private Const LB_SELITEMRANGE& = &H19B&
Private Const LB_SETANCHORINDEX& = &H19C&
Private Const LB_GETANCHORINDEX& = &H19D&
Private Const LB_SETCARETINDEX& = &H19E&
Private Const LB_GETCARETINDEX& = &H19F&
Private Const LB_SETITEMHEIGHT& = &H1A0&
Private Const LB_GETITEMHEIGHT& = &H1A1&
Private Const LB_FINDSTRINGEXACT& = &H1A2&
Private Const LB_SETLOCALE& = &H1A5&
Private Const LB_GETLOCALE& = &H1A6&
Private Const LB_SETCOUNT& = &H1A7&
Private Const LB_INITSTORAGE& = &H1A8&
Private Const LB_ITEMFROMPOINT& = &H1A9&
Private Const LB_MULTIPLEADDSTRING& = &H1B1&

Private Const CB_GETEDITSEL& = &H140&
Private Const CB_LIMITTEXT& = &H141&
Private Const CB_SETEDITSEL& = &H142&
Private Const CB_ADDSTRING& = &H143&
Private Const CB_DELETESTRING& = &H144&
Private Const CB_DIR& = &H145&
Private Const CB_GETCOUNT& = &H146&
Private Const CB_GETCURSEL& = &H147&
Private Const CB_GETLBTEXT& = &H148&
Private Const CB_GETLBTEXTLEN& = &H149&
Private Const CB_INSERTSTRING& = &H14A&
Private Const CB_RESETCONTENT& = &H14B&
Private Const CB_FINDSTRING& = &H14C&
Private Const CB_SELECTSTRING& = &H14D&
Private Const CB_SETCURSEL& = &H14E&
Private Const CB_SHOWDROPDOWN& = &H14F&
Private Const CB_GETITEMDATA& = &H150&
Private Const CB_SETITEMDATA& = &H151&
Private Const CB_GETDROPPEDCONTROLRECT& = &H152&
Private Const CB_SETITEMHEIGHT& = &H153&
Private Const CB_GETITEMHEIGHT& = &H154&
Private Const CB_SETEXTENDEDUI& = &H155&
Private Const CB_GETEXTENDEDUI& = &H156&
Private Const CB_GETDROPPEDSTATE& = &H157&
Private Const CB_FINDSTRINGEXACT& = &H158&
Private Const CB_SETLOCALE& = &H159&
Private Const CB_GETLOCALE& = &H15A&
Private Const CB_GETTOPINDEX& = &H15B&
Private Const CB_SETTOPINDEX& = &H15C&
Private Const CB_GETHORIZONTALEXTENT& = &H15D&
Private Const CB_SETHORIZONTALEXTENT& = &H15E&
Private Const CB_GETDROPPEDWIDTH& = &H15F&
Private Const CB_SETDROPPEDWIDTH& = &H160&
Private Const CB_INITSTORAGE& = &H161&
Private Const CB_MULTIPLEADDSTRING& = &H163&
Private Const CB_GETCOMBOBOXINFO& = &H164&

Private Const EM_GETSEL& = &HB0
Private Const EM_SETSEL& = &HB1
Private Const EM_GETRECT& = &HB2
Private Const EM_SETRECT& = &HB3
Private Const EM_SETRECTNP& = &HB4
Private Const EM_SCROLL& = &HB5
Private Const EM_LINESCROLL& = &HB6
Private Const EM_SCROLLCARET& = &HB7
Private Const EM_GETMODIFY& = &HB8
Private Const EM_SETMODIFY& = &HB9
Private Const EM_GETLINECOUNT& = &HBA
Private Const EM_LINEINDEX& = &HBB
Private Const EM_SETHANDLE& = &HBC
Private Const EM_GETHANDLE& = &HBD
Private Const EM_GETTHUMB& = &HBE
Private Const EM_LINELENGTH& = &HC1
Private Const EM_REPLACESEL& = &HC2
Private Const EM_GETLINE& = &HC4
Private Const EM_LIMITTEXT& = &HC5
Private Const EM_CANUNDO& = &HC6
Private Const EM_UNDO& = &HC7
Private Const EM_FMTLINES& = &HC8
Private Const EM_LINEFROMCHAR& = &HC9
Private Const EM_SETTABSTOPS& = &HCB
Private Const EM_SETPASSWORDCHAR& = &HCC
Private Const EM_EMPTYUNDOBUFFER& = &HCD
Private Const EM_GETFIRSTVISIBLELINE& = &HCE
Private Const EM_SETREADONLY& = &HCF
Private Const EM_SETWORDBREAKPROC& = &HD0
Private Const EM_GETWORDBREAKPROC& = &HD1
Private Const EM_GETPASSWORDCHAR& = &HD2
Private Const EM_SETMARGINS& = &HD3
Private Const EM_GETMARGINS& = &HD4
Private Const EM_SETLIMITTEXT& = EM_LIMITTEXT   ' ;win40 Name change */
Private Const EM_GETLIMITTEXT& = &HD5
Private Const EM_POSFROMCHAR& = &HD6
Private Const EM_CHARFROMPOS& = &HD7

Private Const BM_GETCHECK& = &HF0&
Private Const BM_SETCHECK& = &HF1&
Private Const BM_GETSTATE& = &HF2&
Private Const BM_SETSTATE& = &HF3&
Private Const BM_SETSTYLE& = &HF4&
Private Const BM_CLICK& = &HF5&
Private Const BM_GETIMAGE& = &HF6&
Private Const BM_SETIMAGE& = &HF7&

Private Const EN_SETFOCUS& = &H100&
Private Const EN_KILLFOCUS& = &H200&
Private Const EN_CHANGE& = &H300&
Private Const EN_UPDATE& = &H400&
Private Const EN_ERRSPACE& = &H500&
Private Const EN_MAXTEXT& = &H501&
Private Const EN_HSCROLL& = &H601&
Private Const EN_VSCROLL& = &H602&
Private Const EN_ALIGN_LTR_EC& = &H700&
Private Const EN_ALIGN_RTL_EC& = &H701&
Private Const EN_BEFORE_PASTE& = &H800&
Private Const EN_AFTER_PASTE& = &H801&

Private Const BN_CLICKED& = 0&
Private Const BN_PAINT& = 1&
Private Const BN_HILITE& = 2&
Private Const BN_UNHILITE& = 3&
Private Const BN_DISABLE& = 4&
Private Const BN_DOUBLECLICKED& = 5&
Private Const BN_PUSHED& = BN_HILITE
Private Const BN_UNPUSHED& = BN_UNHILITE
Private Const BN_DBLCLK& = BN_DOUBLECLICKED
Private Const BN_SETFOCUS& = 6&
Private Const BN_KILLFOCUS& = 7&

Private Const LBN_SELCHANGE& = 1&
Private Const LBN_DBLCLK& = 2&
Private Const LBN_SELCANCEL& = 3&
Private Const LBN_SETFOCUS& = 4&
Private Const LBN_KILLFOCUS& = 5&

Private Const CBN_SELCHANGE& = 1&
Private Const CBN_DBLCLK& = 2&
Private Const CBN_SETFOCUS& = 3&
Private Const CBN_KILLFOCUS& = 4&
Private Const CBN_EDITCHANGE& = 5&
Private Const CBN_EDITUPDATE& = 6&
Private Const CBN_DROPDOWN& = 7&
Private Const CBN_CLOSEUP& = 8&
Private Const CBN_SELENDOK& = 9&
Private Const CBN_SELENDCANCEL& = 10&

Private Const DC_Pen As Long = 19&
Private Const DC_BRUSH As Long = 18&
Private Const HOLLOW_BRUSH& = 5&
Private Const HOLLOW_PEN& = 8&
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_CENTER As Long = &H1
Private Const DT_RIGHT As Long = &H2
Private Const DT_CALCRECT As Long = &H400&
Private Const DT_RTLREADING As Long = &H20000

Private Const SWP_NOACTIVATE& = &H10&
Private Const SWP_NOOWNERZORDER& = &H200&
Private Const SWP_NOMOVE& = &H2&
Private Const SWP_NOSIZE& = &H1&
Private Const SWP_NOZORDER& = &H4&
Private Const SWP_FRAMECHANGED& = &H20&
Private Const SWP_SHOWWINDOW& = &H40&

Private Const LOCALE_USER_DEFAULT& = &H400&
Private Const SORT_DIGITSASNUMBERS& = &H8&
Private Const NORM_IGNORECASE& = &H1&

Sub Init(frm As Form)
Dim n As Long
Dim i As Long
Dim c
Dim clsName As String
Dim actl() As Object
Dim cnt As Long

m_hwndParent = frm.hWnd
cnt = frm.Controls.Count
ReDim actl(0 To cnt)
For Each c In frm.Controls
    clsName = ""
    On Error Resume Next
    clsName = LCase(c.LinkItem)
    If clsName <> "" Then n = n + 1: Set actl(c.TabIndex) = c
Next
pvSubclass m_pSubclassForm, frm.hWnd
m_CtlCount = n
If n = 0 Then Exit Sub
ReDim m_Controls(0 To n - 1)
n = 0
For i = 0 To cnt
    If ObjPtr(actl(i)) Then
        pvAddControl actl(i), n
        n = n + 1
    End If
Next
End Sub

Sub AddControl(ctl As Control)
m_CtlCount = m_CtlCount + 1
ReDim Preserve m_Controls(0 To m_CtlCount - 1)
pvAddControl ctl, m_CtlCount - 1
End Sub

Private Sub pvAddControl(ctl As Control, ByVal ix As Long)
Dim c As Control
Dim n As Long
Dim p As Long
Dim win As Long
Dim clsName As String, s As String
Dim l&, t&, w&, h&, nStyle&, exStyle&, Style2&
Dim flags As String
Dim fnt As IFont
Dim sm As Long
Dim Container As Object
Dim hPar As Long
Dim combo As tCOMBOBOXINFO
Dim iIndex As Long

combo.bSize = Len(combo)
Set c = ctl
n = ix
clsName = LCase(c.LinkItem)
flags = ""
p = InStr(1, clsName, "-")
If p > 0 Then flags = Mid(clsName, p + 1): clsName = VBA.left(clsName, p - 1)
s = c.Caption
iIndex = -1
On Error Resume Next
iIndex = c.Index
sm = -1
Set Container = c.Container
Do
    On Error Resume Next
    sm = Container.ScaleMode
    If sm <> -1 Then Exit Do
    Set Container = Container.Container
Loop
l = Container.ScaleX(c.left, sm, vbPixels)
t = Container.ScaleY(c.top, sm, vbPixels)
w = Container.ScaleX(c.Width, sm, vbPixels)
h = Container.ScaleY(c.Height, sm, vbPixels)
hPar = Container.hWnd
nStyle = 0
exStyle = 0
Style2 = 0
If c.Visible Then nStyle = nStyle Or WS_VISIBLE
If c.Appearance = 1 Then exStyle = exStyle Or WS_EX_CLIENTEDGE
If c.BorderStyle = 1 And (exStyle And WS_EX_CLIENTEDGE) = 0 Then nStyle = nStyle Or WS_BORDER
If c.RightToLeft Then exStyle = exStyle Or WS_EX_RTLREADING
nStyle = nStyle Or WS_TABSTOP
Select Case clsName
Case "static": nStyle = nStyle Or SS_NOTIFY
               If c.Alignment = 1 Then nStyle = nStyle Or SS_RIGHT
               If c.Alignment = 2 Then nStyle = nStyle Or SS_CENTER
               If InStr(1, flags, "i") Then nStyle = nStyle Or eSS_ICON
               If InStr(1, flags, "b") Then nStyle = nStyle Or eSS_BITMAP
               If InStr(1, flags, "o") Then nStyle = nStyle Or eSS_OWNERDRAW
               RaiseEvent CreateStatic(c.Name, iIndex, hPar, Style2, nStyle, exStyle)
               nStyle = nStyle Or Style2
Case "edit": If c.Alignment = 1 Then nStyle = nStyle Or ES_RIGHT: exStyle = exStyle Or WS_EX_RIGHT
             If c.Alignment = 2 Then nStyle = nStyle Or ES_CENTER
             If c.RightToLeft Then exStyle = exStyle Or WS_EX_RTLREADING Or WS_EX_LEFTSCROLLBAR
             nStyle = nStyle Or ES_AUTOHSCROLL Or ES_AUTOVSCROLL Or ES_WANTRETURN
             If InStr(1, flags, "l") Then nStyle = nStyle Or eEDIT_LOCKED
             If InStr(1, flags, "m") Then nStyle = nStyle Or eEDIT_MULTILINE
             If InStr(1, flags, "n") Then nStyle = nStyle Or eEDIT_NUMBER
             RaiseEvent CreateEdit(c.Name, iIndex, hPar, nStyle, exStyle)
Case "button": nStyle = nStyle Or BS_NOTIFY Or BS_MULTILINE
             If InStr(1, flags, "o") Then nStyle = nStyle Or BS_OWNERDRAW
             RaiseEvent CreateButton(c.Name, iIndex, hPar, nStyle, exStyle)
Case "frame": clsName = "button": nStyle = nStyle Or BS_NOTIFY Or BS_GROUPBOX Or BS_MULTILINE
             If c.Alignment = 1 Then nStyle = nStyle Or BS_RIGHTBUTTON Or BS_RIGHT
             RaiseEvent CreateFrame(c.Name, iIndex, hPar, nStyle, exStyle)
Case "listbox": nStyle = nStyle Or LBS_NOTIFY Or WS_VSCROLL Or LBS_HASSTRINGS Or LBS_NOINTEGRALHEIGHT
            If c.Alignment = 1 Then exStyle = exStyle Or EX_STYLES.WS_EX_RIGHT
            If c.RightToLeft Then exStyle = exStyle Or WS_EX_RTLREADING Or WS_EX_LEFTSCROLLBAR
            If InStr(1, flags, "s") Then nStyle = nStyle Or eLB_SORT
            If InStr(1, flags, "m") Then nStyle = nStyle Or eLB_MULTIPLESEL
            If InStr(1, flags, "c") Then nStyle = nStyle Or eLB_MULTICOLUMN
            If InStr(1, flags, "e") Then nStyle = nStyle Or eLB_EXTENDEDSEL
            If InStr(1, flags, "o") Then nStyle = nStyle Or eLB_OWNERDRAWFIXED
            If InStr(1, flags, "v") Then nStyle = nStyle Or eLB_OWNERDRAWVARIABLE
            RaiseEvent CreateListBox(c.Name, iIndex, hPar, nStyle, exStyle)
Case "combobox": nStyle = nStyle Or WS_VSCROLL Or CBS_HASSTRINGS Or CBS_NOINTEGRALHEIGHT
            If c.Alignment = 1 Then exStyle = exStyle Or EX_STYLES.WS_EX_RIGHT
            If c.RightToLeft Then exStyle = exStyle Or WS_EX_RTLREADING Or WS_EX_LEFTSCROLLBAR
            If InStr(1, flags, "s") Then nStyle = nStyle Or eCB_SORT
            If InStr(1, flags, "o") Then nStyle = nStyle Or eCB_OWNERDRAWFIXED
            If InStr(1, flags, "v") Then nStyle = nStyle Or eCB_OWNERDRAWVARIABLE
            If InStr(1, flags, "a") Then nStyle = nStyle Or eCB_AUTOHSCROLL
            If InStr(1, flags, "c") Then nStyle = nStyle Or eCB_DROPEDIT
            If InStr(1, flags, "d") Then nStyle = nStyle Or eCB_DROPSTATIC
            If InStr(1, flags, "p") Then nStyle = nStyle Or eCB_SIMPLE
            RaiseEvent CreateComboBox(c.Name, iIndex, hPar, nStyle, exStyle)
            If (nStyle And 3) = 0 Then nStyle = nStyle Or CBS_DROPDOWN
Case "checkbox": clsName = "button": nStyle = nStyle Or BS_AUTOCHECKBOX Or BS_MULTILINE Or BS_NOTIFY
            If c.Alignment = 1 Then nStyle = nStyle Or BS_RIGHTBUTTON Or BS_RIGHT
            If InStr(1, flags, "g") Then nStyle = nStyle Or eBUTTON_GRAPHICAL
            If InStr(1, flags, "o") Then nStyle = nStyle Or BS_OWNERDRAW
            RaiseEvent CreateCheckBox(c.Name, iIndex, hPar, nStyle, exStyle)
Case "3state": clsName = "button": nStyle = nStyle Or BS_AUTO3STATE Or BS_MULTILINE Or BS_NOTIFY
            If c.Alignment = 1 Then nStyle = nStyle Or BS_RIGHTBUTTON Or BS_RIGHT
            If InStr(1, flags, "g") Then nStyle = nStyle Or eBUTTON_GRAPHICAL
            If InStr(1, flags, "o") Then nStyle = nStyle Or BS_OWNERDRAW
            RaiseEvent Create3State(c.Name, iIndex, hPar, nStyle, exStyle)
Case "radiobutton": clsName = "button": nStyle = nStyle Or BS_AUTORADIOBUTTON Or BS_MULTILINE Or BS_NOTIFY
            If c.Alignment = 1 Then nStyle = nStyle Or BS_RIGHTBUTTON Or BS_RIGHT
            If InStr(1, flags, "g") Then nStyle = nStyle Or eBUTTON_GRAPHICAL
            If InStr(1, flags, "o") Then nStyle = nStyle Or BS_OWNERDRAW
            RaiseEvent CreateRadioButton(c.Name, iIndex, hPar, nStyle, exStyle)
Case "radiobuttongroup": clsName = "button": nStyle = nStyle Or BS_AUTORADIOBUTTON Or WS_GROUP Or BS_MULTILINE Or BS_NOTIFY
            If c.Alignment = 1 Then nStyle = nStyle Or BS_RIGHTBUTTON Or BS_RIGHT
            If InStr(1, flags, "g") Then nStyle = nStyle Or eBUTTON_GRAPHICAL
            If InStr(1, flags, "o") Then nStyle = nStyle Or BS_OWNERDRAW
            RaiseEvent CreateRadioButton(c.Name, iIndex, hPar, nStyle, exStyle)
Case Else:
    RaiseEvent Create(c.Name, iIndex, hPar, clsName, nStyle, exStyle)
End Select
nStyle = nStyle Or WS_CHILD
win = CreateWindowExW(exStyle, StrPtr(clsName), StrPtr(s), nStyle, l, t, w, h, hPar, n + 1, App.hInstance, 0)
If clsName = "combobox" Then
    GetComboBoxInfo win, VarPtr(combo)
    m_Controls(n).hList = combo.hwndList
    m_Controls(n).hEdit = combo.hwndItem
    If combo.hwndList Then
        SetWindowLongW combo.hwndList, -12&, n + 1
        'SetProp combo.hwndList, "_ID_", n + 1
        'pvSubclassItem m_Controls(n).m_SubclassList, combo.hwndList
    End If
    If combo.hwndItem Then
        SetWindowLongW combo.hwndItem, -12&, n + 1
        'SetProp combo.hwndItem, "_ID_", n + 1
        pvSubclassItem m_Controls(n).m_SubclassItem, combo.hwndItem
    End If
End If
Set fnt = c.Font
SendMessageW win, WM_SETFONT, fnt.hFont, 1
m_Controls(n).hWnd = win
m_Controls(n).sName = c.Name
m_Controls(n).sClass = clsName
m_Controls(n).BackColor = FixColor(c.BackColor)
m_Controls(n).ForeColor = FixColor(c.ForeColor)
m_Controls(n).BackStyle = c.BackStyle + 1
m_Controls(n).hCursor = -1
m_Controls(n).ToolTipText = c.ToolTipText
If Len(c.Tag) Then m_Controls(n).Tag = c.Tag
pvSubclass2 m_Controls(n).m_Subclass, win
m_Controls(n).iIndex = iIndex
c.Visible = False
c.LinkTimeout = n + 1
End Sub

Property Let NoRedraw(ctl As Object, b As Boolean)
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
SendMessageW win, WM_SETREDRAW, IIf(b, 0, 1), 0
End Property

Property Get Tag(ctl As Object)
Dim ix As Long
ix = GetIx(ctl): If ix < 0 Then Exit Property
Tag = m_Controls(ix).Tag
End Property

Property Let Tag(ctl As Object, v)
Dim ix As Long
ix = GetIx(ctl): If ix < 0 Then Exit Property
m_Controls(ix).Tag = v
End Property

Property Get CueBanner(ctl As Object)
Dim ix As Long
ix = GetIx(ctl): If ix < 0 Then Exit Property
CueBanner = m_Controls(ix).CueBanner
End Property

Property Let CueBanner(ctl As Object, v)
Dim ix As Long
ix = GetIx(ctl): If ix < 0 Then Exit Property
m_Controls(ix).CueBanner = v
Refresh ctl
End Property

Property Get ToolTipText(ctl As Object)
Dim ix As Long
ix = GetIx(ctl): If ix < 0 Then Exit Property
ToolTipText = m_Controls(ix).ToolTipText
End Property

Property Let ToolTipText(ctl As Object, s)
Dim ix As Long
ix = GetIx(ctl): If ix < 0 Then Exit Property
m_Controls(ix).ToolTipText = s
If ix + 1 = m_idTTT Then pvShowToolTip ix
End Property

Property Let ToolTipFlags(ctl As Object, f As eToolTipFlags)
Dim ix As Long
Dim flag As Long
ix = GetIx(ctl): If ix < 0 Then Exit Property
'flag = f And 3
'If f And eTTFRTL Then flag = flag Or DT_RTLREADING
m_Controls(ix).ToolTipFlag = f
End Property

Property Get Text(ctl As Object)
Dim win As Long
Dim s As String
Dim l As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
l = GetWindowTextLength(win)
If l = 0 Then Exit Property
s = AllocStr(l)
GetWindowTextW win, StrPtr(s), l + 1
Text = s
End Property

Property Let Text(ctl As Object, s)
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
SetWindowTextW win, StrPtr(s)
End Property

Sub AddItem(ctl As Object, s, Optional ByVal iPos As Long = -1, Optional ByVal nItemData As Long)
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Sub
ix = GetIx(ctl): If ix < 0 Then Exit Sub
If m_Controls(ix).sClass = "listbox" Then
    iPos = SendMessageW(win, LB_INSERTSTRING, iPos, StrPtr(s))
    If nItemData Then SendMessageW win, LB_SETITEMDATA, iPos, nItemData
ElseIf m_Controls(ix).sClass = "combobox" Then
    iPos = SendMessageW(win, CB_INSERTSTRING, iPos, StrPtr(s))
    If nItemData Then SendMessageW win, CB_SETITEMDATA, iPos, nItemData
End If
End Sub

Sub RemoveItem(ctl As Object, ByVal iPos As Long)
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Sub
ix = GetIx(ctl): If ix < 0 Then Exit Sub
If m_Controls(ix).sClass = "listbox" Then
    Call SendMessageW(win, LB_DELETESTRING, iPos, 0)
ElseIf m_Controls(ix).sClass = "combobox" Then
    SendMessageW win, CB_DELETESTRING, iPos, 0
End If
End Sub

Property Get ItemData(ctl As Object, ByVal n As Long) As Long
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "listbox" Then
    ItemData = SendMessageW(win, LB_GETITEMDATA, n, 0)
ElseIf m_Controls(ix).sClass = "combobox" Then
    ItemData = SendMessageW(win, CB_GETITEMDATA, n, 0)
End If
End Property

Property Let ItemData(ctl As Object, ByVal n As Long, ByVal nData As Long)
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "listbox" Then
    Call SendMessageW(win, LB_SETITEMDATA, n, nData)
ElseIf m_Controls(ix).sClass = "combobox" Then
    Call SendMessageW(win, CB_SETITEMDATA, n, nData)
End If
End Property

Property Get List(ctl As Object, ByVal n As Long)
Dim win As Long
Dim ix As Long
Dim s As String
Dim l As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "listbox" Then
    l = SendMessageW(win, LB_GETTEXTLEN, n, 0)
    If l <= 0 Then Exit Property
    s = AllocStr(l)
    Call SendMessageW(win, LB_GETTEXT, n, StrPtr(s))
    List = s
ElseIf m_Controls(ix).sClass = "combobox" Then
    l = SendMessageW(win, CB_GETLBTEXTLEN, n, 0)
    If l <= 0 Then Exit Property
    s = AllocStr(l)
    Call SendMessageW(win, CB_GETLBTEXT, n, StrPtr(s))
    List = s
End If
End Property

Property Let List(ctl As Object, ByVal n As Long, s)
Dim win As Long
Dim ix As Long, nData As Long
Dim cur As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "listbox" Then
    cur = SendMessageW(win, LB_GETCURSEL, 0, 0)
    nData = SendMessageW(win, LB_GETITEMDATA, n, 0)
    SendMessageW win, LB_DELETESTRING, n, 0
    SendMessageW win, LB_INSERTSTRING, n, StrPtr(s)
    SendMessageW win, LB_SETITEMDATA, n, nData
    SendMessageW win, LB_SETCURSEL, cur, 0
ElseIf m_Controls(ix).sClass = "combobox" Then
    cur = SendMessageW(win, CB_GETCURSEL, 0, 0)
    nData = SendMessageW(win, CB_GETITEMDATA, n, 0)
    SendMessageW win, CB_DELETESTRING, n, 0
    SendMessageW win, CB_INSERTSTRING, n, StrPtr(s)
    SendMessageW win, CB_SETITEMDATA, n, nData
    SendMessageW win, CB_SETCURSEL, cur, 0
End If
End Property

Property Get ListCount(ctl As Object) As Long
Dim win As Long
Dim ix As Long, nData As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "listbox" Then
    ListCount = SendMessageW(win, LB_GETCOUNT, 0, 0)
ElseIf m_Controls(ix).sClass = "combobox" Then
    ListCount = SendMessageW(win, CB_GETCOUNT, 0, 0)
End If
End Property

Property Get ListIndex(ctl As Object) As Long
Dim win As Long
Dim ix As Long, nData As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "listbox" Then
   ListIndex = SendMessageW(win, LB_GETCURSEL, 0, 0)
ElseIf m_Controls(ix).sClass = "combobox" Then
    ListIndex = SendMessageW(win, CB_GETCURSEL, 0, 0)
End If
End Property

Property Let ListIndex(ctl As Object, ByVal n As Long)
Dim win As Long
Dim ix As Long, nData As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "listbox" Then
   SendMessageW win, LB_SETCURSEL, n, 0
ElseIf m_Controls(ix).sClass = "combobox" Then
    SendMessageW win, CB_SETCURSEL, n, 0
End If
End Property

Property Get TopIndex(ctl As Object) As Long
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "listbox" Then
   TopIndex = SendMessageW(win, LB_GETTOPINDEX, 0, 0)
ElseIf m_Controls(ix).sClass = "combobox" Then
    TopIndex = SendMessageW(win, CB_GETTOPINDEX, 0, 0)
End If
End Property

Property Let TopIndex(ctl As Object, ByVal n As Long)
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "listbox" Then
   Call SendMessageW(win, LB_SETTOPINDEX, n, 0)
ElseIf m_Controls(ix).sClass = "combobox" Then
    Call SendMessageW(win, CB_SETTOPINDEX, n, 0)
End If
End Property

Property Let DropWidth(ctl As Object, ByVal n As Long)
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "listbox" Then
   'Call SendMessageW(win, LB_SETTOPINDEX, n, 0)
ElseIf m_Controls(ix).sClass = "combobox" Then
    Call SendMessageW(win, CB_SETDROPPEDWIDTH, n, 0)
End If
End Property

Property Let ListWidth(ctl As Object, ByVal n As Long)
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "listbox" Then
   Call SendMessageW(win, LB_SETHORIZONTALEXTENT, n, 0)
ElseIf m_Controls(ix).sClass = "combobox" Then
    Call SendMessageW(win, CB_SETHORIZONTALEXTENT, n, 0)
End If
End Property

Property Let ItemHeight(ctl As Object, ByVal n As Long)
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "listbox" Then
   Call SendMessageW(win, LB_SETITEMHEIGHT, n, 0)
ElseIf m_Controls(ix).sClass = "combobox" Then
    Call SendMessageW(win, CB_SETITEMHEIGHT, n, 0)
End If
End Property

Property Get ItemHeight(ctl As Object) As Long
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "listbox" Then
   ItemHeight = SendMessageW(win, LB_GETITEMHEIGHT, 0, 0)
ElseIf m_Controls(ix).sClass = "combobox" Then
    ItemHeight = SendMessageW(win, CB_GETITEMHEIGHT, 0, 0)
End If
End Property

Sub ClearList(ctl As Object)
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Sub
ix = GetIx(ctl): If ix < 0 Then Exit Sub
If m_Controls(ix).sClass = "listbox" Then
   Call SendMessageW(win, LB_RESETCONTENT, 0, 0)
ElseIf m_Controls(ix).sClass = "combobox" Then
    Call SendMessageW(win, CB_RESETCONTENT, 0, 0)
End If
End Sub

Public Function FindItem(ctl As Object, pattern, ByVal iStart As Long) As Long
Dim i As Long
Dim s As String
Dim cnt As Long
FindItem = -1
cnt = ListCount(ctl)
If cnt = 0 Then Exit Function
s = UCase(pattern)
i = iStart
Do
    If UCase(List(ctl, i)) Like s Then FindItem = i: Exit Function
    i = i + 1
    If i >= cnt Then i = 0
    If i = iStart Then Exit Do
Loop
End Function

Private Function EscapeLikePattern(ByVal Text As String) As String
    Dim i As Integer
    Dim Result As String
    Dim char As String
    
    ' Loop through each character and wrap special Like characters in []
    For i = 1 To Len(Text)
        char = Mid(Text, i, 1)
        Select Case char
            Case "*", "#", "?", "["
                Result = Result & "[" & char & "]"
            Case Else
                Result = Result & char
        End Select
    Next i
    
    EscapeLikePattern = Result
End Function

Public Function MouseInWindow(ctl As Object) As Boolean
Dim pt(0 To 1) As Long
Dim rc As RECTL
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Function
GetCursorPos VarPtr(pt(0))
GetWindowRect win, rc
MouseInWindow = PtInRect(VarPtr(rc), pt(0), pt(1)) <> 0
End Function

Public Property Get VisibleItems(ctl As Object) As Long
VisibleItems = ClientHeight(ctl) \ ItemHeight(ctl)
End Property

Property Get IsItemVisible(ctl As Object, ByVal n As Long) As Boolean
Dim t As Long
If n < 0 Then Exit Property
If n >= ListCount(ctl) Then Exit Property
t = TopIndex(ctl)
If n < t Then Exit Property
IsItemVisible = n < t + VisibleItems(ctl)
End Property

Property Get MouseRow(ctl As Object) As Long
Dim win As Long
Dim ix As Long
Dim n As Long
Dim rc As RECTL
MouseRow = -1&
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "combobox" Then win = m_Controls(ix).hList: If win = 0 Then Exit Property
Dim ar(0 To 1) As Long
GetCursorPos VarPtr(ar(0))
ScreenToClient win, VarPtr(ar(0))
GetClientRect win, rc
If ar(0) < 0 Or ar(1) < 0 Then Exit Property
If ar(0) > rc.right Or ar(1) > rc.bottom Then Exit Property
n = TopIndex(ctl) + ar(1) \ ItemHeight(ctl)
If n >= 0 And n < ListCount(ctl) Then MouseRow = n
End Property

Property Get TextWidth(ctl As Object, txt, Optional ByRef nHeight As Long) As Long
Dim win As Long
Dim dc As Long
Dim rc As RECTL
Dim hFont As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
hFont = SendMessageW(win, WM_GETFONT, 0, 0)
dc = GetDC(0)
SelectObject dc, hFont
DrawText dc, StrPtr(txt), Len(txt), VarPtr(rc), DT_CALCRECT
ReleaseDC 0, dc
nHeight = rc.bottom - rc.top
TextWidth = rc.right - rc.left
End Property

Property Get TextHeight(ctl As Object, txt) As Long
Dim win As Long
Dim dc As Long
Dim rc As RECTL
Dim hFont As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
hFont = SendMessageW(win, WM_GETFONT, 0, 0)
dc = GetDC(0)
SelectObject dc, hFont
DrawText dc, StrPtr(txt), Len(txt), VarPtr(rc), DT_CALCRECT
ReleaseDC 0, dc
TextHeight = rc.bottom - rc.top
End Property

Property Get SelStart(ctl As Object) As Long
Dim win As Long
Dim ix As Long, nData As Long
Dim nStart As Long, nEnd As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "edit" Then
    Call SendMessageW(win, EM_GETSEL, VarPtr(nStart), VarPtr(nEnd))
    SelStart = nStart
ElseIf m_Controls(ix).sClass = "combobox" Then
    Call SendMessageW(win, CB_GETEDITSEL, VarPtr(nStart), VarPtr(nEnd))
    SelStart = nStart
End If
End Property

Property Get SelLen(ctl As Object) As Long
Dim win As Long
Dim ix As Long, nData As Long
Dim nStart As Long, nEnd As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "edit" Then
    Call SendMessageW(win, EM_GETSEL, VarPtr(nStart), VarPtr(nEnd))
    SelLen = nEnd - nStart
ElseIf m_Controls(ix).sClass = "combobox" Then
    Call SendMessageW(win, CB_GETEDITSEL, VarPtr(nStart), VarPtr(nEnd))
    SelLen = nEnd - nStart
End If
End Property

Sub SelSet(ctl As Object, ByVal nStart As Long, ByVal nLen As Long)
Dim win As Long
Dim ix As Long, nData As Long
Dim nEnd As Long
Dim n As Long
nStart = nStart - 1
If nLen < 0 Then nEnd = -1 Else nEnd = nStart + nLen
win = GetWindow(ctl): If win = 0 Then Exit Sub
ix = GetIx(ctl): If ix < 0 Then Exit Sub
If m_Controls(ix).sClass = "edit" Then
    SendMessageW win, EM_SETSEL, nStart, nEnd
Else
    CopyMemory VarPtr(n), VarPtr(nStart), 2
    CopyMemory VarPtr(n) + 2, VarPtr(nEnd), 2
    SendMessageW win, CB_SETEDITSEL, 0, n
End If
End Sub

Property Get SelText(ctl As Object)
Dim win As Long
Dim ix As Long, nData As Long
Dim nStart As Long, nEnd As Long
Dim s As String
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).sClass = "edit" Then
    Call SendMessageW(win, EM_GETSEL, VarPtr(nStart), VarPtr(nEnd))
    s = Text(ctl)
    If nEnd < 0 Then nEnd = Len(s)
    If nStart >= 0 And nEnd > nStart Then SelText = Mid(s, nStart + 1, nEnd - nStart)
ElseIf m_Controls(ix).sClass = "combobox" Then
    Call SendMessageW(win, CB_GETEDITSEL, VarPtr(nStart), VarPtr(nEnd))
    s = Text(ctl)
    If nEnd < 0 Then nEnd = Len(s)
    If nStart >= 0 And nEnd > nStart Then SelText = Mid(s, nStart + 1, nEnd - nStart)
End If
End Property

Property Let SelText(ctl As Object, txt)
Dim win As Long
Dim ix As Long, nData As Long
Dim nStart As Long, nEnd As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).hEdit Then win = m_Controls(ix).hEdit
Call SendMessageW(win, EM_REPLACESEL, 1, StrPtr(txt))
End Property

Function CurLine(ctl As Object, Optional ByVal charPos As Long = -1&) As Long
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Function
ix = GetIx(ctl): If ix < 0 Then Exit Function
CurLine = SendMessageW(win, EM_LINEFROMCHAR, charPos, 0)
End Function

Function LineStart(ctl As Object, Optional ByVal nLine As Long = -1&) As Long
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Function
ix = GetIx(ctl): If ix < 0 Then Exit Function
LineStart = SendMessageW(win, EM_LINEINDEX, nLine, 0)
End Function

Function LineLen(ctl As Object, Optional ByVal nLine As Long = -1&) As Long
Dim win As Long
Dim ix As Long
Dim nStart As Long
win = GetWindow(ctl): If win = 0 Then Exit Function
ix = GetIx(ctl): If ix < 0 Then Exit Function
nStart = LineStart(ctl, nLine)
If nStart < 0 Then LineLen = -1: Exit Function
LineLen = SendMessageW(win, EM_LINELENGTH, nStart, 0)
End Function

Function GetLine(ctl As Object, Optional ByVal nLine As Long = -1&) As String
Dim win As Long
Dim ix As Long
Dim s As String
Dim l As Long
win = GetWindow(ctl): If win = 0 Then Exit Function
ix = GetIx(ctl): If ix < 0 Then Exit Function
If nLine < 0 Then nLine = CurLine(ctl)
l = SendMessageW(win, EM_LINELENGTH, nLine, 0)
If l <= 0 Then Exit Function
s = AllocStr(l)
s = SendMessageW(win, EM_GETLINE, nLine, StrPtr(s))
GetLine = s
End Function

Function CharPoint(ctl As Object, ByVal charIndex As Long, X As Long, Y As Long) As Long
Dim win As Long
Dim ix As Long
Dim p As Long
Dim x0 As Integer, y0 As Integer
win = GetWindow(ctl): If win = 0 Then Exit Function
ix = GetIx(ctl): If ix < 0 Then Exit Function
p = SendMessageW(win, EM_POSFROMCHAR, charIndex, 0)
CopyMemory VarPtr(x0), VarPtr(p), 2
CopyMemory VarPtr(y0), VarPtr(p) + 2, 2
X = x0
Y = y0
CharPoint = p
End Function

Function CharAtPoint(ctl As Object, ByVal X As Long, ByVal Y As Long) As Long
Dim win As Long
Dim ix As Long
Dim p As Long
Dim n As Long
Dim low As Integer
win = GetWindow(ctl): If win = 0 Then Exit Function
ix = GetIx(ctl): If ix < 0 Then Exit Function
CopyMemory VarPtr(n), VarPtr(X), 2
CopyMemory VarPtr(n) + 2, VarPtr(Y), 2
p = SendMessageW(win, EM_CHARFROMPOS, 0, n)
CopyMemory VarPtr(low), VarPtr(p), 2
CharAtPoint = low
End Function

Property Get Visible(ctl As Object) As Boolean
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Function
Visible = IsWindowVisible(win) <> 0
End Property

Property Let Visible(ctl As Object, b As Boolean)
Dim win As Long
Dim n As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
If b Then n = 8 ' SW_SHOWNA
ShowWindow win, n
End Property

Property Get Enable(ctl As Object) As Boolean
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
Enable = IsWindowEnabled(win) <> 0
End Property

Property Let Enable(ctl As Object, b As Boolean)
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
EnableWindow win, IIf(b, 1, 0)
End Property

Property Get Checked(ctl As Object) As Long
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
Checked = SendMessageW(win, BM_GETCHECK, 0, 0)
End Property

Property Let Checked(ctl As Object, ByVal c As Long)
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
SendMessageW win, BM_SETCHECK, c, 0
End Property

Property Get left(ctl As Object) As Long
Dim win As Long
Dim rc As RECTL
Dim par As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
GetWindowRect win, rc
If TypeOf ctl Is Form Then left = rc.left: Exit Property
par = GetParent(win)
ScreenToClient par, VarPtr(rc)
left = rc.left
End Property

Property Let left(ctl As Object, ByVal c As Long)
Dim win As Long
Dim rc As RECTL
win = GetWindow(ctl): If win = 0 Then Exit Property
GetWindowRect win, rc
If TypeOf ctl Is Form Then rc.left = c Else ScreenToClient GetParent(win), VarPtr(rc): rc.left = c
SetWindowPos win, 0, rc.left, rc.top, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOZORDER
End Property

Property Get top(ctl As Object) As Long
Dim win As Long
Dim rc As RECTL
Dim par As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
GetWindowRect win, rc
If TypeOf ctl Is Form Then top = rc.top: Exit Property
par = GetParent(win)
ScreenToClient par, VarPtr(rc)
top = rc.top
End Property

Property Let top(ctl As Object, ByVal c As Long)
Dim win As Long
Dim rc As RECTL
win = GetWindow(ctl): If win = 0 Then Exit Property
GetWindowRect win, rc
If TypeOf ctl Is Form Then rc.top = c Else ScreenToClient GetParent(win), VarPtr(rc): rc.top = c
SetWindowPos win, 0, rc.left, rc.top, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOZORDER
End Property

Property Get Width(ctl As Object) As Long
Dim win As Long
Dim rc As RECTL
Dim par As Long
Dim w As Long
Dim h As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
GetWindowRect win, rc
w = rc.right - rc.left
h = rc.bottom - rc.top
Width = w
End Property

Property Get Height(ctl As Object) As Long
Dim win As Long
Dim rc As RECTL
Dim par As Long
Dim w As Long
Dim h As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
GetWindowRect win, rc
w = rc.right - rc.left
h = rc.bottom - rc.top
Height = h
End Property

Property Let Width(ctl As Object, ByVal c As Long)
Dim win As Long
Dim rc As RECTL
Dim w As Long
Dim h As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
GetWindowRect win, rc
w = rc.right - rc.left
h = rc.bottom - rc.top
SetWindowPos win, 0, 0, 0, c, h, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOZORDER
End Property

Property Let Height(ctl As Object, ByVal c As Long)
Dim win As Long
Dim rc As RECTL
Dim w As Long
Dim h As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
GetWindowRect win, rc
w = rc.right - rc.left
h = rc.bottom - rc.top
SetWindowPos win, 0, 0, 0, w, c, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOZORDER
End Property

Property Get ClientWidth(ctl As Object) As Long
Dim win As Long
Dim rc As RECTL
Dim par As Long
Dim w As Long
Dim h As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
GetClientRect win, rc
ClientWidth = rc.right
End Property

Property Get ClientHeight(ctl As Object) As Long
Dim win As Long
Dim rc As RECTL
Dim par As Long
Dim w As Long
Dim h As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
GetClientRect win, rc
ClientHeight = rc.bottom
End Property

Property Let ClientWidth(ctl As Object, ByVal c As Long)
Dim win As Long
Dim rc As RECTL
Dim rc2 As RECTL
Dim w As Long
Dim h As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
GetWindowRect win, rc
GetClientRect win, rc2
w = rc.right - rc.left
h = rc.bottom - rc.top
SetWindowPos win, 0, 0, 0, c + w - rc2.right, h, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOZORDER
End Property

Property Let ClientHeight(ctl As Object, ByVal c As Long)
Dim win As Long
Dim rc As RECTL
Dim rc2 As RECTL
Dim w As Long
Dim h As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
GetWindowRect win, rc
GetClientRect win, rc2
w = rc.right - rc.left
h = rc.bottom - rc.top
SetWindowPos win, 0, 0, 0, w, c + h - rc2.bottom, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOZORDER
End Property

Sub Move(ctl As Object, Optional nLeft, Optional nTop, Optional nWidth, Optional nHeight)
Dim win As Long
Dim l&, t&, w&, h&
Dim rc As RECTL
win = GetWindow(ctl): If win = 0 Then Exit Sub
GetWindowRect win, rc
w = rc.right - rc.left
h = rc.bottom - rc.top
If TypeOf ctl Is Form Then Else ScreenToClient GetParent(win), VarPtr(rc)
l = rc.left
t = rc.top
If Not IsMissing(nLeft) Then l = nLeft
If Not IsMissing(nTop) Then t = nTop
If Not IsMissing(nWidth) Then w = nWidth
If Not IsMissing(nHeight) Then h = nHeight
SetWindowPos win, 0, l, t, w, h, SWP_NOACTIVATE Or SWP_NOZORDER
End Sub

Property Get DropState(ctl As Object) As Boolean
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
DropState = SendMessageW(win, CB_GETDROPPEDSTATE, 0, 0) <> 0
End Property

Property Let DropState(ctl As Object, b As Boolean)
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
If b Then SendMessageW win, CB_SHOWDROPDOWN, 1, 0
If Not b Then SendMessageW win, CB_SHOWDROPDOWN, 0, 0
End Property

Sub SetPasswordChar(ctl As Object, Optional ByVal ch As Long = &H25CF&)
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Sub
ix = GetIx(ctl): If ix < 0 Then Exit Sub
If m_Controls(ix).hEdit Then win = m_Controls(ix).hEdit
AddStyle ctl, ES_PASSWORD
SendMessageW win, EM_SETPASSWORDCHAR, ch, 0
End Sub

Property Get BackColor(ctl As Object) As Long
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
BackColor = m_Controls(ix).BackColor
End Property

Property Let BackColor(ctl As Object, ByVal n As Long)
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
m_Controls(ix).BackColor = n
Refresh ctl
End Property

Property Get ForeColor(ctl As Object) As Long
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
ForeColor = m_Controls(ix).ForeColor
End Property

Property Let ForeColor(ctl As Object, ByVal n As Long)
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
m_Controls(ix).ForeColor = n
Refresh ctl
End Property

Property Get BackStyle(ctl As Object) As Long
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
BackStyle = m_Controls(ix).BackStyle
End Property

Property Let BackStyle(ctl As Object, ByVal n As Long)
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
m_Controls(ix).BackStyle = n
Refresh ctl
End Property

Property Let HideSelection(ctl As Object, b As Boolean)
Dim Style As Long
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).hEdit Then win = m_Controls(ix).hEdit
Style = GetWindowLongW(win, -16&)
If b = False Then Style = Style Or ES_NOHIDESEL Else Style = Style And Not ES_NOHIDESEL
SetWindowLongW win, -16&, Style
End Property

Property Get HideSelection(ctl As Object) As Boolean
Dim Style As Long
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).hEdit Then win = m_Controls(ix).hEdit
Style = GetWindowLongW(win, -16&)
HideSelection = (Style And ES_NOHIDESEL) = 0
End Property

Property Let Locked(ctl As Object, b As Boolean)
Dim Style As Long
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).hEdit Then win = m_Controls(ix).hEdit
Style = GetWindowLongW(win, -16&)
If b = False Then Style = Style Or ES_READONLY Else Style = Style And Not ES_READONLY
SetWindowLongW win, -16&, Style
End Property

Property Get Locked(ctl As Object) As Boolean
Dim Style As Long
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).hEdit Then win = m_Controls(ix).hEdit
Style = GetWindowLongW(win, -16&)
Locked = (Style And ES_READONLY) = 0
End Property

Property Let MaxLength(ctl As Object, n As Long)
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
If m_Controls(ix).hEdit Then win = m_Controls(ix).hEdit
SendMessageW win, EM_LIMITTEXT, n, 0
End Property

'Property Let MultiLine(ctl As Object, b As Boolean)
'Dim win As Long
'Dim ix As Long
'Dim Style As Long
'win = GetWindow(ctl): If win = 0 Then Exit Property
'ix = GetIx(ctl): If ix < 0 Then Exit Property
'Style = GetWindowLongW(win, -16&)
'If b Then Style = Style Or ES_MULTILINE Else Style = Style Or Not ES_MULTILINE
'SetWindowLongW win, -16&, Style
'End Property

Property Let NumbersOnly(ctl As Object, b As Boolean)
Dim win As Long
Dim ix As Long
If m_Controls(ix).hEdit Then win = m_Controls(ix).hEdit
Dim Style As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
Style = GetWindowLongW(win, -16&)
If b Then Style = Style Or ES_NUMBER Else Style = Style Or Not ES_NUMBER
SetWindowLongW win, -16&, Style
End Property

Property Let RTL(ctl As Object, b As Boolean)
Dim win As Long
Dim ix As Long
Dim Style As Long, xStyle As Long
Dim exStyle As Long, xexStyle As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
exStyle = WS_EX_RTLREADING

Select Case m_Controls(ix).sClass
Case "button": Style = BS_RIGHTBUTTON
Case "listbox", "combobox": exStyle = exStyle Or WS_EX_RIGHT Or WS_EX_LEFTSCROLLBAR
Case "edit": exStyle = exStyle Or WS_EX_RIGHT Or WS_EX_LEFTSCROLLBAR
End Select
If b Then
    AddStyle ctl, Style, xStyle
    AddExStyle ctl, exStyle, xexStyle
Else
    AddStyle ctl, xStyle, Style
    AddExStyle ctl, xexStyle, exStyle
End If
End Property

Property Let Alignment(ctl As Object, ByVal a As Long)
Dim win As Long
Dim ix As Long
Dim Style As Long, xStyle As Long
Dim exStyle As Long, xexStyle As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
ix = GetIx(ctl): If ix < 0 Then Exit Property
Select Case m_Controls(ix).sClass
Case "static": Style = IIf(a = 0, 0, IIf(a = 1, SS_RIGHT, SS_CENTER)): xStyle = Style Xor 3&
Case "edit": Style = IIf(a = 0, 0, IIf(a = 1, ES_RIGHT, ES_CENTER)): xStyle = Style Xor 3&: If a = 1 Then exStyle = exStyle Or WS_EX_RIGHT Else xexStyle = xexStyle Or WS_EX_RIGHT
Case "listbox", "combobox": If a = 1 Then exStyle = exStyle Or WS_EX_RIGHT Else xexStyle = xexStyle Or WS_EX_RIGHT
Case "button": Style = IIf(a = 0, BS_LEFT, IIf(a = 1, BS_RIGHT, BS_CENTER)): xStyle = Style Xor &H300&
End Select
AddStyle ctl, Style, xStyle
AddExStyle ctl, exStyle, xexStyle
End Property

Property Let ScrollBars(ctl As Object, ByVal v As Long)
Dim win As Long
Dim Style As Long, NoStyle As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
If v = 0 Then NoStyle = Style Or WS_HSCROLL Or WS_VSCROLL Or &HC0&
If v = 1 Then Style = WS_HSCROLL Or &H80&: NoStyle = WS_VSCROLL Or &H40&
If v = 2 Then Style = WS_VSCROLL Or &H40&: NoStyle = WS_HSCROLL Or &H80&
If v = 3 Then Style = Style Or WS_HSCROLL Or WS_VSCROLL Or &HC0&
Style = Style Or GetWindowLongW(win, -16&)
Style = Style And Not NoStyle
SetWindowLongW win, -16&, Style
SetWindowPos win, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
End Property

Sub SetFocus(ctl As Object)
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Sub
SetFocusApi win
End Sub

Sub ZOrder(ctl As Object)
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Sub
SetWindowPos win, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
End Sub

Property Let BorderStyle(ctl As Object, ByVal v As Long)
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
If v Then AddStyle ctl, WS_BORDER Else AddStyle ctl, 0, WS_BORDER
SetWindowPos win, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
End Property

Property Let Appearance(ctl As Object, ByVal v As Long)
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
If v Then AddExStyle ctl, WS_EX_CLIENTEDGE Else AddStyle ctl, 0, WS_EX_CLIENTEDGE
SetWindowPos win, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
End Property

Property Let AcceptFiles(ctl As Object, b As Boolean)
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Property
DragAcceptFiles win, IIf(b, 1, 0)
End Property

Sub AddStyle(ctl As Object, Optional ByVal nStyle&, Optional ByVal xStyle As Long)
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Sub
SetWindowLongW win, -16&, (GetWindowLongW(win, -16&) Or nStyle) And Not xStyle
End Sub

Sub AddExStyle(ctl As Object, Optional ByVal nStyle&, Optional ByVal xStyle As Long)
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Sub
SetWindowLongW win, -20&, (GetWindowLongW(win, -20&) Or nStyle) And Not xStyle
End Sub

Sub Refresh(ctl As Object)
Dim win As Long
win = GetWindow(ctl): If win = 0 Then Exit Sub
InvalidateRectV win, 0, 1
End Sub

Sub SetFont(ctl As Object, fnt)
Dim win As Long
Dim f As IFont
Dim hFont As Long
If IsObject(fnt) Then Set f = fnt: hFont = f.hFont Else hFont = fnt
win = GetWindow(ctl): If win = 0 Then Exit Sub
SendMessageW win, WM_SETFONT, hFont, 1
End Sub

Sub SetMouseCursor(ctl As Control, ByVal cursor As eCursorsCtl)
Dim win As Long
Dim ix As Long
win = GetWindow(ctl): If win = 0 Then Exit Sub
ix = GetIx(ctl): If ix < 0 Then Exit Sub
m_Controls(ix).hCursor = cursor
If cursor > 0 And cursor < 32768 Then m_Controls(ix).hCursor = LoadCursorV(0, cursor)
End Sub

Public Function GetWindow(ctl As Object)
Dim id As Long
If TypeOf ctl Is Form Then GetWindow = ctl.hWnd: Exit Function
id = ctl.LinkTimeout - 1
If id < 0 Or id >= m_CtlCount Then Exit Function
GetWindow = m_Controls(id).hWnd
End Function

Private Function GetIx(ctl As Object)
Dim id As Long
If TypeOf ctl Is Form Then Exit Function
id = ctl.LinkTimeout - 1
If id < 0 Or id >= m_CtlCount Then GetIx = -1: Exit Function
GetIx = id
End Function

Private Function AllocStr(ByVal iLen As Long, Optional Characters As Boolean = True) As String
If Characters Then
    PutMem4 VarPtr(AllocStr), SysAllocStringLen(0, iLen)
Else
    PutMem4 VarPtr(AllocStr), SysAllocStringByteLen(0, iLen)
End If
End Function

Private Function FixColor(ByVal iColor As Long) As Long
Dim c As Long
If iColor >= 0 Then FixColor = iColor: Exit Function
c = iColor And &H7FFFFFFF
c = GetSysColor(c)
FixColor = c
End Function

Private Function GetShift(Optional ByVal wParam) As Integer
Dim Shift As Integer
If IsMissing(wParam) Then
    If GetKeyState(vbKeyShift) < 0 Then Shift = Shift Or vbShiftMask
    If GetKeyState(vbKeyControl) < 0 Then Shift = Shift Or vbCtrlMask
    If GetKeyState(vbKeyMenu) < 0 Then Shift = Shift Or vbAltMask
Else
    If wParam And 4 Then Shift = Shift Or vbShiftMask
    If wParam And 8 Then Shift = Shift Or vbCtrlMask
    If GetKeyState(vbKeyMenu) < 0 Then Shift = Shift Or vbAltMask
End If
GetShift = Shift
End Function

Private Function Low16(ByVal n As Long) As Integer
GetMem2 VarPtr(n), Low16
End Function

Private Function High16(ByVal n As Long) As Integer
GetMem2 VarPtr(n) + 2, High16
End Function

Private Function GetId(ByVal hWnd As Long) As Long
Dim id As Long
'id = GetProp(hwnd, "_ID_")
id = GetDlgCtrlID(hWnd)
If id <= 0 Or id > m_CtlCount Then Exit Function
If m_Controls(id - 1).hWnd = hWnd Then GetId = id: Exit Function
If m_Controls(id - 1).sClass <> "combobox" Then Exit Function
If m_Controls(id - 1).hList <> hWnd And m_Controls(id - 1).hEdit <> hWnd Then Exit Function
GetId = id
End Function

Private Sub pvShowToolTip(ByVal id As Long)
Dim txt As String
Dim flag As Long
Dim pt(0 To 1) As Long
m_idTTT = 0
If id < 0 Then ShowWindow m_hTTT, 0: Exit Sub
m_idTTT = id + 1
txt = m_Controls(id).ToolTipText
flag = m_Controls(id).ToolTipFlag
If Len(txt) = 0 Then ShowWindow m_hTTT, 0: Exit Sub
GetCursorPos VarPtr(pt(0))
ShowToolTip txt, flag, pt(0), pt(1)
End Sub

Public Sub HideToolTip()
ShowWindow m_hTTT, 0
End Sub

Public Sub ShowToolTip(txt, ByVal flag As eToolTipFlags, ByVal X As Long, ByVal Y As Long)
Static hFont As Long
Dim rc As RECTL
Dim dc As Long
Dim l&, t&, w&, h&
Dim scrW&, scrH&
Dim f As Long
Dim hMonitor As Long, monInfo As tMONITORINFO
Dim minX&, minY&, maxX&, maxY&

If Len(txt) = 0 Then HideToolTip: Exit Sub
monInfo.cbSize = Len(monInfo)
hMonitor = MonitorFromWindow(m_hwndParent, MONITOR_DEFAULTTONEAREST)
GetMonitorInfoA hMonitor, VarPtr(monInfo)
minX = monInfo.rcWork.left
minY = monInfo.rcWork.top
maxX = monInfo.rcWork.right
maxY = monInfo.rcWork.bottom
f = flag And 3
If flag And eTTFRTL Then f = f Or DT_RTLREADING
dc = GetDC(0)
If hFont = 0 Then
    hFont = CreateFontA(Round(-9 * GetDeviceCaps(dc, 90) / 72), 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 0, 0, "Segoe UI")
End If
rc.right = 700
rc.bottom = 1000
SelectObject dc, hFont
DrawText dc, StrPtr(txt), Len(txt), VarPtr(rc), DT_CALCRECT Or DT_WORDBREAK Or f
ReleaseDC 0, dc
scrW = GetSystemMetrics(0)
scrH = GetSystemMetrics(1)
w = rc.right + 4
h = rc.bottom + 4
l = X - w \ 2
If l + w > maxX Then l = maxX - w
If l < minX Then l = minX
t = Y + 16
If t + h >= maxY Then t = Y - h - 16
If t < minY Then t = minY
SetWindowPos m_hTTT, 0, l, t, w, h, SWP_NOACTIVATE Or SWP_SHOWWINDOW
'ShowWindow m_hTTT, VbAppWinStyle.vbNormalNoFocus
dc = GetDC(m_hTTT)
SelectObject dc, hFont
SelectObject dc, GetStockObject(DC_BRUSH)
SetDCBrushColor dc, GetSysColor(24)
PatBlt dc, 0, 0, w, h, vbPatCopy
SetTextColor dc, 0
SetBkColor dc, GetSysColor(24)
OffsetRect VarPtr(rc), 1, 1
DrawText dc, StrPtr(txt), Len(txt), VarPtr(rc), DT_WORDBREAK Or f
DeleteDC dc
End Sub



Public Function SubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
Static bTrack As Boolean
Dim id As Long
Dim n As Long
Dim ar(0 To 4) As Long
Dim i As Integer
Dim cl As Collection
Dim cnt As Long
Dim s As String
Dim sName As String
Dim iIndex As Long
Dim tme As tTRACKMOUSEEVENT
Dim r As Long
Dim clas As String
Dim win As Long
Dim dc As Long
Dim lf As LOGFONTW
Dim hFont As Long
Dim rc As RECTL, rc2 As RECTL
Dim flag As Long
Dim Style As Long, exStyle As Long
Dim cbi As tCOMBOBOXINFO
Dim ps As PAINTSTRUCT

id = GetId(hWnd)
If id = 0 Then Exit Function

id = id - 1
sName = m_Controls(id).sName
iIndex = m_Controls(id).iIndex
clas = m_Controls(id).sClass

RaiseEvent Message(sName, iIndex, wMsg, wParam, lParam, r, Handled)
If Handled Then SubclassProc = r: Exit Function

Select Case wMsg
Case WM_KILLFOCUS:
    If clas <> "edit" Then Exit Function
    If m_Controls(id).CueBanner = "" Then Exit Function
    If GetWindowTextLength(hWnd) Then Exit Function
    InvalidateRectV hWnd, 0, 1
Case WM_PAINT:
    SubclassProc = 1
    If clas <> "edit" And clas <> "combobox" Then Exit Function
    If clas = "combobox" And m_Controls(id).hEdit <> 0 Then Exit Function
    If m_Controls(id).CueBanner = "" Then Exit Function
    If GetWindowTextLength(hWnd) Then Exit Function
    'Stop
    SubclassProc = DefSubclassProc413(hWnd, wMsg, wParam, lParam)
    Handled = True
    SubclassProc = 0
    'PostMessageW hWnd, WM_CUEBANNER, 0, 0
    win = hWnd
    GetClientRect win, rc2
    rc = rc2
    If clas = "combobox" Then
        cbi.bSize = Len(cbi)
        GetComboBoxInfo hWnd, VarPtr(cbi)
        rc = cbi.rcItem
    End If
    IntersectClipRect dc, rc.left, rc.top, rc.right, rc.bottom
    Style = GetWindowLongW(win, -16&)
    exStyle = GetWindowLongW(win, -20&)
    hFont = SendMessageW(hWnd, WM_GETFONT, 0, 0)
    GetObjectAPI hFont, Len(lf), lf
    lf.lfItalic = 1
    lf.lfWeight = 0
    hFont = CreateFontIndirectW(VarPtr(lf))
    'dc = BeginPaint(win, ps)
    dc = GetDC(win)
    SelectObject dc, hFont
    SelectObject dc, GetStockObject(DC_BRUSH)
    SetDCBrushColor dc, m_Controls(id).BackColor
    SetBkColor dc, m_Controls(id).BackColor
    SetTextColor dc, &H7F7F7F
    If IsHeb(m_Controls(id).CueBanner) Then flag = flag Or DT_RTLREADING
    If clas <> "combobox" Then
        If Style And ES_RIGHT Then flag = flag Or DT_RIGHT
        If Style And ES_CENTER Then flag = flag Or DT_CENTER
    Else
        If exStyle And WS_EX_RIGHT Then flag = flag Or DT_RIGHT
    End If
    'Stop
    DrawText dc, StrPtr(m_Controls(id).CueBanner), -1, VarPtr(rc), flag
    'If Style And ES_RIGHT Then OffsetRect VarPtr(rc), rc2.Right - rc.Right, 0
    'If Style And ES_CENTER Then OffsetRect VarPtr(rc), (rc2.Right - rc.Right) \ 2, 0
    'If rc.Right > rc2.Right Then rc.Right = rc2.Right
    'DrawText dc, StrPtr(m_Controls(id).CueBanner), -1, VarPtr(rc), flag
    'EndPaint win, ps
    ReleaseDC win, dc
    DeleteObject hFont
    m_Controls(id).bCue = True
    Exit Function
Case WM_SETCURSOR:
    n = m_Controls(id).hCursor
    If n = CUR_DEFAULT Then Exit Function
    Handled = True
    SubclassProc = 1
    SetCursor n
Case WM_MOUSEMOVE:
    If bTrack = False Then
    If clas <> "combobox" Or ObjPtr(m_Controls(id).m_SubclassItem) = 0 Then
        tme.cbSize = Len(tme)
        tme.dwFlags = 3
        tme.hwndTrack = hWnd
        tme.dwHoverTime = -1&
        TrackMouseEvent VarPtr(tme)
        bTrack = True
        RaiseEvent MouseEnter(sName, iIndex)
    End If
    End If
    RaiseEvent MouseMove(sName, iIndex, wParam And 3, GetShift(wParam), Low16(lParam), High16(lParam))
Case WM_MOUSELEAVE:
    bTrack = False: pvShowToolTip -1: RaiseEvent MouseLeave(sName, iIndex)
Case WM_MOUSEHOVER:
    pvShowToolTip id
    RaiseEvent MouseHover(sName, iIndex, wParam And 3, GetShift(wParam), Low16(lParam), High16(lParam))
    tme.cbSize = Len(tme)
    tme.dwFlags = 2
    tme.hwndTrack = hWnd
    tme.dwHoverTime = -1&
    TrackMouseEvent VarPtr(tme)
    bTrack = True
Case WM_LBUTTONDOWN: RaiseEvent MouseDown(sName, iIndex, vbLeftButton, GetShift(wParam), Low16(lParam), High16(lParam))
                     If clas = "edit" Then RaiseEvent Click(sName, iIndex)
Case WM_RBUTTONDOWN: RaiseEvent MouseDown(sName, iIndex, vbRightButton, GetShift(wParam), Low16(lParam), High16(lParam))
Case WM_LBUTTONUP: RaiseEvent MouseUp(sName, iIndex, vbLeftButton, GetShift(wParam), Low16(lParam), High16(lParam))
Case WM_RBUTTONUP: RaiseEvent MouseUp(sName, iIndex, vbRightButton, GetShift(wParam), Low16(lParam), High16(lParam))
Case WM_LBUTTONDBLCLK: If clas = "edit" Then RaiseEvent DblClick(sName, iIndex)
Case WM_MOUSEWHEEL:
    ar(0) = Low16(lParam)
    ar(1) = High16(lParam)
    ScreenToClient (hWnd), VarPtr(ar(0))
    RaiseEvent MouseWheel(sName, iIndex, High16(wParam), GetShift(Low16(wParam)), ar(0), ar(1))
Case WM_KEYDOWN: i = wParam: RaiseEvent KeyDown(sName, iIndex, i, GetShift)
    wParam = i
    If wParam = vbKeyA Then If GetShift = vbCtrlMask Then SendMessageW hWnd, EM_SETSEL, 0&, -1&: wParam = 0: Handled = True: Exit Function
Case WM_KEYUP: i = wParam: RaiseEvent KeyUp(sName, iIndex, i, GetShift)
    wParam = i
Case WM_CHAR: i = wParam: RaiseEvent KeyPress(sName, iIndex, i)
    If i = 1 Then i = 0
    wParam = i
    If i = 0 Then Handled = True
Case WM_DROPFILES:
    cnt = DragQueryFileW(wParam, -1&, 0&, 0&)
    If cnt > 0 Then
        Set cl = New Collection
        s = AllocStr(1024)
        For n = 0 To cnt - 1
            cl.Add VBA.left(s, DragQueryFileW(wParam, n, StrPtr(s), Len(s)))
        Next
        RaiseEvent DropFiles(sName, iIndex, cl)
    End If
    DragFinish wParam
End Select
    '--- note: performance optimization for design-time subclassing
'    If Not Handled And ThunkPrivateData(m_pSubclass) = 0 Then
'        Handled = True
'        SubclassProc = DefSubclassProc413(hwnd, wMsg, wParam, lParam)
'    End If
End Function

Public Function SubclassFormProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
Static bTrack As Boolean
Dim id As Integer
Dim note As Integer
Dim sName As String
Dim iIndex As Long
Dim tme As tTRACKMOUSEEVENT
Dim ar(0 To 4) As Long
Dim di As tDRAWITEMSTRUCT
Dim l&, t&, w&, h&
Dim r&
Dim rc As RECTL
Dim dc As Long
Dim iMsg As Msg

RaiseEvent Message("", -1&, wMsg, wParam, lParam, r, Handled)
If Handled Then SubclassFormProc = r: Exit Function

If wMsg = WM_COMMAND Then
If GetAsyncKeyState(vbKeyTab) < 0 Then
'Debug.Print lParam, GetFocus, High16(wParam)
If lParam <> GetFocus Then
    NextTab lParam
End If
End If
End If

Select Case wMsg
Case &H132& To &H138&:
    id = GetId(lParam)
    If id = 0 Then Exit Function
    id = id - 1
    'SelectObject wParam, GetStockObject(DC_BRUSH)
    SetTextColor wParam, m_Controls(id).ForeColor
    SetBkMode wParam, m_Controls(id).BackStyle
    If m_Controls(id).BackStyle = 2 Then SetBkColor wParam, m_Controls(id).BackColor
    SetDCBrushColor wParam, m_Controls(id).BackColor
    SubclassFormProc = GetStockObject(DC_BRUSH)
    Handled = True
Case WM_DRAWITEM:
    id = wParam - 1
    If id < 0 Or id >= m_CtlCount Then Exit Function
    CopyMemory VarPtr(di), lParam, Len(di)
    If m_Controls(id).hWnd <> di.hwndItem Then Exit Function
    l = di.rcItem.left: t = di.rcItem.top: w = di.rcItem.right - l: h = di.rcItem.bottom - t
    RaiseEvent OwnerDraw(m_Controls(id).sName, m_Controls(id).iIndex, di.hdc, di.itemID, l, t, w, h, di.itemAction, di.itemState)
    Handled = True
    SubclassFormProc = 1
Case WM_COMMAND:
    CopyMemory VarPtr(id), VarPtr(wParam), 2
    CopyMemory VarPtr(note), VarPtr(wParam) + 2, 2
    If id <= 0 Or id > m_CtlCount Then Exit Function
    id = id - 1
    If m_Controls(id).hWnd <> lParam Then Exit Function
    sName = m_Controls(id).sName
    iIndex = m_Controls(id).iIndex
    If m_Controls(id).sClass = "edit" Then
        Select Case note
        Case EN_AFTER_PASTE: RaiseEvent AfterPaste(sName, iIndex)
        Case EN_BEFORE_PASTE: RaiseEvent BeforePaste(sName, iIndex)
        Case EN_CHANGE:
            If GetWindowTextLength(lParam) And m_Controls(id).bCue = True Then
                m_Controls(id).bCue = False
                'GetClientRect lParam, rc
                'dc = GetDC(lParam)
                'SelectObject dc, GetStockObject(DC_BRUSH)
                'SetDCBrushColor dc, m_Controls(id).BackColor
                'PatBlt dc, 0, 0, rc.Right, rc.Bottom, vbPatCopy
                'ReleaseDC lParam, dc
                'GdiFlush
                InvalidateRectV lParam, 0, 1
            End If
            RaiseEvent Change(sName, iIndex)
        Case EN_KILLFOCUS: RaiseEvent LostFocus(sName, iIndex)
        Case EN_SETFOCUS: RaiseEvent GotFocus(sName, iIndex)
        Case EN_VSCROLL: RaiseEvent Scroll(sName, iIndex)
        End Select
    ElseIf m_Controls(id).sClass = "button" Then
        Select Case note
        Case BN_CLICKED: RaiseEvent Click(sName, iIndex)
        Case BN_DBLCLK: RaiseEvent DblClick(sName, iIndex)
        Case BN_SETFOCUS: RaiseEvent GotFocus(sName, iIndex)
        Case BN_KILLFOCUS: RaiseEvent LostFocus(sName, iIndex)
        End Select
    ElseIf m_Controls(id).sClass = "listbox" Then
        Select Case note
        Case LBN_DBLCLK: RaiseEvent DblClick(sName, iIndex)
        Case LBN_SETFOCUS: RaiseEvent GotFocus(sName, iIndex)
        Case LBN_KILLFOCUS: RaiseEvent LostFocus(sName, iIndex)
        Case LBN_SELCHANGE: RaiseEvent SelChange(sName, iIndex)
        End Select
    ElseIf m_Controls(id).sClass = "combobox" Then
        Select Case note
        Case CBN_DBLCLK: RaiseEvent DblClick(sName, iIndex)
        Case CBN_SETFOCUS: RaiseEvent GotFocus(sName, iIndex)
        Case CBN_KILLFOCUS: RaiseEvent LostFocus(sName, iIndex)
        Case CBN_SELCHANGE: RaiseEvent SelChange(sName, iIndex)
        Case CBN_EDITCHANGE: RaiseEvent Change(sName, iIndex)
        Case CBN_DROPDOWN: RaiseEvent DropDown(sName, iIndex)
        Case CBN_CLOSEUP: RaiseEvent DropClose(sName, iIndex)
        End Select
    ElseIf m_Controls(id).sClass = "static" Then
        Select Case note
        Case 0: RaiseEvent Click(sName, iIndex)
        Case 1: RaiseEvent DblClick(sName, iIndex)
        End Select
    End If
Case WM_MOUSEWHEEL:
    ar(0) = Low16(lParam)
    ar(1) = High16(lParam)
    ScreenToClient (hWnd), VarPtr(ar(0))
    RaiseEvent MouseWheel("", -1&, High16(wParam), GetShift(Low16(wParam)), ar(0), ar(1))
Case WM_MOUSEMOVE:
    If bTrack = False Then
        tme.cbSize = Len(tme)
        tme.dwFlags = 3
        tme.hwndTrack = hWnd
        tme.dwHoverTime = -1&
        TrackMouseEvent VarPtr(tme)
        bTrack = True
        RaiseEvent MouseEnter("", -1&)
    End If
Case WM_MOUSELEAVE: bTrack = False: RaiseEvent MouseLeave("", -1&)
Case WM_MOUSEHOVER:
    RaiseEvent MouseHover("", -1&, wParam And 3, GetShift(wParam), Low16(lParam), High16(lParam))
    tme.cbSize = Len(tme)
    tme.dwFlags = 2
    tme.hwndTrack = hWnd
    tme.dwHoverTime = -1&
    TrackMouseEvent VarPtr(tme)
    bTrack = True
End Select
    '--- note: performance optimization for design-time subclassing
'    If Not Handled And ThunkPrivateData(m_pSubclass) = 0 Then
'        Handled = True
'        SubclassProc = DefSubclassProc413(hwnd, wMsg, wParam, lParam)
'    End If
End Function

Public Function SubclassTTTProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
Select Case wMsg

End Select
End Function

Public Function SubclassComboProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
Static bTrack As Boolean
Dim id As Long
Dim n As Long
Dim ar(0 To 4) As Long
Dim i As Integer
Dim cl As Collection
Dim cnt As Long
Dim s As String
Dim sName As String
Dim tme As tTRACKMOUSEEVENT
Dim r As Long
Dim clas As String
Dim dc As Long
Dim lf As LOGFONTW
Dim hFont As Long
Dim rc As RECTL, rc2 As RECTL
Dim flag As Long
Dim Style As Long, exStyle As Long
Dim cbi As tCOMBOBOXINFO
Dim ps As PAINTSTRUCT

id = GetId(hWnd)
If id = 0 Then Exit Function

id = id - 1
sName = m_Controls(id).sName
Select Case wMsg
Case WM_KILLFOCUS:
    If m_Controls(id).CueBanner = "" Then Exit Function
    If GetWindowTextLength(hWnd) Then Exit Function
    InvalidateRectV hWnd, 0, 1
Case WM_PAINT:
    SubclassComboProc = 1
    If m_Controls(id).CueBanner = "" Then Exit Function
    If GetWindowTextLength(hWnd) Then Exit Function
    cbi.bSize = Len(cbi)
    GetComboBoxInfo m_Controls(id).hWnd, VarPtr(cbi)
    Handled = True
    SubclassComboProc = 0
    GetClientRect hWnd, rc2
    rc = rc2
    'rc = cbi.rcItem
    Style = GetWindowLongW(hWnd, -16&)
    exStyle = GetWindowLongW(m_Controls(id).hWnd, -20&)
    hFont = SendMessageW(hWnd, WM_GETFONT, 0, 0)
    GetObjectAPI hFont, Len(lf), lf
    lf.lfItalic = 1
    lf.lfWeight = 0
    hFont = CreateFontIndirectW(VarPtr(lf))
    dc = BeginPaint(hWnd, ps)
    SelectObject dc, hFont
    SelectObject dc, GetStockObject(DC_BRUSH)
    SetDCBrushColor dc, m_Controls(id).BackColor
    PatBlt dc, 0, 0, rc2.right, rc2.bottom, vbPatCopy
    SetBkColor dc, m_Controls(id).BackColor
    SetTextColor dc, &H7F7F7F
    If IsHeb(m_Controls(id).CueBanner) Then flag = flag Or DT_RTLREADING
    If exStyle And WS_EX_RIGHT Then flag = flag Or DT_RIGHT
    'DrawText dc, StrPtr(m_Controls(id).CueBanner), -1, VarPtr(rc), flag Or DT_CALCRECT
    'If Style And ES_RIGHT Then OffsetRect VarPtr(rc), rc2.right - rc.right, 0
    'If Style And ES_CENTER Then OffsetRect VarPtr(rc), (rc2.right - rc.right) \ 2, 0
    'If rc.right > rc2.right Then rc.right = rc2.right
    DrawText dc, StrPtr(m_Controls(id).CueBanner), -1, VarPtr(rc), flag
    EndPaint hWnd, ps
    DeleteObject hFont
    Exit Function
Case WM_MOUSEMOVE:
    If bTrack = False Then
        tme.cbSize = Len(tme)
        tme.dwFlags = 3
        tme.hwndTrack = hWnd
        tme.dwHoverTime = -1&
        TrackMouseEvent VarPtr(tme)
        bTrack = True
        RaiseEvent MouseEnter(sName, m_Controls(id).iIndex)
    End If
Case WM_MOUSELEAVE:
    bTrack = False: pvShowToolTip -1: RaiseEvent MouseLeave(sName, m_Controls(id).iIndex)
Case WM_MOUSEHOVER:
    pvShowToolTip id
    RaiseEvent MouseHover(sName, m_Controls(id).iIndex, wParam And 3, GetShift(wParam), Low16(lParam), High16(lParam))
    tme.cbSize = Len(tme)
    tme.dwFlags = 2
    tme.hwndTrack = hWnd
    tme.dwHoverTime = -1&
    TrackMouseEvent VarPtr(tme)
    bTrack = True
End Select
End Function

Private Sub pvSubclass(o As IUnknown, ByVal hWnd As Long)
    Set o = InitSubclassingThunk(hWnd, Me, InitAddressOfMethod(Me, 5).SubclassFormProc(0, 0, 0, 0, 0))
End Sub

Private Sub pvSubclass2(o As IUnknown, ByVal hWnd As Long)
    Set o = InitSubclassingThunk(hWnd, Me, InitAddressOfMethod(Me, 5).SubclassProc(0, 0, 0, 0, 0))
End Sub

Private Sub pvSubclass3(o As IUnknown, ByVal hWnd As Long)
    Set o = InitSubclassingThunk(hWnd, Me, InitAddressOfMethod(Me, 5).SubclassTTTProc(0, 0, 0, 0, 0))
End Sub

Private Sub pvSubclassItem(o As IUnknown, ByVal hWnd As Long)
    Set o = InitSubclassingThunk(hWnd, Me, InitAddressOfMethod(Me, 5).SubclassComboProc(0, 0, 0, 0, 0))
End Sub

Private Sub pvUnsubclass(o As IUnknown)
    TerminateSubclassingThunk o, Me
End Sub

Private Function GetModuleProc(ModuleName As String, ProcName As String) As Long
GetModuleProc = GetProcAddress(GetModuleHandle(ModuleName), ProcName)
End Function

Private Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As UniForm
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = ThunkAllocate(STR_THUNK, THUNK_SIZE)
    If hThunk = 0 Then
        Exit Function
    End If
    lSize = CallWindowProcW(hThunk, ObjPtr(pObj), MethodParamCount, GetModuleProc("kernel32", "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Public Function InitSubclassingThunk(ByVal hWnd As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgepwEBAAV1aLdCQUg8YIgz4AdC+L+oHHKBIQAIvCBQwREACri8IFSBEQAKuLwgVYERAAq4vCBYAREACruQkAAADzpYHCKBIQAFJqHP9SEFqL+IvCq7gBAAAAqzPAq4tEJAyri3QkFKWlM8Crg+8cagBX/3IM/3cM/1IYi0QkGIk4Xl+4XBIQAC1wEBAAwhAADx8Ai0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1GIsKUv9xDP9yDP9RHItUJASLClL/URQzwMIEAJBVi+yLVRj/QgT/QhiLQhg7QgR0b4tCEIXAdGiLCotBLIXAdDdS/9BaiUIIg/gBd1OFwHUJgX0MAwIAAHRGiwpS/1EwWoXAdTuLClJq8P9xJP9RKFqpAAAACHUoUjPAUFCNRCQEUI1EJARQ/3UU/3UQ/3UM/3UI/3IQ/1IUWVhahcl1E1KLCv91FP91EP91DP91CP9RIFr/ShhQUug4////WF3CGAAPHwA=" ' 9.6.2020 13:56:03
    Const THUNK_SIZE    As Long = 492
    Static hThunk       As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitSubclassingThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = ThunkAllocate(STR_THUNK, THUNK_SIZE)
        If hThunk = 0 Then
            Exit Function
        End If
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        Call DefSubclassProc413(0, 0, 0, 0)                                            '--- load comctl32
        aParams(4) = GetProcByOrdinal(GetModuleHandle("comctl32"), 410)             '--- 410 = SetWindowSubclass ordinal
        aParams(5) = GetProcByOrdinal(GetModuleHandle("comctl32"), 412)             '--- 412 = RemoveWindowSubclass ordinal
        aParams(6) = GetProcByOrdinal(GetModuleHandle("comctl32"), 413)             '--- 413 = DefSubclassProc ordinal
        '--- for IDE protection
        Debug.Assert pvThunkIdeOwner(aParams(7))
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitSubclassingThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProcW(hThunk, hWnd, 0, VarPtr(aParams(0)), VarPtr(InitSubclassingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Public Function TerminateSubclassingThunk(pSubclass As IUnknown, pObj As Object) As IUnknown
    If Not pSubclass Is Nothing Then
        Debug.Assert ThunkPrivateData(pSubclass, 2) = ObjPtr(pObj)
        ThunkPrivateData(pSubclass, 2) = 0
        Set pSubclass = Nothing
    End If
End Function

Property Get ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long) As Long
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(VarPtr(ThunkPrivateData), ByVal (lPtr Xor SIGN_BIT) + 8 + Index * 4 Xor SIGN_BIT, 4)
    End If
End Property

Property Let ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long, ByVal lValue As Long)
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ByVal (lPtr Xor SIGN_BIT) + 8 + Index * 4 Xor SIGN_BIT, VarPtr(lValue), 4)
    End If
End Property

Private Function pvThunkIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowExA(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvThunkIdeOwner = True
End Function

Function ThunkAllocate(sText As String, Optional ByVal Size As Long) As Long
    Static Map(0 To &H3FF) As Long
    Dim baInput()       As Byte
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lPtr            As Long
    Dim p               As Long
    
    ThunkAllocate = VirtualAlloc(0, IIf(Size > 0, Size, (Len(sText) \ 4) * 3), MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If ThunkAllocate = 0 Then
        Exit Function
    End If
    '--- init decoding maps
    If Map(65) = 0 Then
        baInput = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
        For lIdx = 0 To UBound(baInput)
            lChar = baInput(lIdx)
            Map(&H0 + lChar) = lIdx * (2 ^ 2)
            Map(&H100 + lChar) = (lIdx And &H30) \ (2 ^ 4) Or (lIdx And &HF) * (2 ^ 12)
            Map(&H200 + lChar) = (lIdx And &H3) * (2 ^ 22) Or (lIdx And &H3C) * (2 ^ 6)
            Map(&H300 + lChar) = lIdx * (2 ^ 16)
        Next
    End If
    '--- base64 decode loop
    baInput = StrConv(Replace(Replace(sText, vbCr, vbNullString), vbLf, vbNullString), vbFromUnicode)
    lPtr = ThunkAllocate
    p = VarPtr(lChar)
    For lIdx = 0 To UBound(baInput) - 3 Step 4
        lChar = Map(baInput(lIdx + 0)) Or Map(&H100 + baInput(lIdx + 1)) Or Map(&H200 + baInput(lIdx + 2)) Or Map(&H300 + baInput(lIdx + 3))
        Call CopyMemory(ByVal lPtr, p, 3)
        lPtr = (lPtr Xor SIGN_BIT) + 3 Xor SIGN_BIT
    Next
End Function

#If ImplSelfContained Then
Private Property Get pvThunkGlobalData(sKey As String) As Long
    Dim sBuffer     As String
    
    sBuffer = String$(50, 0)
    Call GetEnvironmentVariableA("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvThunkGlobalData = Val(left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvThunkGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariableA("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, lValue)
End Property
#End If


Private Sub Class_Initialize()
Dim clas As String
Dim wcex As WNDCLASSEX
clas = "ToolTipForControls"
wcex.cbSize = Len(wcex)
wcex.hbrBackground = GetStockObject(DC_BRUSH)
wcex.hCursor = LoadCursorV(0, CUR_ARROW)
wcex.hInstance = App.hInstance
wcex.lpfnWndProc = GetProcAddress(GetModuleHandle("user32"), "DefWindowProcW")
wcex.lpszClassName = clas
RegisterClassEx VarPtr(wcex)
m_hTTT = CreateWindowExW(WS_EX_NOACTIVATE, StrPtr(clas), 0, WS_BORDER Or WS_POPUP, 0, 0, 1, 1, 0, 0, App.hInstance, 0)
'pvSubclass3 m_pSubclassTTT, m_hTTT
'WM_CUEBANNER = RegisterWindowMessageA("WM_CUEBANNER")
End Sub

Private Sub Class_Terminate()
Dim i As Long
For i = 0 To m_CtlCount - 1
    pvUnsubclass m_Controls(i).m_Subclass
    pvUnsubclass m_Controls(i).m_SubclassItem
    pvUnsubclass m_Controls(i).m_SubclassList
Next
pvUnsubclass m_pSubclassForm
'pvUnsubclass m_pSubclassTTT
DestroyWindow m_hTTT
End Sub

Private Function CompareItems(ctl As Object, sName As String, ByVal iIndex As Long, SortBy As eSortList, ByVal p1 As Long, ByVal p2 As Long, s1 As String, s2 As String, ByVal d1 As Long, ByVal d2 As Long, ByVal Param) As Long
Dim t As eSortList
t = SortBy And 15
If t = eSortCustom Then
    RaiseEvent Compare(sName, iIndex, p1, p2, CompareItems, Param)
    Exit Function
End If
If t = eSortByStringBinary Then
    CompareItems = StrComp(s1, s2, vbBinaryCompare)
ElseIf t = eSortByStringText Then
    CompareItems = StrComp(s1, s2, vbTextCompare)
ElseIf t = eSortAsNumbers Then
    CompareItems = CompareString(LOCALE_USER_DEFAULT, SORT_DIGITSASNUMBERS Or NORM_IGNORECASE, StrPtr(s1), -1&, StrPtr(s2), -1&) - 2
ElseIf t = eSortByItemData Then
    CompareItems = Sgn(d1 - d2)
End If
If SortBy And eSortDescending Then CompareItems = -CompareItems
End Function

Sub SortList(ctl As Control, SortBy As eSortList, Optional ByVal Param)
Dim win As Long
Dim ix As Long
Dim cnt As Long
Dim grp As Long
Dim aList1() As Long
Dim aList2() As Long
Dim ptr1 As Long, ptr2 As Long
Dim dat1 As Long, dat2 As Long
Dim i As Long, i1 As Long, i2 As Long
Dim p As Long, q As Long, pe As Long, qe As Long
Dim c As Long, k As Long
Dim sName As String
Dim sList() As String
Dim aData() As Long
Dim iIndex As Long

win = GetWindow(ctl): If win = 0 Then Exit Sub
ix = GetIx(ctl): If ix < 0 Then Exit Sub
cnt = ListCount(ctl)
If cnt < 2 Then Exit Sub
iIndex = m_Controls(ix).iIndex

NoRedraw(ctl) = True
sName = m_Controls(ix).sName
ReDim sList(0 To cnt - 1)
ReDim aData(0 To cnt - 1)
For i = 0 To cnt - 1
    sList(i) = List(ctl, i)
    aData(i) = ItemData(ctl, i)
Next
ReDim aList1(0 To cnt - 1)
ReDim aList2(0 To cnt - 1)
ptr1 = ArraySA(aList1) + 12
ptr2 = ArraySA(aList2) + 12
GetMem4 ptr1, dat1
GetMem4 ptr2, dat2
For i = 0 To cnt - 1: aList1(i) = i: Next
grp = 1
Do While grp < cnt ' pass the list until all is sorted
    Do While p < cnt ' pass every sub-list pairs in the list
        k = p
        q = p + grp
        If q >= cnt Then CopyMemory dat2 + k * 4, dat1 + p * 4, 4 * (cnt - p): Exit Do
        pe = q
        qe = q + grp
        If qe > cnt Then qe = cnt
        Do ' merge p and q
            i1 = aList1(p)
            i2 = aList1(q)
            c = CompareItems(ctl, sName, iIndex, SortBy, i1, i2, sList(i1), sList(i2), aData(i1), aData(i2), Param)
            If c <= 0 Then
                aList2(k) = aList1(p)
                p = p + 1
                k = k + 1
                If p = pe Then CopyMemory dat2 + k * 4, dat1 + q * 4, 4 * (qe - q): Exit Do
            ElseIf c > 0 Then
                aList2(k) = aList1(q)
                q = q + 1
                k = k + 1
                If q = qe Then CopyMemory dat2 + k * 4, dat1 + p * 4, 4 * (pe - p): Exit Do
            End If
        Loop
        p = qe
    Loop
    p = 0
    k = dat1: dat1 = dat2: dat2 = k
    PutMem4 ptr1, dat1
    PutMem4 ptr2, dat2
    grp = grp + grp
Loop
ClearList ctl
For i = 0 To cnt - 1
    p = aList1(i)
    AddItem ctl, sList(p), -1, aData(p)
Next
NoRedraw(ctl) = False
End Sub

Private Function ArrayDataPtr(arr) As Long
Dim ptr As Long
ptr = ArraySA(arr)
If ptr = 0 Then Exit Function
GetMem4 ptr + 12, ArrayDataPtr
End Function

Private Function ArraySA(arr) As Long
Dim typ As Long
Dim ptr As Long
GetMem4 VarPtr(arr), typ
If typ = vbLong Then ptr = arr: If ptr Then GetMem4 ptr, ArraySA: Exit Function
If (typ And vbArray) = 0 Then Exit Function
GetMem4 VarPtr(arr) + 8, ptr
If (typ And &H4000) Then GetMem4 ptr, ptr
ArraySA = ptr
End Function

Function GetSortedIndexOfNewItemAssumingListIsSorted(ctl As Object, SortBy As eSortList, s As String, Optional nItemData As Long, Optional Param) As Long
Dim win As Long
Dim ix As Long
Dim cnt As Long
Dim low As Long, high As Long, med As Long
Dim c As Long
Dim sName As String
Dim iIndex As Long

win = GetWindow(ctl): If win = 0 Then Exit Function
ix = GetIx(ctl): If ix < 0 Then Exit Function
cnt = ListCount(ctl)
If cnt = 0 Then Exit Function
sName = m_Controls(ix).sName
iIndex = m_Controls(ix).iIndex

low = 0
high = cnt - 1

Do While low <= high
    med = (low + high) \ 2
    c = CompareItems(ctl, sName, iIndex, SortBy, med, -1, List(ctl, med), s, ItemData(ctl, med), nItemData, Param)
    If c = 0 Then GetSortedIndexOfNewItemAssumingListIsSorted = med: Exit Function
    If c < 0 Then low = med + 1 Else high = med - 1
Loop
GetSortedIndexOfNewItemAssumingListIsSorted = low
End Function

Private Function IsHeb(s) As Boolean
Dim i As Integer
Dim c As Integer
Dim ptr As Long

ptr = StrPtr(s)
For i = 1 To Len(s)
    GetMem2 ptr, c
    ptr = ptr + 2
    If c >= 1488 And c <= 1514 Then IsHeb = True: Exit For
Next
End Function

Private Function NextTab(ByVal hWnd As Long) As Long
Static active As Boolean
Dim frm As Form
Dim c
Dim ti As Long
Dim a() As Long
Dim w As Long
Dim cti As Long
Dim d As Long
Dim cnt As Long
Set frm = Screen.ActiveForm
If frm Is Nothing Then Exit Function
If active Then Exit Function
active = True
cnt = frm.Controls.Count
ReDim a(0 To cnt)
cti = -1
For Each c In frm.Controls
    ti = -1
    w = 0
    On Error Resume Next
    ti = c.TabIndex
    If ti >= 0 Then
        If TypeOf c Is Label Then
            If c.LinkItem <> "" Then
                w = m_Controls(c.LinkTimeout - 1).hWnd
            End If
        Else
            On Error Resume Next
            If c.TabStop Then w = c.hWnd
        End If
        If w <> 0 Then
            a(ti) = w
            If w = hWnd Then cti = ti
        End If
    End If
Next
If cti = -1 Then active = False: Exit Function
If GetAsyncKeyState(vbKeyShift) < 0 Then d = -1 Else d = 1
ti = cti + d
Do While ti <> cti
    If ti < 0 Then ti = cnt Else If ti > cnt Then ti = 0
    If ti = cti Then Exit Do
    If a(ti) Then NextTab = a(ti): SetFocusApi a(ti): active = False: Exit Function
    ti = ti + d
Loop
active = False
End Function
