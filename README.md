
## UniForm: Modern Win32 Controls & Unicode support for VB6 (In a Single Class)

**UniForm** is a lightweight, self-contained VB6 class framework that replaces bloated UserControl libraries and external dependencies. 
By using standard VB `Label` controls as design-time placeholders, UniForm transforms them at runtime into fully-featured, Unicode-native Win32 controls with automatic subclassing, owner-draw support, and modern mouse events.

---

## 🌟 Why UniForm?

* **Zero Dependencies:** Just drop `UniForm.cls` into your project. No OCXs, no Type Libraries (`.tlb`), no installation headaches.
* **True Unicode Support:** Bypass classic VB's ANSI bottlenecks for international text and modern system rendering.
* **WYSIWYG Form Designer:** Leverage the convenient of the native VB6 IDE designer using standard `Label` controls for layout and positioning.
* **Centralized Subclassing:** Say goodbye to IDE crashes caused by multi-control subclassing hooks. One central message router handles everything cleanly.

---

## 🚀 3-Step Quick Start

### 1. Design Your Form
1. Open your Form in the VB6 designer.
2. Draw standard **`Label`** controls where you want your windows/controls to appear.
3. Name them appropriately (e.g., `txtAddress`, `lstFiles`, `cmdPlay`).
4. Set their design properties (Font, BackColor, ForeColor, Alignment, etc.)—UniForm copies these directly to the underlying Win32 control.
5. In the label's **`LinkItem`** property, enter the target Win32 class name (e.g., `Edit`, `ListBox`, `Button`, `Static`, `ComboBox`, `CheckBox`, etc.).

### 2. Declare the Class in Your Form Code
In your Form's code module, declare the class with events:

```vb
Option Explicit

' UniForm instance supporting global events
Public WithEvents Uni As UniForm

Private Sub Form_Load()
    Set Uni = New UniForm
    Uni.Init Me  ' This creates the Win32 controls and hides the placeholder labels
End Sub
```
### 3. Manage Controls & Handle Events
Interact with your controls using UniForm's clean property and method wrappers:
```vb
Private Sub Form_Load()
    ' Set properties using pixel values and standard wrappers
    Uni.ToolTipText(lstFiles) = "Drag files to the list here"
    Uni.Text(txtAddress) = "https://example.com"
End Sub

' Unified event handling across all controls
Private Sub Uni_Click(ByVal ControlName As String)
    If ControlName = "cmdPlay" Then
        MsgBox "Play button clicked!"
    End If
End Sub

Private Sub Uni_MouseWheel(ByVal ControlName As String, ByVal Delta As Integer, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
    ' Smooth mouse wheel handling out of the box!
    If ControlName = "lstFiles" Then
        ' Custom wheel logic if needed
    End If
End Sub
```

## 🛠️ Supported Windows Classes
* Edit (TextBox)
* ListBox
* ComboBox
* Button (CommandButton)
* Static (Labels and Images)
* CheckBox / 3State / RadioButton / RadioButtonGroup
* Any valid Win32 window class can also be targeted via `LinkItem`.

## 📂 Project Structure & Distribution
Because UniForm is entirely self-contained, distribution is as simple as copying a single file:
* `UniForm.cls` — The complete engine, subclassing router, and wrapper library.

## - Full features and usage description in class comments -
