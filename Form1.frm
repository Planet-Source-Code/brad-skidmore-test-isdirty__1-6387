VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   8280
      TabIndex        =   18
      Top             =   6480
      Width           =   1455
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Index           =   2
      Left            =   6360
      TabIndex        =   17
      Tag             =   "Combo 6"
      Text            =   "Combo1"
      Top             =   5760
      Width           =   2175
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Index           =   1
      Left            =   6360
      TabIndex        =   16
      Tag             =   "Combo 5"
      Text            =   "Combo1"
      Top             =   5160
      Width           =   2175
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Index           =   0
      Left            =   6360
      TabIndex        =   15
      Tag             =   "Combo 4"
      Text            =   "Combo1"
      Top             =   4560
      Width           =   2175
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   6240
      TabIndex        =   14
      Tag             =   "Combo 3"
      Text            =   "Combo1"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   6240
      TabIndex        =   13
      Tag             =   "Combo 2"
      Text            =   "Combo1"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6240
      TabIndex        =   12
      Tag             =   "Combo 1"
      Text            =   "Combo1"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check3"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check2"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   2280
      Value           =   2  'Grayed
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check1"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1560
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check5"
      Height          =   495
      Index           =   2
      Left            =   2880
      TabIndex        =   10
      Top             =   5160
      Value           =   2  'Grayed
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check6"
      Height          =   495
      Index           =   1
      Left            =   2880
      TabIndex        =   11
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check4"
      Height          =   495
      Index           =   0
      Left            =   2880
      TabIndex        =   9
      Top             =   4560
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1320
      Width           =   1935
   End
   Begin MSComctlLib.Toolbar tbrReset 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1058
      ButtonWidth     =   2487
      ButtonHeight    =   953
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "UnDo Change"
            Key             =   "ResetActive"
            Object.ToolTipText     =   "Button"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "UnDo All Changes"
            Key             =   "ResetAll"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "CheckBox Ctl Array "
      Height          =   375
      Left            =   5880
      TabIndex        =   25
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Check Box Not Ctl Array "
      Height          =   375
      Left            =   6000
      TabIndex        =   24
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "CheckBox Ctl Array "
      Height          =   375
      Left            =   2880
      TabIndex        =   23
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Check Box Not Ctl Array "
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "TextBox Ctl Array "
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "TextBox Not Ctl Array "
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''''''
'This Goes in your Form as a Mod level V
'     ariable. it will be used to Store
'All the Values of TextBoxes, CheckBoxes
'     , and ComboBoxes on Load
Private mValuesOnLoad() As Variant



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Button"
            'ToDo: Add 'Button' button code.
            MsgBox "Add 'Button' button code."
    End Select
End Sub

Private Sub cmdSave_Click()
    If IsDirty(Me, mValuesOnLoad) Then
        'Put Save Call in here
        'End Save Call
        'Have to Reset the mValuesOnLoad to New Saved Values
        Erase mValuesOnLoad
        FormatData Me, mValuesOnLoad
        MsgBox "Save Successful"
    Else
        'Do Whatever if No Data Changes
        MsgBox "All Data is the Same No Save Made"
    End If
    
End Sub

Private Sub Form_Load()
    Dim MyCOntrol As Control
    'Fill Comboboxes And Set the Listindex
    For Each MyCOntrol In Me.Controls
        If TypeOf MyCOntrol Is ComboBox Then
            FillCombo MyCOntrol
        End If
    Next
    
    FormatData Me, mValuesOnLoad
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsDirty(Me, mValuesOnLoad) Then
        If MsgBox("Data has changed. Do You want to Save Changes ?", vbYesNo) = vbYes Then
            'Put Save Call Here
            MsgBox "Data Saved"
        Else
            MsgBox "Data not Saved"
        End If
    End If
        
End Sub

''''''''''''''''''''''''''''''''''''''''
'
'This is the Click event for a ToolBar w
'     ith Buttons you could use on your form
'I used a tool bar because the Active Co
'     ntrol such as a Textbox or whatever will
'
'Remain Active even though you click on
'     the ToolBar Button. This is Handy to kno
'     w
'if you want to reset Just the Active Te
'     xtbox to its Original Value.


Private Sub tbrReset_ButtonClick(ByVal Button As MSComctlLib.Button)
    'BGS 8/17/99
    
    On Error GoTo EH
    


    Select Case Button.Key
        Case "ResetAll"


            If IsDirty(Me, mValuesOnLoad) Then
    
    
                Select Case MsgBox("Are you sure you want To Reset All Values ?", vbYesNo + vbQuestion, " Reset to Previous Values")
                    Case vbYes
                    Call IsDirty(Me, mValuesOnLoad, RESET_VALUES)
                    Case vbNo
                    Exit Sub
                End Select
            Else
                'MsgBox "Could not find Any Changes to R
                '     eset", vbInformation, "Reset"
            End If
            
        Case "ResetActive"
            Call IsDirty(Me, mValuesOnLoad, , RESET_ACTIVE_CONTROL, Me.ActiveControl)
    End Select


Exit Sub
EH:
MsgBox Err.Description & " In Form " & Me.Name, , "ResetToolBar_ButtonClick"
End Sub

Private Sub FillCombo(MyComboBox As ComboBox)
    Dim iCount As Integer
    
    With MyComboBox
        .AddItem "Combo 1"
        .AddItem "Combo 2"
        .AddItem "Combo 3"
        .AddItem "Combo 4"
        .AddItem "Combo 5"
        .AddItem "Combo 6"
        For iCount = 0 To .ListCount - 1
            .ListIndex = iCount
            If .Text = .Tag Then
                Exit Sub
            End If
        Next
    End With
End Sub


