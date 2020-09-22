Attribute VB_Name = "Module1"
Option Explicit

'This Code Is placed in a Common Moduleso All forms Can Access it
'This Enum used to Navigate the MyFormValuesOnLoad Array (-1,0,1,2)

Public Enum ValueType
    NotControlArray = -1
    MyName
    MyTextOrValue
    MyIndex
End Enum

'These are Constants for Use in calling IsDirty
Public Const RESET_VALUES As Boolean = True
Public Const RESET_ACTIVE_CONTROL As Boolean = True

'This Code Is placed in a Common Module
'so All forms Can Access it
Public Sub FormatData(MyForm As Form, MyFormValuesOnLoad As Variant)
    'BGS 8/10/1999
    'A. formats data in all controls for MyForm
    'depending upon the control type and what its tag property says
    
    'B. Then it places all the control names and their values into
    'a dynamic two dimensional variant array MyFormValuesOnLoad to be used later.
    'The IsDirty boolean function will use this variant array to tell whether
    'changes were made, as well as reset the values on the form if the user
    'desires to do so.
    
    
    On Error GoTo EH
    
    Dim MyCOntrol As Control
    Dim MyControlCount As Integer
    
    MyControlCount = 0
    
    'A. formats data in all controls for MyForm
    'depending upon the control type and what its tag property says
    
    Screen.MousePointer = vbHourglass
    
    For Each MyCOntrol In MyForm.Controls
        'Put data formating code here
        '
        '
        '
        '
        'End Format Code
        If TypeOf MyCOntrol Is TextBox Or TypeOf MyCOntrol Is CheckBox Or TypeOf MyCOntrol Is ComboBox Then
            MyControlCount = MyControlCount + 1
        End If
    Next
    
    'B. Then it places all the control names and their values into
    'a dynamic two dimensional variant array MyFormValuesOnLoad to be used later.
    'The IsDirty boolean function will use this variant array to tell whether
    'changes were made, as well as reset the values on the form if the user
    'desires to do so.
    
    ReDim MyFormValuesOnLoad(MyName To MyIndex, 1 To MyControlCount)
    
    MyControlCount = 0

    For Each MyCOntrol In MyForm.Controls
        If TypeOf MyCOntrol Is TextBox Or TypeOf MyCOntrol Is CheckBox Or TypeOf MyCOntrol Is ComboBox Then
            MyControlCount = MyControlCount + 1
            MyFormValuesOnLoad(MyName, MyControlCount) = MyCOntrol.Name
            If TypeOf MyCOntrol Is TextBox Then
                MyFormValuesOnLoad(MyTextOrValue, MyControlCount) = MyCOntrol.Text
            ElseIf TypeOf MyCOntrol Is CheckBox Then
                MyFormValuesOnLoad(MyTextOrValue, MyControlCount) = MyCOntrol.Value
            ElseIf TypeOf MyCOntrol Is ComboBox Then
                MyFormValuesOnLoad(MyTextOrValue, MyControlCount) = MyCOntrol.ListIndex
            End If

            If isControlArray(MyForm, MyCOntrol) Then
                MyFormValuesOnLoad(MyIndex, MyControlCount) = MyCOntrol.Index
            Else
                MyFormValuesOnLoad(MyIndex, MyControlCount) = NotControlArray
            End If
        End If
    Next
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description & " In Form " & MyForm.Name, , "FormatData"
End Sub


'This Code Is placed in a Common Module so All forms Can Access it

Public Function isControlArray(MyForm As Form, MyCOntrol As Control) As Boolean
    
    'BGS 8/1/1999 Added this function to determin if a Control is part of
    'a control array or not. I had to do this because VB does not have a
    'function that figures this out. (IsArray does not work on Control Arrays)

    On Error GoTo EH
    Dim MyCount As Integer
    Dim CheckMyControl As Control
    
    For Each CheckMyControl In MyForm.Controls
        If CheckMyControl.Name = MyCOntrol.Name Then
            MyCount = MyCount + 1
        End If
    Next
    
    isControlArray = MyCount - 1
    Exit Function
EH:
    MsgBox Err.Description & "in Form " & MyForm.Name, , "isControlArray"
End Function
    
    'This Code Is placed in a Common Module so All forms Can Access it
Public Function IsDirty(MyForm As Form, MyFormValuesOnLoad As Variant, Optional Reset As Boolean, Optional ResetActiveControl As Boolean, Optional MyActiveControl As Control) As Boolean
    'BGS 8/8/1999 IsDirty for Forms with Tex tBoxes, CheckBoxes, and ComboBoxes
    'Checks all the Controls on Myform and compares their values to what is in
    'MyFormValuesOnLoad Variant Array.
    
    'First the Function checks the type of each Control, if they are a TexBox CheckBox
    'or ComboBox then it will continue on. Continuing, it will check to see if the
    'Control in question is a Control array or not. IF it is then the function will
    'compare each Name in the MyFormValuesOnLoad Variant array, When then name matches
    'the one in the Array, then it will compare the Index. When both name and the Index
    'match , then it will check the TypeOf of the Control in Question. If it is a TexBox
    'then the function will compare the .Text to the MyTextOrValue in the Array. If
    'it matches then It is "Not Dirty" so the Boolean variable bIsDirty remains False. (***Note if the Boolean Variable
    'Reset is set to True Then All Controls will be set back to their previous value stored in the Array.
    'Or if ResetActiveControl is Set to True, Then ONLY the Control which currently has Focus would be reset to
    'the previous value stored in the Array. ***) The function does the exact same thing for
    'the CheckBox and ComboBox controls but uses the .Value and .ListIndex instead of the .Text .
    
    'IF the Control in question is not a control array then the function does the exact same
    'thing as above but leaves out checking to make sure the index matches the Array since it
    'does not have that property.
    
    On Error GoTo EH
    
    Dim MyCOntrol As Control
    Dim MyControlCount As Integer
    Dim MyActCtrlName As String
    Dim MyActCtrlIndex As Integer
    Dim bIsDirty As Boolean
    
    Screen.MousePointer = vbHourglass
    
    If ResetActiveControl Then
        If isControlArray(MyForm, MyActiveControl) Then
            MyActCtrlIndex = MyActiveControl.Index
        End If
        MyActCtrlName = MyActiveControl.Name
    End If

    For Each MyCOntrol In MyForm.Controls
        If TypeOf MyCOntrol Is TextBox Or TypeOf MyCOntrol Is CheckBox Or TypeOf MyCOntrol Is ComboBox Then
            With MyCOntrol
                If isControlArray(MyForm, MyCOntrol) Then
                    For MyControlCount = 1 To UBound(MyFormValuesOnLoad, 2)
                        If MyFormValuesOnLoad(MyName, MyControlCount) = .Name Then
                            If MyFormValuesOnLoad(MyIndex, MyControlCount) = .Index Then
                                If TypeOf MyCOntrol Is TextBox Then
                                    If MyFormValuesOnLoad(MyTextOrValue, MyControlCount) <> .Text Then
                                        bIsDirty = True
                                        If Reset Then
                                            .Text = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
                                        End If
                                        If ResetActiveControl Then
                                            If .Name = MyActCtrlName And .Index = MyActCtrlIndex Then
                                                .Text = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
                                                Screen.MousePointer = vbDefault
                                                Exit Function
                                            End If
                                        End If
                                        Exit For
                                    End If
                                ElseIf TypeOf MyCOntrol Is CheckBox Then
                                    If MyFormValuesOnLoad(MyTextOrValue, MyControlCount) <> .Value Then
                                        bIsDirty = True
                                        If Reset Then
                                            .Value = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
                                        End If
                                        If ResetActiveControl Then
                                            If .Name = MyActCtrlName And .Index = MyActCtrlIndex Then
                                                .Value = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
                                                Screen.MousePointer = vbDefault
                                                Exit Function
                                            End If
                                        End If
                                        Exit For
                                    End If
                                ElseIf TypeOf MyCOntrol Is ComboBox Then
                                    If MyFormValuesOnLoad(MyTextOrValue, MyControlCount) <> .ListIndex Then
                                        bIsDirty = True
                                        If Reset Then
                                            .ListIndex = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
                                        End If
                                        If ResetActiveControl Then
                                            If .Name = MyActCtrlName And .Index = MyActCtrlIndex Then
                                                .ListIndex = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
                                                Screen.MousePointer = vbDefault
                                                Exit Function
                                            End If
                                        End If
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next
                Else
                    For MyControlCount = 1 To UBound(MyFormValuesOnLoad, 2)
                        If MyFormValuesOnLoad(MyName, MyControlCount) = .Name Then
                            If TypeOf MyCOntrol Is TextBox Then
                                If MyFormValuesOnLoad(MyTextOrValue, MyControlCount) <> .Text Then
                                    bIsDirty = True
                                    If Reset Then
                                        .Text = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
                                    End If
                                    If ResetActiveControl Then
                                        If .Name = MyActCtrlName Then
                                            .Text = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
                                            Screen.MousePointer = vbDefault
                                            Exit Function
                                        End If
                                    End If
                                    Exit For
                                End If
                            ElseIf TypeOf MyCOntrol Is CheckBox Then
                                If MyFormValuesOnLoad(MyTextOrValue, MyControlCount) <> .Value Then
                                    bIsDirty = True
                                    If Reset Then
                                        .Value = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
                                    End If
                                    If ResetActiveControl Then
                                        If .Name = MyActCtrlName Then
                                            .Value = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
                                            Screen.MousePointer = vbDefault
                                            Exit Function
                                        End If
                                    End If
                                    Exit For
                                End If
                            ElseIf TypeOf MyCOntrol Is ComboBox Then
                                If MyFormValuesOnLoad(MyTextOrValue, MyControlCount) <> .ListIndex Then
                                    bIsDirty = True
                                    If Reset Then
                                        .ListIndex = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
                                    End If
                                    If ResetActiveControl Then
                                        If .Name = MyActCtrlName Then
                                            .ListIndex = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
                                            Screen.MousePointer = vbDefault
                                            Exit Function
                                        End If
                                    End If
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
            End With
        End If
    Next
    Screen.MousePointer = vbDefault
    IsDirty = bIsDirty
    
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description & " In Form " & MyForm.Name, , "IsDirty"
    
End Function

