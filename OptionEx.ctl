VERSION 5.00
Begin VB.UserControl OptionEx 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   FillStyle       =   0  'Solid
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "OptionEx.ctx":0000
   Begin VB.Label ctlCaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "text here"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   1590
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgDisabled 
      Height          =   195
      Left            =   360
      Picture         =   "OptionEx.ctx":0312
      Top             =   2760
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image imgUnchecked 
      Height          =   195
      Left            =   360
      Picture         =   "OptionEx.ctx":03AC
      Top             =   2520
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image imgChecked 
      Height          =   195
      Left            =   360
      Picture         =   "OptionEx.ctx":0505
      Top             =   2280
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image ctlImg 
      Height          =   195
      Left            =   1080
      Top             =   840
      Width           =   195
   End
End
Attribute VB_Name = "OptionEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************
'*UserControl: OptionEx                            DESCRIPTION:                         *
'*File: OptionEx.ctl                                    Skinnable OptionButton          *                                                                                      *
'*Version: 1.0                                                                          *
'*                                                                                      *
'*Copyright (C) Pierre Wennerlund 2002                                                  *
'*                                                                                      *
'*Author: Pierre Wennerlund                                                             *
'*                                                                                      *
'*CHANGE HISTORY:                                                                       *
'*--------------------------------------------------------------------------------------*
'*  Juli 2002                   Pierre Wennerlund:      Initial Code Ver. 1.0           *
'*                                                                                      *
'*                                                                                      *
'*                                                                                      *
'*                                                                                      *
'****************************************************************************************



Private IsEnabled  As Boolean
Private OptionState As Long
Private bChecked As Boolean
Private lPrevState As Long

Public Event Click()
Public Event DblClick()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim m_PicUnchecked As Picture
Dim m_PicChecked As Picture
Dim m_PicDisabled As Picture
Dim m_ForeColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR

Private Sub ctlImg_DblClick()
If IsEnabled = False Then Exit Sub
RaiseEvent DblClick
End Sub

Private Sub ctlImg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If IsEnabled = False Then Exit Sub
RaiseEvent MouseDown(Button, Shift, X, Y)

If Not Button = vbLeftButton Then Exit Sub

If OptionState = 1 Then Exit Sub

   OptionState = 1
   Call ContainerCheck(False)

Call DrawOption(OptionState)

End Sub

Private Sub ContainerCheck(ctlValue As Boolean)

For Each Control In UserControl.Parent.Controls
    If Not Control.Container Is UserControl.Parent Then
        If Control.Name = UserControl.Ambient.DisplayName Then
            Call CheckContainerControls(Control.Container, ctlValue)
            Exit Sub
        End If
    End If
Next Control

Call CheckContainerControls(UserControl.Parent, ctlValue)
End Sub

Private Sub CheckContainerControls(cContainer As Object, ctlValue As Boolean)

For Each Control In UserControl.Parent.Controls
    If Control.Container Is cContainer Then
        If TypeOf Control Is OptionEx Then
            If Not Control.Name = UserControl.Ambient.DisplayName Then
                If Control.Value = True Then
                    Control.Value = ctlValue
                End If
            End If
        End If
    End If
Next
    
End Sub

Private Sub ctlImg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IsEnabled = False Then Exit Sub
RaiseEvent Click
RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub ctlCaption_DblClick()
If IsEnabled = False Then Exit Sub
RaiseEvent DblClick
End Sub

Private Sub ctlCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ctlImg_MouseDown Button, Shift, X, Y
End Sub

Private Sub ctlCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ctlImg_MouseUp Button, Shift, X, Y
End Sub

Public Property Let Value(lChecked As Boolean)

Call ContainerCheck(False)

bChecked = lChecked
If bChecked = True Then
    OptionState = 1
Else
    OptionState = 0
End If



If Enabled = True Then
    Call DrawOption(OptionState)
End If

End Property

Public Property Get Value() As Boolean
Value = bChecked
End Property

Public Property Get PicChecked() As Picture
Set PicChecked = m_PicChecked
End Property

Public Property Set PicChecked(ByVal NewValue As Picture)
    Set m_PicChecked = NewValue
    Call DrawOption(OptionState)
    
    PropertyChanged "PicChecked"
    
End Property

Public Property Get PicDisabled() As Picture
Set PicDisabled = m_PicDisabled
End Property

Public Property Set PicDisabled(ByVal NewValue As Picture)
    Set m_PicDisabled = NewValue
    
    Call DrawOption(OptionState)
    
    PropertyChanged "PicDisabled"
    
End Property

Public Property Get PicUnchecked() As Picture
Set PicUnchecked = m_PicUnchecked
End Property

Public Property Set PicUnchecked(ByVal NewValue As Picture)
    Set m_PicUnchecked = NewValue
    Call DrawOption(OptionState)
    
    PropertyChanged "PicUnchecked"
    
End Property

Private Sub UserControl_InitProperties()

Set PicChecked = imgChecked.Picture
Set PicDisabled = imgDisabled.Picture
Set PicUnchecked = imgUnchecked.Picture
Set Font = Parent.Font

Enabled = True
Value = False
Caption = Name
ForeColor = Parent.ForeColor
BackColor = Parent.BackColor

End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ctlImg_MouseDown Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ctlImg_MouseUp Button, Shift, X, Y
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Value = PropBag.ReadProperty("Value", False)
Caption = PropBag.ReadProperty("Caption", Name)
Enabled = PropBag.ReadProperty("Enabled", True)
Set Font = PropBag.ReadProperty("Font", Parent.Font)


Set PicChecked = PropBag.ReadProperty("PicChecked", Nothing)
Set PicDisabled = PropBag.ReadProperty("PicDisabled", Nothing)
Set PicUnchecked = PropBag.ReadProperty("PicUnchecked", Nothing)

m_ForeColor = PropBag.ReadProperty("ForeColor", Parent.ForeColor)
m_BackColor = PropBag.ReadProperty("BackColor", Parent.BackColor)

UserControl.BackColor = m_BackColor
ctlCaption.ForeColor = m_ForeColor
End Sub

Private Sub UserControl_Resize()
ctlImg.Top = UserControl.ScaleTop
ctlImg.Left = UserControl.ScaleLeft
ctlCaption.Top = UserControl.ScaleTop
ctlCaption.Left = (ctlImg.Left + ctlImg.Width) + 100
ctlCaption.Width = Width

If ctlImg.Height > ctlCaption.Height Then
    Height = ctlImg.Height
    ctlCaption.Top = (ctlImg.Height / 2) - (ctlCaption.Height / 2)
Else
    Height = ctlCaption.Height
    ctlImg.Top = (ctlCaption.Height / 2) - (ctlImg.Height / 2)
End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Value", bChecked, False)
Call PropBag.WriteProperty("Caption", ctlCaption.Caption, Name)
Call PropBag.WriteProperty("Enabled", IsEnabled, True)
Call PropBag.WriteProperty("Font", ctlCaption.Font, Parent.Font)

Call PropBag.WriteProperty("PicChecked", m_PicChecked, Nothing)
Call PropBag.WriteProperty("PicDisabled", m_PicDisabled, Nothing)
Call PropBag.WriteProperty("PicUnchecked", m_PicUnchecked, Nothing)
Call PropBag.WriteProperty("ForeColor", m_ForeColor, Parent.ForeColor)
Call PropBag.WriteProperty("BackColor", m_BackColor, Parent.BackColor)

End Sub

Public Property Set Font(NewFont As IFontDisp)
Set ctlCaption.Font = NewFont
    UserControl_Resize
PropertyChanged "Font"
End Property

Public Property Get Font() As IFontDisp
Set Font = ctlCaption.Font
End Property

Public Property Let Caption(NewCap As String)
ctlCaption.Caption = NewCap

PropertyChanged "Caption"
End Property

Public Property Get Caption() As String
Caption = ctlCaption.Caption
End Property

Property Let Enabled(ByVal lEnabled As Boolean)
IsEnabled = lEnabled
UserControl.Enabled = lEnabled

If IsEnabled = True Then
    If CheckOtherControl = True Then
        If bChecked = True Then
            OptionState = 1
        Else
            OptionState = 0
        End If
        Call DrawOption(OptionState)
    Else
        OptionState = lPrevState
        Call DrawOption(lPrevState)
    End If
   ctlCaption.Enabled = True
Else
    OptionState = 2
    Call DrawOption(OptionState)
    ctlCaption.Enabled = False
End If

PropertyChanged "Enabled"

End Property

Private Function CheckOtherControl() As Boolean
Dim tmpControl As Object
Dim tControl As Control

On Error Resume Next

For Each Control In UserControl.Parent.Controls
    If TypeOf Control Is OptionEx And Control.Name = UserControl.Ambient.DisplayName Then
        For Each tControl In UserControl.Parent.Controls
            If tControl.Container Is Control.Container Then
                If Not tControl.Name = UserControl.Ambient.DisplayName Then
                    If TypeOf tControl Is OptionEx And tControl.Value = True Then
                        CheckOtherControl = True
                        Exit For
                    End If
                End If
            End If
        Next tControl
    End If
Next Control

End Function

Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property

Private Sub DrawOption(lState As Long)

Select Case lState
    Case 0
        If m_PicUnchecked Is Nothing Then
            ctlImg.Picture = imgUnchecked.Picture
        Else
            ctlImg.Picture = m_PicUnchecked
        End If
        lPrevState = OptionState
        bChecked = False
    Case 1
        If m_PicChecked Is Nothing Then
            ctlImg.Picture = imgChecked.Picture
        Else
            ctlImg.Picture = m_PicChecked
        End If
        lPrevState = OptionState
        bChecked = True
    Case 2
        If m_PicDisabled Is Nothing Then
            ctlImg.Picture = imgDisabled.Picture
        Else
            ctlImg.Picture = m_PicDisabled
        End If
        bChecked = False
End Select

Height = ctlImg.Height

End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    
    ctlCaption.ForeColor = m_ForeColor
    
    PropertyChanged "ForeColor"
End Property


Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    
    UserControl.BackColor = m_BackColor
    
    PropertyChanged "BackColor"
End Property

