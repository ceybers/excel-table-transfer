VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KeyMapperView 
   Caption         =   "Table Transfer Tool"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "KeyMapperView.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "KeyMapperView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder "MVVM.Views"
Option Explicit
Implements IView

Private Type TState
    Context As IAppContext
    ViewModel As KeyMapperViewModel
    Result As TtViewResult
    DoubleClickHelper As ListViewDoubleClickHelper
End Type
Private This As TState

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = This.ViewModel
End Property

Public Property Get ViewModel() As KeyMapperViewModel
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal vNewValue As KeyMapperViewModel)
    Set This.ViewModel = vNewValue
End Property

Public Property Get Context() As IAppContext
    Set Context = This.Context
End Property

Public Property Set Context(ByVal vNewValue As IAppContext)
    Set This.Context = vNewValue
End Property

Private Sub cboBack_Click()
    This.Result = vrBack
    Me.Hide
End Sub

Private Sub cboNext_Click()
    This.Result = vrNext
    Me.Hide
End Sub

Private Sub cboCancel_Click()
    This.Result = vrCancel
    Me.Hide
End Sub

Private Sub lvSrcKeys_DblClick()
    This.DoubleClickHelper.OnDblClick
End Sub

Private Sub lvDstKeys_DblClick()
    This.DoubleClickHelper.OnDblClick
End Sub

Private Sub lvSrcKeys_ItemClick(ByVal Item As MSComctlLib.ListItem)
    This.ViewModel.TryEvaluateMatch
End Sub

Private Sub lvDstKeys_ItemClick(ByVal Item As MSComctlLib.ListItem)
    This.ViewModel.TryEvaluateMatch
End Sub

Private Sub lvSrcKeys_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
    TryAutoFocusNext
    Dim ListItem As ListItem
    If This.DoubleClickHelper.TryGetDoubleClickedListItem(Me.lvSrcKeys, x, y, ListItem) Then
        Me.ViewModel.Destination.TrySelect ListItem.Key
    End If
End Sub

Private Sub lvDstKeys_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
    TryAutoFocusNext
    Dim ListItem As ListItem
    If This.DoubleClickHelper.TryGetDoubleClickedListItem(Me.lvDstKeys, x, y, ListItem) Then
        Me.ViewModel.Source.TrySelect ListItem.Key
    End If
End Sub

Private Sub cmbSrcQuality_DropButtonClick()
    Me.cboCancel.SetFocus
    DoShowQuality ttSource
End Sub

Private Sub cmbDstQuality_DropButtonClick()
    Me.cboCancel.SetFocus
    DoShowQuality ttDestination
End Sub

Private Sub cmbMatchQuality_DropButtonClick()
    Me.cboCancel.SetFocus
    DoShowMatchQuality
End Sub

Private Sub IView_Show()
    IView_ShowDialog
End Sub
 
Private Sub IView_Hide()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        This.Result = vrCancel
    End If
End Sub

Private Sub lblHeaderIcon_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    frmAbout.Show
End Sub

Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As KeyMapperViewModel) As IView
    Dim Result As KeyMapperView
    Set Result = New KeyMapperView
    
    Set Result.Context = Context
    Set Result.ViewModel = ViewModel

    Set Create = Result
End Function

Private Function IView_ShowDialog() As TtViewResult
    Me.lblHeaderText.Caption = HDR_TXT_KEY_MAPPER
    Set This.DoubleClickHelper = New ListViewDoubleClickHelper
    
    BindControls
    If Me.cboNext.Enabled Then Me.cboNext.SetFocus
    
    Me.Show
    
    IView_ShowDialog = This.Result
End Function

Private Sub BindControls()
    With Context.BindingManager
        .BindPropertyPath ViewModel, "Source", Me.lvSrcKeys, "ListItems", TwoWayBinding, KeyColumnsToListView
        .BindPropertyPath ViewModel, "Destination", Me.lvDstKeys, "ListItems", TwoWayBinding, KeyColumnsToListView
        
        .BindPropertyPath ViewModel, "Source.Caption", Me.cmbSrcQuality, "Text", OneWayBinding
        .BindPropertyPath ViewModel, "Destination.Caption", Me.cmbDstQuality, "Text", OneWayBinding
        .BindPropertyPath ViewModel, "MatchQualityCaption", Me.cmbMatchQuality, "Text", OneWayBinding
        
        .BindPropertyPath ViewModel, "CanNext", Me.cboNext, "Enabled", OneWayBinding
    End With
End Sub

Private Sub TryAutoFocusNext()
    If Me.cboNext.Enabled = True Then
        Me.cboNext.SetFocus
    End If
End Sub

Private Sub DoShowQuality(ByVal Direction As TtDirection)
    Dim KeyQualityViewModel As KeyQualityViewModel
    Set KeyQualityViewModel = New KeyQualityViewModel
    
    If Direction = ttSource Then
        KeyQualityViewModel.Load This.ViewModel.Source.Selected
    Else
        KeyQualityViewModel.Load This.ViewModel.Destination.Selected
    End If
    
    Dim View As IView
    Set View = KeyQualityView.Create(This.Context, KeyQualityViewModel)
    
    View.Show
End Sub

Private Sub DoShowMatchQuality()
    Dim View As IView
    Set View = MatchQualityView.Create(This.Context, This.ViewModel)
    View.Show
End Sub
