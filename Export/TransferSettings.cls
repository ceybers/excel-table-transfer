VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Model.Settings"
Option Explicit

Private Const SETTINGS_UUID As String = "{e0daadf3-7046-4d50-95c4-5df5ad6b2bba}"
Private Const SETTINGS_FILENAME As String = "tabletransfertool.ini"
Private Const XML_ROOT_NODE_NAME As String = "TableTransferTool"
Private Const MRU_NAME_RECENT_TRANSFERS As String = "RecentTransfers"

Private Type TState
    User As ISettings
    Workbooks As Object
    ListObjects As Object
    
    RecentTransfers As IMostRecentlyUsed
End Type
Private This As TState

Public Property Get User() As ISettings
    Set User = This.User
End Property

Public Property Get RecentTransfers() As IMostRecentlyUsed
    Set RecentTransfers = This.RecentTransfers
End Property

Private Sub Class_Initialize()
    Set This.User = MyDocSettings.Create( _
        UUID:=SETTINGS_UUID, _
        Filename:=SETTINGS_FILENAME)
    Set This.Workbooks = CreateObject("Scripting.Dictionary")
    Set This.ListObjects = CreateObject("Scripting.Dictionary")
    Set This.RecentTransfers = New MostRecentlyUsed
End Sub

Public Function GetWorkbookSettings(ByVal Workbook As Workbook) As ISettings
    Dim WorkbookSettings As XMLSettings
    If Not This.Workbooks.Exists(Workbook) Then
        Set WorkbookSettings = XMLSettingsFactory.CreateWorkbookSettings( _
          Workbook:=Workbook, _
          RootNode:=XML_ROOT_NODE_NAME)
        This.Workbooks.Add Key:=Workbook, Item:=WorkbookSettings
    Else
        Set WorkbookSettings = This.Workbooks.Item(Workbook)
    End If
    
    Set GetWorkbookSettings = WorkbookSettings
End Function

Public Function GetListObjectSettings(ByVal ListObject As ListObject) As ISettings
    Dim WorkbookSettings As XMLSettings
    Set WorkbookSettings = GetWorkbookSettings(ListObject.Parent.Parent)
    
    Dim ListObjectSettings As XMLSettings
    If Not This.ListObjects.Exists(ListObject) Then
        Set ListObjectSettings = XMLSettingsFactory.CreateTableSettings( _
          TableName:=ListObject.Name, _
          WorkbookSettings:=WorkbookSettings)
        This.ListObjects.Add Key:=ListObject, Item:=ListObjectSettings
    Else
        Set ListObjectSettings = This.ListObjects.Item(ListObject)
    End If
    
    Set GetListObjectSettings = ListObjectSettings
End Function

Public Sub LoadMRU()
    This.RecentTransfers.FromCollection This.User.GetCollection(MRU_NAME_RECENT_TRANSFERS)
End Sub

Public Sub SaveMRU()
    This.User.SetCollection MRU_NAME_RECENT_TRANSFERS, This.RecentTransfers.ToCollection
End Sub


