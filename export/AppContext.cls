VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Infrastructure"
Option Explicit
Implements IAppContext

Private Const UUID As String = "{eb37a119-14e2-47c5-ab73-e46804ff84b5}"
Private Const SETTINGS_FILE As String = "tabletransfertool.ini"
Private Const SETTINGS_NODE As String = "TableTransferTool"

Private Type TAppContext
    BindingManager As BindingManager
    CommandManager As CommandManager
    Settings As ISettingsModel
End Type
Private This As TAppContext

Public Property Get Settings() As ISettingsModel
    Set Settings = This.Settings
End Property

Private Property Get IAppContext_BindingManager() As BindingManager
    Set IAppContext_BindingManager = This.BindingManager
End Property

Private Property Get IAppContext_CommandManager() As CommandManager
    Set IAppContext_CommandManager = This.CommandManager
End Property

Private Sub Class_Initialize()
    Set This.BindingManager = New BindingManager
    Set This.BindingManager.Context = Me
    Set This.CommandManager = New CommandManager
End Sub

Public Sub RegisterViewModel(ByVal ViewModel As INotifyPropertyChanged)
    ViewModel.RegisterHandler This.BindingManager
End Sub

Public Sub LoadSettings(ByVal Workbook As Workbook)
    Set This.Settings = SettingsModel.Create() _
        .AddUserSettings(MyDocSettings.Create(UUID:=UUID, Filename:=SETTINGS_FILE)) _
        .AddWorkbookSettings(XMLSettings.Create(Workbook, SETTINGS_NODE))
End Sub
