VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule EncapsulatePublicField
'@Folder "Logging.Model"
Option Explicit

Public LogLevel As LogLevel
Public Topic As String
Public Color As String
Public Timestamp As String
Public Message As String
