Attribute VB_Name = "RunExample2"
'@Folder("MVVM.Example2")
Option Explicit

Private Const DO_DEBUG As Boolean = False

'@EntryPoint "DoTest"
Public Sub DoTest()
    Dim ctx As IAppContext
    Set ctx = New AppContext
    
    Dim VM As CountryViewModel
    Set VM = New CountryViewModel
    
    Dim View As IView
    Set View = GeographyView.Create(ctx, VM)
    
    With View
        If .ShowDialog() Then
            If DO_DEBUG Then Debug.Print "GeographyView.ShowDialog(vm) returned True"
        Else
            If DO_DEBUG Then Debug.Print "GeographyView.ShowDialog(vm) returned False"
        End If
    End With
End Sub
