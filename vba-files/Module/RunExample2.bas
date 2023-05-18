Attribute VB_Name = "RunExample2"
'@Folder("MVVM.Example2")
Option Explicit

Private Const DO_DEBUG As Boolean = False

'@EntryPoint "DoTest"
Public Sub DoTest()
    Dim ctx As IAppContext
    Set ctx = New AppContext
    
    Dim vm As CountryViewModel
    Set vm = New CountryViewModel
    
    Dim view As IView
    Set view = GeographyView.Create(ctx, vm)
    
    With view
        If .ShowDialog() Then
            If DO_DEBUG Then Debug.Print "GeographyView.ShowDialog(vm) returned True"
        Else
            If DO_DEBUG Then Debug.Print "GeographyView.ShowDialog(vm) returned False"
        End If
    End With
End Sub
