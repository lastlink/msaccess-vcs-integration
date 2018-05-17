Attribute VB_Name = "Test_VCS_Loader"
Option Compare Database

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod
Public Sub Test_loadVCS() 'test that the loader is working
    On Error GoTo TestFail
    
    'Arrange:
    loadVCS
    'Act:

    'Assert:
    Assert.IsNotNothing ExportAllSource, "Export sub should exist"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod
Public Sub Test_displayFormVersion() 'test the version modal popup
    On Error GoTo TestFail
    
    'Arrange:
    Fake.MsgBox.Returns 42
    Debug.Print displayFormVersion
    
   ' Debug.Print MsgBox("Flabbergast yet?", vbYesNo, "Rubberduck")
    
       ' .Parameter "prompt", "Flabbergasted yet?"
    With Fakes.MsgBox.Verify
        .Parameter "buttons", vbOK
    End With
    

    'Act:

    'Assert:
    Assert.Inconclusive

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


