Option Explicit

' VBA���W���[�����
Const StarndardModule = 1
Const ClassModule = 2
Const UserForm = 3

' VBA���W���[���t�@�C���̏o�͎��̊g���q/�f�B���N�g����
Dim ModuleSuffix 
Set ModuleSuffix = CreateObject("Scripting.Dictionary")
ModuleSuffix.Add StarndardModule, "bas"
ModuleSuffix.Add ClassModule, "cls"
ModuleSuffix.Add UserForm, "frm"


Const OutputDir = "output"
    
Sub Main()
    Dim path
    Dim access
    Dim project
    Dim module

    WScript.Echo "start"
    path = InputBox("Access�v���W�F�N�g�t�@�C���̃p�X����͂��Ă��������B")
    'path = "D:\Database\Access\Database1.adp"
    WScript.Echo "vba project file = " & path

    Set access = CreateObject("Access.Application")
    access.OpenAccessProject path, True

    Set project = access.CurrentProject
    For Each module In project.AllModules
        WScript.Echo module.Name
    Next
End Sub

Main

