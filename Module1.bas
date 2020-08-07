Attribute VB_Name = "Module1"
Type enreg
        nom As String
        prenom As String
        add As String
        codpost As String
        tel As String
        comment As String
        photo As String
End Type
Public max As Integer
Public enregistrement(1 To 999) As enreg
Type config
        image As Boolean
        scan As String
        imprimauto As Boolean
        direct As String
End Type
Public config1 As config

Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwndApp&, ByVal wPixTypes&) As Long
Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwndApp&) As Long

Type file1
        path1 As String
        filename As String
End Type
Public fileamoi As file1

Public Sub Sauver(file As String)

Open file For Output As #1
Write #1, max
For i = 1 To max
Write #1, enregistrement(i).nom, enregistrement(i).prenom, enregistrement(i).add, enregistrement(i).codpost, enregistrement(i).tel, enregistrement(i).photo, enregistrement(i).comment
Next i
Close #1


file = App.Path & "\Abihome.ini"
Open file For Output As #1
Write #1, config1.image, config1.scan, config1.imprimauto, config1.direct

Close #1
End Sub
Public Sub Charger(file As String)
Open file For Input As #1
Input #1, max
For i = 1 To max
Input #1, enregistrement(i).nom, enregistrement(i).prenom, enregistrement(i).add, enregistrement(i).codpost, enregistrement(i).tel, enregistrement(i).photo, enregistrement(i).comment
Next i
Close #1
frmprinc.VScroll1.Value = 1
frmprinc.VScroll1.max = max
End Sub
