VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmprinc 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Futurs clients ?"
   ClientHeight    =   4125
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   8430
   Icon            =   "frmprinc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Enregistrements : "
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin MSComDlg.CommonDialog cdchph 
         Left            =   3840
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   327681
         Filter          =   "*.bmo|*.bmp"
      End
      Begin MSComDlg.CommonDialog cd2 
         Left            =   480
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   327681
         Filter          =   "*.csv|*.csv"
      End
      Begin MSComDlg.CommonDialog cd1 
         Left            =   480
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   327681
         Filter          =   "*.csv|*.csv"
      End
      Begin VB.TextBox txtcommentaire 
         Height          =   975
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   3000
         Width           =   2895
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Importer une photo"
         Height          =   375
         Left            =   4200
         TabIndex        =   17
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton cmdsup 
         Caption         =   "Supprimer"
         Height          =   375
         Left            =   4200
         TabIndex        =   16
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton cmdsuivant 
         Caption         =   "suivant"
         Height          =   375
         Left            =   4200
         TabIndex        =   15
         Top             =   2280
         Width           =   1455
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Left            =   1920
         Max             =   999
         Min             =   1
         TabIndex        =   13
         Top             =   2400
         Value           =   1
         Width           =   255
      End
      Begin VB.TextBox txtnum 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   12
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txttel 
         Height          =   375
         Left            =   5760
         TabIndex        =   10
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtcodpost 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox txtadd 
         Height          =   645
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtprenom 
         Height          =   285
         Left            =   5760
         TabIndex        =   5
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtnom 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Commentaires:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Photo"
         Height          =   255
         Left            =   4680
         TabIndex        =   14
         Top             =   1680
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   1695
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Numéro : "
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Téléphone :"
         Height          =   255
         Left            =   4680
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Code postal :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Adresse :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Prénom :"
         Height          =   255
         Left            =   4680
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nom : "
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Menu mnufichier 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuopen 
         Caption         =   "0uvrir"
      End
      Begin VB.Menu mnusaveas 
         Caption         =   "Sauver sous"
      End
      Begin VB.Menu mnuconfig 
         Caption         =   "&Configuration"
      End
      Begin VB.Menu mnuimport 
         Caption         =   "Importer"
         Begin VB.Menu mnuselect 
            Caption         =   "Selectionner une source"
         End
         Begin VB.Menu mnuimporter 
            Caption         =   "Importer"
         End
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Im&primer"
         Begin VB.Menu mnuprintlist 
            Caption         =   "&Liste"
         End
         Begin VB.Menu mnuphoto 
            Caption         =   "&Photos"
         End
      End
      Begin VB.Menu mnuquit 
         Caption         =   "Quitter"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmprinc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim finpage As Boolean

Private Sub cmdImport_Click()
If config1.image = False Then
    r = TWAIN_AcquireToClipboard(Me.hWnd, 0)
    Image1.Picture = Clipboard.GetData(vbCFDIB)
    Clipboard.Clear
Else
    cdchph.DialogTitle = "Choisissez un BitMap à Insérer"
    cdchph.filename = ""
    cdchph.Action = 1
    Image1.Picture = LoadPicture(cdchph.filename)
End If

If Image1.Picture <> 0 Then
    file = fileamoi.path1 & "\" & VScroll1.Value & ".bmp"
    enregistrement(VScroll1.Value).photo = file
    SavePicture Image1.Picture, file
    If config1.imprimauto = True Then
    Printer.PaintPicture Image1.Picture, 10, 10, 30, 30, 40, 40, 30, 30
    If finpage = True Then
        Printer.NewPage
        finpage = False
        
    Else
        finpage = True
    End If
        
    
    End If
    End If

End Sub

Private Sub Form_Load()
txtnum.Text = VScroll1.Value
file = App.Path & "\pat.csv"
Charger (file)



file = App.Path & "\Abihome.ini"
Open file For Input As #1
Input #1, config1.image, config1.scan, config1.imprimauto

Close #1


End Sub
Private Sub cmdsup_Click()
If MsgBox("Voulez-vous supprimer l'enregistrement ?", vbOKCancel, "Ecraser ?") = vbOK Then
    For i = VScroll1.Value To max
        enregistrement(i).nom = enregistrement(i + 1).nom
        enregistrement(i).prenom = enregistrement(i + 1).prenom
        enregistrement(i).add = enregistrement(i + 1).add
        enregistrement(i).codpost = enregistrement(i + 1).codpost
        enregistrement(i).tel = enregistrement(i + 1).tel
        enregistrement(i).photo = enregistrement(i + 1).photo
        enregistrement(i).comment = enregistrement(i + 1).comment
    Next i
        max = max - 1
        VScroll1.max = max
        VScroll1.Value = VScroll1.Value - 1
End If

End Sub


Private Sub cmdsuivant_Click()
max = max + 1
VScroll1.max = max
VScroll1.Value = max
End Sub

Private Sub Form_Unload(Cancel As Integer)
Sauver (fileamoi.path1 & fileamoi.filename)
End
End Sub

Private Sub mnuconfig_Click()
Frmconfig.Show 1
End Sub

Private Sub mnuimporter_Click()
cmdImport_Click
End Sub

Private Sub mnuopen_Click()
cd2.DialogTitle = "Ouverture"
cd2.Action = 1
If Not (cd2.CancelError) Then
    file = cd2.filename
    Open file For Input As #1
    Input #1, max
    For i = 1 To max
    Input #1, enregistrement(i).nom, enregistrement(i).prenom, enregistrement(i).add, enregistrement(i).codpost, enregistrement(i).tel, enregistrement(i).photo, enregistrement(i).comment
    Next i
    Close #1
    VScroll1.Value = 1
    VScroll1.max = max
    
   ' fileamoi.path = mid
End If

End Sub

Private Sub mnuprintlist_Click()
If MsgBox("Etes-vous sur de vouloir imprimer la liste des clients ?", vbYesNo, "impression") = vbYes Then

Printer.Font = "Comic sans ms"
'Printer.CurrentY = Printer.CurrentY + 10
Printer.FontSize = 10
Printer.FontItalic = True
Printer.FontBold = True
Printer.FontUnderline = True
Printer.Print "N°  Nom        Prenom    Addresse         Code postal" ', "Téléphone"
Printer.FontItalic = False
Printer.FontBold = False
Printer.FontUnderline = False
Printer.FontSize = 10
For i = 1 To max
Printer.Print i; " "; enregistrement(i).nom, enregistrement(i).prenom, enregistrement(i).add, enregistrement(i).codpost, enregistrement(i).tel, enregistrement(i).comment
'Printer.Print enregistrement(i).photo

Next i
Printer.EndDoc
End If
    
End Sub

Private Sub mnuquit_Click()

Sauver (App.Path & "\pat.csv")
End
End Sub

Private Sub mnusaveas_Click()
cd1.DialogTitle = "Sauver sous"
cd1.Action = 2

    Sauver (cd1.filename)


End Sub

Private Sub mnuselect_Click()
r = TWAIN_SelectImageSource(Me.hWnd)
End Sub

Private Sub txtadd_Change()
enregistrement(VScroll1.Value).add = txtadd.Text
End Sub

Private Sub txtcodpost_Change()
enregistrement(VScroll1.Value).codpost = txtcodpost.Text
End Sub

Private Sub txtcommentaire_Change()
enregistrement(VScroll1.Value).comment = txtcommentaire.Text

End Sub

Private Sub txtnom_Change()
enregistrement(VScroll1.Value).nom = txtnom.Text
End Sub

Private Sub txtprenom_Change()
enregistrement(VScroll1.Value).prenom = txtprenom.Text
End Sub

Private Sub txttel_Change()
enregistrement(VScroll1.Value).tel = txttel.Text
End Sub

Private Sub VScroll1_Change()
        txtnum.Text = VScroll1.Value
        txtnom.Text = enregistrement(VScroll1.Value).nom
        txtprenom.Text = enregistrement(VScroll1.Value).prenom
        txtadd.Text = enregistrement(VScroll1.Value).add
        txtcodpost.Text = enregistrement(VScroll1.Value).codpost
        txttel.Text = enregistrement(VScroll1.Value).tel
        txtcommentaire.Text = enregistrement(VScroll1.Value).comment
        If enregistrement(VScroll1.Value).photo <> "" Then
        Image1.Picture = LoadPicture(enregistrement(VScroll1.Value).photo)
        Else
            Image1.Picture = LoadPicture("")
        End If
End Sub
