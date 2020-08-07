VERSION 5.00
Begin VB.Form Frmconfig 
   Caption         =   "Configuration"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Imprimer : "
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   8535
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimer automatiquement les images après récupération ?"
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.CommandButton cmdaide 
      Caption         =   "A&ide"
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdannuler 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Images :"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.OptionButton optscan 
         Caption         =   "Un Scanner ou Appareil photo numérique."
         Height          =   195
         Left            =   1080
         TabIndex        =   3
         Top             =   480
         Width           =   3375
      End
      Begin VB.OptionButton opthdd 
         Caption         =   "un fichier sur disque dur."
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Importer d' : "
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "Frmconfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
config1.imprimauto = (Check1.Value = 1)
End Sub

Private Sub cmdaide_Click()
MsgBox "Désolé Programme en cours de Conception"
End Sub

Private Sub cmdannuler_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
file = App.Path & "\Abihome.ini"
Open file For Output As #1
Write #1, config1.image, config1.imprimauto

Close #1
Unload Me
End Sub


Private Sub Form_Load()
If config1.image = True Then
    opthdd = True
 Else
    optscan = True
End If
If config1.imprimauto = True Then
    Check1.Value = 1
End If
End Sub

Private Sub opthdd_Click()
config1.image = True
End Sub

Private Sub optscan_Click()
config1.image = False
End Sub
