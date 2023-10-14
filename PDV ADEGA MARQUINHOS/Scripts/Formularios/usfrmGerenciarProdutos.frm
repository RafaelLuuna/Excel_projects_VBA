VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfrmGerenciarProdutos 
   Caption         =   "Gerenciar Produtos"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6420
   OleObjectBlob   =   "usfrmGerenciarProdutos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfrmGerenciarProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()
    usfrmCompraMercadoria.Show
End Sub

Private Sub CommandButton4_Click()
    usfrmCadastroProduto.Show
End Sub
