VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit:
Private DATA_TIPO As String
Private DATA_NOME As String
Private DATA_QNT As Integer
Private DATA_LINHA As Integer
Private DATA_VALOR As Double
Private DATA_PRECO_VENDA As Double
Private DATA_DESCONTO As Double
Private DATA_PAGAMENTO As String
Private DATA_TAMANHO As String
'

Property Get Tipo()
    Tipo = DATA_TIPO
End Property

Property Let Tipo(Value)
    DATA_TIPO = Value
End Property

Property Get Nome()
    Nome = DATA_NOME
End Property

Property Let Nome(Value)
    DATA_NOME = Value
End Property

Property Get Qnt()
    Qnt = DATA_QNT
End Property

Property Let Qnt(Value)
    DATA_QNT = Value
End Property

Property Get Valor()
    Valor = DATA_VALOR
End Property
Property Let Valor(Value)
    DATA_VALOR = Value
End Property

Property Get Linha()
    Linha = DATA_LINHA
End Property

Property Let Linha(Value)
    DATA_LINHA = Value
End Property

Property Get PrecoVenda()
    PrecoVenda = DATA_PRECO_VENDA
End Property

Property Let PrecoVenda(Value)
    DATA_PRECO_VENDA = Value
End Property

Property Get Desconto()
    Desconto = DATA_DESCONTO
End Property

Property Let Desconto(Value)
    DATA_DESCONTO = Value
End Property

Property Get MetodoPagamento()
    MetodoPagamento = DATA_PAGAMENTO
End Property

Property Let MetodoPagamento(Value)
    DATA_PAGAMENTO = Value
End Property

Property Get Tamanho()
    Tamanho = DATA_TAMANHO
End Property

Property Let Tamanho(Value)
    DATA_TAMANHO = Value
End Property


Property Get ValorUnitario()
    ValorUnitario = DATA_VALOR / DATA_QNT
End Property


