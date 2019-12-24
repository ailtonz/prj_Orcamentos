Attribute VB_Name = "basType"
Public banco(2) As infBanco
Public proposta(8) As infProposta
Public Orcamento As infOrcamento
Public grand As infGrand

Type infGrand

    strProfissao As String
    strNome As String
    strValorLiquido As String
    strOperacao As String
    
End Type

Type infFiltro

    strCampo As String
    strValor As String
        
End Type

Type infBanco

    strSource As String
    strDriver As String
    strLocation As String
    strDatabase As String
    strUser As String
    strPassword As String
    strPort As String
    
    strTabela As String
    
    strFiltro As infFiltro
    
End Type


Type infProposta

    strControle As String
    strCliente As String
    strResponsavel As String
    strProjeto As String
    strJournal As String
    strAutor As String
    strPublisher As String
        
End Type

Type infOrcamento

    strOperator As String
    strControle As String
    strVendedor As String
    strStatus As String
        
End Type
