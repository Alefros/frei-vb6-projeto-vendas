Attribute VB_Name = "Vendas"
Global CPF, StrConf, strCampo As String


Global Tab_Vend As New ADODB.Recordset

Global status As String             'TRABALHA COM A MSGBOX = box_1

Global conectar As New ADODB.Connection ' variável que trabalha com a conexão a base de dados'

Global tab_cat As New ADODB.Connection

Global tab_ufs As New ADODB.Recordset ' variável que trabalha com ajuste de registro'

Global tab_cid As New ADODB.Recordset   'ajuste o registro'

Global tab_bar As New ADODB.Recordset

Global tab_loca As New ADODB.Recordset

Global caminho As String                'TRABALHA COM A CONEXÃO A BASE DE DADOS'

Global tabcli As New ADODB.Recordset
Function box_1()
            MsgBox "Informações " & status & "com sucesso", vbExclamation
End Function
Function abrir_banco()
            caminho = "Provider=microsoft.jet.oledb.4.0;data source="
            'caminho = "Provider=microsoft.jet.oledb.4.0;data source=\\172.16.10.62\infob$\infob02\Vendas\Dados\vendas.mdb"
            'caminho = "Provider=microsoft.jet.oledb.4.0;data source=V:\Amauri\Vendas\Dados\vendas.mdb"
            caminho = caminho + App.Path & "\vendas.mdb"
            conectar.Open (caminho)
            
End Function

Function CalculaCPF()
         Dim I As Integer
         Dim strCaracter As String
         Dim intNumero As Integer
         Dim intMais As Integer
         Dim lngSoma As Long
         Dim dblDivisao As Double
         Dim lngInteiro As Long
         Dim intResto As Integer
         Dim intDig1 As Integer
         Dim intDig2 As Integer

         lngSoma = 0
         intNumero = 0
         intMais = 0
         
         'Inicia cálculos do 1º dígito
         For I = 2 To 10
             strCaracter = Right(strCampo, I - 1)
             intNumero = Left(strCaracter, 1)
             intMais = intNumero * I
             lngSoma = lngSoma + intMais
        Next I
        dblDivisao = lngSoma / 11

        lngInteiro = Int(dblDivisao) * 11
        intResto = lngSoma - lngInteiro
        If intResto = 0 Or intResto = 1 Then
           intDig1 = 0
        Else
           intDig1 = 11 - intResto
        End If

        strCampo = strCampo & intDig1
        lngSoma = 0
        intNumero = 0
        intMais = 0

        'Inicia cálculos do 2º dígito
        For I = 2 To 11
            strCaracter = Right(strCampo, I - 1)
            intNumero = Left(strCaracter, 1)
            intMais = intNumero * I
            lngSoma = lngSoma + intMais
        Next I
        dblDivisao = lngSoma / 11
        lngInteiro = Int(dblDivisao) * 11
        intResto = lngSoma - lngInteiro
        If intResto = 0 Or intResto = 1 Then
           intDig2 = 0
        Else
           intDig2 = 11 - intResto
        End If
        StrConf = intDig1 & intDig2
End Function

