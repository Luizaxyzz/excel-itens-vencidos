Sub IdentificarItensVencidos()
'
' IdentificarItensVencidos - Macro para identificar itens vencidos
' Formatação condicional baseada em datas
'

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dataVencimento As Date
    Dim hoje As Date
    Dim diasRestantes As Long
    
    Set ws = ThisWorkbook.Sheets("Controle")
    hoje = Date
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' Limpar formatação anterior
    ws.Range("A2:G" & lastRow).Interior.ColorIndex = xlNone
    
    ' Processar cada linha
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value <> "" Then
            dataVencimento = ws.Cells(i, 3).Value
            diasRestantes = DateDiff("d", hoje, dataVencimento)
            
            ' Atualizar coluna de status
            ws.Cells(i, 6).Value = diasRestantes
            
            ' Aplicar formatação condicional
            If diasRestantes < 0 Then
                ' VENCIDO - Vermelho
                ws.Range("A" & i & ":G" & i).Interior.Color = RGB(255, 0, 0)
                ws.Range("A" & i & ":G" & i).Font.Color = RGB(255, 255, 255)
                ws.Range("A" & i & ":G" & i).Font.Bold = True
                ws.Cells(i, 7).Value = "VENCIDO"
                
            ElseIf diasRestantes <= 7 Then
                ' VENCIMENTO PRÓXIMO - Amarelo
                ws.Range("A" & i & ":G" & i).Interior.Color = RGB(255, 255, 0)
                ws.Range("A" & i & ":G" & i).Font.Color = RGB(0, 0, 0)
                ws.Range("A" & i & ":G" & i).Font.Bold = True
                ws.Cells(i, 7).Value = "ATENÇÃO"
                
            ElseIf diasRestantes <= 30 Then
                ' VENCIMENTO EM BREVE - Laranja
                ws.Range("A" & i & ":G" & i).Interior.Color = RGB(255, 192, 0)
                ws.Range("A" & i & ":G" & i).Font.Color = RGB(255, 255, 255)
                ws.Cells(i, 7).Value = "PRÓXIMO"
                
            Else
                ' OK - Verde
                ws.Range("A" & i & ":G" & i).Interior.Color = RGB(0, 176, 80)
                ws.Range("A" & i & ":G" & i).Font.Color = RGB(255, 255, 255)
                ws.Cells(i, 7).Value = "OK"
            End If
        End If
    Next i
    
    MsgBox "✓ Análise concluída!" & vbCrLf & "Itens vencidos e próximos de vencer foram destacados.", vbInformation

End Sub

Sub LimparFormatacao()
'
' LimparFormatacao - Remove formatações e cores
'
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("Controle")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ws.Range("A2:G" & lastRow).Interior.ColorIndex = xlNone
    ws.Range("A2:G" & lastRow).Font.Color = RGB(0, 0, 0)
    ws.Range("A2:G" & lastRow).Font.Bold = False
    
    MsgBox "Formatação removida com sucesso!", vbInformation
End Sub

Sub ExibirItensPróximosVencer()
'
' ExibirItensPróximosVencer - Filtra e exibe apenas itens próximos de vencer
'
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim contador As Long
    
    Set ws = ThisWorkbook.Sheets("Controle")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    contador = 0
    
    ' Remover filtro anterior
    On Error Resume Next
    ws.AutoFilter.ShowAllData
    On Error GoTo 0
    
    ' Aplicar novo filtro
    ws.Range("A1:G" & lastRow).AutoFilter
    ws.Range("A1:G" & lastRow).AutoFilter Field:=7, Criteria1:="<>OK"
    
    ' Contar itens
    For i = 2 To lastRow
        If ws.Cells(i, 7).Value <> "OK" Then
            contador = contador + 1
        End If
    Next i
    
    MsgBox "Total de itens que requerem atenção: " & contador, vbInformation
End Sub
