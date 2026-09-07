Dim ws As Worksheet
Dim lastRow As Long
Dim i As Long

Sub AdicionarBotao()
'
' AdicionarBotao - Cria botões na planilha
'
    Set ws = ThisWorkbook.Sheets("Controle")
    
    ' Botão 1: Analisar Vencimentos
    With ws.Shapes.AddShape(msoShapeRoundedRectangle, 10, 10, 200, 40)
        .Name = "BtnAnalisar"
        .TextFrame.Characters.Text = "ANALISAR VENCIMENTOS"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 12
        .Fill.Color = RGB(0, 176, 80)
        .Line.Color = RGB(0, 0, 0)
        .OnAction = "IdentificarItensVencidos"
    End With
    
    ' Botão 2: Limpar Formatação
    With ws.Shapes.AddShape(msoShapeRoundedRectangle, 220, 10, 200, 40)
        .Name = "BtnLimpar"
        .TextFrame.Characters.Text = "LIMPAR FORMATAÇÃO"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 12
        .Fill.Color = RGB(255, 192, 0)
        .Line.Color = RGB(0, 0, 0)
        .OnAction = "LimparFormatacao"
    End With
    
    ' Botão 3: Filtrar Vencimentos
    With ws.Shapes.AddShape(msoShapeRoundedRectangle, 430, 10, 200, 40)
        .Name = "BtnFiltrar"
        .TextFrame.Characters.Text = "FILTRAR PRÓXIMOS VENCER"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 12
        .Fill.Color = RGB(255, 0, 0)
        .Line.Color = RGB(0, 0, 0)
        .OnAction = "ExibirItensPróximosVencer"
    End With
    
    MsgBox "Botões criados com sucesso!", vbInformation
End Sub

' Chamar esta sub ao abrir a planilha
Sub Workbook_Open()
    Call AdicionarBotao
End Sub
