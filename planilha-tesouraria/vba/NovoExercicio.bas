Attribute VB_Name = "ModNovoExercicio"
' Macro Novo Exercício — Planilha de Tesouraria Igreja
' Importar: Alt+F11 → Arquivo → Importar → selecione este arquivo

Sub NovoExercicio()
    Dim resp As VbMsgBoxResult
    Dim novoAno As Long
    Dim ws As Worksheet
    Dim d As Date
    Dim row As Long
    Dim domingoNum As Long

    resp = MsgBox( _
        "ATENÇÃO: Esta ação irá limpar todos os lançamentos do ano atual." & vbCrLf & vbCrLf & _
        "Serão mantidas: fórmulas, gráficos, Dashboard, categorias e configurações." & vbCrLf & vbCrLf & _
        "Deseja iniciar um novo exercício financeiro?", _
        vbYesNo + vbExclamation, "Novo Exercício")

    If resp = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo ErroHandler

    ' Arquivar totais no comparativo (ano atual → coluna anterior)
    With Sheets("Comparativo Anual")
        .Range("C3:C8").Value = .Range("D3:D8").Value
        .Range("D3:D8").ClearContents
    End With

    ' Limpar lançamentos
    Sheets("Dízimos Mensais").Range("A3:H502").ClearContents
    Sheets("Doações").Range("A3:G502").ClearContents
    Sheets("Arrecadação das Missas").Range("A3:G502").ClearContents
    Sheets("Gastos").Range("A3:I502").ClearContents
    Sheets("Conciliação Bancária").Range("A3:F502").ClearContents

    ' Incrementar ano
    novoAno = Sheets("Configurações").Range("B3").Value + 1
    Sheets("Configurações").Range("B3").Value = novoAno

    ' Zerar saldos iniciais (ajuste manualmente após)
    Sheets("Configurações").Range("B8:B9").ClearContents

    ' Preencher domingos do novo ano
    Set ws = Sheets("Arrecadação das Missas")
    ws.Range("A3:G502").ClearContents
    d = DateSerial(novoAno, 1, 1)
    Do While Weekday(d) <> vbSunday
        d = d + 1
    Loop
    domingoNum = 1
    row = 3
    Do While Year(d) = novoAno
        ws.Cells(row, 1).Value = d
        ws.Cells(row, 2).Value = "Domingo " & domingoNum
        row = row + 1
        d = d + 7
        domingoNum = domingoNum + 1
    Loop

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Novo exercício " & novoAno & " iniciado com sucesso!" & vbCrLf & vbCrLf & _
           "Próximos passos:" & vbCrLf & _
           "1. Informe os saldos iniciais (Caixa e Banco)" & vbCrLf & _
           "2. Defina as metas mensais de arrecadação" & vbCrLf & _
           "3. Cadastre os dizimistas", _
           vbInformation, "Concluído"
    Exit Sub

ErroHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Erro ao iniciar novo exercício: " & Err.Description, vbCritical
End Sub
