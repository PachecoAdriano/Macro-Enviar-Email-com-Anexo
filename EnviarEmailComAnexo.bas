Attribute VB_Name = "Módulo1"
Sub ENVIAR_EMAIL()
    Dim MyOlapp     As Object, MeuItem As Object
    Dim valorData   As String
    Dim Cliente     As String
    Dim Email       As String
    Dim Linha       As Integer
    Dim PauseTime   As Integer
    Dim Start       As Single
    
    Linha = Sheets("Planilha1").Cells(Sheets("Planilha1").Rows.Count, 1).End(xlUp).Row
    valorData = Range("F5").Value
    
    Set MyOlapp = CreateObject("Outlook.Application")
    PauseTime = Range("F7")
    
    On Error Resume Next
    
    Do While Linha >= 2
        Cliente = Range("A" & Linha)
        Email = Range("C" & Linha)
    
        Set MeuItem = MyOlapp.CreateItem(olMailItem)
        With MeuItem
            
            .to = Email
            .CC = Range("F3").Value
            .Subject = "RELATORIO MENSAL " & Cliente
            .Attachments.Add Range("F9") & Trim$(Cliente) & Trim$(valorData) & ".PDF"
            
            If (Err.Number <> 0) Then
                Range("H" & Linha).Value = Cliente
                Err.Clear
            Else
                .Display
                .Send
            End If
        End With
    
        
        Start = Timer    ' Set start time.
        Do While Timer < Start + PauseTime
            DoEvents    ' Yield to other processes.
        Loop
        
    
        Linha = Linha - 1
    Loop
    
    MsgBox "Ufa, acabou!"

End Sub



