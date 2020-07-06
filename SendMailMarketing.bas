Public WrkB                As Workbook
Public WrkS                As Worksheet

Public IntervaloMailing    As Range
Public Celula              As Range

Public AppOutk As Outlook.Application
Public MailOutk As Outlook.MailItem
Dim email As String

'Dim Account As String


Declare Sub Sleep Lib "kernel32" _
(ByVal dwMilliseconds As Long)
Public Sub MandarEmail()

Set WrkB = ThisWorkbook
Set WrkS = WrkB.Sheets("Enviar_Email")

Set IntervaloMailing = WrkS.Range("A6:A100000")


With WrkS
    .Select
        For Each Celula In IntervaloMailing
            Call CriaEmail
            Sleep (8000)
            Next
        
End With

End Sub
Sub CriaEmail()
On Error GoTo Erro
Set AppOutk = New Outlook.Application
Set MailOutk = AppOutk.CreateItem(olMailItem)
If WrkS.Cells(Celula.Row, 2) = 0 Then
MsgBox "ENVIO DOS E-MAIL'S CONCLUÍDO !", vbInformation, "Envio Concluído"
End
End If
email = WrkS.Cells(2, 2)
With MailOutk
    .Display
        .SentOnBehalfOfName = email
        .To = WrkS.Cells(Celula.Row, 2).Value
        .CC = WrkS.Cells(Celula.Row, 3).Value
        .BCC = WrkS.Cells(Celula.Row, 4).Value
        .Subject = WrkS.Cells(Celula.Row, 5).Value
        '.Body = "Prezado(a) Cliente," & vbCrLf & "Sua fatura da Conta Contrato:" & WrkS.Cells(Celula.Row, 7) & "será entregue via e-mail no endereço:" & WrkS.Cells(Celula.Row, 8) & "e sua frase de segurança é:" & WrkS.Cells(Celula.Row, 9) & "." & vbCrLf & "Caso deseje o descadastro, favor responder esse e-mail." & Signature
        .HTMLBody = "<font size=2 color=5C881A face=arial>Prezado(a) Cliente,<br >" & vbNewLine & "<font size=2  color=5C881A face=arial>Sua fatura da CC:" & " " & "<b>" & WrkS.Cells(Celula.Row, 7) & "</b>" & " " & "<font size=2  color=5C881A face=arial>será entregue via e-mail no endereço:" & " " & "<b>" & WrkS.Cells(Celula.Row, 8) & "</b>" & " " & "<font size=2  color=5C881A face=arial>e a sua frase de segurança é:" & " " & "<b>" & WrkS.Cells(Celula.Row, 9) & "." & "</b>" & vbCrLf & "<P><font size=2 color=5C881A face=arial>Caso deseje o descadastro, favor responder esse e-mail.<br></P>" & .HTMLBody

        '.HTMLBody = .HTMLBody & "<font size=2  color=5C881A face=arial>Prezado(a) Cliente,<br>" & vbCrLf

'        .HTMLBody = _
'              "<HTML>" & vbNewLine & _
'              "<BODY style=font-size:11pt;font-family:Calibri> " & vbNewLine & _
'         "<P>Prezado,</P>" & WrkS.Cells(Celula.Row, 7) & vbNewLine & _
'         "<P>Boa tarde</P>" & vbNewLine & _
'         "<P></P> " & vbNewLine & _
'         "<P></P> " & vbNewLine & _
'         "<P>Seguem documentos.</P>" & vbNewLine & _
'      "</BODY>" & vbNewLine & _
'   "</HTML>"
                    
        .Attachments.Add WrkS.Cells(Celula.Row, 10).Value
        .Attachments.Add WrkS.Cells(Celula.Row, 11).Value
        .Attachments.Add WrkS.Cells(Celula.Row, 12).Value
        .Attachments.Add WrkS.Cells(Celula.Row, 13).Value
        .Attachments.Add WrkS.Cells(Celula.Row, 14).Value
        .Attachments.Add WrkS.Cells(Celula.Row, 15).Value
Erro:
.Importance = olImportanceHigh
.Send
End With

Set MailOutk = Nothing
Set AppOutk = Nothing

End Sub
