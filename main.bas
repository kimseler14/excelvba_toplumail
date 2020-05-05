Sub mailgonder()
If MsgBox("Kaydetmek istediğinden emin misin? İşlem bitene kadar durduramazsın.", vbYesNo) = vbNo Then Exit Sub
    Dim rng As Range
    Dim filename As String
    Dim filename2 As String
    Dim strflm As String
    Dim sayı As Integer
    Dim saveLocation As String
    Dim newMail As CDO.Message
    Dim mailConfiguration As CDO.Configuration
    Dim fields As Variant
    Dim msConfigURL As String
    Dim mailadresi As String
    Dim ay As String
    
    
    saveLocation = ActiveWorkbook.Path & "\"
    
    For sayı = 1 To 5
        
        Range("M3").Select
        ActiveCell.FormulaR1C1 = sayı
        
        ay = Range("N2").Text
        filename = Range("B3").Text
        mailadresi = Range("K20").Text
        If filename = "#YOK" Then
            Do
                sayı = sayı + 1
                Range("M3").Select
                ActiveCell.FormulaR1C1 = sayı
                filename = Range("B3").Text
                mailadresi = Range("K20").Text
            Loop Until filename <> "#YOK"
        End If
        
        filename2 = Range("A2").Text
        strflm = filename2 & " " & filename & ".pdf"
    
        Set rng = Range("A1:D22")

        rng.ExportAsFixedFormat Type:=xlTypePDF, _
        filename:=saveLocation & strflm, IgnorePrintAreas:=True, Quality:=xlQualityStandard
        
        
        
        
        
        
        Set newMail = New CDO.Message
        Set mailConfiguration = New CDO.Configuration
        
        mailConfiguration.Load -1
        
        Set fields = mailConfiguration.fields
        
        With newMail
            .Subject = "HABER KAĞIDI " & ay
            .From = "<optional>"
            .To = mailadresi
            .cc = ""
            .BCC = ""
            ' To set email body as HTML, use .HTMLBody
            ' To send a complete webpage, use .CreateMHTMLBody
            .TextBody = "Haber Kağıdı Ektedir. " & strflm
            .AddAttachment saveLocation & strflm
    
    
        End With
        
        msConfigURL = "http://schemas.microsoft.com/cdo/configuration"
        
        With fields
            .Item(msConfigURL & "/smtpusessl") = True
            .Item(msConfigURL & "/smtpauthenticate") = 1
            
            .Item(msConfigURL & "/smtpserver") = "smtp.gmail.com"
            .Item(msConfigURL & "/smtpserverport") = 465
            .Item(msConfigURL & "/sendusing") = 2
            
            .Item(msConfigURL & "/sendusername") = ".....................@gmail.com"
            .Item(msConfigURL & "/sendpassword") = "..........................."
            
            .Update
        
        End With
        
        newMail.Configuration = mailConfiguration
        newMail.Send
        
        MsgBox "Mail Gönderildi" & strflm, vbInformation

        
        

        

        

Next sayı
End Sub

