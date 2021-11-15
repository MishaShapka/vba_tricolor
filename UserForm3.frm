VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11640
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long





'---------------------------------------------------------------------------------------
' Module      : Final
' Author      : Mikhail Shapka/mishashapka@icloud.com
' Date        : 26.01.2021
'---------------------------------------------------------------------------------------





Private Sub CommandButton1_Click()
    Sheets.Add.Name = "Общий"
        For i = 1 To Sheets.Count
            If Sheets(i).Name <> "Общий" Then
               myR_Total = Sheets("Общий").Range("A" & Sheets("Общий").Rows.Count).End(xlUp).Row
               myR_i = Sheets(i).Range("A" & Sheets(i).Rows.Count).End(xlUp).Row
               Sheets(i).Rows("2:" & myR_i).Copy Destination:=Sheets("Общий").Range("A" & myR_Total + 1)
            End If
        Next
End Sub

Private Sub CommandButton100_Click()
Sheets("Лист2").Cells().Copy
    Range("a3").PasteSpecial Paste:=xlPasteValues
End Sub

Private Sub CommandButton101_Click()

'n = Workbooks("Table.xlsx").Sheets("отправления").Cells(Rows.Count, 1).End(xlUp).Row
'f = Cells(Rows.Count, 1).End(xlUp).Row
'
'
'For i = 1 To f
'    n = n + 1
'    Rows(i).Copy
'    Workbooks("Table.xlsx").Sheets("отправления").Rows(n).PasteSpecial Paste:=xlPasteValues
'
'Next i


n = Workbooks("Table.xlsx").Sheets("отправления").Cells(Rows.Count, 1).End(xlUp).Row
f = Cells(Rows.Count, 1).End(xlUp).Row

    Range("a1:o" & f).Copy
    Workbooks("Table.xlsx").Sheets("отправления").Range("a" & n + 1).PasteSpecial Paste:=xlPasteValues

    
End Sub

Private Sub CommandButton102_Click()
    f = Cells(Rows.Count, 2).End(xlUp).Row
    
    For i = 1 To f
        If Range("c" & i).Interior.Pattern = xlNone Then
            If Range("d" & i) = "Вернуть" Then
                Set objOL = CreateObject("Outlook.Application")
                Set objMail = objOL.CreateItem(olMailItem)
                With objMail
                .Display
                .To = "oa.pichmanova@ponyexpress.ru; ii.bayramgulova@ponyexpress.ru"
                .CC = "ChuchalovVY@monobrand-tt.ru"
                .Subject = Range("c" & i)
                .HTMLBody = "<p>Возвращаем.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                '.DeferredDeliveryTime = Date + 17 / 24
                .send
                End With
                Set objMail = Nothing
                Set objOL = Nothing
                
                Range("c" & i).Interior.Color = RGB(146, 208, 80)
                
                
                
            ElseIf Range("d" & i) = "Когда доставят" Then
                Set objOL = CreateObject("Outlook.Application")
                Set objMail = objOL.CreateItem(olMailItem)
                With objMail
                .Display
                .To = "oa.pichmanova@ponyexpress.ru; ii.bayramgulova@ponyexpress.ru"
                .CC = "ChuchalovVY@monobrand-tt.ru"
                .Subject = Range("c" & i)
                .HTMLBody = "<p>Подскажите, когда планируется доставка?</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                '.DeferredDeliveryTime = Date + 17 / 24
                .send
                End With
                Set objMail = Nothing
                Set objOL = Nothing
                Range("c" & i).Interior.Color = RGB(146, 208, 80)
                
                
                
                
            ElseIf Left(Range("d" & i), 1) = "7" Then
                Set objOL = CreateObject("Outlook.Application")
                Set objMail = objOL.CreateItem(olMailItem)
                With objMail
                    .Display
                    .To = "oa.pichmanova@ponyexpress.ru; ii.bayramgulova@ponyexpress.ru"
                    .CC = "ChuchalovVY@monobrand-tt.ru"
                    .Subject = Range("c" & i)
                    .HTMLBody = "<p>Верный номер телефона - " & Range("d" & i) & " </p>" _
                    & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                    & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                    & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                    & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                    & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                    & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                    & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                    & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                    & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                    & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                    & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                    & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
    
                    '.DeferredDeliveryTime = Date + 17 / 24
                    .send
                End With
                Set objMail = Nothing
                Set objOL = Nothing
            
            
                Range("c" & i).Interior.Color = RGB(146, 208, 80)
            
            
            Else
                Set objOL = CreateObject("Outlook.Application")
                Set objMail = objOL.CreateItem(olMailItem)
                With objMail
                .Display
                .To = "oa.pichmanova@ponyexpress.ru; ii.bayramgulova@ponyexpress.ru"
                .CC = "ChuchalovVY@monobrand-tt.ru"
                .Subject = Range("c" & i)
                .HTMLBody = "<p>Ольга, поступила информация от Кц:</p>" _
                & "<p>" & Range("d" & i) & "</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                '.DeferredDeliveryTime = Date + 17 / 24
                '.Send
                End With
                Set objMail = Nothing
                Set objOL = Nothing
            
                Range("c" & i).Interior.Color = RGB(146, 208, 80)
            
            
            End If
        End If
    Next i
    



End Sub

Private Sub CommandButton103_Click()


f = Workbooks("Table.xlsx").Sheets("отправления").Cells(Rows.Count, 1).End(xlUp).Row
For i = 1 To f
    If Range("b1") = Workbooks("Table.xlsx").Sheets("отправления").Range("g" & i) And Workbooks("Table.xlsx").Sheets("отправления").Range("b" & i) = "Отправление" Then
        datazakaza = Workbooks("Table.xlsx").Sheets("отправления").Range("f" & i) - 1
        imyazakaza = Workbooks("Table.xlsx").Sheets("отправления").Range("i" & i)
        dataotgruzki = Workbooks("Table.xlsx").Sheets("отправления").Range("f" & i)
    End If
    
    If Range("b1") = Workbooks("Table.xlsx").Sheets("отправления").Range("g" & i) And Workbooks("Table.xlsx").Sheets("отправления").Range("b" & i) = "Возврат" Then
        datavozvrata = Workbooks("Table.xlsx").Sheets("отправления").Range("f" & i)
        vozvratnakladnaya = Workbooks("Table.xlsx").Sheets("отправления").Range("h" & i)
    End If
Next i
        
        
        
g = Workbooks("22-50242.xlsx").Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
For i = 1 To g
    If Workbooks("22-50242.xlsx").Sheets(1).Range("b" & i) = Range("b2") Then
        dataschetotpravlenie = Workbooks("22-50242.xlsx").Sheets(1).Range("c" & i)
        
        If Workbooks("22-50242.xlsx").Sheets(1).Range("ao" & i) <> "3,36" Then
            nomerscheta2250242 = Workbooks("22-50242.xlsx").Sheets(1).Range("ap" & i)
            nakladnayaotpravleniya = Range("b2")
            summotpravleniya = Workbooks("22-50242.xlsx").Sheets(1).Range("ao" & i)
        
        End If
        
        klient = Workbooks("22-50242.xlsx").Sheets(1).Range("m" & i)
        adressdostavki = Workbooks("22-50242.xlsx").Sheets(1).Range("f" & i)
        
    End If
Next i



g = Workbooks("22-50447.xlsx").Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
For i = 1 To g
    If Workbooks("22-50447.xlsx").Sheets(1).Range("b" & i) = "26-3514-4094" Then
        dataschetvozvrata = Workbooks("22-50447.xlsx").Sheets(1).Range("c" & i)
        nomerscheta2250447 = Workbooks("22-50447.xlsx").Sheets(1).Range("ap" & i)
        summvozvrata = Workbooks("22-50447.xlsx").Sheets(1).Range("ao" & i)
    End If
Next i


Range("a1") = "Номер заказа"
Range("a2") = "Номер накладной"

Range("a3") = datazakaza
Range("b3") = "Заказчик передал исполнителю заявку на доставку отправления: " & imyazakaza & ", с указанием необходимости принять денежные средства за отправление. Получатель " & klient & ", " & adressdostavki

Range("a4") = dataotgruzki
Range("b4") = "Отправление было отгружено."

Range("a5") = dataschetotpravlenie
Range("b5") = "Исполнителем был выставлен счёт №" & nomerscheta2250242 & ", включающий в себя оплату за доставку по накладной " & nakladnayaotpravleniya & " в размере " & summotpravleniya & "."

Range("b6") = "Поступило обращение №: "

Range("b7") = "Запросили информацию у специалиста по работе с крупными клиентами, Пичмановой Ольги, касательно этого заказа."

Range("b8") = "Поступил ответ от специалиста по работе с крупными клиентами, Пичмановой Ольги: "

Range("a9") = datavozvrata
Range("b9") = "Отправление было возвращено."


Range("a10") = dataschetvozvrata
Range("b10") = "Исполнителем был выставлен счёт №" & nomerscheta2250447 & ", включающий в себя оплату за возврат по накладной " & vozvratnakladnaya & " в размере " & summvozvrata & "."

Range("a11") = "Просим вернуть"
Range("b11") = summvozvrata + summotpravleniya
















    
    
End Sub

Private Sub CommandButton104_Click()

End Sub

Private Sub CommandButton105_Click()
    Application.DisplayAlerts = False
    For i = Sheets.Count To 1 Step -1
        If Sheets(i).Name <> "main" Then
                Sheets(i).Delete
         End If
    Next
    Application.DisplayAlerts = True
End Sub

Private Sub CommandButton106_Click()



Dim objOutlook As Object, objNamespace As Object
    Dim objFolder As Object, objMail As Object
    Dim iRow&, iCount&, IdMail$
    
    iRow = Cells(Rows.Count, "A").End(xlUp).Row
    iCount = Application.Max(Range("A:A"))
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(6).Folders("Pony Express") '6=olFolderInbox
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    
    
    For Each objMail In objFolder.Items
        IdMail = objMail.EntryID
    
        f = Cells(Rows.Count, 3).End(xlUp).Row
        
        For i = 1 To f
            If Range("c" & i).Interior.Pattern = xlNone Then
            
                If objMail.Subject = "RE: " & Range("c" & i) Or objMail.Subject = Range("c" & i) Then
                    Range("c" & i).Interior.Color = RGB(255, 255, 0)
                    
                    
                    
                    
                    
                    
                    Set objOL = CreateObject("Outlook.Application")
                    Set objMail = objOL.CreateItem(olMailItem)
                    
                    Set replyall = objOL.mail.replyall
'                   With replyall
                        With replyall
                            .Display
                            .To = "oa.pichmanova@ponyexpress.ru; ii.bayramgulova@ponyexpress.ru"
                            .CC = "ChuchalovVY@monobrand-tt.ru"
                            .Subject = ActiveCell
                            .HTMLBody = "<p>Ольга, добрый день.</p>" _
                            & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                            & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                            & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                            & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                            & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                            & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                            & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                            & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                            & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                            & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                            & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                            & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
            
                            '.DeferredDeliveryTime = Date + 17 / 24
                            '.Send
                        End With
                    Set objMail = Nothing
                    Set objOL = Nothing
                    
                    
                    
                    
                  
                End If
        
            End If
        
        Next i

    
    Next
    
objOutlook.Quit
    
Application.ScreenUpdating = True




























'Dim mail 'object/mail item iterator
'Dim replyall 'object which will represent the reply email
'
'For Each mail In Outlook.Application.ActiveExplorer.Selection
'    If mail.Class = olMail Then
'        Set replyall = mail.replyall
'        With replyall
'            .Body = "26-3437-4930"  '<-- uncomment and it will delete the thread"
'            .Display
'        End With
'    End If
'Next

End Sub

Private Sub CommandButton107_Click()
    f = Cells(Rows.Count, 1).End(xlUp).Row
    ff = Workbooks("main.xlsb").Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
    
    
    For n = 1 To ff
    
        For i = 1 To f
            If Range("c" & i) = Workbooks("main.xlsb").Sheets(1).Range("a" & n) Then
                Range("c" & i).Rows.Clear
            End If
        Next i
    
    Next n
    
    

    Range("c1:c" & f).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
End Sub

Private Sub CommandButton108_Click()
    f = Cells(Rows.Count, 1).End(xlUp).Row
    ff = Workbooks("main.xlsb").Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
    
    
    For n = 1 To ff
    
        For i = 6 To f
            If Range("l" & i) = Workbooks("main.xlsb").Sheets(1).Range("a" & n) Then
                Range("l" & i).Rows.Clear
            End If
        Next i
    
    Next n
    
    

    Range("l6:l" & f).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
End Sub

Private Sub CommandButton109_Click()

 Const strStartDir = "c:\test" 'папка, с которой начать обзор файлов  Const strSaveDir = "c:\test\result" 'папка, в которую будет предложено сохранить результат  Dim wbTarget As New Workbook, wbSrc As Workbook, shSrc As Worksheet, shTarget As Worksheet, arFiles, _  i As Integer, stbar As Boolean  On Error Resume Next 'если указанный путь не существует, обзор начнется с пути по умолчанию  ChDir strStartDir  On Error GoTo 0  With Application 'меньше писанины  arFiles = .GetOpenFilename("Excel Files (*.xls), *.xls", , "Объединить файлы", , True)  If Not IsArray(arFiles) Then End 'если не выбрано ни одного файла  Set wbTarget = Workbooks.Add(template:=xlWorksheet)
 .ScreenUpdating = False
 stbar = .DisplayStatusBar
 .DisplayStatusBar = True
 .DisplayAlerts = False
 For i = 1 To UBound(arFiles)
 .StatusBar = "Обработка файла " & i & " из " & UBound(arFiles)
 Set wbSrc = Workbooks.Open(arFiles(i), ReadOnly:=True)
 For Each shSrc In wbSrc.Worksheets
 If IsNull(shSrc.UsedRange.Text) Then
 Set shTarget = wbTarget.Sheets.Add(After:=wbTarget.Sheets(wbTarget.Sheets.Count))
 shTarget.Name = shSrc.Name & "-" & i
 shSrc.Cells.Copy shTarget.Range("A1")
 End If
 Next
 wbSrc.Close False 'закрыть без запроса на сохранение  Next  .ScreenUpdating = True  .DisplayStatusBar = stbar  .StatusBar = False  If wbTarget.Sheets.Count = 1 Then 'не добавлено ни одного листа  MsgBox "В указанных книгах нет непустых листов, сохранять нечего!"
 wbTarget.Close False
 End
 Else
 .DisplayAlerts = False
 wbTarget.Sheets(1).Delete
 .DisplayAlerts = True
 End If
 On Error Resume Next 'если указанный путь не существует и его не удается создать,  'обзор начнется с последней использованной папки  If Dir(strSaveDir, vbDirectory) = Empty Then MkDir strSaveDir  ChDir strSaveDir  On Error GoTo 0  arFiles = .GetSaveAsFilename("Результат", "Excel Files (*.xls), *.xls", , "Сохранить объединенную книгу")  If VarType(arFiles) = vbBoolean Then 'если не выбрано имя  GoTo save_err  Else  On Error GoTo save_err  wbTarget.SaveAs arFiles  End If  End
save_err:
 MsgBox "Книга не сохранена!", vbCritical
 End With

End Sub

Private Sub CommandButton110_Click()


f = Cells(Rows.Count, 1).End(xlUp).Row
    

For i = 1 To f

    If IsEmpty(Range("f" & i)) = True Then
    
    Else
        Range("e" & i) = Range("f" & i)
        Range("f" & i).Clear
    End If
    
    
    If IsEmpty(Range("g" & i)) = True Then
    Else
        Range("e" & i) = Range("g" & i)
        Range("g" & i).Clear
    
    End If
    If IsEmpty(Range("h" & i)) = True Then
    Else
        Range("e" & i) = Range("h" & i)
        Range("h" & i).Clear
    
    End If
    
    
    
    
    Range("f" & i) = Range("e" & i)
    
    If Range("e" & i) = "Не связался. Сброс на 60 сек. " Or Range("e" & i) = "Не связался. Абонент временно недоступен. " Or Range("e" & i) = "Не связался. Автоответчик. " Or Range("e" & i) = "Не связался. Занято. " Or Range("e" & i) = "Не связался. Неправильный номер телефона. " Or Range("e" & i) = "Не связался. Сброс на 40 сек. " Then
        Range("f" & i) = "Возвращаем."

        
    End If
    
    
    

Next i

    Cells.Replace What:="Связался. ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        Columns(7).Delete



End Sub

Private Sub CommandButton111_Click()

X = Range("a1")
ActiveWorkbook.SaveAs FileName:="C:\Users\ShapkaMY\Desktop\Прозвон\" & X & " прозвон.xlsx"

Workbooks(X & " прозвон.xlsx").Close
End Sub

Private Sub CommandButton112_Click()

f = Cells(Rows.Count, 1).End(xlUp).Row


For i = 1 To f
     Range("f" & i) = Range("e" & i)
     
     
     If Range("f" & i) = "указан корректно" Then
     

        Range("f" & i).FormulaR1C1 = "=VLOOKUP(RC[-4],Статистика.csv!C6:C11,6,0)"
     End If
     
Next i




 
 
 
 
 
End Sub

Private Sub CommandButton113_Click()
X = Range("a1")
ActiveWorkbook.SaveAs FileName:="C:\Users\ShapkaMY\Desktop\Прозвон\" & X & " актуализация номеров.xlsx"

Workbooks(X & " актуализация номеров.xlsx").Close
End Sub

Private Sub CommandButton114_Click()

f = Cells(Rows.Count, 1).End(xlUp).Row

Dim rArea As Range

    For Each rArea In Range("f1:f" & f).Areas
    rArea.FormulaLocal = rArea.FormulaLocal
    Next


End Sub

Private Sub CommandButton115_Click()
ActiveCell = "Возвращаем"
ActiveCell.Offset(1).Select
End Sub

Private Sub CommandButton116_Click()

'
' Макрос1 Макрос
'

'

'    ActiveCell.FormulaR1C1 = "= & range('b'&2)"
    
Range("j1").FormulaR1C1 = "=" & Range("g3")



End Sub

Private Sub CommandButton117_Click()
Rows(1).Insert
    Range("a1") = Date
End Sub

Private Sub CommandButton118_Click()
 dp = TextBox14.Text

    Dim objOutlook As Object, objNamespace As Object
    Dim objFolder As Object, objMail As Object
    Dim iRow&, iCount&, IdMail$
    Dim X As Date
    
    iRow = Cells(Rows.Count, "A").End(xlUp).Row
    iCount = Application.Max(Range("A:A"))
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(6).Folders("Pony Express") '6=olFolderInbox
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    For Each objMail In objFolder.Items
    IdMail = objMail.EntryID
'    MsgBox (objMail.SenderName)
'    MsgBox (objMail.ReceivedTime)


    X = TextBox9.Text

    If objMail.SenderName = "Пичманова Ольга Александровна" Or objMail.SenderName = "Байрамгулова Ирина Игоревна" Then
        If objMail.ReceivedTime > X Then
            If Application.CountIf(Range("G:G"), IdMail) = 0 Then
                iRow = iRow + 1: iCount = iCount + 1
                Cells(iRow, 1) = iCount
                Cells(iRow, 2) = objMail.SenderName
                Cells(iRow, 3) = objMail.ReceivedTime
                'Cells(iRow, 3) = objMail.SenderEmailAddress
                Cells(iRow, 4) = objMail.Subject
                'Cells(iRow, 6) = objMail.CreationTime
                Cells(iRow, 5) = Left(objMail.body, 200)
                'Cells(iRow, 7) = IdMail '"'" & IdMail
                'MsgBox (objMail.CreationTime)
                
            End If
        End If
    End If
    Next
    
    objOutlook.Quit
    
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton119_Click()

    Columns("A:M").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns(1).ColumnWidth = 6
    Columns(2).ColumnWidth = 18
    Columns(3).ColumnWidth = 18
    Columns(4).ColumnWidth = 18
    Columns(5).ColumnWidth = 40
End Sub

Private Sub CommandButton12_Click()
Application.ScreenUpdating = False
Dim t As Date



    f = Cells(Rows.Count, 1).End(xlUp).Row + 10
    For i = 1 To f + 1
    
        
        
        
        
        
        If Range("i" & i) = "Комплект Звук без проводов Триколор+ Подарок (3 шт.световозвращателя)" _
        Or Range("i" & i) = "Комплект Звук без проводов Триколор+ Подарок (3 шт.световозвращателя)" _
            Then
                Range("i" & i) = "Комплект «Звук без проводов Триколор + 3 световозвращателя Триколор»"
            
            ElseIf Range("i" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7, 2 Mpix, Full HD, ИК 10м, WiFi)" Or Range("i" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (ИСХ)" _
            Then
                Range("i" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD, ИК 10м, WiFi)"
            
            ElseIf _
                Range("i" & i) = "Видеокамера IP уличная Триколор Умный дом SCO-2 (1/2,7, 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)" Or _
                Range("i" & i) = "Видеокамера IP уличная Триколор Умный дом SCO-2 (1/2,7, 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)" _
            Then
                Range("i" & i) = "Видеокамера IP уличная Триколор Умный дом SCO-2 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)"
            
            ElseIf _
                Range("i" & i) = "Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)" _
            Then
                Range("i" & i) = "Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)"
            
           
            ElseIf _
                Range("i" & i) = "Комплект усилитель сотовой связи 900/2100, Триколор, TR-900/2100-50-kit+ Подарок Органайзер для пультов ДУ и прессы" _
            Then
            Range("i" & i) = "Комплект усилитель сотовой связи 900/2100, Триколор, TR-900/2100-50-kit"
            Rows(i).Copy
            Rows(i + 1).Insert
            Rows(i + 1).Select
            Range("k" & i) = "11790"
            Range("i" & i + 1) = "Органайзер для пультов ДУ и прессы"
            Range("k" & i + 1) = "200"
    
        
            ElseIf _
                Range("i" & i) = "Комплект усилитель мобильного интернета, ""Триколор ТВ"", DS-4G-5kit+ Подарок Органайзер для пультов ДУ и прессы" _
            Then
                Range("i" & i) = "Комплект усилитель мобильного интернета, " & Chr(34) & "Триколор ТВ" & Chr(34) & ", DS-4G-5kit"
                Rows(i).Copy
                Rows(i + 1).Insert
                Rows(i + 1).Select
                Range("k" & i) = "10790"
                Range("i" & i + 1) = "Органайзер для пультов ДУ и прессы"
                Range("k" & i + 1) = "200"
            
            'Лот 1
            ElseIf _
                Range("i" & i) = "Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi), Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7, 2 Mpix, Full HD, ИК 10м, WiFi)" _
                Or Range("i" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7, 2 Mpix, Full HD, ИК 10м, WiFi), Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)" _
                Or Range("i" & i) = "Комплект камер Триколор" _
                Or Range("i" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7, 2 Mpix, Full HD, ИК 10м, WiFi), Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)" _
                Or Range("i" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7, 2 Mpix, Full HD, ИК 10м, WiFi), Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi) " _
            Then
                Range("i" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD, ИК 10м, WiFi)"
                
            If Range("j" & i) = "2" Then
                Range("j" & i) = "1"
            End If
                    
            If Range("j" & i) = "4" Then
                Range("j" & i) = "2"
            End If
            
                Rows(i).Copy
                Rows(i + 1).Insert
                Rows(i + 1).Select
                Range("k" & i) = "2400"
                Range("i" & i + 1) = "Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)"
                Range("k" & i + 1) = "3500"
                
            'Лот 2
            ElseIf _
                Range("i" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7, 2 Mpix, Full HD, ИК 10м, WiFi), Видеокамера IP уличная Триколор Умный дом SCO-2 (1/2,7, 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)" _
                Or Range("i" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7, 2 Mpix, Full HD, ИК 10м, WiFi), Видеокамера IP уличная Триколор Умный дом SCO-2 (1/2,7, 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi) " _
                Or Range("i" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7, 2 Mpix, Full HD, ИК 10м, WiFi), Видеокамера IP уличная Триколор Умный дом SCO-2 (1/2,7, 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)" _
            Then
                Range("i" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD, ИК 10м, WiFi)"
                If Range("j" & i) = "2" Then
                    Range("j" & i) = "1"
                End If
                If Range("j" & i) = "4" Then
                    Range("j" & i) = "2"
                End If
                Rows(i).Copy
                Rows(i + 1).Insert
                Rows(i + 1).Select
                Range("k" & i) = "2400"
                Range("i" & i + 1) = "Видеокамера IP уличная Триколор Умный дом SCO-2 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)"
                Range("k" & i + 1) = "3500"
                
                
            'Лот 3
            ElseIf _
                Range("i" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-1 (1/2,7"", 2 Mpix, Full HD, ИК 10м, WiFi), Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7"", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)" _
                Or Range("i" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-1 (1/2,7"", 2 Mpix, Full HD, ИК 10м, WiFi), Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7"", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi) " _
            Then
                Range("i" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD, ИК 10м, WiFi)"
                If Range("j" & i) = "2" Then
                    Range("j" & i) = "1"
                End If
                If Range("j" & i) = "4" Then
                    Range("j" & i) = "2"
                End If
                Rows(i).Copy
                Rows(i + 1).Insert
                Rows(i + 1).Select
                Range("k" & i) = "2400"
                Range("i" & i + 1) = "Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)"
                Range("k" & i + 1) = "3500"
                
        End If
    Next i
    
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton120_Click()

    f = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To f
        If Left(Range("d" & i), 3) = "RE:" Or Left(Range("d" & i), 3) = "FW:" Or Left(Range("d" & i), 9) = "Automatic" Then
        Else
        Range("d" & i).Rows.Clear
        End If
    Next i
    

    
    Range("d1:d" & f).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
'    For i = 1 To f
'        If Left(Range("d" & i), 1) = " " Then
'            Range("d" & i).Rows.Clear
'            Right(Range("d" & i),Len(str)-5)
'        End If
'    Next i
End Sub

Private Sub CommandButton121_Click()
    
f = Workbooks("Table.xlsx").Sheets("отправления").Cells(Rows.Count, 1).End(xlUp).Row
    Dim X As Date
    
    
    X = "01.04.2021"
    
    For i = 1 To f
    
    
    If Workbooks("Table.xlsx").Sheets("отправления").Range("f" & i) > X Then
'        If Workbooks("Table.xlsx").Sheets("отправления").Range("k" & i) = 0 Then
        Workbooks("Table.xlsx").Sheets("отправления").Rows(i).Copy
        Rows(1).Insert
        
'         MsgBox (Workbooks("Table.xlsx").Sheets("отправления").Rows(i))
        
'        End If
    End If
    
    
    
    
    

    Next i
    
    

    
    
End Sub

Private Sub CommandButton122_Click()


For n = 2 To 300

    For i = 1 To 300
        If Workbooks("EXPORT.xls").Sheets(1).Range("a" & n) = Workbooks("15.04.2021 ТРС Тула (реестр отправлений).xlsx").Sheets(1).Range("c" & i) Then
            Workbooks("15.04.2021 ТРС Тула (реестр отправлений).xlsx").Sheets(1).Rows(i).Copy
            Workbooks("15.04.2021 ТРС Тула (реестр отправлений).xlsx").Sheets(2).Rows(1).Insert
        End If

    Next i
Next n


End Sub

Private Sub CommandButton123_Click()



'
'For n = 1 To 466
'    For i = 1 To 466
'        If Workbooks("13.04.2021 ТРС Тула (реестр отправлений).xlsx").Sheets(1).Range("c" & i) > 0 Then
'            If Dir("C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\13.04.2021\Тула\Чеки\" & Range("c" & i) & ".pdf") = Range("c" & n) & ".pdf" Then
'            Name "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\13.04.2021\Тула\Чеки\" & Range("c" & i) & ".pdf" As ""
'            End If
'
'        End If
'    Next i
'Next n


'
'
'For i = 1 To 280
'    x = Range("b" & i)
'    Name "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\22.04.2021\Тула\Накладные\THERMOPRINT_Часть" & i & ".pdf" As "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\22.04.2021\Тула\Накладные\000" & x & ".pdf"
'
'Next i




For i = 1 To 500
    X = Range("b" & i)
    Name "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\30.04.2021\Тула\Накладные\THERMOPRINT_Часть" & i & ".pdf" As "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\30.04.2021\Тула\Накладные\000" & X & ".pdf"
Next i



'Name "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\13.04.2021\Тула\Чеки\285927.pdf" As "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\13.04.2021\Тула\Чеки\2_285927.pdf"






End Sub

Private Sub CommandButton124_Click()
    Set s = Workbooks("main.xlsb").Sheets(1)
    ddt = TextBox8.Text

    If CheckBox1.Value = True Then
        ekat = "<p>Екатеринбург:<br>Количество заказов:" & s.Range("b1") & "<br>Адрес: 620024, г. Екатеринбург, по ул. Бисертской, 145 (литер АА1).</p>"
    End If
    
    If CheckBox2.Value = True Then
        spb = "<p>Санкт-Петербург:<br>Количество заказов:" & s.Range("b2") & "<br>Адрес: 196084, г. Санкт-Петербург, Витебский пр., д. 3, лит. Б1.</p>"
    End If
    
    If CheckBox3.Value Then
        nino = "<p>Нижний Новгород:<br>Количество заказов:" & s.Range("b3") & "<br>Адрес: 603127, г.Нижний Новгород, Сормовский район, 7-й Микрорайон, Сормовский промузел, ул. Коновалова, д.10/1.</p>"
    End If
    
    If CheckBox4.Value = True Then
        novo = "<p>Новосибирск:<br>Количество заказов:" & s.Range("b4") & "<br>Адрес: 630088, г. Новосибирск, ул. Петухова, дом. 35, корпус 6.</p>"
    End If
    
    If CheckBox5.Value Then
        tula = "<p>Тула:<br>Количество заказов:" & s.Range("b5") & "<br>Адрес: 301107, Ленинский район, сельское поселение Шатское, поселок Шатск, строение 2/1.</p>"
    End If
    
    If CheckBox6.Value = True Then
        rostov = "<p>Ростов-на-Дону:<br>Количество заказов:" & s.Range("b6") & "<br>Адрес: 344092, г. Ростов-на-Дону, Стартовая,д. 3/11, Литер 'Л'.</p>"
    End If
    
    If CheckBox7.Value = True Then
        sar = "Саратов"
    End If



        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Екатеринбург"
        
            With objMail
                    .Display
                    .To = "Ksenia.Starostina@russianpost.ru;"
                    .CC = "ChuchalovVY@monobrand-tt.ru;"
                    .Subject = "Отгрузка " & TextBox1
                    .HTMLBody = "<p style='font-size: 11pt;'>Добрый день.</p>" _
                    & "<p style='font-size: 11pt;'>Отгрузка на " & TextBox1 & "</p>" _
                    & tula _
                    & spb _
                    & nino _
                    & novo _
                    & ekat _
                    & rostov _
                    & sar _
                    & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                    & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                    & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                    & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                    & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                    & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                    & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                    & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                    & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                    & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                    & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                    & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                    If CheckBox8.Value Then
                        .DeferredDeliveryTime = Date + ddt / 24
                    End If
                
                
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing






End Sub
Private Sub CommandButton126_Click()
    Dim d As Date
    
    d = "01.04.21"
    
    For pochta = 1 To 30
        X = 0
        
        For i = 1 To 30000
        If Workbooks("Table.xlsx").Sheets("отправления").Range("f" & i) > d Then
            If Workbooks("Table.xlsx").Sheets("отправления").Range("v" & i) = "Почта" Then
            
                If Workbooks("Table.xlsx").Sheets("отправления").Range("a" & i) = Range("a" & pochta) Then
                    X = X + 1
                End If
            End If
        End If
        
        Next i
        
        
    If Range("b" & pochta) = Sheets("Наименования").Range("a2") Then
        Range("e" & pochta) = X
    End If
    
    
    Next pochta
End Sub

Private Sub CommandButton125_Click()
For i = 1 To 13000
    If Left(Range("a" & i), 2) = "28" Or Left(Range("a" & i), 2) = "29" Or Left(Range("a" & i), 4) = "T-04" Then
    Else
    
    Range("a" & i).Rows.Clear
    End If
Next i
Range("A1:A13000").SpecialCells(xlCellTypeBlanks).EntireRow.Delete




End Sub

Private Sub CommandButton127_Click()
    For i = 3 To 200
        
    
        Rows(i).Insert
        
        Rows(i).Insert
        Rows(i).Insert
        Rows(i).Insert
        Rows(i).Insert

        
        i = i + 5
        

    Next i
    

    
End Sub

Private Sub CommandButton128_Click()
Range("B:B").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
End Sub

Private Sub CommandButton129_Click()
    
    Range("").Copy
        Rows(y + 1).Insert
        Rows(y + 1).Select
    
    
End Sub

Private Sub CommandButton13_Click()
Application.ScreenUpdating = False
     f = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To 300
        Range("m" & i).Copy
        Range("k" & i).Select
    Next i

    For y = 1 To 300

        
        If Range("j" & y) = 2 Then
        Range("j" & y) = "1"
        
        Rows(y).Copy
        Rows(y + 1).Insert
        Rows(y + 1).Select
        
        ElseIf Range("j" & y) = 3 Then
        Range("j" & y) = "1"
        Rows(y).Copy
        Rows(y + 1).Insert
        Rows(y + 1).Select
        Rows(y).Copy
        Rows(y + 1).Insert
        Rows(y + 1).Select
        
        ElseIf Range("j" & y) = 4 Then
        Range("j" & y) = "1"
        
        Rows(y).Copy
        Rows(y + 1).Insert
        Rows(y + 1).Select
        Rows(y).Copy
        
        Rows(y + 1).Insert
        Rows(y + 1).Select
        Rows(y).Copy
        
        Rows(y + 1).Insert
        Rows(y + 1).Select
        
        ElseIf Range("j" & y) = 5 Then
        Range("j" & y) = "1"
        
        Rows(y).Copy
        Rows(y + 1).Insert
        Rows(y + 1).Select
        Rows(y).Copy
        
        Rows(y + 1).Insert
        Rows(y + 1).Select
        Rows(y).Copy
        
        Rows(y + 1).Insert
        Rows(y + 1).Select
        
        Rows(y).Copy
        Rows(y + 1).Insert
        Rows(y + 1).Select
        End If
    Next y
      
Application.ScreenUpdating = True
End Sub
Private Sub CommandButton130_Click()
 dp = TextBox19.Text

    Dim objOutlook As Object, objNamespace As Object
    Dim objFolder As Object, objMail As Object
    Dim iRow&, iCount&, IdMail$
    Dim X As Date
    
    iRow = Cells(Rows.Count, "A").End(xlUp).Row
    iCount = Application.Max(Range("A:A"))
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(6) '.Folders("КС") '6=olFolderInbox
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    For Each objMail In objFolder.Items
    IdMail = objMail.EntryID
'    MsgBox (objMail.SenderName)
'    MsgBox (objMail.ReceivedTime)


    X = TextBox19.Text

    If objMail.SenderName = "Пичманова Ольга Александровна" Or objMail.SenderName = "Байрамгулова Ирина Игоревна" Or objMail.SenderName = "Старостина Ксения Александровна" Then
        If objMail.ReceivedTime > X Then
            If Application.CountIf(Range("G:G"), IdMail) = 0 Then
                iRow = iRow + 1: iCount = iCount + 1
                Cells(iRow, 1) = iCount
                Cells(iRow, 2) = objMail.SenderName
                Cells(iRow, 3) = objMail.ReceivedTime
                'Cells(iRow, 3) = objMail.SenderEmailAddress
                Cells(iRow, 4) = objMail.Subject
                'Cells(iRow, 6) = objMail.CreationTime
                Cells(iRow, 5) = Left(objMail.body, 200)
                'Cells(iRow, 7) = IdMail '"'" & IdMail
                'MsgBox (objMail.CreationTime)
                
            End If
        End If
    End If
    Next
    
    objOutlook.Quit
    
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton131_Click()
    X = ActiveWorkbook.Name
    Workbooks.Add
    Workbooks(X).Sheets(2).Copy before:=Sheets(1)
    
    Rows(1).Insert
    
    
    Range("a1") = "№"
    Range("b1") = "Номер заказа"
    Range("c1") = "Номер накладной Pony Express"
    Range("d1") = "Комментарий Pony Express"
    Range("e1") = "Попытка 1"
    Range("f1") = "Попытка 2"
    Range("g1") = "Попытка 3"
    
    f = Cells(Rows.Count, 2).End(xlUp).Row
    
    For i = 1 To f
        If Range("c" & i).Interior.Pattern = xlNone Then
        Else
        Range("b" & i).Rows.Clear

        End If
    Next i
    
    
    Range("B1:B" & f).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    Rows(1).Insert
    Range("a1") = Date
    
    Columns(1).ColumnWidth = 6
    Columns(2).ColumnWidth = 12
    Columns(3).ColumnWidth = 30
    Columns(4).ColumnWidth = 30
    Columns(5).ColumnWidth = 30
    Columns(6).ColumnWidth = 30
    Columns(7).ColumnWidth = 30
    
'    ActiveWorkbook.SaveAs FileName:="C:\Users\ShapkaMY\Desktop\Прозвон\" & Date & " прозвон.xlsx"
    
    
End Sub

Private Sub CommandButton132_Click()
Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
            With objMail
                .Display
                .To = "mihajlov@cc.tricolor.tv; moysya@cc.tricolor.tv; simkina@cc.tricolor.tv"
                .CC = "ChuchalovVY@monobrand-tt.ru;"
                .Subject = "Прозвон от " & Date
                .HTMLBody = "<p>Коллеги, добрый день!</p>" _
                & "<p>Прошу актуализировать данные, на запросы от КС до <b>" & Date + 1 & " 18:00</b><br>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                
                .Attachments.Add "C:\Users\ShapkaMY\Desktop\Прозвон\" & Date & " прозвон.xlsx" 'указывается полный путь к файлу
                '.DeferredDeliveryTime = Date + 17 / 24
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
End Sub

Private Sub CommandButton133_Click()
Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
            With objMail
                .Display
                .To = "mihajlov@cc.tricolor.tv; dubkova@cc.tricolor.tv; druzhinina@cc.tricolor.tv; moysya@cc.tricolor.tv; simkina@cc.tricolor.tv"
                .CC = "ChuchalovVY@monobrand-tt.ru;"
                .Subject = "Кривые заказы " & Date
                .HTMLBody = "<p>Коллеги, добрый день!</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                
'                .Attachments.Add "C:\Users\ShapkaMY\Desktop\Прозвон\" & Date & " прозвон.xlsx" 'указывается полный путь к файлу
                '.DeferredDeliveryTime = Date + 17 / 24
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
End Sub

Private Sub CommandButton134_Click()
If CheckBox9.Value = True Then
        pochta = "Почта России"
    Else
        pochta = "Pony Express"
    End If
    
    


    Trsdate = TextBox1.Text
    ddt = TextBox8.Text
    
    Trsaddress1 = "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & pochta & " Екатеринбург"
    Trsaddress2 = "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & pochta & " Санкт-Петербург"
    Trsaddress3 = "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & pochta & " Нижний Новгород"
    Trsaddress4 = "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & pochta & " Новосибирск"
    Trsaddress5 = "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & pochta & " Тула"
    Trsaddress6 = "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & pochta & " Ростов-на-Дону"
    Trsaddress7 = "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & pochta & " Саратов"
  
    
    If CheckBox9.Value Then
    Opochta = "Отгружаем через Почту России"
    
    
    End If
    
    
    
    
    If CheckBox1.Value Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Екатеринбург"
            With objMail
                .Display
                .To = "sklad1@rd.e-burg.n-l-e.ru; logist@rd.e-burg.n-l-e.ru; sklad@rd.e-burg.n-l-e.ru"
                .CC = "antipova@n-l-e.ru; ChuchalovVY@monobrand-tt.ru; BelyaevskiyKO@monobrand-tt.ru; BocharovAV@tricolor.tv"
                .Subject = "ОТПРАВКА ИНТЕРНЕТ-МАГАЗИН ООО <ТТ> " & Trsdate & " " & city & " " & pochta
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p style='background: #00FFFF;'>" & Opochta & "</p>" _
                & "<p>Прошу подготовить к отправке ТМЦ согласно вложенному реестру отправлений.<br>" _
                & "Прилагаю:</p>" _
                & "<ul><li>Реестр отправлений</li><li>Накладные</li><li>Товарные чеки</li></ul>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                
                
                'указывается текст письма
                .Attachments.Add Trsaddress1 & "\" & Trsdate & " ТРС " & city & " (реестр отправлений) " & pochta & ".xlsx" 'указывается полный путь к файлу
                If CheckBox9.Value Then
                    .Attachments.Add Trsaddress1 & "\doc\F103.pdf"
                End If
                .Attachments.Add Trsaddress1 & "\Накладные.7z"
                .Attachments.Add Trsaddress1 & "\Чеки.7z"
                
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox2.Value = True Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Санкт-Петербург"
            With objMail
                .Display
                .To = "nachskl@trs.spb.n-l-e.ru; skl@trs.spb.n-l-e.ru"
                .CC = "antipova@n-l-e.ru; ChuchalovVY@monobrand-tt.ru; BelyaevskiyKO@monobrand-tt.ru; BocharovAV@tricolor.tv"
                .Subject = "ОТПРАВКА ИНТЕРНЕТ-МАГАЗИН ООО <ТТ> " & Trsdate & " " & city & " " & pochta
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p style='background: #00FFFF;'>" & Opochta & "</p>" _
                & "<p>Прошу подготовить к отправке ТМЦ согласно вложенному реестру отправлений.<br>" _
                & "Прилагаю:</p>" _
                & "<ul><li>Реестр отправлений</li><li>Реестр накладных</li><li>Наклейки</li><li>Товарные чеки</li></ul>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                
                .Attachments.Add Trsaddress2 & "\" & Trsdate & " ТРС " & city & " (реестр отправлений) " & pochta & ".xlsx" 'указывается полный путь к файлу
                If CheckBox9.Value Then
                    .Attachments.Add Trsaddress2 & "\doc\F103.pdf"
                Else
                    .Attachments.Add Trsaddress2 & "\" & Trsdate & " ТРС " & city & " для Pony Express.xlsx" 'указывается полный путь к файлу
                End If
                
                
                
                .Attachments.Add Trsaddress2 & "\Накладные.7z"
                .Attachments.Add Trsaddress2 & "\Чеки.7z"
                
                 If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                '.Send
                
            End With
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox3.Value Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Нижний Новгород"
            With objMail
                .Display
                .To = "logist@rd.nnov.n-l-e.ru; sklad@rd.nnov.n-l-e.ru; operator@rd.nnov.n-l-e.ru"
                .CC = "antipova@n-l-e.ru; ChuchalovVY@monobrand-tt.ru; BelyaevskiyKO@monobrand-tt.ru; BocharovAV@tricolor.tv"
                .Subject = "ОТПРАВКА ИНТЕРНЕТ-МАГАЗИН ООО <ТТ> " & Trsdate & " " & city & " " & pochta
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p style='background: #00FFFF;'>" & Opochta & "</p>" _
                & "<p>Прошу подготовить к отправке ТМЦ согласно вложенному реестру отправлений.<br>" _
                & "Прилагаю:</p>" _
                & "<ul><li>Реестр отправлений</li><li>Накладные</li><li>Товарные чеки</li></ul>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                
                
                .Attachments.Add Trsaddress3 & "\" & Trsdate & " ТРС " & city & " (реестр отправлений) " & pochta & ".xlsx" 'указывается полный путь к файлу
                  If CheckBox9.Value Then
                    .Attachments.Add Trsaddress3 & "\doc\F103.pdf"
                Else
                    .Attachments.Add Trsaddress3 & "\" & Trsdate & " ТРС " & city & " для Pony Express.xlsx" 'указывается полный путь к файлу
                End If
                .Attachments.Add Trsaddress3 & "\Накладные.7z"
                .Attachments.Add Trsaddress3 & "\Чеки.7z"
                
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox4.Value = True Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Новосибирск"
            With objMail
                .Display
                .To = "sklad@trs.nvsb.n-l-e.ru; logist@trs.nvsb.n-l-e.ru; director@trs.nvsb.n-l-e.ru"
                .CC = "antipova@n-l-e.ru; ChuchalovVY@monobrand-tt.ru; BelyaevskiyKO@monobrand-tt.ru; BocharovAV@tricolor.tv"
                .Subject = "ОТПРАВКА ИНТЕРНЕТ-МАГАЗИН ООО <ТТ> " & Trsdate & " " & city & " " & pochta
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p style='background: #00FFFF;'>" & Opochta & "</p>" _
                & "<p>Прошу подготовить к отправке ТМЦ согласно вложенному реестру отправлений.<br>" _
                & "Прилагаю:</p>" _
                & "<ul><li>Реестр отправлений</li><li>Накладные</li><li>Товарные чеки</li></ul>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                
                .Attachments.Add Trsaddress4 & "\" & Trsdate & " ТРС " & city & " (реестр отправлений) " & pochta & ".xlsx" 'указывается полный путь к файлу
                If CheckBox9.Value Then
                    .Attachments.Add Trsaddress4 & "\doc\F103.pdf"
                Else
                    .Attachments.Add Trsaddress4 & "\" & Trsdate & " ТРС " & city & " для Pony Express.xlsx" 'указывается полный путь к файлу
                End If
                .Attachments.Add Trsaddress4 & "\Накладные.7z"
                .Attachments.Add Trsaddress4 & "\Чеки.7z"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox5.Value Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Тула"
            With objMail
                .Display
                .To = "logist@ts.tula.n-l-e.ru; logist2@ts.tula.n-l-e.ru; logist3@ts.tula.n-l-e.ru; operator1@ts.tula.n-l-e.ru; operator2@ts.tula.n-l-e.ru"
                .CC = "antipova@n-l-e.ru; ChuchalovVY@monobrand-tt.ru; BelyaevskiyKO@monobrand-tt.ru; BocharovAV@tricolor.tv"
                .Subject = "ОТПРАВКА ИНТЕРНЕТ-МАГАЗИН ООО <ТТ> " & Trsdate & " " & city & " " & pochta
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p style='background: #00FFFF;'>" & Opochta & "</p>" _
                & "<p>Прошу подготовить к отправке ТМЦ согласно вложенному реестру отправлений.<br>" _
                & "Прилагаю:</p>" _
                & "<ul><li>Реестр отправлений</li><li>Реестр накладных</li><li>Накладные</li><li>Товарные чеки</li></ul>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                
                .Attachments.Add Trsaddress5 & "\" & Trsdate & " ТРС " & city & " (реестр отправлений) " & pochta & ".xlsx" 'указывается полный путь к файлу
                If CheckBox9.Value Then
                    .Attachments.Add Trsaddress5 & "\doc\F103.pdf"
                Else
                    .Attachments.Add Trsaddress5 & "\" & Trsdate & " ТРС " & city & " для Pony Express.xlsx" 'указывается полный путь к файлу
                End If
                .Attachments.Add Trsaddress5 & "\Накладные.7z"
                .Attachments.Add Trsaddress5 & "\Чеки.7z"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
   
    
    If CheckBox6.Value = True Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Ростов-на-Дону"
            With objMail
                .Display
                .To = "logist@rd.rostov.n-l-e.ru; tovaroved@trs1.rostov.n-l-e.ru; sklad@trs1.rostov.n-l-e.ru"
                .CC = "antipova@n-l-e.ru; ChuchalovVY@monobrand-tt.ru; BelyaevskiyKO@monobrand-tt.ru; BocharovAV@tricolor.tv"
                .Subject = "ОТПРАВКА ИНТЕРНЕТ-МАГАЗИН ООО <ТТ> " & Trsdate & " " & city & " " & pochta
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p style='background: #00FFFF;'>" & Opochta & "</p>" _
                & "<p>Прошу подготовить к отправке ТМЦ согласно вложенному реестру отправлений.<br>" _
                & "Прилагаю:</p>" _
                & "<ul><li>Реестр отправлений</li><li>Накладные</li><li>Товарные чеки</li></ul>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                .Attachments.Add Trsaddress6 & "\" & Trsdate & " ТРС " & city & " (реестр отправлений) " & pochta & ".xlsx" 'указывается полный путь к файлу
                If CheckBox9.Value Then
                    .Attachments.Add Trsaddress6 & "\doc\F103.pdf"
                Else
                    .Attachments.Add Trsaddress6 & "\" & Trsdate & " ТРС " & city & " для Pony Express.xlsx" 'указывается полный путь к файлу
                End If
                .Attachments.Add Trsaddress6 & "\Накладные.7z"
                .Attachments.Add Trsaddress6 & "\Чеки.7z"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox7.Value = True Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Саратов"
            With objMail
                .Display
                .To = "sklad1@trs1.saratov.n-l-e.ru; sklad2@trs1.saratov.n-l-e.ru; sklad@trs1.saratov.n-l-e.ru"
                .CC = "antipova@n-l-e.ru; ChuchalovVY@monobrand-tt.ru; BelyaevskiyKO@monobrand-tt.ru; BocharovAV@tricolor.tv"
                .Subject = "ОТПРАВКА ИНТЕРНЕТ-МАГАЗИН ООО <ТТ> " & Trsdate & " " & city & " " & pochta
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p style='background: #00FFFF;'>" & Opochta & "</p>" _
                & "<p>Прошу подготовить к отправке ТМЦ согласно вложенному реестру отправлений.<br>" _
                & "Прилагаю:</p>" _
                & "<ul><li>Реестр отправлений</li><li>Накладные</li><li>Товарные чеки</li></ul>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                .Attachments.Add Trsaddress7 & "\" & Trsdate & " ТРС " & city & " (реестр отправлений) " & pochta & ".xlsx" 'указывается полный путь к файлу
                If CheckBox9.Value Then
                    .Attachments.Add Trsaddress7 & "\doc\F103.pdf"
                Else
                    .Attachments.Add Trsaddress7 & "\" & Trsdate & " ТРС " & city & " для Pony Express.xlsx" 'указывается полный путь к файлу
                End If
                .Attachments.Add Trsaddress7 & "\Накладные.7z"
                .Attachments.Add Trsaddress7 & "\Чеки.7z"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
    End If
End Sub

Private Sub CommandButton135_Click()

f = Cells(Rows.Count, 2).End(xlUp).Row

Set X = Workbooks("Прозвон 13.05.xlsx").Sheets(1)



'Range("e1:e" & f).FormulaR1C1 = _
'        "=VLOOKUP(RC[-3],'[Прозвон 13.05.xlsx]Лист1'!C2:C5,4,0)"
'
'
'
       
For i = 2 To f
If IsEmpty(X.Range("f" & i)) = True Then
    Range("e" & i) = X.Range("e" & i)

Else

    Range("e" & i).FormulaR1C1 = _
        "=VLOOKUP(RC[-3],'[Прозвон 13.05.xlsx]Лист1'!C2:C6,5,0)"
End If
Next i


End Sub

Private Sub CommandButton136_Click()


    X = Int((999999999 - 1 + 1) * Rnd + 1)
    X = Time + X - Date
Range("a1") = X

End Sub

Private Sub CommandButton137_Click()

          Selection.FormulaArray = _
        "=INDEX([main.xlsb]Итог!C15,MATCH(RC[1]&RC[-6],[main.xlsb]Итог!C16&[main.xlsb]Итог!C9,0))"
End Sub

Private Sub CommandButton138_Click()

    Trsdate = TextBox1.Text
    
    If CheckBox9.Value = True Then
        pochta = "Почта России"
    Else
        pochta = "Pony Express"
    End If


    
    
'    i = 1
'
'    If CheckBox & i & .Value Then
''        cityks = pochta & " " & city
''        Workbooks.Open FileName:="C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & cityks & "\" & Trsdate & " ТРС " & city & " (реестр отправлений) " & pochta & ".xlsx"
''    Call CommandButton25_Click
''        Windows(Trsdate & " ТРС " & city & " (реестр отправлений) " & pochta & ".xlsx").Close True
'    MsgBox ("ok")
'
'    End If



'For Each chk In ActiveSheet.CheckBoxes
'    MsgBox chk.Name
'Next


For i = 1 To 5


If CheckBox1.Value Then
ch = CheckBox1
End If



ch = CheckBox1


    If ch.Value Then
        MsgBox (i)
    End If
Next i

End Sub

Private Sub CommandButton139_Click()

    f = Cells(Rows.Count, 2).End(xlUp).Row
    
    For i = 1 To f
        If Range("c" & i).Interior.Pattern = xlNone Then
                Set objOL = CreateObject("Outlook.Application")
                Set objMail = objOL.CreateItem(olMailItem)
                    With objMail
                        .Display
                        
                        
                        If Range("a" & i) = "кц" Then
                            .To = "simkina@cc.tricolor.tv; moysya@cc.tricolor.tv; mihajlov@cc.tricolor.tv"
                        ElseIf Range("a" & i) = "др" Then
                            .To = "dubkova@cc.tricolor.tv; druzhinina@cc.tricolor.tv; simkina@cc.tricolor.tv; i.smirnova@cc.tricolor.tv"
                        
                        Else
                            If Left(Range("c" & i), 3) = "26-" Then
                                .To = "oa.pichmanova@ponyexpress.ru"
                            ElseIf Left(Range("c" & i), 3) = "800" Then
                                .To = "Ksenia.Starostina@russianpost.ru; Marina.Darovskaya@russianpost.ru; Biryukova.Julia@russianpost.ru"
                            End If
                        
                        End If
                        
                        


                        X = "<p>Номер заказа: " & Range("b" & i) & "<br>Номер накладной: " & Range("c" & i) & "<br>Описание ситуации: " & Range("d" & i) & "</p>"
                        
                        .CC = "ChuchalovVY@monobrand-tt.ru"
                        .Subject = Range("b" & i) & "/" & Range("c" & i)
                        .HTMLBody = X _
                        & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                        & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                        & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                        & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                        & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                        & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                        & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                        & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                        & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                        & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                        & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                        & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
    
                        '.DeferredDeliveryTime = Date + 17 / 24
                        '.Send
    
                    End With
                
                Set objMail = Nothing
                Set objOL = Nothing
                
                Range("c" & i).Interior.Color = RGB(146, 208, 80)
    
        End If
    Next i
End Sub

Private Sub CommandButton140_Click()
Dim MyFolder As String
Dim MyFile As String
MyFolder = "C:\Users\ShapkaMY\Desktop\Реестры ТРС\10 Октябрь\24.05.2021"
MyFile = Dir(MyFolder & "\*.xlsx")
Do While MyFile <> ""
Workbooks.Open FileName:=MyFolder & "\" & MyFile

    f = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 1 To f
        If IsEmpty(Range("k" & i)) = True Then
            Range("k" & i) = "б/н"
        End If
        
        If IsEmpty(Range("i" & i)) = True Then
            Range("i" & i) = "1"
        End If
        If IsEmpty(Range("j" & i)) = True Then
            Range("j" & i) = Range("b" & i)
        End If
    Next i


Windows(MyFile).Close True

MyFile = Dir
Loop
End Sub

Private Sub CommandButton141_Click()
  f = Cells(Rows.Count, 1).End(xlUp).Row
  
    For i = 2 To f
        If Range("i" & i) = "Подтвержден" Then
        Else
            Range("i" & i).Rows.Clear
        End If
  
        If Range("q" & i) = "В фирменном салоне" Then
            Range("q" & i).Rows.Clear
        End If
    Next i
    
    On Error Resume Next
    Range("q1:q" & f).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Range("i1:i" & f).SpecialCells(xlCellTypeBlanks).EntireRow.Delete

    
End Sub

Private Sub CommandButton142_Click()

'f = Cells(Rows.Count, 1).End(xlUp).Row
'    For i = 2 To f
'        If Range("k" & i).Interior.Pattern = xlNone Then
'            Range("k" & i).Rows.Clear
'        End If
'
'    Next i
'    Range("k1:k" & f).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
MsgBox (Range("k2").Font.Bold)

If Range("k2").Font.Bold = True Then
MsgBox ("ok")
End If
    
End Sub

Private Sub CommandButton143_Click()
'    Columns("K:K").Select
'    Selection.FormatConditions.AddUniqueValues
'    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
'    Selection.FormatConditions(1).DupeUnique = xlDuplicate
'    With Selection.FormatConditions(1).Interior
'        .PatternColorIndex = xlAutomatic
'        '.Color = 65535
'        .TintAndShade = 0
'        .Color = RGB(146, 208, 80)
'        .Font.Bold
'    End With

f = Cells(Rows.Count, 1).End(xlUp).Row
  
  
'For n = 2 To f
'
'    For i = 2 To f
'        If Range("k" & n) = Range("k" & i) Then
'            Range("k" & i).Interior.Color = RGB(146, 208, 80)
'            Range("k" & n).Interior.Color = RGB(146, 208, 80)
'        End If
'    Next i
'
'Next n


    On Error Resume Next
    ' ?????? ??????, ???????????? ??? ??????? ?????-??????????
    Colors = Array(12900829)

    Dim coll As New Collection, dupes As New Collection, _
        cols As New Collection, ra As Range, cell As Range, n&
    Err.Clear: Set ra = Intersect(Selection, ActiveSheet.UsedRange)
    If Err Then Exit Sub

    ra.Interior.ColorIndex = xlColorIndexNone: Application.ScreenUpdating = False
    For Each cell In ra.Cells ' ?????????? ???????? ?????????? ? ????????? dupes
        Err.Clear: If Len(Trim(cell)) Then coll.Add CStr(cell.Value), CStr(cell.Value)
        If Err Then dupes.Add CStr(cell.Value), CStr(cell.Value)
    Next cell
    
    For i& = 1 To dupes.Count ' ????????? ????????? cols ??????? ??? ?????? ??????????
        n = n Mod (UBound(Colors) + 1): cols.Add Colors(n), dupes(i): n = n + 1
    Next
    
    For Each cell In ra.Cells ' ?????????? ??????, ???? ??? ?? ???????? ???????? ????
        cell.Interior.Color = cols(CStr(cell.Value))
    Next cell
    Application.ScreenUpdating = True




End Sub

Private Sub CommandButton144_Click()
    f = Cells(Rows.Count, 2).End(xlUp).Row
    n = Range("b" & f - 1)
    w = Range("b" & f)

    X = 0
    For i = 1 To f
        If Range("d" & i) = "19500 гр." Then
            n = n - 1
            X = X + 1
        End If
    Next i
    
    

    
    
    
    
    
    
    
    
    
    Workbooks("main.xlsb").Sheets(1).Range("a" & b) = city
    Workbooks("main.xlsb").Sheets(1).Range("b" & b) = n
    Workbooks("main.xlsb").Sheets(1).Range("c" & b) = w
    Workbooks("main.xlsb").Sheets(1).Range("d" & b) = X
    
    If w > 10000 Then
    Workbooks("main.xlsb").Sheets(1).Range("e" & b) = "Обратите внимание на кол-во и вес. Возможно потребуется машина."
    End If
End Sub

Private Sub CommandButton145_Click()
    f = Cells(Rows.Count, 1).End(xlUp).Row
    
    

    If Range("a2") = "ТРС Екатеринбург" Then
    city = "Екатеринбург"
    b = 1
    ElseIf Range("a2") = "ТРС Санкт-Петербург" Then
    city = "Санкт-Петербург"
    b = 2
    ElseIf Range("a2") = "ТРС Нижний Новгород" Then
    city = "Нижний Новгород"
    b = 3
    ElseIf Range("a2") = "ТРС Новосибирск" Then
    city = "Новосибирск"
    b = 4
    ElseIf Range("a2") = "ТРС Тула" Then
    city = "Тула"
    b = 5
    ElseIf Range("a2") = "ТРС Ростов-на-Дону" Then
    city = "Ростов-на-Дону"
    b = 6
    ElseIf Range("a2") = "ТРС Саратов" Then
    city = "Саратов"
    b = 7
    End If
    
    
    
    
    
    
    For i = 2 To f
    
        Set X = Range("c" & i)
        Set y = Range("c" & i + 1)
        
        If X = y Then
        Else
            n = n + 1
        End If
        
        If Range("e" & i) = "Индивидуальный абонентский Терминал (И-АТ) Gemini I S2X" Then
            k = k + 1
        
        End If
        

    Next i

Set Z = Workbooks("main.xlsb").Sheets(1)
Z.Range("a" & b) = city 'ТРС
Z.Range("b" & b) = n - k 'Кол-во заказов
Z.Range("c" & b) = k 'Кол-во коробок



End Sub

Private Sub CommandButton146_Click()
On Error GoTo Instr
'Dim myWord As New Word.Application
'Dim myDocument As Word.Document
'Создаем новый документ по шаблону
  Set myDocument = _
  myWord.Documents.Add("C:\Users\ShapkaMY\Desktop\Образец.docm")
  myWord.Visible = True
With myDocument
'Замещаем текст закладок
  .Bookmarks("rpo").Range = "г. Омск"
'Удаляем границы ячеек
  .Tables(1).Borders.OutsideLineStyle = wdLineStyleNone
  .Tables(1).Borders.InsideLineStyle = wdLineStyleNone
End With
'Освобождаем переменные
Set myDocument = Nothing
Set myWord = Nothing
'Завершаем процедуру
Exit Sub
'Обработка ошибок
Instr:
If Err.Description <> "" Then
  MsgBox "Произошла ошибка: " & Err.Description
End If
If Not myWord Is Nothing Then
  myWord.Quit
  Set myDocument = Nothing
  Set myWord = Nothing
End If
End Sub

Private Sub CommandButton147_Click()
    Dim objWrdApp As Object, objWrdDoc As Object
    'создаем новое приложение Word
    Set objWrdApp = CreateObject("Word.Application")
    'Можно так же сделать приложение Word видимым. По умолчанию открывается в скрытом режиме
    'objWrdApp.Visible = True
    'открываем документ Word - документ "Doc1.doc" должен существовать
    Set objWrdDoc = objWrdApp.Documents.Open("C:\Users\ShapkaMY\Desktop\test.docx")
    'Копируем из Excel диапазон "A1:A10"
'    Range("A1").Copy
    
    
    'вставляем скопированные ячейки в Word - в начала документа
'   objWrdDoc.Range(0).Paste
    objWrdDoc.Bookmarks("qwe").Range.Text = "777"

    'закрываем документ Word с сохранением
    objWrdDoc.Close True    ' False - без сохранения
    'закрываем приложение Word - обязательно!
    objWrdApp.Quit
    'очищаем переменные Word - обязательно!
    Set objWrdDoc = Nothing: Set objWrdApp = Nothing
End Sub

Private Sub CommandButton148_Click()
Dim MyFolder As String
Dim MyFile As String
MyFolder = "C:\Users\ShapkaMY\Desktop\Реестры ТРС\10 Октябрь\" & TextBox1 & ""
MyFile = Dir(MyFolder & "\*.xlsx")
Do While MyFile <> ""
Workbooks.Open FileName:=MyFolder & "\" & MyFile

    f = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 1 To f
        If IsEmpty(Range("k" & i)) = True Then
            Range("k" & i) = "б/н"
        End If
        
        If IsEmpty(Range("i" & i)) = True Then
            Range("i" & i) = "1"
        End If
        If IsEmpty(Range("j" & i)) = True Then
            Range("j" & i) = Range("b" & i)
        End If
    Next i


Windows(MyFile).Close True

MyFile = Dir
Loop
End Sub

Private Sub CommandButton149_Click()
Range("A:AA").Copy 'Копируем содержимое листа
Sheets.Add.Name = "Остатки" 'Создаем лист "Возврат".
Range("A1").PasteSpecial Paste:=xlPasteValues 'Вставляем как значение


'For i = 2 To 100
'
'n = Range("n" & i)
'
'if n = n






'Next i
















End Sub

Private Sub CommandButton15_Click()
 Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
            With objMail
                .Display
                .To = "mihajlov@cc.tricolor.tv; moysya@cc.tricolor.tv; simkina@cc.tricolor.tv"
                .CC = "ChuchalovVY@monobrand-tt.ru; Butko@monobrand-tt.ru"
                .Subject = "Прозвон от " & Date
                .HTMLBody = "<p>Коллеги, добрый день!</p>" _
                & "<p>Прошу актуализировать данные, на запросы от КС до <b>" & Date + 1 & " 18:00</b><br>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                '.DeferredDeliveryTime = Date + 17 / 24
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
End Sub

Private Sub CommandButton150_Click()
Range("n:n").RemoveDuplicates 1, xlYes
End Sub

Private Sub CommandButton151_Click()
For i = 1 To 13
Columns(1).Delete
Next i

For i = 1 To 12
Columns(2).Delete
Next i

End Sub

Private Sub CommandButton152_Click()

End Sub

Private Sub CommandButton153_Click()

Set X = Workbooks("Статистика.csv").Sheets("Статистика")
f = X.Cells(Rows.Count, 11).End(xlUp).Row

For n = 2 To 300
    For i = 2 To f
        If X.Range("q" & i) = "Курьером" Then
            Range("n" & i).FormulaR1C1 = "=COUNTIF(Статистика.csv!C14,RC[-3])"
        End If
    Next i

Next n




End Sub

Private Sub CommandButton155_Click()
f = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To f

    Range("e" & i).FormulaR1C1 = _
        "=COUNTIFS(Статистика.csv!C9,""Подтвержден"",Статистика.csv!C17,""Курьером"",Статистика.csv!C14,RC[-3])"
        
        
    Range("f" & i).FormulaR1C1 = _
        "=COUNTIFS(Статистика.csv!C9,""Подтвержден"",Статистика.csv!C17,""Почта России (ОПС)"",Статистика.csv!C14,RC[-4])+COUNTIFS(Статистика.csv!C9,""Подтвержден"",Статистика.csv!C17,""Почта России (курьер)"",Статистика.csv!C14,RC[-4])"



Range("g" & i).FormulaR1C1 = _
        "=SUMIFS('[Table.xlsx]сводные остатки'!C9,'[Table.xlsx]сводные остатки'!C1,""Санкт-Петербург"",'[Table.xlsx]сводные остатки'!C2,RC[-3])"
Range("h" & i).FormulaR1C1 = _
        "=SUMIFS('[Table.xlsx]сводные остатки'!C9,'[Table.xlsx]сводные остатки'!C1,""Тула"",'[Table.xlsx]сводные остатки'!C2,RC[-4])"
Range("i" & i).FormulaR1C1 = _
        "=SUMIFS('[Table.xlsx]сводные остатки'!C9,'[Table.xlsx]сводные остатки'!C1,""Новосибирск"",'[Table.xlsx]сводные остатки'!C2,RC[-5])"
Range("j" & i).FormulaR1C1 = _
        "=SUMIFS('[Table.xlsx]сводные остатки'!C9,'[Table.xlsx]сводные остатки'!C1,""Нижний Новгород"",'[Table.xlsx]сводные остатки'!C2,RC[-6])"
Range("k" & i).FormulaR1C1 = _
        "=SUMIFS('[Table.xlsx]сводные остатки'!C9,'[Table.xlsx]сводные остатки'!C1,""Ростов-на-Дону"",'[Table.xlsx]сводные остатки'!C2,RC[-7])"
Range("l" & i).FormulaR1C1 = _
        "=SUMIFS('[Table.xlsx]сводные остатки'!C9,'[Table.xlsx]сводные остатки'!C1,""Екатеринбург"",'[Table.xlsx]сводные остатки'!C2,RC[-8])"
Range("m" & i).FormulaR1C1 = _
        "=SUMIFS('[Table.xlsx]сводные остатки'!C9,'[Table.xlsx]сводные остатки'!C1,""Саратов"",'[Table.xlsx]сводные остатки'!C2,RC[-9])"
Next i





End Sub

Private Sub CommandButton156_Click()
Worksheets(1).Range("i1").AutoFilter Field:=9, Criteria1:="Подтвержден"
 
Range("q1").AutoFilter Field:=17, Criteria1:= _
        "=Почта России (курьер)", Operator:=xlOr, Criteria2:="=Почта России (ОПС)"

 
End Sub

Private Sub CommandButton157_Click()
Application.ScreenUpdating = False
    f = Cells(Rows.Count, 11).End(xlUp).Row
    For i = 1 To f
        X = Range("i" & i)
        
        y1 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b2")
        y2 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b3")
        y3 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b4")
        y4 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b5")
        y5 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b6")
        y6 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b7")
        y7 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b8")
        y8 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b9")
        y9 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b10")
        y10 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b11")
        y11 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b12")
        y11 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b13")
        y12 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b14")
        y13 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b15")
        y14 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b16")
        y15 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b17")
        y16 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b18")
        y17 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b19")
        y19 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b20")
        y20 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b21")
        y21 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b22")
        y22 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b23")
        y23 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b24")
        y24 = Workbooks("TableHSR").Sheets("сводные остатки").Range("b25")
            

        If X = y Or X = y1 Or X = y2 Or X = y3 Or X = y4 Or X = y5 Or X = y6 Or X = y7 Or X = y8 Or X = y9 Or X = y15 Or X = y11 Or X = y12 Or X = y13 Or X = y14 Or X = y15 Or X = y16 Or X = y17 Or X = y18 Or X = y19 Or X = y20 Or X = y21 Or X = y22 Or X = y21 Or X = y22 Or X = y23 Or X = y24 Or X = y25 Then
            Range("l" & i) = "ok"
        Else
            Range("l" & i) = "error"
        End If

        Range("m" & i).FormulaR1C1 = "=RC[-2]/RC[-3]"
        Range("m" & i).Value = Range("m" & i).Value

        Range("m" & i).Copy
        Range("k" & i).PasteSpecial Paste:=xlPasteValues
    Next i
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton158_Click()
     Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
            With objMail
                .Display
                .To = "DorofeevaAV@tricolor.tv; PetrovaE@tricolor.tv"
                .CC = ""
                .Subject = "Отчёт от " & Date
                .HTMLBody = "<p>Добрый день.</p>" _
                & "<p>Во вложении отчёт от " & Date & "</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                
                
                'указывается текст письма
                .Attachments.Add "C:\Users\ShapkaMY\Desktop\Table.xlsx" 'указывается полный путь к файлу
                .DeferredDeliveryTime = Date + 12 / 24
                
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
End Sub

Private Sub CommandButton159_Click()
 Range("D1").FormulaArray = _
        "=INDEX([Table.xlsx]отправления!C4,MATCH(RC[3]&RC[5],[Table.xlsx]отправления!C7&[Table.xlsx]отправления!C9,0))"
End Sub

Private Sub CommandButton16_Click()
    f = Cells(Rows.Count, 11).End(xlUp).Row
    X = 0
    y = 0
    Z = 0
    


    'b2b
    b2b1 = TextBox2.Text
    b2b2 = TextBox3.Text
    b2b3 = TextBox4.Text
    b2b4 = TextBox5.Text
    b2b5 = TextBox6.Text
    b2b6 = TextBox7.Text
    
    For i = 4 To f
        If Range("ag" & i) = 25 Then
            X = Range("an" & i) + X
            xs = Range("am" & i) + xs
            xt = Range("aa" & i) + Range("ac" & i) + xt
            
            
        ElseIf Range("ag" & i) = 39.99 Then
            y = Range("an" & i) + y
            ys = Range("am" & i) + ys
            yt = Range("aa" & i) + Range("ac" & i) + yt
            
        'Изменить и искать по номеру заказа.
        ElseIf _
            Range("v" & i) = b2b1 Or _
            Range("v" & i) = b2b2 Or _
            Range("v" & i) = b2b3 Or _
            Range("v" & i) = b2b4 Or _
            Range("v" & i) = b2b5 Or _
            Range("v" & i) = b2b6 _
            Then
                b = Range("an" & i) + b
                bs = Range("am" & i) + bs
                bt = Range("aa" & i) + Range("ac" & i) + bt
            
        ElseIf _
            Range("ag" & i) <> 25 Or _
            Range("ag" & i) <> 39.99 Or _
            Range("v" & i) <> b2b1 Or _
            Range("v" & i) <> b2b2 Or _
            Range("v" & i) <> b2b3 Or _
            Range("v" & i) <> b2b4 Or _
            Range("v" & i) <> b2b5 Or _
            Range("v" & i) <> b2b6 _
            Then
            Z = Range("an" & i) + Z
            zs = Range("am" & i) + zs
            zt = Range("aa" & i) + Range("ac" & i) + zt
        End If
    Next i
    
    f = Cells(Rows.Count, 40).End(xlUp).Row
    proverka = Range("an" & f)
    proverk2 = Range("ao" & f)
    
    proverka3 = Range("aa" & f)
    proverka4 = Range("ac" & f)
    proverka5 = Range("am" & f)
    
    
    
    

    Sheets.Add
    Range("a2") = "без НДС"
    Range("a3") = "с НДС"
    Range("b1") = "ИМ"
    Range("c1") = "Обмен с доставкой"
    Range("d1") = "Меняйся за 2500"
    Range("e1") = "B2B"
    
    
    Range("d2") = X '2500
    Range("c2") = y 'Обмен
    Range("b2") = Z 'им
    Range("e2") = b 'b2b
    
    Range("d3") = X * 1.2
    Range("c3") = y * 1.2
    Range("b3") = Z * 1.2
    Range("e3") = b * 1.2
    
    'Проверка
    
    Range("f2") = X + y + Z + b
    Range("g2") = proverka
    Range("h2") = X + y + Z + b - proverka
    
    
    
    
    

    Range("b5") = "ИМ"
    Range("d5") = "Обмен с доставкой"
    Range("f5") = "Меняйся за 2500"
    Range("h5") = "B2B"

    
    
    Range("b6") = "Страховка"
    Range("c6") = "Тариф"
    Range("d6") = "Страховка"
    Range("e6") = "Тариф"
    Range("f6") = "Страховка"
    Range("g6") = "Тариф"
    Range("h6") = "Страховка"
    Range("i6") = "Тариф"
    
    
    Range("a7") = "без НДС"
    Range("a8") = "с НДС"
    
    Range("f7") = xs
    Range("g7") = xt
    
    Range("d7") = ys
    Range("e7") = yt
    
    Range("b7") = zs
    Range("c7") = zt
    
    Range("h7") = bs
    Range("i7") = bt
    
    
    
    Range("f8") = xs * 1.2
    Range("g8") = xt * 1.2
    
    Range("d8") = ys * 1.2
    Range("e8") = yt * 1.2
    
    Range("b8") = zs * 1.2
    Range("c8") = zt * 1.2
    
    Range("h8") = bs * 1.2
    Range("i8") = bt * 1.2
    
    
    Range("j7") = xs + ys + zs + bs + xt + yt + zt + bt
    Range("k7") = proverka3 + proverka4 + proverka5
    Range("l7") = Range("j7") - Range("k7")
    
    
End Sub

Private Sub CommandButton160_Click()

'Trsdatedbrf = TextBox20.Text
'MsgBox (Trsdatedbrf)

f = Workbooks("Table.xlsx").Sheets("отправления").Cells(Rows.Count, 1).End(xlUp).Row



Trsdatedbrf = "21.07.2021"

'MsgBox (Trsdatedbrf + 1)


For i = 2 To f
    If Workbooks("Table.xlsx").Sheets("отправления").Range("f" & i) = "21.07.2021" Then
        Workbooks("Table.xlsx").Sheets("отправления").Rows(i).Copy
        Workbooks("Main.xlsb").Sheets(1).Rows(1).Insert
    End If
Next i



End Sub

Private Sub CommandButton161_Click()
    asn = ActiveSheet.Name
    Sheets.Add.Name = "Итог"
    
    Sheets(asn).Range("a:a").Copy
    Range("b:b").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("b:b").Copy
    Range("a:a").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("f:f").Copy
    Range("c:c").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("g:g").Copy
    Range("d:d").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("h:h").Copy
    Range("e:e").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("i:i").Copy
    Range("g:g").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("j:j").Copy
    Range("h:h").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("k:k").Copy
    Range("i:i").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("o:o").Copy
    Range("j:j").PasteSpecial Paste:=xlPasteValues
    
    
    Range("c:c").NumberFormat = "m/d/yyyy"
    Range("e:e").NumberFormat = "0"
    Range("j:j").NumberFormat = "0"

End Sub

Private Sub CommandButton162_Click()

End Sub

Private Sub CommandButton163_Click()
    f = Cells(Rows.Count, 1).End(xlUp).Row
    
    

    
    For i = 1 To f
    
        If Range("j" & i) = "б/н" Then
            Range("j" & i).Rows.Clear
        ElseIf Range("j" & i) = "1" Then
            Range("j" & i).Rows.Clear
        End If
    
    
        If Range("b" & i) = "Екатеринбург" Then
        Range("b" & i) = "ТРС Екатеринбург"
        ElseIf Range("b" & i) = "Нижний Новгород" Then
        Range("b" & i) = "ТРС Нижний Новгород"
        ElseIf Range("b" & i) = "Ростов-на-Дону" Then
        Range("b" & i) = "ТРС Ростов-на-Дону"
        ElseIf Range("b" & i) = "Тула" Then
        Range("b" & i) = "ТРС Тула"
        ElseIf Range("b" & i) = "Санкт-Петербург" Then
        Range("b" & i) = "ТРС Санкт-Петербург"
        ElseIf Range("b" & i) = "Новосибирск" Then
        Range("b" & i) = "ТРС Новосибирск"
        ElseIf Range("b" & i) = "Саратов" Then
        Range("b" & i) = "ТРС Саратов"
        End If
        
        
        If Left(Range("e" & i), 1) = "2" Then
            Range("f" & i) = "Комиссионная торговля/Оплата наличными при получении/Белявский Кирилл Олегович/ТОРГОВЫЕ ТЕХНОЛОГИИ ООО/РУБ"
        ElseIf Left(Range("e" & i), 3) = "800" Then
            Range("f" & i) = "Комиссионная торговля/Оплата при получении Почта/Головинова Юлия Альбертовна/ТОРГОВЫЕ ТЕХНОЛОГИИ ООО/РУБ"
        End If


        If Range("g" & i) = "Индивидуальный абонентский Терминал (И-АТ) Gemini I S2X" Then
            Range("g" & i) = "Терминал индивидуальный абонентский (И-АТ)  SkyEdgeII-c Gemini-i S2X (tr)"
        End If
    Next i
    



End Sub

Private Sub CommandButton164_Click()


'Шапка

tema = "Нарушение контрольных сроков"


Range("a1") = "Запрос № 2 от " & Date
Range("d1") = "по теме: " & tema

Range("a5") = "На розыск внутренних партионных регистрируемых почтовых отправлений 'Тороговые технологии', пересылаемых в рамках договора"
Range("a6") = "№ 0000258 от 06.04.2020,заключенного между ООО 'Торговые технологии' и ФГУП 'Почта России',"

Range("a8") = "Наложенный платеж прошу перечислить в адрес Торговых технологий"
Range("a9") = "Реквизиты для перечисления наложенного платежа: Банковские реквизиты: Ф. ОПЕРУ Банка ВТБ (ПАО) в Санкт-Петербурге, Санкт-Петербург БИК 44030704 Расчетный счет 40702810980040000420 Корреспондентский счет 30101810200000000704"


'Range("a11") = "Место сдачи: 300008 - Путейская, д.6, Тула, Тульская область"
'Range("a12") = "Дата сдачи: 05.05.2021"
Range("a13") = "Обратный адрес для ответа: ShapkaMY@monobrand-tt.ru"


f = Cells(Rows.Count, 2).End(xlUp).Row
    X = 1
    For i = 17 To f - 1
        Range("a" & i) = X
        X = X + 1
        
        
        
        Range("e" & i).FormulaR1C1 = _
        "=VLOOKUP(RC[-3],'[Report_010521-120821.xlsx]Лист1'!C8:C14,7,0)"
        
        Range("f" & i).FormulaR1C1 = _
        "=VLOOKUP(RC[-4],'[Report_010521-120821.xlsx]Лист1'!C8:C13,6,0)"
        
        Range("j" & i).FormulaR1C1 = _
        "=VLOOKUP(RC[-8],'[Report_010521-120821.xlsx]Лист1'!C8:C39,32,0)"
        
        Range("k" & i).FormulaR1C1 = _
        "=VLOOKUP(RC[-9],'[Report_010521-120821.xlsx]Лист1'!C8:C25,18,0)"
        
        Range("l" & i).FormulaR1C1 = _
        "=VLOOKUP(RC[-10],'[Report_010521-120821.xlsx]Лист1'!C8:C23,16,0)"
        
        Range("ad" & i).FormulaR1C1 = _
        "=VLOOKUP(RC[-28],'[Report_010521-120821.xlsx]Лист1'!C8:C33,26,0)"
        
        Range("ae" & i) = Range("ad" & i)
        
    Next i














End Sub

Private Sub CommandButton165_Click()
  n = TextBox21.Text

    
    
    
    For i = 1 To n
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next i
    
    
    
End Sub

Private Sub CommandButton166_Click()
Range("b17:b300").NumberFormat = "#"
End Sub

Private Sub CommandButton167_Click()
    Application.ScreenUpdating = False
    Dim FilesToOpen
    Dim X As Integer
    FilesToOpen = Application.GetOpenFilename _
      (FileFilter:="All files (*.*), *.*", _
      MultiSelect:=True, Title:="Files to Merge")
    If TypeName(FilesToOpen) = "Boolean" Then
        MsgBox "Не выбрано ни одного файла!"
        Exit Sub
    End If
    X = 1
    While X <= UBound(FilesToOpen)
        Set importWB = Workbooks.Open(FileName:=FilesToOpen(X))
        Sheets().Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        importWB.Close savechanges:=False
        X = X + 1
    Wend
       Application.ScreenUpdating = True
End Sub

Private Sub CommandButton168_Click()
    Range("A1:AV30000").Copy
    Sheets.Add.Name = "Итог"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Private Sub CommandButton169_Click()
Range("b:b").NumberFormat = "#"
End Sub

Private Sub CommandButton17_Click()

    f = Cells(Rows.Count, 11).End(xlUp).Row
    X = 0
    y = 0
    Z = 0
    For i = 4 To f
        If Range("p" & i) = 0.75 Then
            X = Range("an" & i) + X
        ElseIf Range("p" & i) = 0.9 Then
            y = Range("an" & i) + y
        ElseIf Range("p" & i) <> 0.9 Or Range("p" & i) <> 0.75 Then
            Z = Range("an" & i) + Z
        End If
    Next i
    
    
    Sheets.Add
    Range("a2") = "без НДС"
    Range("a3") = "с НДС"
    Range("b1") = "ИМ"
    Range("c1") = "Обмен с доставкой"
    Range("d1") = "Меняйся за 2500"
    
    Range("d2") = X
    Range("c2") = y
    Range("b2") = Z
    
    Range("d3") = X * 1.2
    Range("c3") = y * 1.2
    Range("b3") = Z * 1.2
    
End Sub

Private Sub CommandButton18_Click()

    If CheckBox9.Value = True Then
        pochta = "Почта России"
    Else
        pochta = "Pony Express"
    End If
    
    


    Trsdate = TextBox1.Text
    ddt = TextBox8.Text
    
    Trsaddress1 = "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\Екатеринбург"
    Trsaddress2 = "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\Санкт-Петербург"
    Trsaddress3 = "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\Нижний Новгород"
    Trsaddress4 = "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\Новосибирск"
    Trsaddress5 = "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\Тула"
    Trsaddress6 = "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\Ростов-на-Дону"
    Trsaddress7 = "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\Саратов"
  
    
    If CheckBox9.Value Then
    Opochta = "Отгружаем через Почту России"
    
    
    End If
    
    
    
    
    If CheckBox1.Value Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Екатеринбург"
            With objMail
                .Display
                .To = "sklad1@rd.e-burg.n-l-e.ru; logist@rd.e-burg.n-l-e.ru; sklad@rd.e-burg.n-l-e.ru"
                .CC = "antipova@n-l-e.ru; ChuchalovVY@monobrand-tt.ru; BelyaevskiyKO@monobrand-tt.ru; BocharovAV@tricolor.tv"
                .Subject = "ОТПРАВКА ИНТЕРНЕТ-МАГАЗИН ООО <ТТ> " & Trsdate & " " & city & pochta
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p style='background: #00FFFF;'>" & Opochta & "</p>" _
                & "<p>Прошу подготовить к отправке ТМЦ согласно вложенному реестру отправлений.<br>" _
                & "Прилагаю:</p>" _
                & "<ul><li>Реестр отправлений</li><li>Накладные</li><li>Товарные чеки</li></ul>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                
                
                'указывается текст письма
                .Attachments.Add Trsaddress1 & "\" & Trsdate & " ТРС " & city & " (реестр отправлений).xlsx" 'указывается полный путь к файлу
                .Attachments.Add Trsaddress1 & "\Накладные.7z"
                .Attachments.Add Trsaddress1 & "\Чеки.7z"
                
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox2.Value = True Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Санкт-Петербург"
            With objMail
                .Display
                .To = "nachskl@trs.spb.n-l-e.ru; skl@trs.spb.n-l-e.ru"
                .CC = "antipova@n-l-e.ru; ChuchalovVY@monobrand-tt.ru; BelyaevskiyKO@monobrand-tt.ru; BocharovAV@tricolor.tv"
                .Subject = "ОТПРАВКА ИНТЕРНЕТ-МАГАЗИН ООО <ТТ> " & Trsdate & " " & city & pochta
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p style='background: #00FFFF;'>" & Opochta & "</p>" _
                & "<p>Прошу подготовить к отправке ТМЦ согласно вложенному реестру отправлений.<br>" _
                & "Прилагаю:</p>" _
                & "<ul><li>Реестр отправлений</li><li>Реестр накладных</li><li>Наклейки</li><li>Товарные чеки</li></ul>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                .Attachments.Add Trsaddress2 & "\" & Trsdate & " ТРС " & city & " (реестр отправлений).xlsx" 'указывается полный путь к файлу
                .Attachments.Add Trsaddress2 & "\" & Trsdate & " ТРС " & city & " для Pony Express.xlsx" 'указывается полный путь к файлу
                .Attachments.Add Trsaddress2 & "\Накладные.7z"
                .Attachments.Add Trsaddress2 & "\Чеки.7z"
                
                 If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                '.Send
                
            End With
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox3.Value Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Нижний Новгород"
            With objMail
                .Display
                .To = "logist@rd.nnov.n-l-e.ru; sklad@rd.nnov.n-l-e.ru; operator@rd.nnov.n-l-e.ru"
                .CC = "antipova@n-l-e.ru; ChuchalovVY@monobrand-tt.ru; BelyaevskiyKO@monobrand-tt.ru; BocharovAV@tricolor.tv"
                .Subject = "ОТПРАВКА ИНТЕРНЕТ-МАГАЗИН ООО <ТТ> " & Trsdate & " " & city & pochta
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p style='background: #00FFFF;'>" & Opochta & "</p>" _
                & "<p>Прошу подготовить к отправке ТМЦ согласно вложенному реестру отправлений.<br>" _
                & "Прилагаю:</p>" _
                & "<ul><li>Реестр отправлений</li><li>Накладные</li><li>Товарные чеки</li></ul>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                .Attachments.Add Trsaddress3 & "\" & Trsdate & " ТРС " & city & " (реестр отправлений).xlsx" 'указывается полный путь к файлу
                .Attachments.Add Trsaddress3 & "\Накладные.7z"
                .Attachments.Add Trsaddress3 & "\Чеки.7z"
                
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox4.Value = True Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Новосибирск"
            With objMail
                .Display
                .To = "sklad@trs.nvsb.n-l-e.ru; logist@trs.nvsb.n-l-e.ru; director@trs.nvsb.n-l-e.ru"
                .CC = "antipova@n-l-e.ru; ChuchalovVY@monobrand-tt.ru; BelyaevskiyKO@monobrand-tt.ru; BocharovAV@tricolor.tv"
                .Subject = "ОТПРАВКА ИНТЕРНЕТ-МАГАЗИН ООО <ТТ> " & Trsdate & " " & city & pochta
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p style='background: #00FFFF;'>" & Opochta & "</p>" _
                & "<p>Прошу подготовить к отправке ТМЦ согласно вложенному реестру отправлений.<br>" _
                & "Прилагаю:</p>" _
                & "<ul><li>Реестр отправлений</li><li>Накладные</li><li>Товарные чеки</li></ul>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                .Attachments.Add Trsaddress4 & "\" & Trsdate & " ТРС " & city & " (реестр отправлений).xlsx" 'указывается полный путь к файлу
                .Attachments.Add Trsaddress4 & "\Накладные.7z"
                .Attachments.Add Trsaddress4 & "\Чеки.7z"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox5.Value Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Тула"
            With objMail
                .Display
                .To = "logist@ts.tula.n-l-e.ru; logist2@ts.tula.n-l-e.ru; logist3@ts.tula.n-l-e.ru; operator1@ts.tula.n-l-e.ru; operator2@ts.tula.n-l-e.ru"
                .CC = "antipova@n-l-e.ru; ChuchalovVY@monobrand-tt.ru; BelyaevskiyKO@monobrand-tt.ru; BocharovAV@tricolor.tv"
                .Subject = "ОТПРАВКА ИНТЕРНЕТ-МАГАЗИН ООО <ТТ> " & Trsdate & " " & city & pochta
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p style='background: #00FFFF;'>" & Opochta & "</p>" _
                & "<p>Прошу подготовить к отправке ТМЦ согласно вложенному реестру отправлений.<br>" _
                & "Прилагаю:</p>" _
                & "<ul><li>Реестр отправлений</li><li>Реестр накладных</li><li>Накладные</li><li>Товарные чеки</li></ul>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                .Attachments.Add Trsaddress5 & "\" & Trsdate & " ТРС " & city & " (реестр отправлений).xlsx" 'указывается полный путь к файлу
                .Attachments.Add Trsaddress5 & "\" & Trsdate & " ТРС " & city & " для Pony Express.xlsx" 'указывается полный путь к файлу
                .Attachments.Add Trsaddress5 & "\Накладные.7z"
                .Attachments.Add Trsaddress5 & "\Чеки.7z"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox10.Value Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Тула"
            With objMail
                .Display
                .To = "logist@ts.tula.n-l-e.ru; logist2@ts.tula.n-l-e.ru; logist3@ts.tula.n-l-e.ru; operator1@ts.tula.n-l-e.ru; operator2@ts.tula.n-l-e.ru"
                .CC = "antipova@n-l-e.ru; ChuchalovVY@monobrand-tt.ru; BelyaevskiyKO@monobrand-tt.ru; BocharovAV@tricolor.tv"
                .Subject = "ОТПРАВКА ИНТЕРНЕТ-МАГАЗИН ООО <ТТ> " & Trsdate & " " & city & " Почта России"
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p style='background: #00FFFF;'>" & Opochta & "</p>" _
                & "<p>Прошу подготовить к отправке ТМЦ согласно вложенному реестру отправлений.<br>" _
                & "Прилагаю:</p>" _
                & "<ul><li>Реестр отправлений</li><li>Реестр накладных</li><li>Накладные</li><li>Товарные чеки</li></ul>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                .Attachments.Add Trsaddress8 & "\" & Trsdate & " ТРС " & city & " (реестр отправлений) Почта России.xlsx" 'указывается полный путь к файлу
                .Attachments.Add Trsaddress8 & "\" & Trsdate & " ТРС " & city & " для Pony Express.xlsx" 'указывается полный путь к файлу
                .Attachments.Add Trsaddress8 & "\Накладные.7z"
                .Attachments.Add Trsaddress8 & "\Чеки.7z"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox6.Value = True Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Ростов-на-Дону"
            With objMail
                .Display
                .To = "logist@rd.rostov.n-l-e.ru; tovaroved@trs1.rostov.n-l-e.ru; sklad@trs1.rostov.n-l-e.ru"
                .CC = "antipova@n-l-e.ru; ChuchalovVY@monobrand-tt.ru; BelyaevskiyKO@monobrand-tt.ru; BocharovAV@tricolor.tv"
                .Subject = "ОТПРАВКА ИНТЕРНЕТ-МАГАЗИН ООО <ТТ> " & Trsdate & " " & city & pochta
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p style='background: #00FFFF;'>" & Opochta & "</p>" _
                & "<p>Прошу подготовить к отправке ТМЦ согласно вложенному реестру отправлений.<br>" _
                & "Прилагаю:</p>" _
                & "<ul><li>Реестр отправлений</li><li>Накладные</li><li>Товарные чеки</li></ul>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                .Attachments.Add Trsaddress6 & "\" & Trsdate & " ТРС " & city & " (реестр отправлений).xlsx" 'указывается полный путь к файлу
                .Attachments.Add Trsaddress6 & "\Накладные.7z"
                .Attachments.Add Trsaddress6 & "\Чеки.7z"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox7.Value = True Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Саратов"
            With objMail
                .Display
                .To = "sklad1@trs1.saratov.n-l-e.ru; sklad2@trs1.saratov.n-l-e.ru; sklad@trs1.saratov.n-l-e.ru"
                .CC = "antipova@n-l-e.ru; ChuchalovVY@monobrand-tt.ru; BelyaevskiyKO@monobrand-tt.ru; BocharovAV@tricolor.tv"
                .Subject = "ОТПРАВКА ИНТЕРНЕТ-МАГАЗИН ООО <ТТ> " & Trsdate & " " & city & pochta
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p style='background: #00FFFF;'>" & Opochta & "</p>" _
                & "<p>Прошу подготовить к отправке ТМЦ согласно вложенному реестру отправлений.<br>" _
                & "Прилагаю:</p>" _
                & "<ul><li>Реестр отправлений</li><li>Накладные</li><li>Товарные чеки</li></ul>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                .Attachments.Add Trsaddress7 & "\" & Trsdate & " ТРС " & city & " (реестр отправлений).xlsx" 'указывается полный путь к файлу
                .Attachments.Add Trsaddress7 & "\Накладные.7z"
                .Attachments.Add Trsaddress7 & "\Чеки.7z"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
    End If
End Sub

Private Sub CommandButton170_Click()



Range("h:h").Copy
Range("b:b").PasteSpecial Paste:=xlPasteValues

End Sub

Private Sub CommandButton171_Click()
'Шапка

tema = "Нарушение контрольных сроков"


Range("a1") = "Запрос № 2 от " & Date
Range("d1") = "по теме: " & tema

Range("a5") = "На розыск внутренних партионных регистрируемых почтовых отправлений 'Тороговые технологии', пересылаемых в рамках договора"
Range("a6") = "№ 0000258 от 06.04.2020,заключенного между ООО 'Торговые технологии' и ФГУП 'Почта России',"

Range("a8") = "Наложенный платеж прошу перечислить в адрес Торговых технологий"
Range("a9") = "Реквизиты для перечисления наложенного платежа: Банковские реквизиты: Ф. ОПЕРУ Банка ВТБ (ПАО) в Санкт-Петербурге, Санкт-Петербург БИК 44030704 Расчетный счет 40702810980040000420 Корреспондентский счет 30101810200000000704"


'Range("a11") = "Место сдачи: 300008 - Путейская, д.6, Тула, Тульская область"
'Range("a12") = "Дата сдачи: 05.05.2021"
Range("a13") = "Обратный адрес для ответа: ShapkaMY@monobrand-tt.ru"


f = Cells(Rows.Count, 2).End(xlUp).Row
    X = 1
    For i = 17 To f - 1
        Range("a" & i) = X
        X = X + 1
        
        
   
    Range("c" & i).FormulaR1C1 = "=VLOOKUP(RC[-1],[main.xlsb]Итог!C2:C4,3,0)"

    Range("d" & i).FormulaR1C1 = "=VLOOKUP(RC[-2],[main.xlsb]Итог!C2:C7,6,0)"

    Range("e" & i).FormulaR1C1 = "=VLOOKUP(RC[-3],[main.xlsb]Итог!C2:C14,13,0)"

    Range("f" & i).FormulaR1C1 = "=VLOOKUP(RC[-4],[main.xlsb]Итог!C2:C13,12,0)"

    Range("g" & i).FormulaR1C1 = "=VLOOKUP(RC[-5],[main.xlsb]Итог!C2:C15,14,0)"

    Range("h" & i).FormulaR1C1 = "=VLOOKUP(RC[-6],[main.xlsb]Итог!C2:C33,32,0)"

    Range("j" & i).FormulaR1C1 = "=VLOOKUP(RC[-8],[main.xlsb]Итог!C2:C25,24,0)"

    Range("k" & i).FormulaR1C1 = "=VLOOKUP(RC[-9],[main.xlsb]Итог!C2:C25,24,0)"

    Range("l" & i).FormulaR1C1 = "=VLOOKUP(RC[-10],[main.xlsb]Итог!C2:C23,22,0)"

    Range("m" & i).FormulaR1C1 = "=VLOOKUP(RC[-11],[main.xlsb]Итог!C2:C21,20,0)"

    Range("n" & i).FormulaR1C1 = "=VLOOKUP(RC[-12],[main.xlsb]Итог!C2:C21,20,0)"

    Range("r" & i).FormulaR1C1 = "=VLOOKUP(RC[-16],[main.xlsb]Итог!C2:C6,5,0)"

    Range("s" & i).FormulaR1C1 = "=VLOOKUP(RC[-17],[main.xlsb]Итог!C2:C4,3,0)"

    Range("aa" & i).FormulaR1C1 = "=VLOOKUP(RC[-25],[main.xlsb]Итог!C2:C22,21,0)"

    Range("ac" & i).FormulaR1C1 = "=VLOOKUP(RC[-27],[main.xlsb]Итог!C2:C13,12,0)"

    Range("ad" & i).FormulaR1C1 = "=VLOOKUP(RC[-28],[main.xlsb]Итог!C2:C33,32,0)"

    Range("ae" & i).FormulaR1C1 = "=VLOOKUP(RC[-29],[main.xlsb]Итог!C2:C33,32,0)"

        
    Next i

Range("c:c").NumberFormat = "m/d/yyyy"
Range("s:s").NumberFormat = "m/d/yyyy"
Range("aa:aa").NumberFormat = "m/d/yyyy"






End Sub

Private Sub CommandButton172_Click()
     Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
            With objMail
                .Display
                .To = "Federal.Client.zapros@russianpost.ru"
                .CC = "ChuchalovVY@monobrand-tt.ru"
                .Subject = "Претензия" & Date
                .HTMLBody = "<p>Добрый день.</p>" _
                & "<p>Просьба рассмотреть претензию и дать обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

            
                
                'указывается текст письма
                .Attachments.Add "C:\Users\ShapkaMY\Desktop\backup\Претензии ПР\" & Date & " Претензии.xlsx" 'указывается полный путь к файлу
                .DeferredDeliveryTime = Date + 12 / 24
                
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
End Sub

Private Sub CommandButton173_Click()
    X = ActiveWorkbook.Name
    y = "Шаблон"
    
    Workbooks.Add
    
    Workbooks(X).Sheets(y).Copy before:=Sheets(1)
    ActiveWorkbook.SaveAs FileName:="C:\Users\ShapkaMY\Desktop\backup\Претензии ПР\" & Date & " Претензии.xlsx"
End Sub

Private Sub CommandButton174_Click()
    asn = ActiveSheet.Name
    Sheets.Add.Name = "Итог"
    
    Sheets(asn).Range("a:g").Copy
    Range("a:g").PasteSpecial Paste:=xlPasteValues
    
    f = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    
    
    
    f = Cells(Rows.Count, 1).End(xlUp).Row
    Range("h1:h" & f) = "отправка"
    Range("k1:k" & f).FormulaR1C1 = "=VLOOKUP(RC[-6],[Table3.xlsx]Наименования!C1:C2,2,0)"

End Sub

Private Sub CommandButton175_Click()



    f = Cells(Rows.Count, 8).End(xlUp).Row


    'Очищаем все ячейки в столбце "AA", где есть символ "0".
    For i = 1 To f
        If Range("H" & i) = "Накладная (листовка) к заказу ХШР Медиа" Then
            Range("h" & i).Rows.Clear
        End If
    Next i
    Range("h1:h" & f).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    
    
    
    
    
    
End Sub

Private Sub CommandButton277_Click()
f = Cells(Rows.Count, 1).End(xlUp).Row
    Range("b1:b" & f) = "Возврат"
    Range("e1:e" & f).FormulaR1C1 = "=WEEKNUM(RC[1],11)"
    
    For i = 1 To f
        If Range("a" & i) = "ТРС Екатеринбург" Then
        Range("a" & i) = "Екатеринбург"
        ElseIf Range("a" & i) = "ТРС Нижний Новгород" Then
        Range("a" & i) = "Нижний Новгород"
        ElseIf Range("a" & i) = "ТРС Ростов-на-Дону" Then
        Range("a" & i) = "Ростов-на-Дону"
        ElseIf Range("a" & i) = "ТРС Тула" Then
        Range("a" & i) = "Тула"
        ElseIf Range("a" & i) = "ТРС Санкт-Петербург" Then
        Range("a" & i) = "Санкт-Петербург"
        ElseIf Range("a" & i) = "ТРС Новосибирск" Then
        Range("a" & i) = "Новосибирск"
        ElseIf Range("a" & i) = "ТРС Саратов" Then
        Range("a" & i) = "Саратов"
        ElseIf Range("a" & i) = "Склад" Or Range("a" & i) = "Склад*" Then
        Range("a" & i).Rows.Clear
        End If
    Next i
    On Error Resume Next
    Range("A1:A" & f).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    f = Cells(Rows.Count, 1).End(xlUp).Row
     
    Columns("D:D").Select
    Selection.NumberFormat = "General"
    
    
    For i = 1 To f
        Range("k" & i).FormulaArray = _
            "=INDEX([Table.xlsx]отправления!C11,MATCH(RC[-4]&RC[-2],[Table.xlsx]отправления!C7&[Table.xlsx]отправления!C9,0))"
        Range("D" & i).FormulaArray = _
        "=INDEX([Table.xlsx]отправления!C4,MATCH(RC[3]&RC[5],[Table.xlsx]отправления!C7&[Table.xlsx]отправления!C9,0))"
    Next i
    
    f = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 1 To f
    If Range("o" & i) = "" Then
        Range("o" & i) = "б/н"
    End If
    
    If Range("l" & i) = "" Or Range("l" & i) = "упаковано в пакет Pony" Or Range("l" & i) = "Возврат" Then
        Range("l" & i) = "норма"
    End If
    
    Range("m" & i).Rows.Clear
    
    If Range("h" & i) = "" Then
        Range("h" & i).FormulaR1C1 = "=VLOOKUP(RC[-1],[Table.xlsx]отправления!C7:C8,2,0)"
    End If
    

    Next i
    
    
    Dim rArea As Range

    For Each rArea In Range("f1:f" & f).Areas
        rArea.FormulaLocal = rArea.FormulaLocal
    Next
End Sub

Private Sub CommandButton176_Click()

End Sub

Private Sub CommandButton178_Click()

        
    ddt = TextBox8.Text
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Екатеринбург"
        

        
            With objMail
                .Display
                .To = "sg.suhova@ponyexpress.ru; ekaterinburg.all@ponyexpress.ru"
                .CC = "ChuchalovVY@monobrand-tt.ru"
                .Subject = "Заказ ИМ на " & Trsdate & " ООО Торговые технологии/дог.22-50242 " & city
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p>Заказ ИМ на " & Trsdate & ", дог. 22-50242<br>" _
                & "Количество пакетов (1 кг) - " & Workbooks("main.xlsb").Sheets(1).Range("b1") & " шт.<br>" _
                & "<span style ='color:red;'>" & X & "</span><br>" _
                & "<span style ='color:red;'>" & Workbooks("main.xlsb").Sheets(1).Range("e1") & "</span><br>" _
                & "Адрес: 620024 г. Екатеринбург, по ул. Бисертской, 145 (литер АА1)</p>" _
                & "<p>Просьба подтвердить получение письма.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                
                '
                '.Send
            End With
        X = 0
        Set objMail = Nothing
        Set objOL = Nothing
End Sub

Private Sub CommandButton177_Click()
    asn = ActiveSheet.Name
    Sheets.Add.Name = "Итог для 1С"
    
    Sheets(asn).Range("a:a").Copy
    Range("b:b").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("b:b").Copy
    Range("a:a").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("f:f").Copy
    Range("c:c").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("g:g").Copy
    Range("d:d").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("h:h").Copy
    Range("e:e").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("i:i").Copy
    Range("g:g").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("j:j").Copy
    Range("h:h").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("k:k").Copy
    Range("i:i").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("o:o").Copy
    Range("j:j").PasteSpecial Paste:=xlPasteValues
    
    
    Range("c:c").NumberFormat = "m/d/yyyy"
    Range("e:e").NumberFormat = "0"
    Range("j:j").NumberFormat = "0"
    
    
     f = Cells(Rows.Count, 1).End(xlUp).Row
    
    

    
    For i = 1 To f
    
        If Range("j" & i) = "б/н" Then
            Range("j" & i).Rows.Clear
        ElseIf Range("j" & i) = "1" Then
            Range("j" & i).Rows.Clear
        End If
    
    
        If Range("b" & i) = "Екатеринбург" Then
        Range("b" & i) = "ТРС Екатеринбург"
        ElseIf Range("b" & i) = "Нижний Новгород" Then
        Range("b" & i) = "ТРС Нижний Новгород"
        ElseIf Range("b" & i) = "Ростов-на-Дону" Then
        Range("b" & i) = "ТРС Ростов-на-Дону"
        ElseIf Range("b" & i) = "Тула" Then
        Range("b" & i) = "ТРС Тула"
        ElseIf Range("b" & i) = "Санкт-Петербург" Then
        Range("b" & i) = "ТРС Санкт-Петербург"
        ElseIf Range("b" & i) = "Новосибирск" Then
        Range("b" & i) = "ТРС Новосибирск"
        ElseIf Range("b" & i) = "Саратов" Then
        Range("b" & i) = "ТРС Саратов"
        End If
        
        
        If Left(Range("e" & i), 1) = "2" Then
            Range("f" & i) = "Комиссионная торговля/Оплата наличными при получении/Белявский Кирилл Олегович/ТОРГОВЫЕ ТЕХНОЛОГИИ ООО/РУБ"
        ElseIf Left(Range("e" & i), 3) = "800" Then
            Range("f" & i) = "Комиссионная торговля/Оплата при получении Почта/Головинова Юлия Альбертовна/ТОРГОВЫЕ ТЕХНОЛОГИИ ООО/РУБ"
        End If


        If Range("g" & i) = "Индивидуальный абонентский Терминал (И-АТ) Gemini I S2X" Then
            Range("g" & i) = "Терминал индивидуальный абонентский (И-АТ)  SkyEdgeII-c Gemini-i S2X (tr)"
        End If
    Next i
    
    
    
    
    
    
End Sub

Private Sub CommandButton179_Click()
Dim j As Integer
Dim objHTTP As Object
Dim Json As String
Dim result As String
Dim URL As String
Dim Token As String
Dim X As String
Dim a As Date
Dim t As Date



'Dim n As Integer



t = Time
f = Cells(Rows.Count, 1).End(xlUp).Row
o = 0
a = "01.06.2021"



For n = 2 To f
    
        If Range("f" & n) > a Then
            If Left(Range("h" & n), 3) = "800" Then
                If Range("q" & n) = "Получено адресатом" Or Range("q" & n) = "Получено отправителем" Then

                Else

                X = Range("h" & n)
                URL = "https://otpravka-api.pochta.ru/1.0/shipment/search?query=" + X
                Token = "ekIcc3ZbbIdgl8TQJLb6KrqYGeNDt8KD"
                Token2 = "ZC5wb2RhQGl0ZWNoLWdyb3VwLnJ1OnRyaWNvbG9yVEVTVA=="

                Set objHTTP = CreateObject("Msxml2.XMLHTTP.6.0")
                    objHTTP.Open "GET", URL, False
                    objHTTP.setRequestHeader "Content-type", "application/json;charset=UTF-8"
                    objHTTP.setRequestHeader "Accept", "application/json"
                    objHTTP.setRequestHeader "Authorization", "AccessToken " + Token
                    objHTTP.setRequestHeader "X-User-Authorization", "Basic  " + Token2
                    objHTTP.send (Json)
                    result = objHTTP.responseText
                    'Range("Q50305").Value = result
                Set objHTTP = Nothing
                'Application.Wait (Now + TimeValue("0:00:01"))
                On Error Resume Next
                hon = Split(Split(result, "human-operation-name" & Chr(34) & " : " & Chr(34) & "")(1), "" & Chr(34) & "")(0)
                Range("Q" & n) = hon
                
                lod = Split(Split(result, "last-oper-date" & Chr(34) & " : " & Chr(34) & "")(1), "T")(0)
                Range("r" & n) = lod

                End If
            End If

         o = o + 1
        End If
        
Next n







' "add-to-mmo" : false,

t = Time - t

MsgBox o
MsgBox t

End Sub

Private Sub CommandButton180_Click()


    Dim sEnv As String, sURL As String
    Dim xmlhtp As Object, xmlDoc As Object, b
    sURL = "https://tracking.russianpost.ru/rtm34"


    '  sEnv = Worksheets(1).Cells(11, 1)

    sEnv = "<?xml version=""1.0"" encoding=""utf-8""?>" ' & vbNewLine
'    sEnv = sEnv & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:ns0=""API"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soapenc=""http://schemas.xmlsoap.org/soap/encoding/"">" ' & vbNewLine
    sEnv = sEnv & "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:oper=""http://russianpost.org/operationhistory"" xmlns:data=""http://russianpost.org/operationhistory/data"" xmlns:ns1=""http://schemas.xmlsoap.org/soap/envelope/"">" ' & vbNewLine


    
    
       sEnv = sEnv & "<soap:Header/>"
       sEnv = sEnv & "<soap:Body>"
       sEnv = sEnv & "<oper:getOperationHistory>"
       sEnv = sEnv & "<!--Optional:-->"
       sEnv = sEnv & "<data:OperationHistoryRequest>"
       sEnv = sEnv & "<data:Barcode>80082062494412</data:Barcode>"
       sEnv = sEnv & "<data:MessageType>0</data:MessageType>"
       sEnv = sEnv & "<!--Optional:-->"
       sEnv = sEnv & "<data:Language>RUS</data:Language>"
       sEnv = sEnv & "</data:OperationHistoryRequest>"
       sEnv = sEnv & "<!--Optional:-->"
       sEnv = sEnv & "<data:AuthorizationHeader ns1:mustUnderstand=""?"">"
       sEnv = sEnv & "<data:login>ykDaLTEChMLavX</data:login>"
       sEnv = sEnv & "<data:password>JPOIsPTd3W03</data:password>"
       sEnv = sEnv & "</data:AuthorizationHeader>"
       sEnv = sEnv & "</oper:getOperationHistory>"
       sEnv = sEnv & "</soap:Body>"
       sEnv = sEnv & "</soap:Envelope>"
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

    Set xmlhtp = CreateObject("Microsoft.XMLHTTP")
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")

    b = Len(sEnv)
    
    With xmlhtp
           .Open "POST", sURL, False
                .setRequestHeader "Content-Length", b
                .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
                .setRequestHeader "soapAction", "API#GetClientInfo"
                .setRequestHeader "Host", "https://tracking.russianpost.ru"
        .send
    End With
    
    With xmlhtp
        .Open "POST", sURL, False
                .setRequestHeader "Content-Length", b
                .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
                .setRequestHeader "soapAction", "API#GetClientInfo"
                .setRequestHeader "Host", "https://tracking.russianpost.ru"
        .send ' sEnv

        xmlDoc.LoadXML .responseText
        MsgBox .responseText
    End With

  

End Sub

Private Sub CommandButton181_Click()

Dim sURL As String
Dim sEnv As String
'Dim xmlhtp As New MSXML2.XMLHTTP
Dim xmlDoc As New DOMDocument

Set xmlhtp = CreateObject("MSXML2.XMLHTTP")
'Set xmlDoc = CreateObject("DOMDocument")

sURL = "https://tracking.russianpost.ru/rtm34"

sEnv = "<?xml version=""1.0"" encoding=""utf-8""?>"
    sEnv = sEnv & "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:oper=""http://russianpost.org/operationhistory"" xmlns:data=""http://russianpost.org/operationhistory/data"" xmlns:ns1=""http://schemas.xmlsoap.org/soap/envelope/"">"
    sEnv = sEnv & "<soap:Header/>"
       sEnv = sEnv & "<soap:Body>"
       sEnv = sEnv & "<oper:getOperationHistory>"
       sEnv = sEnv & "<!--Optional:-->"
       sEnv = sEnv & "<data:OperationHistoryRequest>"
       sEnv = sEnv & "<data:Barcode>80082062494412</data:Barcode>"
       sEnv = sEnv & "<data:MessageType>0</data:MessageType>"
       sEnv = sEnv & "<!--Optional:-->"
       sEnv = sEnv & "<data:Language>RUS</data:Language>"
       sEnv = sEnv & "</data:OperationHistoryRequest>"
       sEnv = sEnv & "<!--Optional:-->"
       sEnv = sEnv & "<data:AuthorizationHeader ns1:mustUnderstand=""?"">"
       sEnv = sEnv & "<data:login>ykDaLTEChMLavX</data:login>"
       sEnv = sEnv & "<data:password>JPOIsPTd3W03</data:password>"
       sEnv = sEnv & "</data:AuthorizationHeader>"
       sEnv = sEnv & "</oper:getOperationHistory>"
       sEnv = sEnv & "</soap:Body>"
       sEnv = sEnv & "</soap:Envelope>"




'sEnv = sEnv & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
'sEnv = sEnv & " <soap:Body>"
'sEnv = sEnv & " <CurrentConvertToEUR xmlns=""http://www.gama-system.com/webservices"">"
'sEnv = sEnv & " <dcmValue>100</dcmValue>"
'sEnv = sEnv & " <strBank>BS</strBank>"
'sEnv = sEnv & " <strCurrency>USD</strCurrency>"
'sEnv = sEnv & " <intRank>1</intRank>"
'sEnv = sEnv & " </CurrentConvertToEUR>"
'sEnv = sEnv & " </soap:Body>"
'sEnv = sEnv & "</soap:Envelope>"

With xmlhtp
.Open "post", sURL, False
.setRequestHeader "Host", "webservices.gama-system.com"
.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
.setRequestHeader "soapAction", "https://tracking.russianpost.ru/rtm34"
.setRequestHeader "Accept-encoding", "zip"
.send sEnv
xmlDoc.LoadXML .responseText
'xmlDoc.Save ThisWorkbook.Path & "\WebQueryResult.xml"
MsgBox .responseText
End With

End Sub

Private Sub CommandButton183_Click()
    asn = ActiveSheet.Name
    Sheets.Add.Name = "Итог"
    
    Sheets(asn).Range("a:a").Copy
    Range("a:a").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("c:c").Copy
    Range("b:b").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("f:f").Copy
    Range("c:c").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("h:h").Copy
    Range("d:d").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("i:i").Copy
    Range("e:e").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("k:k").Copy
    Range("f:f").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("p:p").Copy
    Range("g:g").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("v:v").Copy
    Range("h:h").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("v:v").Copy
    Range("h:h").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("w:w").Copy
    Range("i:i").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("x:x").Copy
    Range("j:j").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("y:y").Copy
    Range("k:k").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("z:z").Copy
    Range("l:l").PasteSpecial Paste:=xlPasteValues
End Sub

Private Sub CommandButton184_Click()
f = Cells(Rows.Count, 1).End(xlUp).Row
 For i = 1 To f
        If Range("h" & i) = "Накладная (листовка) к заказу ХШР Медиа" Or Range("V" & i) = "Накладная" Then
        Range("h" & i).Rows.Clear
        End If
    Next i
    
    f = Cells(Rows.Count, 1).End(xlUp).Row
    Range("h1:k" & f).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
End Sub

Private Sub CommandButton185_Click()
Dim t As Date

t = Time


    f = Cells(Rows.Count, 1).End(xlUp).Row + 3
    For i = 1 To f + 1
    
        
        
        
        
        
        If Range("h" & i) = "Комплект Звук без проводов Триколор+ Подарок (3 шт.световозвращателя)" _
        Or Range("h" & i) = "Комплект Звук без проводов Триколор+ Подарок (3 шт.световозвращателя)" _
            Then
                Range("h" & i) = "Комплект ""Звук без проводов Триколор + 3 световозвращателя Триколор"""
            
            ElseIf Range("h" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7, 2 Mpix, Full HD, ИК 10м, WiFi)" Or Range("h" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (ИСХ)" _
            Then
                Range("h" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD, ИК 10м, WiFi)"
            
            ElseIf _
                Range("h" & i) = "Видеокамера IP уличная Триколор Умный дом SCO-2 (1/2,7, 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)" Or _
                Range("h" & i) = "Видеокамера IP уличная Триколор Умный дом SCO-2 (1/2,7, 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)" _
            Then
                Range("h" & i) = "Видеокамера IP уличная Триколор Умный дом SCO-2 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)"
            
            ElseIf _
                Range("h" & i) = "Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)" _
            Then
                Range("h" & i) = "Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)"
            
           
            ElseIf _
                Range("h" & i) = "Комплект усилитель сотовой связи 900/2100, Триколор, TR-900/2100-50-kit+ Подарок Органайзер для пультов ДУ и прессы" _
            Then
            Range("h" & i) = "Комплект усилитель сотовой связи 900/2100, Триколор, TR-900/2100-50-kit"
            Rows(i).Copy
            Rows(i + 1).Insert
            Rows(i + 1).Select
            Range("i" & i) = "11790"
            Range("h" & i + 1) = "Органайзер для пультов ДУ и прессы"
            Range("i" & i + 1) = "200"
    
        
            ElseIf _
                Range("h" & i) = "Комплект усилитель мобильного интернета, " & Chr(34) & "Триколор ТВ" & Chr(34) & ", DS-4G-5kit+ Подарок Органайзер для пультов ДУ и прессы" _
                Or Range("h" & i) = "Комплект усилитель мобильного интернета, ""Триколор ТВ"", DS-4G-5kit+ Подарок Органайзер для пультов ДУ и прессы" _
                Or Range("h" & i) = "Комплект усилитель мобильного интернета, ""Триколор ТВ"", DS-4G-5kit+ Подарок Органайзер для пультов ДУ и прессы" _
            Then
                Range("h" & i) = "Комплект усилитель мобильного интернета, " & Chr(34) & "Триколор ТВ" & Chr(34) & ", DS-4G-5kit"
                Rows(i).Copy
                Rows(i + 1).Insert
                Rows(i + 1).Select
                Range("i" & i) = "10790"
                Range("h" & i + 1) = "Органайзер для пультов ДУ и прессы"
                Range("i" & i + 1) = "200"
            
            'Лот 1
            ElseIf _
                Range("h" & i) = "Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi), Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7, 2 Mpix, Full HD, ИК 10м, WiFi)" _
                Or Range("h" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7, 2 Mpix, Full HD, ИК 10м, WiFi), Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)" _
                Or Range("h" & i) = "Комплект камер Триколор" _
                Or Range("h" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7, 2 Mpix, Full HD, ИК 10м, WiFi), Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)" _
                Or Range("h" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7, 2 Mpix, Full HD, ИК 10м, WiFi), Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi) " _
                Or Range("h" & i) = "Комплект камер Триколор" _
            Then
                Range("h" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD, ИК 10м, WiFi)"
                
            If Range("j" & i) = "2" Then
                Range("j" & i) = "1"
            End If
                    
            If Range("j" & i) = "4" Then
                Range("j" & i) = "2"
            End If
            
                Rows(i).Copy
                Rows(i + 1).Insert
                Rows(i + 1).Select
                Range("i" & i) = "2400"
                Range("h" & i + 1) = "Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)"
                Range("i" & i + 1) = "3500"
                
            'Лот 2
            ElseIf _
                Range("h" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7, 2 Mpix, Full HD, ИК 10м, WiFi), Видеокамера IP уличная Триколор Умный дом SCO-2 (1/2,7, 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)" _
                Or Range("h" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7, 2 Mpix, Full HD, ИК 10м, WiFi), Видеокамера IP уличная Триколор Умный дом SCO-2 (1/2,7, 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi) " _
                Or Range("h" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7, 2 Mpix, Full HD, ИК 10м, WiFi), Видеокамера IP уличная Триколор Умный дом SCO-2 (1/2,7, 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)" _
                Or Range("h" & i) = "Комплект камер Триколор 2" _
            Then
                Range("h" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-2 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD, ИК 10м, WiFi)"
                If Range("j" & i) = "2" Then
                    Range("j" & i) = "1"
                End If
                If Range("j" & i) = "4" Then
                    Range("j" & i) = "2"
                End If
                Rows(i).Copy
                Rows(i + 1).Insert
                Rows(i + 1).Select
                Range("i" & i) = "2400"
                Range("h" & i + 1) = "Видеокамера IP уличная Триколор Умный дом SCO-2 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)"
                Range("i" & i + 1) = "3500"
                
                
            'Лот 3
            ElseIf _
                Range("h" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-1 (1/2,7"", 2 Mpix, Full HD, ИК 10м, WiFi), Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7"", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)" _
                Or Range("h" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-1 (1/2,7"", 2 Mpix, Full HD, ИК 10м, WiFi), Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7"", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi) " _
                Or Range("h" & i) = "Комплект камер Триколор 3" _
            Then
                Range("h" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD, ИК 10м, WiFi)"
                If Range("j" & i) = "2" Then
                    Range("j" & i) = "1"
                End If
                If Range("j" & i) = "4" Then
                    Range("j" & i) = "2"
                End If
                Rows(i).Copy
                Rows(i + 1).Insert
                Rows(i + 1).Select
                Range("i" & i) = "2400"
                Range("h" & i + 1) = "Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)"
                Range("i" & i + 1) = "3500"
                
        End If
    Next i
    
t = Time - t

MsgBox (t)
End Sub

Private Sub CommandButton186_Click()
 ' Code Snippets : Some of the source code listed below was taken from the following websites and credit show be given to the respective authors
 '#  http://scn.sap.com/community/epm/blog/2012/08/10/how-to-invoke-a-soap-web-service-from-custom-vba-code
 '#  http://www.vbaexpress.com/forum/showthread.php?t=34354
 '#  http://stackoverflow.com/questions/241725/calling-a-webservice-from-vba-using-soap
 '#  http://brettdotnet.posterous.com/excel-vba-using-a-web-service-with-xmlhttp-we
    'Declare our working variables
    Dim sURL As String
    Dim sEnv As String
       
    'Set and Instantiate our working objects
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    sURL = "https://tracking.russianpost.ru/rtm34"
      
    
    ' we create our SOAP envelope for submission to the Web Service
     'sEnv = "<?xml version=""1.0"" encoding=""utf-8""?>"
'     sEnv = sEnv & "<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope"">"
'     sEnv = sEnv & "  <soap:Header>"
'     sEnv = sEnv & "  <soap:Body>"
'     sEnv = sEnv & "   <soap:Request>"
'     sEnv = sEnv & "    <soap:User>username</soap:User>"
'     sEnv = sEnv & "    <soap:Pwd>password</soap:Pwd>"
'     sEnv = sEnv & "    <soap:Sku>KT-21-61261-01</soap:Sku>"
'     sEnv = sEnv & "   </soap:Request>"
'     sEnv = sEnv & "  </soap:Header>"
'     sEnv = sEnv & "  </soap:Body>"
'     sEnv = sEnv & "</soap:Envelope>"
     
     
     sEnv = sEnv & "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:oper=""http://russianpost.org/operationhistory"" xmlns:data=""http://russianpost.org/operationhistory/data"" xmlns:ns1=""http://schemas.xmlsoap.org/soap/envelope/"">"
     sEnv = sEnv & "<soap:Header/>"
     sEnv = sEnv & "<soap:Body>"
     sEnv = sEnv & "<oper:getOperationHistory>"
     sEnv = sEnv & "<data:OperationHistoryRequest>"
     sEnv = sEnv & "<data:Barcode>80082062494412</data:Barcode>"
     sEnv = sEnv & "<data:MessageType>0</data:MessageType>"
     sEnv = sEnv & "<data:Language>RUS</data:Language>"
     sEnv = sEnv & "</data:OperationHistoryRequest>"
     sEnv = sEnv & "<data:AuthorizationHeader ns1:mustUnderstand=""?"">"
     sEnv = sEnv & "<data:login>ykDaLTEChMLavX</data:login>"
     sEnv = sEnv & "<data:password>JPOIsPTd3W03</data:password>"
     sEnv = sEnv & "</data:AuthorizationHeader>"
     sEnv = sEnv & "</oper:getOperationHistory>"
     sEnv = sEnv & "</soap:Body>"
     sEnv = sEnv & "</soap:Envelope>"

     
     
    'we invoke the web service
    'use this code snippet to invoke a web service which requires authentication
    objHTTP.Open "Post", "https://tracking.russianpost.ru/rtm34", False
    objHTTP.setRequestHeader "Content-Type", "text/xml"

    objHTTP.send sEnv
    Range("a1") = objHTTP.responseText
    'clean up code
    Set objHTTP = Nothing
    Set xmlDoc = Nothing
End Sub

Private Sub CommandButton187_Click()
 'Set and instantiate our working objects
    Dim Req As Object
    Dim sEnv As String
    Dim Resp As New MSXML2.DOMDocument60
    Set Req = CreateObject("MSXML2.XMLHTTP")
    Set Resp = CreateObject("MSXML2.DOMDocument.6.0")
    
    Dim a As Date
    f = Cells(Rows.Count, 1).End(xlUp).Row
    
    a = "22.09.2021"
    
For n = 1 To 3
    'If Range("f" & n) > a Then

    
        Req.Open "Post", "https://tracking.russianpost.ru/rtm34", False
         sEnv = sEnv & "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:oper=""http://russianpost.org/operationhistory"" xmlns:data=""http://russianpost.org/operationhistory/data"" xmlns:ns1=""http://schemas.xmlsoap.org/soap/envelope/"">"
         sEnv = sEnv & "<soap:Header/>"
         sEnv = sEnv & "<soap:Body>"
         sEnv = sEnv & "<oper:getOperationHistory>"
         sEnv = sEnv & "<data:OperationHistoryRequest>"
         sEnv = sEnv & "<data:Barcode>" & Range("h" & n) & "</data:Barcode>"
         sEnv = sEnv & "<data:MessageType>0</data:MessageType>"
         sEnv = sEnv & "<data:Language>RUS</data:Language>"
         sEnv = sEnv & "</data:OperationHistoryRequest>"
         sEnv = sEnv & "<data:AuthorizationHeader ns1:mustUnderstand=""?"">"
         sEnv = sEnv & "<data:login>ykDaLTEChMLavX</data:login>"
         sEnv = sEnv & "<data:password>JPOIsPTd3W03</data:password>"
         sEnv = sEnv & "</data:AuthorizationHeader>"
         sEnv = sEnv & "</oper:getOperationHistory>"
         sEnv = sEnv & "</soap:Body>"
         sEnv = sEnv & "</soap:Envelope>"
    ' Send SOAP Request
       Req.send (sEnv)

        
        
        Req2 = Replace(Req.responseText, "ns3:Name", "Name")
    '    Range("a1") = Req2
    
        Dim pDoc As New MSXML2.DOMDocument60
        pDoc.LoadXML Req2
        
        Set nodeXML = pDoc.getElementsByTagName("Name")
        For i = 0 To nodeXML.Length - 2
        Range("q" & n) = nodeXML(i).Text
        Next
    
        
      'clean up code
        Set Req = Nothing
        Set Resp = Nothing

    'End If
    
Next n
    
End Sub

Private Sub CommandButton188_Click()
Dim a As Date
Dim t As Date
t = Time



f = Cells(Rows.Count, 8).End(xlUp).Row
n = ActiveCell.Row

For u = n To f

Debug.Print n




    Dim Req As Object
    Dim sEnv As String
    Dim Resp As New MSXML2.DOMDocument60
    Set Req = CreateObject("MSXML2.XMLHTTP")
    Set Resp = CreateObject("MSXML2.DOMDocument.6.0")
    

    

    
        Req.Open "Post", "https://tracking.russianpost.ru/rtm34", False
         sEnv = sEnv & "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:oper=""http://russianpost.org/operationhistory"" xmlns:data=""http://russianpost.org/operationhistory/data"" xmlns:ns1=""http://schemas.xmlsoap.org/soap/envelope/"">"
         sEnv = sEnv & "<soap:Header/>"
         sEnv = sEnv & "<soap:Body>"
         sEnv = sEnv & "<oper:getOperationHistory>"
         sEnv = sEnv & "<data:OperationHistoryRequest>"
         sEnv = sEnv & "<data:Barcode>" & Range("h" & u) & "</data:Barcode>"
         sEnv = sEnv & "<data:MessageType>0</data:MessageType>"
         sEnv = sEnv & "<data:Language>RUS</data:Language>"
         sEnv = sEnv & "</data:OperationHistoryRequest>"
         sEnv = sEnv & "<data:AuthorizationHeader ns1:mustUnderstand=""?"">"
         sEnv = sEnv & "<data:login>ykDaLTEChMLavX</data:login>"
         sEnv = sEnv & "<data:password>JPOIsPTd3W03</data:password>"
         sEnv = sEnv & "</data:AuthorizationHeader>"
         sEnv = sEnv & "</oper:getOperationHistory>"
         sEnv = sEnv & "</soap:Body>"
         sEnv = sEnv & "</soap:Envelope>"
    ' Send SOAP Request
        Req.send (sEnv)
        'Debug.Print Req
        
        
        Req2 = Replace(Req.responseText, "ns3:OperAttr", "OperAttr")
'        Debug.Print Req2
        
    
        Dim pDoc As New MSXML2.DOMDocument60
        pDoc.LoadXML Req2
        
        Set nodeXML = pDoc.getElementsByTagName("OperAttr")
        
        For i = 1 To nodeXML.Length - 1
            X = nodeXML(i).Text
            If X = "1Вручение адресату" _
            Or X = "8Адресату курьером" _
            Or X = "6Адресату почтальоном" _
            Then
            s = "Получено адресатом"

            ElseIf X = "2Вручение отправителю" _
            Or X = "7Отправителю почтальоном" _
            Or X = "8Отправителю курьером" _
            Then
            s = "Получено отправителем"
            End If
'
        Next i

        Range("q" & u) = X
        
      'clean up code
        Set Req = Nothing
        Set Req2 = Nothing
        Set Resp = Nothing
        Set nodeXML = Nothing
        Set pDoc = Nothing
        sEnv = ""
        
        
        
    
Next u
t = Time - t

End Sub

Private Sub CommandButton189_Click()
Application.ScreenUpdating = False
t = GetTickCount


Dim a As Date


n = ActiveCell.Row

    Dim Req As Object
    Dim sEnv As String
    Dim Resp As New MSXML2.DOMDocument60
    Set Req = CreateObject("MSXML2.XMLHTTP")
    Set Resp = CreateObject("MSXML2.DOMDocument.6.0")
    
        Req.Open "Post", "https://tracking.russianpost.ru/rtm34", False
         sEnv = sEnv & "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:oper=""http://russianpost.org/operationhistory"" xmlns:data=""http://russianpost.org/operationhistory/data"" xmlns:ns1=""http://schemas.xmlsoap.org/soap/envelope/"">"
         sEnv = sEnv & "<soap:Header/>"
         sEnv = sEnv & "<soap:Body>"
         sEnv = sEnv & "<oper:getOperationHistory>"
         sEnv = sEnv & "<data:OperationHistoryRequest>"
         sEnv = sEnv & "<data:Barcode>" & Range("h" & n) & "</data:Barcode>"
         sEnv = sEnv & "<data:MessageType>0</data:MessageType>"
         sEnv = sEnv & "<data:Language>RUS</data:Language>"
         sEnv = sEnv & "</data:OperationHistoryRequest>"
         sEnv = sEnv & "<data:AuthorizationHeader ns1:mustUnderstand=""?"">"
         sEnv = sEnv & "<data:login>ykDaLTEChMLavX</data:login>"
         sEnv = sEnv & "<data:password>JPOIsPTd3W03</data:password>"
         sEnv = sEnv & "</data:AuthorizationHeader>"
         sEnv = sEnv & "</oper:getOperationHistory>"
         sEnv = sEnv & "</soap:Body>"
         sEnv = sEnv & "</soap:Envelope>"
        Req.send (sEnv)

Debug.Print (GetTickCount - t) / 1000, vbInformation

        
        Req2 = Replace(Req.responseText, "ns3:OperAttr", "OperAttr")
        
    
        Dim pDoc As New MSXML2.DOMDocument60
        pDoc.LoadXML Req2
        
        Set nodeXML = pDoc.getElementsByTagName("OperAttr")
        
        For i = 1 To nodeXML.Length - 1
            X = nodeXML(i).Text
            If X = "1Вручение адресату" _
            Or X = "8Адресату курьером" _
            Or X = "6Адресату почтальоном" _
            Then
            s = "Получено адресатом"
                For y = 0 To nodeXML.Length - 1
                X = nodeXML(y).Text
                    If X = "2Вручение отправителю" _
                    Or X = "7Отправителю почтальоном" _
                    Or X = "8Отправителю курьером" _
                Then
                        s = "Получено отправителем"
                     Exit For
                    End If
                Next y
            
            Exit For

            ElseIf X = "2Вручение отправителю" _
            Or X = "7Отправителю почтальоном" _
            Or X = "8Отправителю курьером" _
            Then
            s = "Получено отправителем"
            Exit For
            Else
            s = "В пути"
            End If
           

        Next i

t = GetTickCount


        Range("q" & n) = s
        
      'clean up code
        Set Req = Nothing
        Set Req2 = Nothing
        Set Resp = Nothing
        Set nodeXML = Nothing
        Set pDoc = Nothing
        sEnv = ""
        
        
        
Debug.Print (GetTickCount - t) / 1000, vbInformation
Application.ScreenUpdating = True
    
End Sub

Private Sub CommandButton19_Click()
   
    Trsdate = TextBox1.Text
    ddt = TextBox8.Text
    If CheckBox1.Value Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Екатеринбург"
        
         
        If Workbooks("main.xlsb").Sheets(1).Range("c1") > 0 Then
        X = "Количество коробок (20 кг, дшв 1120х800х190) - " & Workbooks("main.xlsb").Sheets(1).Range("c1")
        ElseIf X = 0 Then
        X = ""
        
        End If
        
            With objMail
                .Display
                .To = "sg.suhova@ponyexpress.ru; ekaterinburg.all@ponyexpress.ru"
                .CC = "ChuchalovVY@monobrand-tt.ru"
                .Subject = "Заказ ИМ на " & Trsdate & " ООО Торговые технологии/дог.22-50242 " & city
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p>Заказ ИМ на " & Trsdate & ", дог. 22-50242<br>" _
                & "Количество пакетов (1 кг) - " & Workbooks("main.xlsb").Sheets(1).Range("b1") & " шт.<br>" _
                & "<span style ='color:red;'>" & X & "</span><br>" _
                & "<span style ='color:red;'>" & Workbooks("main.xlsb").Sheets(1).Range("e1") & "</span><br>" _
                & "Адрес: 620024 г. Екатеринбург, по ул. Бисертской, 145 (литер АА1)</p>" _
                & "<p>Просьба подтвердить получение письма.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                
                '
                '.Send
            End With
        X = 0
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox2.Value Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Санкт-Петербург"
        
        
          If Workbooks("main.xlsb").Sheets(1).Range("c2") > 0 Then
        X = "Количество коробок (20 кг, дшв 1120х800х190) - " & Workbooks("main.xlsb").Sheets(1).Range("c2")
        ElseIf X = 0 Then
        X = ""
        End If
            With objMail
                .Display
                .To = "oa.pichmanova@ponyexpress.ru"
                .CC = "ChuchalovVY@monobrand-tt.ru"
                .Subject = "Заказ ИМ на " & Trsdate & " ООО Торговые технологии/дог.22-50242 " & city
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p>Заказ ИМ на " & Trsdate & ", дог. 22-50242<br>" _
                & "Количество пакетов (1 кг) - " & Workbooks("main.xlsb").Sheets(1).Range("b2") & " шт.<br>" _
                & "<span style ='color:red;'>" & X & "</span><br>" _
                & "<span style ='color:red;'>" & Workbooks("main.xlsb").Sheets(1).Range("e2") & "</span><br>" _
                & "Адрес: 196084, г. Санкт-Петербург, Витебский пр., д. 3, лит. Б1.</p>" _
                & "<p>Просьба подтвердить получение письма.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        X = 0
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox3.Value Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Нижний Новгород"
        
        If Workbooks("main.xlsb").Sheets(1).Range("c3") > 0 Then
        X = "Количество коробок (20 кг, дшв 1120х800х190) - " & Workbooks("main.xlsb").Sheets(1).Range("c3")
        ElseIf X = 0 Then
        X = ""
        End If
        
            With objMail
                .Display
                .To = "ll.sakhokiya@ponyexpress.ru; nizhniynovgorod.all@ponyexpress.ru"
                .CC = "ChuchalovVY@monobrand-tt.ru"
                .Subject = "Заказ ИМ на " & Trsdate & " ООО Торговые технологии/дог.22-50242 " & city
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p>Заказ ИМ на " & Trsdate & ", дог. 22-50242<br>" _
                & "Количество пакетов (1 кг) - " & Workbooks("main.xlsb").Sheets(1).Range("b3") & " шт.<br>" _
                & "<span style ='color:red;'>" & X & "</span><br>" _
                & "<span style ='color:red;'>" & Workbooks("main.xlsb").Sheets(1).Range("e3") & "</span><br>" _
                & "Адрес: 603127, г.Нижний Новгород, Сормовский район, 7-й Микрорайон, Сормовский промузел, ул. Коновалова, д.10/1.</p>" _
                & "<p>Просьба подтвердить получение письма.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        X = 0
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox4.Value Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Новосибирск"
        
        
          If Workbooks("main.xlsb").Sheets(1).Range("c4") > 0 Then
        X = "Количество коробок (20 кг, дшв 1120х800х190) - " & Workbooks("main.xlsb").Sheets(1).Range("c4")
        ElseIf X = 0 Then
        X = ""
        End If
            With objMail
                .Display
                .To = "novosibirsk.order@ponyexpress.ru"
                .CC = "ChuchalovVY@monobrand-tt.ru"
                .Subject = "Заказ ИМ на " & Trsdate & " ООО Торговые технологии/дог.22-50242 " & city
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p>Заказ ИМ на " & Trsdate & ", дог. 22-50242<br>" _
                & "Количество пакетов (1 кг) - " & Workbooks("main.xlsb").Sheets(1).Range("b4") & " шт.<br>" _
                & "<span style ='color:red;'>" & X & "</span><br>" _
                & "<span style ='color:red;'>" & Workbooks("main.xlsb").Sheets(1).Range("e4") & "</span><br>" _
                & "Адрес: 630088, г. Новосибирск, ул. Петухова, дом. 35, корпус 6.</p>" _
                & "<p>Просьба подтвердить получение письма.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        X = 0
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox5.Value Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Тула"
        
          If Workbooks("main.xlsb").Sheets(1).Range("c5") > 0 Then
        X = "Количество коробок (20 кг, дшв 1120х800х190) - " & Workbooks("main.xlsb").Sheets(1).Range("c5")
        ElseIf X = 0 Then
        X = ""
        End If
            With objMail
                .Display
                .To = "no.tyuftyakova@ponyexpress.ru; tula.order@ponyexpress.ru"
                .CC = "ChuchalovVY@monobrand-tt.ru; oa.pichmanova@ponyexpress.ru; ay.popovich@ponyexpress.ru"
                .Subject = "Заказ ИМ на " & Trsdate & " ООО Торговые технологии/дог.22-50242 " & city
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p>Заказ ИМ на " & Trsdate & ", дог. 22-50242<br>" _
                & "Количество пакетов (1 кг) - " & Workbooks("main.xlsb").Sheets(1).Range("b5") & " шт.<br>" _
                & "<span style ='color:red;'>" & X & "</span><br>" _
                & "<span style ='color:red;'>" & Workbooks("main.xlsb").Sheets(1).Range("e5") & "</span><br>" _
                & "Адрес: 301107, Ленинский район, сельское поселение Шатское, поселок Шатск, строение 2/1.</p>" _
                & "<p>Просьба подтвердить получение письма.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        X = 0
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox6.Value Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Ростов-на-Дону"
        
          If Workbooks("main.xlsb").Sheets(1).Range("c6") > 0 Then
        X = "Количество коробок (20 кг, дшв 1120х800х190) - " & Workbooks("main.xlsb").Sheets(1).Range("c6")
        ElseIf X = 0 Then
        X = ""
        End If
            With objMail
                .Display
                .To = "rostov-na-dony.order@ponyexpress.ru; ls.borodina@ponyexpress.ru"
                .CC = "ChuchalovVY@monobrand-tt.ru"
                .Subject = "Заказ ИМ на " & Trsdate & " ООО Торговые технологии/дог.22-50242 " & city
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p>Заказ ИМ на " & Trsdate & ", дог. 22-50242<br>" _
                & "Количество пакетов (1 кг) - " & Workbooks("main.xlsb").Sheets(1).Range("b6") & " шт.<br>" _
                & "<span style ='color:red;'>" & X & "</span><br>" _
                & "<span style ='color:red;'>" & Workbooks("main.xlsb").Sheets(1).Range("e6") & "</span><br>" _
                & "Адрес: 344092, г. Ростов-на-Дону, Стартовая,д. 3/11, Литер (Л).</p>" _
                & "<p>Просьба подтвердить получение письма.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        X = 0
        Set objMail = Nothing
        Set objOL = Nothing
    End If
    
    If CheckBox7.Value Then
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        city = "Саратов"
        
          If Workbooks("main.xlsb").Sheets(1).Range("c7") > 0 Then
        X = "Количество коробок (20 кг, дшв 1120х800х190) - " & Workbooks("main.xlsb").Sheets(1).Range("c7")
        ElseIf X = 0 Then
        X = ""
        End If
        
            With objMail
                .Display
                .To = "saratov.all@ponyexpress.ru"
                .CC = "ChuchalovVY@monobrand-tt.ru"
                .Subject = "Заказ ИМ на " & Trsdate & " ООО Торговые технологии/дог.22-50242 " & city
                .HTMLBody = "<p>Коллеги, добрый день.</p>" _
                & "<p>Заказ ИМ на " & Trsdate & ", дог. 22-50242<br>" _
                & "Количество пакетов (1 кг) - " & Workbooks("main.xlsb").Sheets(1).Range("b7") & " шт.<br>" _
                & "<span style ='color:red;'>" & X & "</span><br>" _
                & "<span style ='color:red;'>" & Workbooks("main.xlsb").Sheets(1).Range("e7") & "</span><br>" _
                & "Адрес: 410047,Саратовская область, г. Саратов, пос. Мирный, Б/Н</p>" _
                & "<p>Просьба подтвердить получение письма.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                If CheckBox8.Value Then
                    .DeferredDeliveryTime = Date + ddt / 24
                End If
                
                '.Send
            End With
        X = 0
        Set objMail = Nothing
        Set objOL = Nothing
    End If
End Sub

Private Sub CommandButton190_Click()
 Application.ScreenUpdating = False
 Application.Calculation = xlCalculationManual

t = GetTickCount

Dim a As Date
Dim b As Date



f = Cells(Rows.Count, 1).End(xlUp).Row
a = TextBox22
b = TextBox23

For u = 2 To f
If Range("f" & u) >= a And Range("f" & u) <= b Then
If Left(Range("h" & u), 3) = "800" Then

t = GetTickCount

    Dim Req As Object
    Dim sEnv As String
    Dim Resp As New MSXML2.DOMDocument60
    Set Req = CreateObject("MSXML2.XMLHTTP")
    Set Resp = CreateObject("MSXML2.DOMDocument.6.0")
    
        Req.Open "Post", "https://tracking.russianpost.ru/rtm34", False
         sEnv = sEnv & "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:oper=""http://russianpost.org/operationhistory"" xmlns:data=""http://russianpost.org/operationhistory/data"" xmlns:ns1=""http://schemas.xmlsoap.org/soap/envelope/"">"
         sEnv = sEnv & "<soap:Header/>"
         sEnv = sEnv & "<soap:Body>"
         sEnv = sEnv & "<oper:getOperationHistory>"
         sEnv = sEnv & "<data:OperationHistoryRequest>"
         sEnv = sEnv & "<data:Barcode>" & Range("h" & u) & "</data:Barcode>"
         sEnv = sEnv & "<data:MessageType>0</data:MessageType>"
         sEnv = sEnv & "<data:Language>RUS</data:Language>"
         sEnv = sEnv & "</data:OperationHistoryRequest>"
         sEnv = sEnv & "<data:AuthorizationHeader ns1:mustUnderstand=""?"">"
         sEnv = sEnv & "<data:login>ykDaLTEChMLavX</data:login>"
         sEnv = sEnv & "<data:password>JPOIsPTd3W03</data:password>"
         sEnv = sEnv & "</data:AuthorizationHeader>"
         sEnv = sEnv & "</oper:getOperationHistory>"
         sEnv = sEnv & "</soap:Body>"
         sEnv = sEnv & "</soap:Envelope>"
        Req.send (sEnv)

        Req2 = Replace(Req.responseText, "ns3:OperAttr", "OperAttr")
   
        Dim pDoc As New MSXML2.DOMDocument60
        pDoc.LoadXML Req2
        
        Set nodeXML = pDoc.getElementsByTagName("OperAttr")
        
        For i = 1 To nodeXML.Length - 1
            X = nodeXML(i).Text
            If X = "1Вручение адресату" _
            Or X = "8Адресату курьером" _
            Or X = "6Адресату почтальоном" _
            Then
            s = "Получено адресатом"
                For y = 0 To nodeXML.Length - 1
                X = nodeXML(y).Text
                    If X = "2Вручение отправителю" _
                    Or X = "7Отправителю почтальоном" _
                    Or X = "8Отправителю курьером" _
                Then
                        s = "Получено отправителем"
                     Exit For
                    End If
                Next y
                Exit For
            
            ElseIf X = "2Вручение отправителю" _
                Or X = "7Отправителю почтальоном" _
                Or X = "8Отправителю курьером" _
            Then
                s = "Получено отправителем"
                Exit For
            Else
            s = "В пути"
            End If
            
        Next i
        
        On Error Resume Next
        Range("q" & u) = s

        Set s = Nothing
        Set X = Nothing
        Set Req = Nothing
        Set Req2 = Nothing
        Set Resp = Nothing
        Set nodeXML = Nothing
        Set pDoc = Nothing
        sEnv = ""

Range("r" & u) = (GetTickCount - t) / 1000
Debug.Print "Заказ" & (GetTickCount - t) / 1000, vbInformation
End If
End If
Next u


Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton191_Click()

Application.ScreenUpdating = False

Dim a As Date
Dim t As Date
t = Time


f = Cells(Rows.Count, 1).End(xlUp).Row
a = "24.09.2021"
'For u = 2 To f
u = 50423

If Range("f" & u) >= a Then
If Left(Range("h" & u), 3) = "800" Then


    Dim Req As Object
    Dim sEnv As String
    Dim Resp As New MSXML2.DOMDocument60
    Set Req = CreateObject("MSXML2.XMLHTTP")
    Set Resp = CreateObject("MSXML2.DOMDocument.6.0")
    

    

    
        Req.Open "Post", "https://tracking.russianpost.ru/rtm34", False
         sEnv = sEnv & "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:oper=""http://russianpost.org/operationhistory"" xmlns:data=""http://russianpost.org/operationhistory/data"" xmlns:ns1=""http://schemas.xmlsoap.org/soap/envelope/"">"
         sEnv = sEnv & "<soap:Header/>"
         sEnv = sEnv & "<soap:Body>"
         sEnv = sEnv & "<oper:getOperationHistory>"
         sEnv = sEnv & "<data:OperationHistoryRequest>"
         sEnv = sEnv & "<data:Barcode>" & Range("h" & u) & "</data:Barcode>"
         sEnv = sEnv & "<data:MessageType>0</data:MessageType>"
         sEnv = sEnv & "<data:Language>RUS</data:Language>"
         sEnv = sEnv & "</data:OperationHistoryRequest>"
         sEnv = sEnv & "<data:AuthorizationHeader ns1:mustUnderstand=""?"">"
         sEnv = sEnv & "<data:login>ykDaLTEChMLavX</data:login>"
         sEnv = sEnv & "<data:password>JPOIsPTd3W03</data:password>"
         sEnv = sEnv & "</data:AuthorizationHeader>"
         sEnv = sEnv & "</oper:getOperationHistory>"
         sEnv = sEnv & "</soap:Body>"
         sEnv = sEnv & "</soap:Envelope>"
    ' Send SOAP Request
        Req.send (sEnv)
'        Debug.Print Req
        
        
        Req2 = Replace(Req.responseText, "ns3:OperationParameters", "OperationParameters")
        
        
    
        Dim pDoc As New MSXML2.DOMDocument60
        pDoc.LoadXML Req2
        
        Set nodeXML = pDoc.getElementsByTagName("OperationParameters")

        For i = 1 To nodeXML.Length - 1
            X = nodeXML(i).Text
            Debug.Print X
        Next i
        Debug.Print nodeXML.Length - 1
        Debug.Print X
        
'        req3 = Replace(x.responseText, "ns3:Name", "Name")
'        Debug.Print req3
'
'        pDoc.LoadXML req3
'        Set nodeXML = pDoc.getElementsByTagName("Name")
'
'        For i = 1 To nodeXML.Length - 1
'            x = nodeXML(i).Text
'        Next i
        
        
        
        Range("q" & u) = X
      'clean up code
        Set Req = Nothing
        Set Req2 = Nothing
        Set Resp = Nothing
        Set nodeXML = Nothing
        Set pDoc = Nothing
        sEnv = ""
        
        
        
        
End If
End If
'Next u
t = Time - t
MsgBox (t)
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton192_Click()
Dim a As Date
Dim t As Date
t = Time


n = ActiveCell.Row




    Dim Req As Object
    Dim sEnv As String
    Dim Resp As New MSXML2.DOMDocument60
    Set Req = CreateObject("MSXML2.XMLHTTP")
    Set Resp = CreateObject("MSXML2.DOMDocument.6.0")
    

    

    
        Req.Open "Post", "https://svc-api.p2e.ru/UI_Service.svc?singleWsdl", False
         sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tem=""http://tempuri.org/"">"
         sEnv = sEnv & "<soapenv:Header/>"
         sEnv = sEnv & "<soapenv:Body>"
         sEnv = sEnv & "<tem:SubmitRequest>"
         sEnv = sEnv & "<tem:accesskey>32f4cd13-e64f-4ae2-8c4b-cdd67bbd491f</tem:accesskey>"
         sEnv = sEnv & "<tem:requestBody>"
         sEnv = sEnv & "<![CDATA[<Request xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xsi:type=""OrderRequest"">"
         sEnv = sEnv & "<Mode>Status</Mode>"
         sEnv = sEnv & "<OrderList>"
         sEnv = sEnv & "<Order>"
         sEnv = sEnv & "<ServiceList>"
         sEnv = sEnv & "<Service xsi:type=""DeliveryService"">"
         sEnv = sEnv & "<Waybill>"
         sEnv = sEnv & "<Number>26-9228-0459</Number>"
         sEnv = sEnv & "</Waybill></Service></ServiceList></Order></OrderList></Request>]]>"
         sEnv = sEnv & "</tem:requestBody>"
         sEnv = sEnv & "</tem:SubmitRequest>"
         sEnv = sEnv & "</soapenv:Body>"
         sEnv = sEnv & "</soapenv:Envelope>"
         
         
   
    ' Send SOAP Request
        Req.send (sEnv)

        
        
'        Req2 = Replace(Req.responseText, "ns3:Name", "Name")
        
        Range("a1") = Req.responseText
        
'        Dim pDoc As New MSXML2.DOMDocument60
'        pDoc.LoadXML Req2
'        Debug.Print Req2
'
'        Set nodeXML = pDoc.getElementsByTagName("Name")
'
'        For i = 1 To nodeXML.Length - 1
'        x = nodeXML(i).Text
''            If x = "Вручение адресату" _
''            Or x = "Адресату курьером" _
''            Or x = "Адресату почтальоном" _
''            Then
''            x = "Получено адресатом"
''            Exit For
''            ElseIf x = "Вручение отправителю" _
''            Or x = "Отправителю почтальоном" _
''            Or x = "Отправителю курьером" _
''            Then
''            x = "Получено отправителем"
''            Exit For
''            End If
'        Next i
'
'        Range("q" & n) = x
        
      'clean up code
        Set Req = Nothing
        Set Req2 = Nothing
        Set Resp = Nothing
        Set nodeXML = Nothing
        Set pDoc = Nothing
        sEnv = ""
        
        
        
    

t = Time - t
MsgBox (t)
End Sub

Private Sub CommandButton193_Click()
Dim a As Date
Dim b As Date
Dim t As Date

t1 = GetTickCount



u = ActiveCell.Row


 
    Dim Req As Object
    Dim sEnv As String
    Dim Resp As New MSXML2.DOMDocument60
    Set Req = CreateObject("MSXML2.XMLHTTP")
    Set Resp = CreateObject("MSXML2.DOMDocument.6.0")
    
    t = Time
    n = ActiveCell.Row
    
        Req.Open "POST", "https://svc-api.p2e.ru/UI_Service.asmx?WSDL", False
        Req.setRequestHeader "Content-Type", "text/xml"
        sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tem=""http://tempuri.org/""><soapenv:Header/><soapenv:Body><tem:SubmitRequest><tem:accessKey>32f4cd13-e64f-4ae2-8c4b-cdd67bbd491f</tem:accessKey><tem:requestBody><![CDATA[<Request xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:type=""OrderRequest""><Mode>Status</Mode><OrderList><Order>"
        sEnv = sEnv & "<ClientsNumber>" & Range("g" & u) & "</ClientsNumber>"
        sEnv = sEnv & "</Order></OrderList></Request>]]></tem:requestBody></tem:SubmitRequest></soapenv:Body></soapenv:Envelope>"
        Req.send (sEnv)

        Req2 = Req.responseText
        

        
        
        Req2 = Replace(Req.responseText, "&gt;", ">")
        Req2 = Replace(Req2, "&lt;", "<")
        Req2 = Replace(Req2, "xsi:type;", "type")
        
        
        Dim xmlDoc As Object, post As Object
        Set xmlDoc = CreateObject("Microsoft.XMLDOM")
        xmlDoc.SetProperty "SelectionLanguage", "XPath"
        xmlDoc.async = False
        
        
        xmlDoc.LoadXML Req2
        
        Set nodeXML = xmlDoc.getElementsByTagName("PegasEventCode")
        For i = 0 To nodeXML.Length - 1
        X = nodeXML(i).Text
        
        
            If X = "17" _
            Or X = "140" _
            Or X = "572" _
            Or X = "574" _
            Or X = "146" _
            Or X = "7" _
            Or X = "147" _
            Then
            s = "Возврат в пути"
                For y = 0 To nodeXML.Length - 1
                X = nodeXML(y).Text
                    If X = "98" _
                    Then
                     s = "Возврат"
                     Exit For
                    End If
                Next y
            
            
            
            Exit For
            
            ElseIf X = "610" _
            Or X = "98" _
            Then
            s = "Доставлен"
            Else
            s = "В пути"
            
            End If
        Next
        Range("q" & u) = s
        


        Set Req = Nothing
        Set Req2 = Nothing
        Set Resp = Nothing
        Set nodeXML = Nothing
        Set pDoc = Nothing
        sEnv = ""
        Set s = Nothing
        Set X = Nothing

        
             
        


t2 = GetTickCount
Debug.Print t2 - t1

End Sub

Private Sub CommandButton194_Click()
Dim a As Date
Dim t As Date
t = Time


n = ActiveCell.Row




    Dim Req As Object
    Dim sEnv As String
    Dim Resp As New MSXML2.DOMDocument60
    Set Req = CreateObject("MSXML2.XMLHTTP")
    Set Resp = CreateObject("MSXML2.DOMDocument.6.0")
    

    

    
        Req.Open "POST", "https://svc-api.p2e.ru/UI_Service.asmx?WSDL", False
        Req.setRequestHeader "Content-Type", "text/xml"
        sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tem=""http://tempuri.org/""><soapenv:Header/><soapenv:Body><tem:SubmitRequest><tem:accessKey>32f4cd13-e64f-4ae2-8c4b-cdd67bbd491f</tem:accessKey><tem:requestBody><![CDATA[<Request xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xsi:type=""OrderRequest""><Mode>Status</Mode> <!--Required--><OrderList> <!--Required--><Order> <!--Required--><ServiceList> <!--Required--><Service xsi:type=""DeliveryService""> <!--Required--><Waybill>"
        sEnv = sEnv & "<Number>26-1469-1552</Number>"
        sEnv = sEnv & "</Waybill></Service></ServiceList></Order></OrderList></Request>]]></tem:requestBody></tem:SubmitRequest></soapenv:Body></soapenv:Envelope>"
    ' Send SOAP Request
        Req.send (sEnv)

        
        Req2 = Req.responseText
        Req2 = Replace(Req.responseText, "&gt;", ">")
        Req2 = Replace(Req2, "&lt;", "<")
        
        
'        req2 = Replace(req2, "ServiceStatus xsi:type=""DeliveryStatus""", "st")
'        req2 = Replace(req2, "ServiceStatus", "st")


            Debug.Print Req2
            Dim y&, f&, w
             For Each w In Split(Req2, ">")
                If w = "Description" Then y = y + 1
            Next
            MsgBox (y)
'
'        vvv = Split(Split(req2, "<Description>")(1), "</Description>")(0)
'
'
'        Debug.Print vvv
'
'        Dim pDoc As New MSXML2.DOMDocument60
'        pDoc.LoadXML req2
'
'        Debug.Print pDoc
        
        'Set nodeXML = pDoc.getElementsByTagName("StatusList")
        'Debug.Print pDoc.SelectSingleNode("StatusList").Text
        
        'Debug.Print nodeXML
        
        
'        For i = 1 To nodeXML.Length - 1
'        x = nodeXML(i).Text
'        Next i
'
'        MsgBox nodeXML.Length
        
        
'       Range("q" & n) = x
        
      'clean up code
        Set Req = Nothing
        Set Req2 = Nothing
        Set Resp = Nothing
        Set nodeXML = Nothing
        Set pDoc = Nothing
        sEnv = ""
        
        
        
    

t = Time - t
MsgBox (t)
End Sub

Private Sub CommandButton195_Click()
    Dim a As Date
    Dim t As Date
    Dim Req As Object
    Dim sEnv As String
    Dim Resp As New MSXML2.DOMDocument60
    Set Req = CreateObject("MSXML2.XMLHTTP")
    Set Resp = CreateObject("MSXML2.DOMDocument.6.0")
    
    t = Time
    n = ActiveCell.Row
    
        Req.Open "POST", "https://svc-api.p2e.ru/UI_Service.asmx?WSDL", False
        Req.setRequestHeader "Content-Type", "text/xml"
        sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tem=""http://tempuri.org/""><soapenv:Header/><soapenv:Body><tem:SubmitRequest><tem:accessKey>32f4cd13-e64f-4ae2-8c4b-cdd67bbd491f</tem:accessKey><tem:requestBody><![CDATA[<Request xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xsi:type=""OrderRequest""><Mode>Status</Mode> <!--Required--><OrderList> <!--Required--><Order> <!--Required--><ServiceList> <!--Required--><Service xsi:type=""DeliveryService""> <!--Required--><Waybill>"
        sEnv = sEnv & "<Number>26-9211-6086</Number>"
        sEnv = sEnv & "</Waybill></Service></ServiceList></Order></OrderList></Request>]]></tem:requestBody></tem:SubmitRequest></soapenv:Body></soapenv:Envelope>"
        Req.send (sEnv)

        Req2 = Req.responseText
        

        
        
        Req2 = Replace(Req.responseText, "&gt;", ">")
        Req2 = Replace(Req2, "&lt;", "<")
        
        
        Dim xmlDoc As Object, post As Object
        Set xmlDoc = CreateObject("Microsoft.XMLDOM")
        xmlDoc.SetProperty "SelectionLanguage", "XPath"
        xmlDoc.async = False
        
        
        xmlDoc.LoadXML Req2
        
        Set nodeXML = xmlDoc.getElementsByTagName("ServiceStatus")
        For i = 0 To nodeXML.Length - 9
        X = nodeXML(i).Text
        Next
        Debug.Print X
        


        Set Req = Nothing
        Set Req2 = Nothing
        Set Resp = Nothing
        Set nodeXML = Nothing
        Set pDoc = Nothing
        sEnv = ""

t = Time - t
MsgBox (t)
End Sub

Private Sub CommandButton196_Click()
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.SetProperty "SelectionLanguage", "XPath"
xmlDoc.async = False






'xmldoc.Load ("C:\Users\ShapkaMY\Desktop\1\test1.xml")
'
'n = ("C:\Users\ShapkaMY\Desktop\1\test1.xml")
'
'
'
'
y = Range("a1")
'
'Range("b1") = y
'
'
xmlDoc.xml
xmlDoc.Load y
'
Set nodeXML = xmlDoc.getElementsByTagName("ServiceStatus")
For i = 0 To nodeXML.Length - 3
X = nodeXML(i).Text
Next
MsgBox X





End Sub

Private Sub CommandButton197_Click()
Range("a2") = Trim(Range("a2"))


End Sub

Private Sub CommandButton198_Click()


End Sub

Private Sub CommandButton199_Click()
 Dim a As Date
    Dim t As Date
    Dim Req As Object
    Dim sEnv As String
    Dim Resp As New MSXML2.DOMDocument60
    Set Req = CreateObject("MSXML2.XMLHTTP")
    Set Resp = CreateObject("MSXML2.DOMDocument.6.0")
    
    t = Time
    n = ActiveCell.Row
    
        Req.Open "POST", "https://svc-api.p2e.ru/UI_Service.asmx?WSDL", False
        Req.setRequestHeader "Content-Type", "text/xml"
        sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tem=""http://tempuri.org/""><soapenv:Header/><soapenv:Body><tem:SubmitRequest><tem:accessKey>32f4cd13-e64f-4ae2-8c4b-cdd67bbd491f</tem:accessKey><tem:requestBody><![CDATA[<Request xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xsi:type=""OrderRequest""><Mode>Status</Mode> <!--Required--><OrderList> <!--Required--><Order> <!--Required--><ServiceList> <!--Required--><Service xsi:type=""DeliveryService""> <!--Required--><Waybill>"
        sEnv = sEnv & "<Number>26-9211-6086</Number>"
        sEnv = sEnv & "</Waybill></Service></ServiceList></Order></OrderList></Request>]]></tem:requestBody></tem:SubmitRequest></soapenv:Body></soapenv:Envelope>"
        Req.send (sEnv)

        Req2 = Req.responseText
        

        
        
        Req2 = Replace(Req.responseText, "&gt;", ">")
        Req2 = Replace(Req2, "&lt;", "<")
        Req2 = Replace(Req2, "xsi:type", "type")
        
        Dim doc_XML As DOMDocument60

        Set doc_XML = New DOMDocument60
        
        'Data = winHttpReq.responseText
        doc_XML.Load Req2
        
        Set List = doc_XML.DocumentElement.ChildNodes
        For Each sub_list In List
            If sub_list.Attributes(0).Text = "Response" Then
                For Each Node In sub_list.ChildNodes(0).ChildNodes
                    If Node.Attributes(0).Text = "DeliveryStatus" Then
                        result = Node.nodeTypedValue
                    End If
                Next Node
            End If
        Next sub_list




        


        Set Req = Nothing
        Set Req2 = Nothing
        Set Resp = Nothing
        Set nodeXML = Nothing
        Set pDoc = Nothing
        sEnv = ""

t = Time - t
MsgBox (t)
End Sub

Private Sub CommandButton2_Click()
    Range("A:AA").Copy
    Sheets.Add.Name = "Отправление"
    Range("A1").PasteSpecial Paste:=xlPasteValues
    
    f = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To f
        If Range("K" & i) = "-" Then
        Range("K" & i).Rows.Clear
        End If
    Next i
    
    For i = 1 To f
        If Range("V" & i) = "Накладная (листовка) к заказу ХШР Медиа" Or Range("V" & i) = "Накладная" Then
        Range("V" & i).Rows.Clear
        End If
    Next i
    
    f = Cells(Rows.Count, 1).End(xlUp).Row
    Range("k1:k" & f).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Range("V1:V" & f).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
End Sub

Private Sub CommandButton20_Click()
    Dim fso As Object, i As Integer
    'Создаем новый экземпляр FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Создаем несколько новых папок
    With fso
        .CreateFolder ("C:\Users\ShapkaMY\Desktop\" & TextBox1.Text)
                .CreateFolder ("C:\Users\ShapkaMY\Desktop\" & TextBox1.Text & "\Екатеринбург.txt")
                .CreateFolder ("C:\Users\ShapkaMY\Desktop\" & TextBox1.Text & "\Нижний Новгород")
                .CreateFolder ("C:\Users\ShapkaMY\Desktop\" & TextBox1.Text & "\Ростов-на-Дону")
                .CreateFolder ("C:\Users\ShapkaMY\Desktop\" & TextBox1.Text & "\Санкт-Петербург")
                .CreateFolder ("C:\Users\ShapkaMY\Desktop\" & TextBox1.Text & "\Тула")

    End With
    
End Sub

Private Sub CommandButton200_Click()
Dim a As Date
    Dim t As Date
    Dim Req As Object
    Dim sEnv As String
    Dim Resp As New MSXML2.DOMDocument60
    Set Req = CreateObject("MSXML2.XMLHTTP")
    Set Resp = CreateObject("MSXML2.DOMDocument.6.0")
    
    t = Time
    n = ActiveCell.Row
    
        Req.Open "POST", "https://svc-api.p2e.ru/UI_Service.asmx?WSDL", False
        Req.setRequestHeader "Content-Type", "text/xml"
        sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tem=""http://tempuri.org/""><soapenv:Header/><soapenv:Body><tem:SubmitRequest><tem:accessKey>32f4cd13-e64f-4ae2-8c4b-cdd67bbd491f</tem:accessKey><tem:requestBody><![CDATA[<Request xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xsi:type=""OrderRequest""><Mode>Status</Mode> <!--Required--><OrderList> <!--Required--><Order> <!--Required--><ServiceList> <!--Required--><Service xsi:type=""DeliveryService""> <!--Required--><Waybill>"
        sEnv = sEnv & "<Number>26-9211-5437</Number>"
        sEnv = sEnv & "</Waybill></Service></ServiceList></Order></OrderList></Request>]]></tem:requestBody></tem:SubmitRequest></soapenv:Body></soapenv:Envelope>"
        Req.send (sEnv)

        Req2 = Req.responseText
        

        
        
        Req2 = Replace(Req.responseText, "&gt;", ">")
        Req2 = Replace(Req2, "&lt;", "<")
        Req2 = Replace(Req2, "xsi:type;", "type")
        
        
        Dim xmlDoc As Object, post As Object
        Set xmlDoc = CreateObject("Microsoft.XMLDOM")
        xmlDoc.SetProperty "SelectionLanguage", "XPath"
        xmlDoc.async = False
        
        
        xmlDoc.LoadXML Req2
        
        Set nodeXML = xmlDoc.getElementsByTagName("PegasEventCode")
        For i = 0 To nodeXML.Length - 1
        X = nodeXML(i).Text
        Next
        Debug.Print X
        


        Set Req = Nothing
        Set Req2 = Nothing
        Set Resp = Nothing
        Set nodeXML = Nothing
        Set pDoc = Nothing
        sEnv = ""

t = Time - t
MsgBox (t)
End Sub

Private Sub CommandButton201_Click()
Dim a As Date
Dim b As Date
Dim t As Date

t = Time
Application.ScreenUpdating = False





f = Cells(Rows.Count, 1).End(xlUp).Row
a = TextBox22
b = TextBox23

For u = 2 To f
If Range("f" & u) >= a And Range("f" & u) <= b Then
If Left(Range("h" & u), 2) = "26" Then
 
    Dim Req As Object
    Dim sEnv As String
    Dim Resp As New MSXML2.DOMDocument60
    Set Req = CreateObject("MSXML2.XMLHTTP")
    Set Resp = CreateObject("MSXML2.DOMDocument.6.0")
    
    t = Time
    n = ActiveCell.Row
    
        Req.Open "POST", "https://svc-api.p2e.ru/UI_Service.asmx?WSDL", False
        Req.setRequestHeader "Content-Type", "text/xml"
        sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tem=""http://tempuri.org/""><soapenv:Header/><soapenv:Body><tem:SubmitRequest><tem:accessKey>32f4cd13-e64f-4ae2-8c4b-cdd67bbd491f</tem:accessKey><tem:requestBody><![CDATA[<Request xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xsi:type=""OrderRequest""><Mode>Status</Mode> <!--Required--><OrderList> <!--Required--><Order> <!--Required--><ServiceList> <!--Required--><Service xsi:type=""DeliveryService""> <!--Required--><Waybill>"
        sEnv = sEnv & "<Number>" & Range("h" & u) & "</Number>"
        sEnv = sEnv & "</Waybill></Service></ServiceList></Order></OrderList></Request>]]></tem:requestBody></tem:SubmitRequest></soapenv:Body></soapenv:Envelope>"
        Req.send (sEnv)

        Req2 = Req.responseText
        

        
        
        Req2 = Replace(Req.responseText, "&gt;", ">")
        Req2 = Replace(Req2, "&lt;", "<")
        Req2 = Replace(Req2, "xsi:type;", "type")
        
        
        Dim xmlDoc As Object, post As Object
        Set xmlDoc = CreateObject("Microsoft.XMLDOM")
        xmlDoc.SetProperty "SelectionLanguage", "XPath"
        xmlDoc.async = False
        
        
        xmlDoc.LoadXML Req2
        
        Set nodeXML = xmlDoc.getElementsByTagName("PegasEventCode")
        For i = 0 To nodeXML.Length - 1
        X = nodeXML(i).Text
        
        
            If X = "572" _
            Or X = "574" _
            Then
            s = "Возврат"
            Exit For
            
            ElseIf X = "98" _
            Or X = "610" _
            Then
            s = "Доставлен"

            End If
        Next
        Range("q" & u) = s
        


        Set Req = Nothing
        Set Req2 = Nothing
        Set Resp = Nothing
        Set nodeXML = Nothing
        Set pDoc = Nothing
        sEnv = ""

        
             
        
End If
End If
Next u


Application.ScreenUpdating = True
t = Time - t
MsgBox (t)
End Sub

Private Sub CommandButton202_Click()
Dim a As Date
Dim b As Date
Dim t As Date

t = Time
Application.ScreenUpdating = False





f = Cells(Rows.Count, 1).End(xlUp).Row
a = TextBox22
b = TextBox23

For u = 2 To f
If Range("f" & u) >= a And Range("f" & u) <= b Then
If Left(Range("h" & u), 2) = "26" Then
    

    Dim Req As Object
    Dim sEnv As String
    Dim Resp As New MSXML2.DOMDocument60
    Set Req = CreateObject("MSXML2.XMLHTTP")
    Set Resp = CreateObject("MSXML2.DOMDocument.6.0")
    
    t = Time
    n = ActiveCell.Row
    
        Req.Open "POST", "https://svc-api.p2e.ru/UI_Service.asmx?WSDL", False
        Req.setRequestHeader "Content-Type", "text/xml"
        sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tem=""http://tempuri.org/""><soapenv:Header/><soapenv:Body><tem:SubmitRequest><tem:accessKey>32f4cd13-e64f-4ae2-8c4b-cdd67bbd491f</tem:accessKey><tem:requestBody><![CDATA[<Request xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:type=""OrderRequest""><Mode>Status</Mode><OrderList><Order>"
        sEnv = sEnv & "<ClientsNumber>" & Range("g" & u) & "</ClientsNumber>"
        sEnv = sEnv & "</Order></OrderList></Request>]]></tem:requestBody></tem:SubmitRequest></soapenv:Body></soapenv:Envelope>"
        Req.send (sEnv)

        Req2 = Req.responseText
        

        
        
        Req2 = Replace(Req.responseText, "&gt;", ">")
        Req2 = Replace(Req2, "&lt;", "<")
        Req2 = Replace(Req2, "xsi:type;", "type")
        
        
        Dim xmlDoc As Object, post As Object
        Set xmlDoc = CreateObject("Microsoft.XMLDOM")
        xmlDoc.SetProperty "SelectionLanguage", "XPath"
        xmlDoc.async = False
        
        
        xmlDoc.LoadXML Req2
        
        Set nodeXML = xmlDoc.getElementsByTagName("PegasEventCode")
        For i = 0 To nodeXML.Length - 1
        X = nodeXML(i).Text
        
        
            If X = "17" _
            Or X = "140" _
            Or X = "572" _
            Or X = "574" _
            Or X = "146" _
            Or X = "7" _
            Or X = "147" _
            Then
            s = "Возврат"
            Exit For
            
            ElseIf X = "610" _
            Or X = "98" _
            Then
            s = "Доставлен"

            End If
        Next
        On Error Resume Next
        Range("q" & u) = s
        


        Set Req = Nothing
        Set Req2 = Nothing
        Set Resp = Nothing
        Set nodeXML = Nothing
        Set pDoc = Nothing
        sEnv = ""
        Set s = Nothing
        Set X = Nothing
        
             
        
End If
End If
Next u


Application.ScreenUpdating = True
t = Time - t
MsgBox (t)
End Sub

Private Sub CommandButton203_Click()
Dim a As Date
Dim b As Date
Dim t As Date

t = Time
Application.ScreenUpdating = False





f = Cells(Rows.Count, 1).End(xlUp).Row
a = TextBox22
b = TextBox23

For u = 2 To f
If Range("f" & u) >= a And Range("f" & u) <= b Then
If Left(Range("h" & u), 2) = "26" Then
    

    Dim Req As Object
    Dim sEnv As String
    Dim Resp As New MSXML2.DOMDocument60
    Set Req = CreateObject("MSXML2.XMLHTTP")
    Set Resp = CreateObject("MSXML2.DOMDocument.6.0")
    
    t = Time
    n = ActiveCell.Row
    
        Req.Open "POST", "https://svc-api.p2e.ru/UI_Service.asmx?WSDL", False
        Req.setRequestHeader "Content-Type", "text/xml"
        sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:tem=""http://tempuri.org/""><soapenv:Header/><soapenv:Body><tem:SubmitRequest><tem:accessKey>32f4cd13-e64f-4ae2-8c4b-cdd67bbd491f</tem:accessKey><tem:requestBody><![CDATA[<Request xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:type=""OrderRequest""><Mode>Status</Mode><OrderList><Order>"
        sEnv = sEnv & "<ClientsNumber>" & Range("g" & u) & "</ClientsNumber>"
        sEnv = sEnv & "</Order></OrderList></Request>]]></tem:requestBody></tem:SubmitRequest></soapenv:Body></soapenv:Envelope>"
        Req.send (sEnv)

        Req2 = Req.responseText
        

        
        
        Req2 = Replace(Req.responseText, "&gt;", ">")
        Req2 = Replace(Req2, "&lt;", "<")
        Req2 = Replace(Req2, "xsi:type;", "type")
        
        
        Dim xmlDoc As Object, post As Object
        Set xmlDoc = CreateObject("Microsoft.XMLDOM")
        xmlDoc.SetProperty "SelectionLanguage", "XPath"
        xmlDoc.async = False
        
        
        xmlDoc.LoadXML Req2
        
        Set nodeXML = xmlDoc.getElementsByTagName("PegasEventCode")
        For i = 0 To nodeXML.Length - 1
        X = nodeXML(i).Text
        
        
            If X = "17" _
            Or X = "140" _
            Or X = "572" _
            Or X = "574" _
            Or X = "146" _
            Or X = "7" _
            Or X = "147" _
            Then
            s = "Возврат в пути"
                For y = 0 To nodeXML.Length - 1
                X = nodeXML(y).Text
                    If X = "98" _
                    Then
                     s = "Возврат"
                     Exit For
                    End If
                Next y
            
            Exit For
            
            ElseIf X = "610" _
            Or X = "98" _
            Then
            s = "Доставлен"
            
            End If
        Next
        On Error Resume Next
        Range("q" & u) = s
        


        Set Req = Nothing
        Set Req2 = Nothing
        Set Resp = Nothing
        Set nodeXML = Nothing
        Set pDoc = Nothing
        sEnv = ""
        Set s = Nothing
        Set X = Nothing
        
             
        
End If
End If
Next u


Application.ScreenUpdating = True
t = Time - t
MsgBox (t)
End Sub

Private Sub CommandButton204_Click()
Dim sw As StopWatch
Set sw = New StopWatch
sw.StartTimer

' Do whatever you want to time here

Debug.Print "That took: " & sw.EndTimer & "milliseconds"
End Sub

Private Sub CommandButton205_Click()

    Set Main = Workbooks("main.xlsb")
    Set Table = Workbooks("Table3.xlsx")



    f = Main.Sheets("Итог").Cells(Rows.Count, 1).End(xlUp).Row
    f2 = Table.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row + 1
    

    Main.Sheets("Итог").Rows("1:" & f).Copy
    
    
    Table.Sheets(1).Range("a" & f2).PasteSpecial Paste:=xlPasteValues
    
    With Table.Sheets(1).Range("a:u")
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    
    
Range("h:h").NumberFormat = "#"
End Sub

Private Sub CommandButton206_Click()


f = Cells(Rows.Count, 1).End(xlUp).Row

'Вприм тарифы ИМ
Range("q1:q" & f).FormulaR1C1 = _
        "=IF(COUNTIF(R1C[-14]:RC[-14],RC[-14])=1,VLOOKUP(RC[-14],Статистика.csv!C6:C25,20,0),"" "")"
        
'Вприм курьерскую службу
Range("n1:n" & f).FormulaR1C1 = "=VLOOKUP(RC[-11],Статистика.csv!C6:C17,12,0)"


'Меняем формулы на значение
Range("k:q").Copy
Range("k:q").PasteSpecial Paste:=xlPasteValues

'Меняем формат накладных на числовой, без запятых
Range("d:d").NumberFormat = "#"

End Sub

Private Sub CommandButton207_Click()
    Dim objOutlApp As Object, oNSpace As Object, oIncoming As Object
    Dim oIncMails As Object, oMail As Object, oAtch As Object
    Dim IsNotAppRun As Boolean
    Dim sFolder As String, s As String
    '?????? ??????? ?????? ????? ? ???????
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = False Then Exit Sub
        sFolder = .SelectedItems(1)
    End With
    sFolder = sFolder & IIf(Right(sFolder, 1) = Application.PathSeparator, "", Application.PathSeparator)
    '????????? ?????????? ??????, ????? ???? ???????? ?? ????????
    Application.ScreenUpdating = False
 
    '???????????? ? Outlook
    On Error Resume Next
    Set objOutlApp = GetObject(, "outlook.Application")
    If objOutlApp Is Nothing Then
        Set objOutlApp = CreateObject("outlook.Application")
        IsNotAppRun = True
    End If
    '???????? ?????? ? ?????? ?????
    Set oNSpace = objOutlApp.GetNamespace("MAPI")
    '???????????? ? ????? ????????, ????????? ????? ?? ?????????
    Set oIncoming = oNSpace.GetDefaultFolder(6).Folders("ТРС")
    
    'Set oIncoming = oNSpace.Folders("Personal Folders").Folders("Inbox").Folders("1")
    '????????? ==> GetDefaultFolder(3)
    '????????? ==> GetDefaultFolder(4)
    '???????????? ==> GetDefaultFolder(5)
    '???????? ==> GetDefaultFolder(6)
 
    '???????? ????????? ????? ????????(??????? ????????)
    Set oIncMails = oIncoming.Items
    '????????????? ?????? ??????
    For Each oMail In oIncMails
        '????????????? ?????? ???????? ??????
        For Each oAtch In oMail.Attachments
            '???????? ?????? ????? Excel
            If oAtch Like "*.xl*" Then
                s = GetAtchName(sFolder & oAtch)
               oAtch.SaveAsFile s
            End If
        Next
    Next
    '???? ?????????? Outlook ???? ??????? ????? - ?????????
    If IsNotAppRun Then
        objOutlApp.Quit
    End If
    '??????? ??????????
    Set oIncMails = Nothing
    Set oIncoming = Nothing
    Set oNSpace = Nothing
    Set objOutlApp = Nothing
    '?????????? ????? ??????????? ?????????? ??????
    Application.ScreenUpdating = True
End Sub
'---------------------------------------------------------------------------------------
' Procedure : GetAtchName
' Purpose   : ??????? ????????? ??????????? ????? ?????
'             ???? ???? ? ?????? s ??? ???? - ????????? ????? ? ???????
'---------------------------------------------------------------------------------------
Function GetAtchName(ByVal s As String)
    Dim s1 As String, s2 As String, sEx As String
    Dim lu As Long, lp As Long
 
    s1 = s
    lp = InStrRev(s, ".", -1, 1)
    If lp Then
        sEx = Mid(s, lp)
        s1 = Mid(s, 1, lp - 1)
    End If
    s2 = s
    lu = 0
    Do While (Dir(s2, 16) <> "")
        lu = lu + 1
        s2 = s1 & "(" & lu & ")" & sEx
    Loop
    GetAtchName = s2
End Function



Private Sub CommandButton209_Click()
   
End Sub

Private Sub CommandButton21_Click()
         Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
            With objMail
                .Display
                .To = "poisk@cc.tricolor.tv"
                .CC = "ChuchalovVY@monobrand-tt.ru; Butko@monobrand-tt.ru"
                .Subject = "Запрос записи " & Date
                .HTMLBody = "<p>Коллеги, добрый день!</p>" _
                & "<p>Предоставьте, пожалуйста, запись разговора с клиентом по номеру заказа:</p>" _
                & "<ul><li></li></ul>" _
                & "<p>Спасибо.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                '.DeferredDeliveryTime = Date + 17 / 24
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
End Sub

Private Sub CommandButton210_Click()
f = Cells(Rows.Count, 1).End(xlUp).Row
   Range("aa1:aa" & f).FormulaR1C1 = _
        "=IF(ISNA(VLOOKUP(C[-24],[TableHSR.xlsx]отправления!C7,1,FALSE)),""0"",VLOOKUP(C[-24],[TableHSR.xlsx]отправления!C7,1,FALSE))"
End Sub

Private Sub CommandButton211_Click()
Application.ScreenUpdating = False
t = GetTickCount


Dim a As Date


n = ActiveCell.Row

    Dim Req As Object
    Dim sEnv As String
    Dim Resp As New MSXML2.DOMDocument60
    Set Req = CreateObject("MSXML2.XMLHTTP")
    Set Resp = CreateObject("MSXML2.DOMDocument.6.0")
    
        Req.Open "Post", "https://tracking.russianpost.ru/rtm34", False
         sEnv = sEnv & "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:oper=""http://russianpost.org/operationhistory"" xmlns:data=""http://russianpost.org/operationhistory/data"" xmlns:ns1=""http://schemas.xmlsoap.org/soap/envelope/"">"
         sEnv = sEnv & "<soap:Header/>"
         sEnv = sEnv & "<soap:Body>"
         sEnv = sEnv & "<oper:getOperationHistory>"
         sEnv = sEnv & "<data:OperationHistoryRequest>"
         sEnv = sEnv & "<data:Barcode>" & Range("h" & n) & "</data:Barcode>"
         sEnv = sEnv & "<data:MessageType>0</data:MessageType>"
         sEnv = sEnv & "<data:Language>RUS</data:Language>"
         sEnv = sEnv & "</data:OperationHistoryRequest>"
         sEnv = sEnv & "<data:AuthorizationHeader ns1:mustUnderstand=""?"">"
         sEnv = sEnv & "<data:login>ykDaLTEChMLavX</data:login>"
         sEnv = sEnv & "<data:password>JPOIsPTd3W03</data:password>"
         sEnv = sEnv & "</data:AuthorizationHeader>"
         sEnv = sEnv & "</oper:getOperationHistory>"
         sEnv = sEnv & "</soap:Body>"
         sEnv = sEnv & "</soap:Envelope>"
        Req.send (sEnv)

Debug.Print (GetTickCount - t) / 1000, vbInformation

        
        Req2 = Replace(Req.responseText, "ns3:OperAttr", "OperAttr")
        
    
        Dim pDoc As New MSXML2.DOMDocument60
        pDoc.LoadXML Req2
        
        Set nodeXML = pDoc.getElementsByTagName("OperAttr")
        
        For i = 1 To nodeXML.Length - 1
            X = nodeXML(i).Text
           
           

        Next i

t = GetTickCount


        Range("q" & n) = X
        
      'clean up code
        Set Req = Nothing
        Set Req2 = Nothing
        Set Resp = Nothing
        Set nodeXML = Nothing
        Set pDoc = Nothing
        sEnv = ""
        
        
        
Debug.Print (GetTickCount - t) / 1000, vbInformation
Application.ScreenUpdating = True
    
End Sub

Private Sub CommandButton22_Click()
    Application.ScreenUpdating = False
    Dim FilesToOpen
    Dim X As Integer
    FilesToOpen = Application.GetOpenFilename _
      (FileFilter:="All files (*.*), *.*", _
      MultiSelect:=True, Title:="Files to Merge")
    If TypeName(FilesToOpen) = "Boolean" Then
        MsgBox "Не выбрано ни одного файла!"
        Exit Sub
    End If
    X = 1
    While X <= UBound(FilesToOpen)
        Set importWB = Workbooks.Open(FileName:=FilesToOpen(X))
        Sheets().Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        importWB.Close savechanges:=False
        X = X + 1
    Wend
       Application.ScreenUpdating = True
End Sub

Private Sub CommandButton23_Click()
    Sheets.Add.Name = "Общий"
    For i = 1 To Sheets.Count
        If Sheets(i).Name <> "Общий" Then
           myR_Total = Sheets("Общий").Range("A" & Sheets("Общий").Rows.Count).End(xlUp).Row
           myR_i = Sheets(i).Range("A" & Sheets(i).Rows.Count).End(xlUp).Row
           Sheets(i).Rows("2:" & myR_i).Copy Destination:=Sheets("Общий").Range("A" & myR_Total + 1)
        End If
    Next
End Sub

Private Sub CommandButton24_Click()
    asn = ActiveSheet.Name
    Sheets.Add.Name = "Итог"
    Range("f:f").NumberFormat = "m/d/yyyy"
    
    Sheets(asn).Range("a:a").Copy
    Range("a:a").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("b:b").Copy
    Range("f:f").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("c:c").Copy
    Range("g:g").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("d:d").Copy
    Range("h:h").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("e:e").Copy
    Range("i:i").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("f:f").Copy
    Range("j:j").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("g:g").Copy
    Range("k:k").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("h:h").Copy
    Range("l:l").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("i:i").Copy
    Range("m:m").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("j:j").Copy
    Range("n:n").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("k:k").Copy
    Range("o:o").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("l:l").Copy
    Range("p:p").PasteSpecial Paste:=xlPasteValues
    
    f = Cells(Rows.Count, 1).End(xlUp).Row
    Range("b1:b" & f) = "Отправление"
    Range("e1:e" & f).FormulaR1C1 = "=WEEKNUM(RC[1],11)"
    
    For i = 1 To f
        If Range("a" & i) = "ТРС Екатеринбург" Then
        Range("a" & i) = "Екатеринбург"
        ElseIf Range("a" & i) = "ТРС Нижний Новгород" Then
        Range("a" & i) = "Нижний Новгород"
        ElseIf Range("a" & i) = "ТРС Ростов-на-Дону" Then
        Range("a" & i) = "Ростов-на-Дону"
        ElseIf Range("a" & i) = "ТРС Тула" Then
        Range("a" & i) = "Тула"
        ElseIf Range("a" & i) = "ТРС Санкт-Петербург" Then
        Range("a" & i) = "Санкт-Петербург"
        ElseIf Range("a" & i) = "ТРС Новосибирск" Then
        Range("a" & i) = "Новосибирск"
        ElseIf Range("a" & i) = "ТРС Саратов" Then
        Range("a" & i) = "Саратов"
        End If
        Range("f" & i).FormulaLocal = Range("f" & i).FormulaLocal
        
        'Range("f" & i) = Range("f" & i) + 1
        
    Next i
    
Range("h:h").NumberFormat = "#"

    
End Sub

Private Sub CommandButton25_Click()
'    Columns("A:M").Select
'    With Selection
'        .HorizontalAlignment = xlGeneral
'        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
'    End With
    Columns(1).ColumnWidth = 20
    Columns(2).ColumnWidth = 14
    Columns(3).ColumnWidth = 14
    Columns(4).ColumnWidth = 14
    Columns(5).ColumnWidth = 40
    Columns(6).ColumnWidth = 8
    Columns(7).ColumnWidth = 14
    Columns(8).ColumnWidth = 20
    Columns(9).ColumnWidth = 20
    Columns(10).ColumnWidth = 14
    Columns(11).ColumnWidth = 20
    Columns(12).ColumnWidth = 20
    Columns(13).ColumnWidth = 20
    
f = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To f
    
        
        
        
        If Range("e" & i) = "КNY_Модуль управления GS SMH-ZW-I1" Then
        Range("e" & i) = "Модуль управления GS SMH-ZW-I1"
        
        ElseIf Range("e" & i) = "КNY_Умная розетка GS SKHMP30-I1" Then
        Range("e" & i) = "Умная розетка GS SKHMP30-I1"
        
        ElseIf Range("e" & i) = "KNY_Умная лампа GS BDHM8E27W70-I1" Then
        Range("e" & i) = "Умная лампа GS BDHM8E27W70-I1"
        
        ElseIf Range("e" & i) = "Домашняя камера SCI-1, новогодняя акция" Then
        Range("e" & i) = "Видеокамера IP домашняя Триколор Умный дом SCI-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD, ИК 10м, WiFi)"
        
        ElseIf Range("e" & i) = "Уличная камера SCO-1, новогодняя акция" Then
        Range("e" & i) = "Видеокамера IP уличная Триколор Умный дом SCO-1 (1/2,7" & Chr(34) & ", 2 Mpix, Full HD 1080p, ИК 30м, IP67, WiFi)"
        
        End If
        
    Next i
    

f = Cells(Rows.Count, 1).End(xlUp).Row



For i = 2 To f

Dim a As Long
Dim b As Long

    Set X = Range("c" & i - 1)
    Set y = Range("c" & i)
    a = RGB(255, 255, 0)
    b = RGB(0, 176, 80)

    Cells(i, 3).Interior.Color = a
    Cells(i, 4).Interior.Color = a
    Cells(i, 5).Interior.Color = a
   

    If X = y Then
        If X.Interior.Color = a Then
        Cells(i, 3).Interior.Color = a
        Cells(i, 4).Interior.Color = a
        Cells(i, 5).Interior.Color = a
        ElseIf X.Interior.Color = b Then
        Cells(i, 3).Interior.Color = b
        Cells(i, 4).Interior.Color = b
        Cells(i, 5).Interior.Color = b
        End If
    Else
        If X.Interior.Color = a Then
        Cells(i, 3).Interior.Color = b
        Cells(i, 4).Interior.Color = b
        Cells(i, 5).Interior.Color = b
        ElseIf X.Interior.Color = b Then
        Cells(i, 3).Interior.Color = a
        Cells(i, 4).Interior.Color = a
        Cells(i, 5).Interior.Color = a
        End If
    End If
Next i

Range("d:d").NumberFormat = "#"



For i = 2 To f

    X = Int((999999999 - 1 + 1) * Rnd + 1)
    X = Time + X - Date
    
    Range("L" & i) = X
    Range("B" & i) = TextBox1.Text
    
Next i


    
End Sub

Private Sub CommandButton26_Click()
    Columns("A:O").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("A:O").ColumnWidth = 20
    Range("a2") = "Заказчик: ООО 'ТОРГОВЫЕ ТЕХНОЛОГИИ', ИНН 7813266200, КПП 781301001, 197101, г. Санкт-Петербург, ул. Большая Монетная, дом №16, корпус 1, лит 5-Н, кв. 411"
    Range("a3") = "Грузоотправитель: Общество с ограниченной ответственностью 'Спутник Трейд'"

    f = Cells(Rows.Count, 7).End(xlUp).Row
    Range("F6:f" & f) = "285x185x40"
    

    
    
End Sub

Private Sub CommandButton27_Click()
X = ActiveWorkbook.Name
    Workbooks.Add
    Workbooks(X).Sheets(1).Copy before:=Sheets(1)
    y = Range("a2")
    Z = TextBox1.Text
    ActiveWorkbook.SaveAs FileName:="C:\Users\ShapkaMY\Desktop\2021\01 Январь\" & Z & "\1.Реестр отправлений\" & Date & " " & y & " (реестр отправлений).xlsx"
    
End Sub

Private Sub CommandButton28_Click()
    Range("a1") = "Наименование"
    Range("b1") = "Стоимость"
    Range("c1") = "Сейчас"
    Range("d1") = "Останется"
    
    Columns(1).ColumnWidth = 20
    Columns(2).ColumnWidth = 20
    Columns(3).ColumnWidth = 20
    Columns(4).ColumnWidth = 20
    
    Range("d2").FormulaR1C1 = "=RC[-1]-SUM(RC[-2]:R[98]C[-2])"
    
    
End Sub

Private Sub CommandButton3_Click()
    Range("A:AA").Copy 'Копируем содержимое листа
    Sheets.Add.Name = "Возврат" 'Создаем лист "Возврат".
    Range("A1").PasteSpecial Paste:=xlPasteValues 'Вставляем как значение
    
    'Удаляем в столбце F все строки, где есть Выгружено Триколор-Бета
    f = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To f
        If Range("F" & i) = "Выгружено Триколор-Бета" Then
        Range("F" & i).Rows.Clear
        End If
    Next i

    Range("F:F").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    f = Cells(Rows.Count, 11).End(xlUp).Row
    Range("aa1:aa" & f).FormulaR1C1 = _
    "=COUNTIFS([TableHSR.xlsx]отправления!C2,""Возврат"",[TableHSR.xlsx]отправления!C1,""HSR МСК"",[TableHSR.xlsx]отправления!C7,C[-24])"

    'Очищаем все ячейки в столбце "AA", где есть символ "0".
    For i = 1 To f
        If Range("AA" & i) = "0" Then
            Range("ab1") = "ok"
        Else
            Range("AA" & i).Rows.Clear
        End If
    Next i

    Range("AA:AA").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    'Добавляем в столбец AB "Возврат"
    f = Cells(Rows.Count, 11).End(xlUp).Row
    For i = 1 To f
        Range("AB" & i) = "Возврат"
    Next i
End Sub

Private Sub CommandButton30_Click()
  Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
            With objMail
                .Display
                .To = "poisk@cc.tricolor.tv"
                .CC = "ChuchalovVY@monobrand-tt.ru"
                .Subject = "Претензия " & Date
                .HTMLBody = "<p>Ольга, добрый день!</p>" _
                & "<p>Отправляю вам скан претензии.</p>" _
                & "<p>Просьба изучить и дать обратную связь.</p>" _
                & "<p>Спасибо.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                '.DeferredDeliveryTime = Date + 17 / 24
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
End Sub

Private Sub CommandButton31_Click()

    X = Sheets.Count
    For i = 1 To X
        Worksheets(i).AutoFilterMode = False
    Next i
    
End Sub

Private Sub CommandButton32_Click()
    asn = ActiveSheet.Name
    Sheets.Add.Name = "Итог"
    
    Sheets(asn).Range("a:a").Copy
    Range("a:a").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("b:b").Copy
    Range("f:f").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("c:c").Copy
    Range("g:g").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("d:d").Copy
    Range("h:h").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("e:e").Copy
    Range("i:i").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("f:f").Copy
    Range("j:j").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("g:g").Copy
    Range("l:l").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("h:h").Copy
    Range("m:m").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("i:i").Copy
    Range("o:o").PasteSpecial Paste:=xlPasteValues
    
    
    f = Cells(Rows.Count, 1).End(xlUp).Row
    Range("b1:b" & f) = "Возврат"
    Range("e1:e" & f).FormulaR1C1 = "=WEEKNUM(RC[1],11)"
    
    For i = 1 To f
        If Range("a" & i) = "ТРС Екатеринбург" Then
        Range("a" & i) = "Екатеринбург"
        ElseIf Range("a" & i) = "ТРС Нижний Новгород" Then
        Range("a" & i) = "Нижний Новгород"
        ElseIf Range("a" & i) = "ТРС Ростов-на-Дону" Then
        Range("a" & i) = "Ростов-на-Дону"
        ElseIf Range("a" & i) = "ТРС Тула" Then
        Range("a" & i) = "Тула"
        ElseIf Range("a" & i) = "ТРС Санкт-Петербург" Then
        Range("a" & i) = "Санкт-Петербург"
        ElseIf Range("a" & i) = "ТРС Новосибирск" Then
        Range("a" & i) = "Новосибирск"
        ElseIf Range("a" & i) = "ТРС Саратов" Then
        Range("a" & i) = "Саратов"
        ElseIf Range("a" & i) = "Склад" Or Range("a" & i) = "Склад*" Then
        Range("a" & i).Rows.Clear
        End If
    Next i
    On Error Resume Next
    Range("A1:A" & f).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    f = Cells(Rows.Count, 1).End(xlUp).Row
     
    Columns("D:D").Select
    Selection.NumberFormat = "General"
    
    
    For i = 1 To f
        Range("k" & i).FormulaArray = _
            "=INDEX([Table.xlsx]отправления!C11,MATCH(RC[-4]&RC[-2],[Table.xlsx]отправления!C7&[Table.xlsx]отправления!C9,0))"
        Range("D" & i).FormulaArray = _
        "=INDEX([Table.xlsx]отправления!C4,MATCH(RC[3]&RC[5],[Table.xlsx]отправления!C7&[Table.xlsx]отправления!C9,0))"
    Next i
    
    f = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 1 To f
    If Range("o" & i) = "" Then
        Range("o" & i) = "б/н"
    End If
    
    If Range("l" & i) = "" Or Range("l" & i) = "упаковано в пакет Pony" Or Range("l" & i) = "Возврат" Then
        Range("l" & i) = "норма"
    End If
    
    Range("m" & i).Rows.Clear
    
    If Range("h" & i) = "" Then
        Range("h" & i).FormulaR1C1 = "=VLOOKUP(RC[-1],[Table.xlsx]отправления!C7:C8,2,0)"
    End If
    

    Next i
    
    
    Dim rArea As Range

    For Each rArea In Range("f1:f" & f).Areas
        rArea.FormulaLocal = rArea.FormulaLocal
    Next
    
End Sub

Private Sub CommandButton33_Click()

'    a = ActiveWorkbook.Name
'    b = ActiveSheet.Name
'
'Workbooks.Open Filename:="C:\Users\ShapkaMY\Desktop\Статистика.csv"
'    Workbooks(a).Sheets(b).Activate
    
       f = Cells(Rows.Count, 1).End(xlUp).Row
       Range("d1:d" & f).FormulaR1C1 = "=VLOOKUP(RC[3],Статистика.csv!C6:C7,2,0)"
'  Windows("Статистика.csv").Close True
End Sub

Private Sub CommandButton34_Click()
    f = Cells(Rows.Count, 1).End(xlUp).Row
    Range("K1:K" & f).FormulaR1C1 = _
        "=VLOOKUP(RC[-2],[Table.xlsx]цены_наименования!C4:C5,2,0)"
End Sub

Private Sub CommandButton35_Click()
    f = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 1 To f
        If IsEmpty(Range("k" & i)) = True Then
            Range("k" & i) = "б/н"
        End If
        
        If IsEmpty(Range("i" & i)) = True Then
            Range("i" & i) = "1"
        End If
        If IsEmpty(Range("j" & i)) = True Then
            Range("j" & i) = Range("b" & i)
        End If
    Next i
    

    
End Sub

Private Sub CommandButton36_Click()
     Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
            With objMail
                .Display
                .To = "BelyaevskiyKO@monobrand-tt.ru"
                .CC = ""
                .Subject = "Отредактированный HSR отчёт от " & Date
                .HTMLBody = "<p>Кирилл Олегович, здравствуйте.</p>" _
                & "<p>Во вложении отредактированный отчёт от " & Date & "</p>" _
                & "<p>В файл 001 занёс.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                
                
                'указывается текст письма
                .Attachments.Add "C:\Users\ShapkaMY\Desktop\backup\HSR отчеты\" & Date & " Hsr отчёт.xlsx" 'указывается полный путь к файлу
                .DeferredDeliveryTime = Date + 12 / 24
                
                
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
End Sub

Private Sub CommandButton37_Click()

        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
        
        Set rng = ActiveCell
            With objMail
                .Display
                .To = "sklad1@rd.e-burg.n-l-e.ru; logist@rd.e-burg.n-l-e.ru; sklad@rd.e-burg.n-l-e.ru"
                .CC = "antipova@n-l-e.ru; ChuchalovVY@monobrand-tt.ru; BelyaevskiyKO@monobrand-tt.ru; BocharovAV@tricolor.tv; Butko@monobrand-tt.ru"
                .Subject = "ОТПРАВКА ИНТЕРНЕТ-МАГАЗИН ООО <ТТ> " & Trsdate & " " & city
                .HTMLBody = rng.Select _
                & "<p><br>" _
                & "Прилагаю:</p>" _
                & "<ul><li>Реестр отправлений</li><li>Накладная для Pony Express</li><li>Товарные чеки</li></ul>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                
                
                
                
                '.Send
            End With
            
            
        Set objMail = Nothing
        Set objOL = Nothing


End Sub

Private Sub CommandButton38_Click()
    f = Cells(Rows.Count, 1).End(xlUp).Row
    Range("b2:b" & f) = "=VLOOKUP(RC[-1],[Table.xlsx]отправления!C7:C8,2,0)"

End Sub

Private Sub CommandButton39_Click()
'    Dim go
'    go = Shell("C:\Users\ShapkaMY\Desktop\Table.xlsx", 1)
'
'
'    Workbooks.Open Filename:="C:\Users\User\Desctop\file.xlsx"
'    Range("A1").Select
'    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
'    Selection.Copy
'
'    Windows("file.xlsx").Close
End Sub

Private Sub CommandButton4_Click()
    Sheets("Отправление").Activate
    f = Cells(Rows.Count, 1).End(xlUp).Row
    Range("aa1:aa" & f).FormulaR1C1 = _
        "=IF(ISNA(VLOOKUP(C[-24],[TableHSR.xlsx]отправления!C7,1,FALSE)),""0"",VLOOKUP(C[-24],[TableHSR.xlsx]отправления!C7,1,FALSE))"
    
    'Очищаем все ячейки в столбце "AA", где есть символ "0".
    For i = 1 To f
        If Range("AA" & i) <> "0" Then
            Range("AA" & i).Rows.Clear
        End If
    Next i
    Range("AA:AA").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    f = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To f
        Range("AB" & i) = "Отправление"
    Next i
End Sub

Private Sub CommandButton40_Click()
'

    Selection.FormulaArray = _
        "=INDEX([main.xlsb]Итог!C15,MATCH(RC[-8]&RC[-6],[main.xlsb]Итог!C7&[main.xlsb]Итог!C9,0))"

End Sub

Private Sub CommandButton41_Click()
 f = Cells(Rows.Count, 1).End(xlUp).Row
 Range("K1:K" & f).FormulaR1C1 = _
       "=VLOOKUP(RC[-2],[Table.xlsx]цены_наименования!C4:C5,2,0)"
End Sub

Private Sub CommandButton42_Click()

    Columns(12).Delete
    Columns(12).Delete
    Rows(1).Delete
End Sub

Private Sub CommandButton43_Click()

    If CheckBox9.Value = True Then
        pochta = "Почта России"
    Else
        pochta = "Pony Express"
    End If

    Trsdate = TextBox1.Text
    
    
    If CheckBox1.Value Then
        city = pochta & " Екатеринбург"
        Workbooks.Open FileName:="C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & city & "\" & Trsdate & " ТРС Екатеринбург (реестр отправлений) " & pochta & ".xlsx"
        Call CommandButton25_Click
        Call CommandButton145_Click
        Windows(Trsdate & " ТРС Екатеринбург (реестр отправлений) " & pochta & ".xlsx").Close True
    End If
    
    If CheckBox2.Value Then
        city = pochta & " Санкт-Петербург"
        Workbooks.Open FileName:="C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & city & "\" & Trsdate & " ТРС Санкт-Петербург (реестр отправлений) " & pochta & ".xlsx"
        Call CommandButton25_Click
        Call CommandButton145_Click
        Windows(Trsdate & " ТРС Санкт-Петербург (реестр отправлений) " & pochta & ".xlsx").Close True
    End If
    
    If CheckBox3.Value Then
        city = pochta & " Нижний Новгород"
        Workbooks.Open FileName:="C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & city & "\" & Trsdate & " ТРС Нижний Новгород (реестр отправлений) " & pochta & ".xlsx"
        Call CommandButton25_Click
        Call CommandButton145_Click
        Windows(Trsdate & " ТРС Нижний Новгород (реестр отправлений) " & pochta & ".xlsx").Close True
    End If
    
    If CheckBox4.Value Then
        city = pochta & " Новосибирск"
        Workbooks.Open FileName:="C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & city & "\" & Trsdate & " ТРС Новосибирск (реестр отправлений) " & pochta & ".xlsx"
        Call CommandButton25_Click
        Call CommandButton145_Click
        Windows(Trsdate & " ТРС Новосибирск (реестр отправлений) " & pochta & ".xlsx").Close True
    End If
    
    If CheckBox5.Value Then
        city = pochta & " Тула"
        Workbooks.Open FileName:="C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & city & "\" & Trsdate & " ТРС Тула (реестр отправлений) " & pochta & ".xlsx"
        Call CommandButton25_Click
        Call CommandButton145_Click
        Windows(Trsdate & " ТРС Тула (реестр отправлений) " & pochta & ".xlsx").Close True
    End If
    
    If CheckBox6.Value Then
        city = pochta & " Ростов-на-Дону"
        Workbooks.Open FileName:="C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & city & "\" & Trsdate & " ТРС Ростов-на-Дону (реестр отправлений) " & pochta & ".xlsx"
        Call CommandButton25_Click
        Call CommandButton145_Click
        Windows(Trsdate & " ТРС Ростов-на-Дону (реестр отправлений) " & pochta & ".xlsx").Close True
    End If
       If CheckBox7.Value Then
        city = pochta & " Саратов"
        Workbooks.Open FileName:="C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & city & "\" & Trsdate & " ТРС Саратов (реестр отправлений) " & pochta & ".xlsx"
        Call CommandButton25_Click
        Call CommandButton145_Click
        Windows(Trsdate & " ТРС Саратов (реестр отправлений) " & pochta & ".xlsx").Close True
    End If
    
    
         

End Sub


Private Sub CommandButton44_Click()
    Trsdate = TextBox1.Text
    
    
    If CheckBox9.Value = True Then
        pochta = "Почта России"
    Else
        pochta = "Pony Express"
    End If
    
    
    If CheckBox1.Value Then
        city = "Екатеринбург"
        cityk = pochta & " " & city
        Workbooks.Open FileName:="C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & cityk & "\" & Trsdate & " ТРС " & city & " для Pony Express.xlsx"
        Call CommandButton26_Click
        Call CommandButton144_Click
        Windows(Trsdate & " ТРС " & city & " для Pony Express.xlsx").Close True
    End If
    
    If CheckBox2.Value Then
        city = "Санкт-Петербург"
        cityk = pochta & " " & city
        Workbooks.Open FileName:="C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & cityk & "\" & Trsdate & " ТРС " & city & " для Pony Express.xlsx"
        Call CommandButton26_Click
        Call CommandButton144_Click
        Windows(Trsdate & " ТРС " & city & " для Pony Express.xlsx").Close True
    End If
    
     If CheckBox3.Value Then
        city = "Нижний Новгород"
        cityk = pochta & " " & city
        Workbooks.Open FileName:="C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & cityk & "\" & Trsdate & " ТРС " & city & " для Pony Express.xlsx"
        Call CommandButton26_Click
        Call CommandButton144_Click
        Windows(Trsdate & " ТРС " & city & " для Pony Express.xlsx").Close True
    End If
    
     If CheckBox4.Value Then
        city = "Новосибирск"
        cityk = pochta & " " & city
        Workbooks.Open FileName:="C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & cityk & "\" & Trsdate & " ТРС " & city & " для Pony Express.xlsx"
        Call CommandButton26_Click
        Call CommandButton144_Click
        Windows(Trsdate & " ТРС " & city & " для Pony Express.xlsx").Close True
    End If
    
     If CheckBox5.Value Then
        city = "Тула"
        cityk = pochta & " " & city
        Workbooks.Open FileName:="C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & cityk & "\" & Trsdate & " ТРС " & city & " для Pony Express.xlsx"
        Call CommandButton26_Click
        Call CommandButton144_Click
        Windows(Trsdate & " ТРС " & city & " для Pony Express.xlsx").Close True
    End If
    
     If CheckBox6.Value Then
        city = "Ростов-на-Дону"
        cityk = pochta & " " & city
        Workbooks.Open FileName:="C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & cityk & "\" & Trsdate & " ТРС " & city & " для Pony Express.xlsx"
        Call CommandButton26_Click
        Call CommandButton144_Click
        Windows(Trsdate & " ТРС " & city & " для Pony Express.xlsx").Close True
    End If
    
     If CheckBox7.Value Then
        city = "Саратов"
        cityk = pochta & " " & city
        Workbooks.Open FileName:="C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & cityk & "\" & Trsdate & " ТРС " & city & " для Pony Express.xlsx"
        Call CommandButton26_Click
        Call CommandButton144_Click
        Windows(Trsdate & " ТРС " & city & " для Pony Express.xlsx").Close True
    End If
    
    
End Sub

Private Sub CommandButton45_Click()

    f = Cells(Rows.Count, 1).End(xlUp).Row
     For i = 1 To f
      If Range("g" & i) = 0 Then
         If Dir("C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & city & "\Чеки\" & Range("c" & i) & ".pdf") = Range("c" & i) & ".pdf" Then
         Kill "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & city & "\Чеки\" & Range("c" & i) & ".pdf"
         End If
    End If
   Next i
End Sub

Private Sub CommandButton52_Click()
 ActiveCell.FormulaR1C1 = _
        "=INDEX([Table.xlsx]отправления!C7,MATCH(RC[1],[Table.xlsx]отправления!C8,0))"
 ActiveCell.Copy
 ActiveCell.PasteSpecial Paste:=xlPasteValues
End Sub

Private Sub CommandButton53_Click()
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],[Table.xlsx]отправления!C7:C8,2,0)"
ActiveCell.Copy
ActiveCell.PasteSpecial Paste:=xlPasteValues
End Sub

Private Sub CommandButton54_Click()
    Range("c:c").Copy
    Range("c:c").PasteSpecial Paste:=xlPasteValues
End Sub

Private Sub CommandButton55_Click()
    X = ActiveCell
    Workbooks.Add
    
    Range("a1") = "Номер договора"
    Range("a2") = "Номер заказ"
    Range("a3") = "Номер накладной"
    Range("a4") = "Дата заявки"
    Range("a5") = "Дата отгрузки"
    Range("a6") = "Дата доставки"
    Range("a7") = "Наименование заказа"
    Range("a8") = "Стоимость заказ"
    Range("a9") = "Стоимость доставки"
    Range("a10") = "ФИО клиента"
    Range("a11") = "Адрес клиента"
    Range("a12") = "В комплект заказ входит"
    
    Range("b1") = "22-50242 от 19.07.2017."
    Range("b2").FormulaR1C1 = _
        "=INDEX([Table1.xlsx]отправления!C7,MATCH(R[1]C,[Table1.xlsx]отправления!C8,0))"
    Range("b3") = X
    Range("b5").FormulaR1C1 = _
        "=INDEX([Table1.xlsx]отправления!C6,MATCH(R[-1]C,[Table1.xlsx]отправления!C8,0))"
    Range("b6") = "Заказ доставлен клиенту"
    Range("b7").FormulaR1C1 = "=VLOOKUP(R[-4]C,'Статистика 2020.csv'!C6:C14,9,0)"
    Range("b8").FormulaR1C1 = "=VLOOKUP(R[-5]C,'Статистика 2020.csv'!C6:C16,11,0)"
    Range("b9") = "Стоимость доставки"
    Range("b10").FormulaR1C1 = "=VLOOKUP(R[-7]C,'Статистика 2020.csv'!C6:C10,5,0)"
    Range("b11").FormulaR1C1 = "=VLOOKUP(R[-8]C,'Статистика 2020.csv'!C6:C23,18,0)"
    
    Columns("A:A").ColumnWidth = 26
    Columns("B:B").ColumnWidth = 26
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    
End Sub

Private Sub CommandButton56_Click()
f = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To f
        If Range("a" & i) = "Н.Новгород ТРС. Волго-Вятский СТ" Then
        Range("a" & i) = "Нижний Новгород"
        ElseIf Range("a" & i) = "Екатеринбург ТРС.Урал СТ" Then
        Range("a" & i) = "Екатеринбург"
        ElseIf Range("a" & i) = "Ростов.ТРС.Юг СТ" Then
        Range("a" & i) = "Ростов-на-Дону"
        ElseIf Range("a" & i) = "Саратов ТРС. Поволжье СТ" Then
        Range("a" & i) = "Саратов"
        ElseIf Range("a" & i) = "СПб.Транзит СТ" Then
        Range("a" & i) = "Санкт-Петербург"
        ElseIf Range("a" & i) = "Тула.ТС СТ" Then
        Range("a" & i) = "Тула"
        ElseIf Range("a" & i) = "Новосибирск ТРС СТ" Then
        Range("a" & i) = "Новосибирск"
        End If
        
    
    Next i
    
End Sub

Private Sub CommandButton57_Click()

f = Cells(Rows.Count, 1).End(xlUp).Row



For i = 2 To f

Dim a As Long
Dim b As Long

    Set X = Range("g" & i - 1)
    Set y = Range("g" & i)
    a = RGB(255, 255, 0)
    b = RGB(0, 176, 80)

    Cells(i, 3).Interior.Color = a
    Cells(i, 8).Interior.Color = a
    Cells(i, 5).Interior.Color = a
   

    If X = y Then
        If X.Interior.Color = a Then
        Cells(i, 3).Interior.Color = a
        Cells(i, 4).Interior.Color = a
        Cells(i, 5).Interior.Color = a
        ElseIf X.Interior.Color = b Then
        Cells(i, 3).Interior.Color = b
        Cells(i, 4).Interior.Color = b
        Cells(i, 5).Interior.Color = b
        End If
    Else
        If X.Interior.Color = a Then
        Cells(i, 3).Interior.Color = b
        Cells(i, 4).Interior.Color = b
        Cells(i, 5).Interior.Color = b
        ElseIf X.Interior.Color = b Then
        Cells(i, 3).Interior.Color = a
        Cells(i, 4).Interior.Color = a
        Cells(i, 5).Interior.Color = a
        End If
    End If

    


Next i








End Sub

Private Sub CommandButton58_Click()

    f = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To f
        Range("k" & i).FormulaArray = _
            "=INDEX([Table.xlsx]отправления!C11,MATCH(RC[-4]&RC[-2],[Table.xlsx]отправления!C7&[Table.xlsx]отправления!C9,0))"
    Next i
End Sub

Private Sub CommandButton59_Click()
 Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
            With objMail
                .Display
                .To = "oa.pichmanova@ponyexpress.ru; ii.bayramgulova@ponyexpress.ru"
                .CC = "ChuchalovVY@monobrand-tt.ru"
                .Subject = ActiveCell
                .HTMLBody = "<p>Возвращаем.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                '.DeferredDeliveryTime = Date + 17 / 24
                .send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
        
        ActiveCell.Interior.Color = RGB(146, 208, 80)
End Sub

Private Sub CommandButton6_Click()
    Application.DisplayAlerts = False
    For i = Sheets.Count To 1 Step -1
        If Sheets(i).Name <> "Отправление" Then
            If Sheets(i).Name <> "Возврат" Then
                Sheets(i).Delete
            End If
         End If
    Next
    Application.DisplayAlerts = True
    
    Sheets.Add.Name = "Общий"
    For i = 1 To Sheets.Count
        If Sheets(i).Name <> "Общий" Then
           myR_Total = Sheets("Общий").Range("A" & Sheets("Общий").Rows.Count).End(xlUp).Row
           myR_i = Sheets(i).Range("A" & Sheets(i).Rows.Count).End(xlUp).Row
           Sheets(i).Rows("1:" & myR_i).Copy Destination:=Sheets("Общий").Range("A" & myR_Total + 1)
        End If
    Next
    
    asn = ActiveSheet.Name
    Sheets.Add.Name = "Итог"
    
    Sheets(asn).Range("ab:ab").Copy
    Range("b:b").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("a:a").Copy
    Range("d:d").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("k:k").Copy
    Range("f:f").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("c:c").Copy
    Range("g:g").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("i:i").Copy
    Range("h:h").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("v:v").Copy
    Range("i:i").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("x:x").Copy
    Range("j:j").PasteSpecial Paste:=xlPasteValues
    
    Sheets(asn).Range("w:w").Copy
    Range("k:k").PasteSpecial Paste:=xlPasteValues
     
    Range("d:d").NumberFormat = "dd/mm/yy"
    Range("f:f").NumberFormat = "dd/mm/yy"
    
    f = Cells(Rows.Count, 11).End(xlUp).Row
    For i = 2 To f
        If Range("b" & i) = "Возврат" Then
        Range("a" & i & ":K" & i).Interior.Color = RGB(192, 192, 192)
        End If
    Next i
    Range("a2:a" & f) = "HSR МСК"
    
    Range("e1:e" & f).FormulaR1C1 = "=WEEKNUM(RC[1],11)"
    
    
End Sub

Private Sub CommandButton60_Click()
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
            With objMail
                .Display
                .To = "oa.pichmanova@ponyexpress.ru; ii.bayramgulova@ponyexpress.ru"
                .CC = "ChuchalovVY@monobrand-tt.ru"
                .Subject = ActiveCell
                .HTMLBody = "<p>Верный номер телефона - " & Cells(Selection.Row(), Selection.Column() + 1) & " </p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                '.DeferredDeliveryTime = Date + 17 / 24
                .send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
        
        
        ActiveCell.Interior.Color = RGB(146, 208, 80)
End Sub

Private Sub CommandButton61_Click()
    Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
            With objMail
                .Display
                .To = "oa.pichmanova@ponyexpress.ru; ii.bayramgulova@ponyexpress.ru"
                .CC = "ChuchalovVY@monobrand-tt.ru"
                .Subject = ActiveCell
                .HTMLBody = "<p>Ольга, добрый день.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                '.DeferredDeliveryTime = Date + 17 / 24
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
End Sub

Private Sub CommandButton62_Click()
     Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
            With objMail
                .Display
                .To = "skuznetsova@cc.tricolor.tv; trifonova@cc.tricolor.tv; dubkova@cc.tricolor.tv"
                .CC = "ChuchalovVY@monobrand-tt.ru; simkina@cc.tricolor.tv; mihajlov@cc.tricolor.tv"
                .Subject = "Неверный номер телефона"
                .HTMLBody = "<p>Коллеги, добрый день!</p>" _
                & "<p>Нужно актуализировать данные по неверным номерам телефонов.<br>" _
                & "Список ниже:</p>" _
                & "<p>Спасибо.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                '.DeferredDeliveryTime = Date + 17 / 24
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
End Sub

Private Sub CommandButton63_Click()
        Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
            With objMail
                .Display
                .To = "oa.pichmanova@ponyexpress.ru; ii.bayramgulova@ponyexpress.ru"
                .CC = "ChuchalovVY@monobrand-tt.ru"
                .Subject = ActiveCell
                .HTMLBody = "<p>Подскажите, когда планируется доставка?</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                '.DeferredDeliveryTime = Date + 17 / 24
                .send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
        
        
        ActiveCell.Interior.Color = RGB(146, 208, 80)
End Sub

Private Sub CommandButton64_Click()




For trs = 2 To 30
    X = 0
    
    For i = 1 To 30
        If Sheets("Отправления").Range("a" & i) = Range("a" & trs) Then
            If Sheets("Отправления").Range("d" & i) = Range("b" & trs) Then
                X = X + Sheets("Отправления").Range("e" & i)
            End If
        End If
    Next i
    
    Range("c" & trs) = X
Next trs
 
 
End Sub

Private Sub CommandButton65_Click()
    f = Cells(Rows.Count, 11).End(xlUp).Row
    For i = 1 To f
        If Range("A" & i) = "Ростов-на-Дону" Or _
            Range("A" & i) = "Тула" Or _
            Range("A" & i) = "Санкт-Петербург" Or _
            Range("A" & i) = "Новосибирск" Or _
            Range("A" & i) = "Саратов" Or _
            Range("A" & i) = "Нижний Новгород" Then
            
            Range("p" & i) = "OK"
            
        Else: Range("p" & i) = "Error"
        'Rows(i).Interior.color = RGB(255, 0, 0)
        End If
        
    Next i



End Sub

Private Sub CommandButton66_Click()
    Dim olookApp As Outlook.Application

    Set olookApp = CreateObject("Outlook.Application")
    
    olookApp.q ' Недокументированный метод.
    
    olookApp.Quit
    Set olookApp = Nothing
End Sub

Private Sub CommandButton67_Click()
Dim objOutlook As Object, objNamespace As Object
Dim objFolder As Object, objMail As Object
Dim iRow&, iCount&, IdMail$

iRow = Cells(Rows.Count, "A").End(xlUp).Row
iCount = Application.Max(Range("A:A"))

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(6) '.Folders("Human Resource Management") '6=olFolderInbox

Application.ScreenUpdating = False

'On Error Resume Next
For Each objMail In objFolder.Items
IdMail = objMail.EntryID
If objMail.SenderName = "Пичманова Ольга Александровна" Then
    If objMail.CreationTime > "21.02.2021" Then
        If Application.CountIf(Range("G:G"), IdMail) = 0 Then
            iRow = iRow + 1: iCount = iCount + 1
            Cells(iRow, 1) = iCount
            Cells(iRow, 2) = objMail.SenderName
            Cells(iRow, 3) = objMail.CreationTime
            'Cells(iRow, 3) = objMail.SenderEmailAddress
            Cells(iRow, 4) = objMail.Subject
            'Cells(iRow, 5) = objMail.CreationTime
            Cells(iRow, 5) = Left(objMail.body, 100)
            'Cells(iRow, 7) = IdMail '"'" & IdMail
            'MsgBox (objMail.CreationTime)
            
        End If
    End If
End If
Next

objOutlook.Quit

Application.ScreenUpdating = True
End Sub

Private Sub CommandButton68_Click()
Application.ScreenUpdating = False
Call CommandButton31_Click
Call CommandButton1_Click
Call CommandButton2_Click
Call CommandButton3_Click
Call CommandButton4_Click
Call CommandButton6_Click
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton69_Click()
Application.ScreenUpdating = False
Call CommandButton42_Click
Call CommandButton8_Click
Call CommandButton36_Click

Application.ScreenUpdating = True
End Sub

Private Sub CommandButton7_Click()
    f = Cells(Rows.Count, 11).End(xlUp).Row
    For i = 1 To f
        X = Range("i" & i)
        y1 = Workbooks("TableHSR").Sheets("HSR24").Range("C3")
        y2 = Workbooks("TableHSR").Sheets("HSR24").Range("C4")
        y3 = Workbooks("TableHSR").Sheets("HSR24").Range("C7")
        y4 = Workbooks("TableHSR").Sheets("HSR24").Range("C10")
        y5 = Workbooks("TableHSR").Sheets("HSR24").Range("C11")
        y6 = Workbooks("TableHSR").Sheets("HSR24").Range("C12")
        y7 = Workbooks("TableHSR").Sheets("HSR24").Range("C13")
        y8 = Workbooks("TableHSR").Sheets("HSR24").Range("C14")
        y9 = Workbooks("TableHSR").Sheets("HSR24").Range("C15")
        y11 = Workbooks("TableHSR").Sheets("HSR24").Range("C16")
        y12 = Workbooks("TableHSR").Sheets("HSR24").Range("C19")
        y13 = Workbooks("TableHSR").Sheets("HSR24").Range("C20")
        y14 = Workbooks("TableHSR").Sheets("HSR24").Range("C21")
        y15 = Workbooks("TableHSR").Sheets("HSR24").Range("C22")

        If X = y Or X = y1 Or X = y2 Or X = y3 Or X = y4 Or X = y5 Or X = y6 Or X = y7 Or X = y8 Or X = y9 Or X = y15 Or X = y11 Or X = y12 Or X = y13 Or X = y14 Then
        Range("l" & i) = "ok"
        Else
        Range("l" & i) = "error"
        End If
        
        Range("m" & i).FormulaR1C1 = "=RC[-2]/RC[-3]"
        Range("m" & i).Value = Range("m" & i).Value
        
        Range("m" & i).Copy
        Range("k" & i).PasteSpecial Paste:=xlPasteValues
    Next i
    
End Sub

Private Sub CommandButton70_Click()

    dp = TextBox9.Text

    Dim objOutlook As Object, objNamespace As Object
    Dim objFolder As Object, objMail As Object
    Dim iRow&, iCount&, IdMail$
    Dim X As Date
    
    iRow = Cells(Rows.Count, "A").End(xlUp).Row
    iCount = Application.Max(Range("A:A"))
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(6).Folders("КС") '6=olFolderInbox
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    For Each objMail In objFolder.Items
    IdMail = objMail.EntryID
'    MsgBox (objMail.SenderName)
'    MsgBox (objMail.ReceivedTime)


    X = TextBox9.Text

    If objMail.SenderName = "Пичманова Ольга Александровна" Or objMail.SenderName = "Байрамгулова Ирина Игоревна" Or objMail.SenderName = "Старостина Ксения Александровна" Then
        If objMail.ReceivedTime > X Then
            If Application.CountIf(Range("G:G"), IdMail) = 0 Then
                iRow = iRow + 1: iCount = iCount + 1
                Cells(iRow, 1) = iCount
                Cells(iRow, 2) = objMail.SenderName
                Cells(iRow, 3) = objMail.ReceivedTime
                'Cells(iRow, 3) = objMail.SenderEmailAddress
                Cells(iRow, 4) = objMail.Subject
                'Cells(iRow, 6) = objMail.CreationTime
                Cells(iRow, 5) = Left(objMail.body, 200)
                'Cells(iRow, 7) = IdMail '"'" & IdMail
                'MsgBox (objMail.CreationTime)
                
            End If
        End If
    End If
    Next
    
    objOutlook.Quit
    
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton71_Click()
Columns("A:M").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Private Sub CommandButton72_Click()

    Columns("A:M").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns(1).ColumnWidth = 6
    Columns(2).ColumnWidth = 18
    Columns(3).ColumnWidth = 18
    Columns(4).ColumnWidth = 18
    Columns(5).ColumnWidth = 40
End Sub

Private Sub CommandButton73_Click()

    f = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To f
        If Left(Range("d" & i), 3) = "RE:" Or Left(Range("d" & i), 3) = "FW:" Or Left(Range("d" & i), 9) = "Automatic" Then
        Range("d" & i).Rows.Clear
        End If
    Next i
    

    
    Range("d1:d" & f).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
'    For i = 1 To f
'        If Left(Range("d" & i), 1) = " " Then
'            Range("d" & i).Rows.Clear
'            Right(Range("d" & i),Len(str)-5)
'        End If
'    Next i
    

End Sub

Private Sub CommandButton74_Click()
q = ActiveCell.Row
Rows(q).Delete
End Sub

Private Sub CommandButton75_Click()
ActiveCell = "Клиент не отвечает по телефону. Уточнить актуальность."
ActiveCell.Offset(1).Select

End Sub

Private Sub CommandButton76_Click()
    X = ActiveWorkbook.Name
    y = "Итог"
    
    Workbooks.Add
    
    Workbooks(X).Sheets(1).Copy before:=Sheets(1)
    Columns(1).ColumnWidth = 6
    Columns(2).ColumnWidth = 14
    Columns(3).ColumnWidth = 14
    Columns(4).ColumnWidth = 26
    Columns(5).ColumnWidth = 26
    Columns(6).ColumnWidth = 26
    Columns(7).ColumnWidth = 26
    
    
    Range("d:d").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    f = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To f
        If Range("d" & i) = "Неверный номер" Or Range("d" & i) = "Номер клиента в сети не зарегестрирован." Then
        Range("d" & i).Rows.Clear
        End If
    Next i
    On Error Resume Next
    Range("d1:d" & f).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    f = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To f
    Range("a" & i) = i - 1
    Next i
    
    Rows(1).Insert
    Range("a1") = Date
    
    ActiveWorkbook.SaveAs FileName:="C:\Users\ShapkaMY\Desktop\Прозвон\" & Date & " прозвон.xlsx"
End Sub

Private Sub CommandButton77_Click()
    f = Cells(Rows.Count, 3).End(xlUp).Row
    asn = ActiveSheet.Name
    Sheets.Add
    
    
    Range("a1") = "№"
    Range("b1") = "Номер заказа"
    Range("c1") = "Номер накладной Pony Express"
    Range("d1") = "Комментарий Pony Express"
    Range("e1") = "Попытка 1"
    Range("f1") = "Попытка 2"
    Range("g1") = "Попытка 3"
    
    
    Sheets(asn).Range("d1:d" & f).Copy
    Range("c2:c" & f + 1).PasteSpecial Paste:=xlPasteValues
    Sheets(asn).Range("e1:e" & f).Copy
    Range("d2:d" & f + 1).PasteSpecial Paste:=xlPasteValues


    f = Cells(Rows.Count, 3).End(xlUp).Row
    For i = 2 To f
        Range("a" & i) = i - 1
        Range("b" & i).FormulaR1C1 = _
        "=INDEX([Table.xlsx]отправления!C7,MATCH(RC[1],[Table.xlsx]отправления!C8,0))"
        Range("b" & i).Copy
        Range("b" & i).PasteSpecial Paste:=xlPasteValues
    Next i
    
    
    
    


'    f = Cells(Rows.Count, 1).End(xlUp).Row
'
'    asne = ActiveSheet.Name
'    Sheets.Add.Name = Date & " прозвон"
'
''    Range("a1") = "№"
''    Range("b1") = "Номер заказа"
''    Range("c1") = "Номер накладной Pony Express"
''    Range("d1") = "Комментарий Pony Express"
''    Range("e1") = "Попытка 1"
''    Range("f1") = "Попытка 2"
''    Range("g1") = "Попытка 3"
'
'
'    Sheets(asne).Range("d1:d" & f).Copy
'    Range("c2:c" & f + 1).PasteSpecial Paste:=xlPasteValue
'
'    Sheets(asne).Range("e1:e" & f).Copy
'    Range("d2:d" & f + 1).PasteSpecial Paste:=xlPasteValue
'
'    f = Cells(Rows.Count, 1).End(xlUp).Row
'
'    For i = 2 To f
'        Range("a" & i) = i
'    Next i
'
'
'    Range("d:d").FormulaR1C1 = _
'        "=INDEX([Table.xlsx]отправления!C7,MATCH(RC[1],[Table.xlsx]отправления!C8,0))"

'
End Sub

Private Sub CommandButton78_Click()
Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
            With objMail
                .Display
                .To = "mihajlov@cc.tricolor.tv; moysya@cc.tricolor.tv; simkina@cc.tricolor.tv"
                .CC = "ChuchalovVY@monobrand-tt.ru;"
                .Subject = "Прозвон от " & Date
                .HTMLBody = "<p>Коллеги, добрый день!</p>" _
                & "<p>Прошу актуализировать данные, на запросы от КС до <b>" & Date + 1 & " 18:00</b><br>" _
                & "<p>По факту отправки, прошу предоставить обратную связь.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"

                
                .Attachments.Add "C:\Users\ShapkaMY\Desktop\Прозвон\" & Date & " прозвон.xlsx" 'указывается полный путь к файлу
                '.DeferredDeliveryTime = Date + 17 / 24
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
        
End Sub

Private Sub CommandButton79_Click()
ActiveCell = "Неверный номер"
ActiveCell.Offset(1).Select
End Sub

Private Sub CommandButton8_Click()
    X = ActiveWorkbook.Name
    y = "Итог"
    
    Workbooks.Add
    
    Workbooks(X).Sheets(y).Copy before:=Sheets(1)
    ActiveWorkbook.SaveAs FileName:="C:\Users\ShapkaMY\Desktop\backup\HSR отчеты\" & Date & " Hsr отчёт.xlsx"
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Frame3_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label12_Click()

End Sub

Private Sub Label13_Click()
    
End Sub

Private Sub Label14_Click()

End Sub

Private Sub Label17_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label22_Click()

End Sub

Private Sub Label23_Click()

End Sub

Private Sub Label24_Click()

End Sub

Private Sub Label26_Click()

End Sub

Private Sub Label27_Click()

End Sub

Private Sub Label32_Click()

End Sub

Private Sub Label34_Click()

End Sub

Private Sub Label8_Click()
    X = ActiveWorkbook.Name
    y = "Итог"
    
    Workbooks.Add
    'ActiveWorkbook.SaveAs Filename:="C:\Users\ShapkaMY\Desktop\" & Date & " Hsr отчёт.xlsx"
    Workbooks(X).Sheets(y).Copy before:=Sheets(1)
    ActiveWorkbook.SaveAs FileName:="C:\Users\ShapkaMY\Desktop\" & Date & " Hsr отчёт.xlsx"
End Sub

Private Sub ListBox1_Click()
    With UserForm1.ListBox1
        .AddItem "???????? 1"
        .AddItem "???????? 2"
        .AddItem "???????? 3"
    End With
End Sub

Private Sub CommandButton80_Click()
    X = ActiveWorkbook.Name
    y = "Итог"
    
    Workbooks.Add
    
    
    
    Workbooks(X).Sheets(1).Copy before:=Sheets(1)
    Columns(1).ColumnWidth = 6
    Columns(2).ColumnWidth = 14
    Columns(3).ColumnWidth = 14
    Columns(4).ColumnWidth = 26
    Columns(5).ColumnWidth = 26
    Columns(6).ColumnWidth = 26
    Columns(7).ColumnWidth = 26
    
    
    Range("d:d").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    f = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To f
        If Range("d" & i) = "Клиент не отвечает по телефону. Уточнить актуальность." Then
        Range("d" & i).Rows.Clear
        End If
    Next i
    Range("d1:d" & f).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    f = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To f
    Range("a" & i) = i - 1
    Next i
    
    Range("e1") = "Верный номер телефона"
    Range("f1").Rows.Clear
    Range("g1").Rows.Clear
    
    Rows(1).Insert
    Range("a1") = Date
    
    
    ActiveWorkbook.SaveAs FileName:="C:\Users\ShapkaMY\Desktop\Прозвон\" & Date & " актуализация номеров.xlsx"
    
End Sub

Private Sub CommandButton81_Click()
    Set objOL = CreateObject("Outlook.Application")
        Set objMail = objOL.CreateItem(olMailItem)
            With objMail
                .Display
                .To = "skuznetsova@cc.tricolor.tv; trifonova@cc.tricolor.tv; dubkova@cc.tricolor.tv"
                .CC = "ChuchalovVY@monobrand-tt.ru; simkina@cc.tricolor.tv; mihajlov@cc.tricolor.tv"
                .Subject = "" & Date & " Неверный номер телефона"
                .HTMLBody = "<p>Коллеги, добрый день!</p>" _
                & "<p>Нужно актуализировать данные по неверным номерам телефонов.<br>" _
                & "Список во вложении</p>" _
                & "<p>Спасибо.</p>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>С уважением,</span><br>" _
                & "<b><span style='font-size:9.0pt; font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Шапка Михаил</span></b><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Администратор интернет-магазина</span><br>" _
                & "<img src='C:\Users\ShapkaMY\AppData\Roaming\Microsoft\Signatures\Шапка.files\image001.png'" & "width=width height=heigth><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>ООО «Торговые технологии»</span><br>" _
                & "<span style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Компания по развитию монобрендовой сети Триколор</span><br>" _
                & "<p style='font-size:9.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>+7 (812) 219 68 68 (4003)</p>" _
                & "<p><span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>Содержание данного сообщения и вложений к нему является конфиденциальной</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>информацией. Оно предназначается только Вам и не подлежит передаче третьим</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>лицам ни в исходном, ни в измененном виде. Если данное сообщение попало к Вам</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>случайно, просим сообщить отправителю об ошибке и уничтожить данное</span><br>" _
                & "<span style='font-size:7.0pt;font-family:&quot;Segoe UI&quot;,&quot;sans-serif&quot;;color:#004DA0'>сообщение из своей почтовой системы</span><br></p>"
                
                
                .Attachments.Add "C:\Users\ShapkaMY\Desktop\Прозвон\" & Date & " актуализация номеров.xlsx" 'указывается полный путь к файлу
                
                
                '.DeferredDeliveryTime = Date + 17 / 24
                '.Send
            End With
        Set objMail = Nothing
        Set objOL = Nothing
End Sub

Private Sub CommandButton82_Click()


    Dim objOutlook As Object, objNamespace As Object
    Dim objFolder As Object, objMail As Object
    Dim iRow&, iCount&, IdMail$
    
    iRow = Cells(Rows.Count, "A").End(xlUp).Row
    iCount = Application.Max(Range("A:A"))
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(6).Folders("Pony Express") '6=olFolderInbox
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    
    
    For Each objMail In objFolder.Items
        IdMail = objMail.EntryID
    
        f = Cells(Rows.Count, 3).End(xlUp).Row
        
        For i = 1 To f
            If Range("c" & i).Interior.Pattern = xlNone Then
            
                If objMail.Subject = "RE: " & Range("c" & i) Or objMail.Subject = Range("c" & i) Then
                    Range("c" & i).Interior.Color = RGB(255, 255, 0)
                End If
        
            End If
        
        Next i

    
    Next
    
objOutlook.Quit
    
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton83_Click()
    
    Dim objOutlook As Object, objNamespace As Object
    Dim objFolder As Object, objMail As Object
    Dim iRow&, iCount&, IdMail$
    
    iRow = Cells(Rows.Count, "A").End(xlUp).Row
    iCount = Application.Max(Range("A:A"))
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(5).Folders("outbox") '6=olFolderInbox
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    For Each objMail In objFolder.Items
    IdMail = objMail.EntryID
'    MsgBox (objMail.SenderName)
'    MsgBox (objMail.ReceivedTime)
'    MsgBox (objMail.Subject)
'    MsgBox ("RE: " & Range("c1"))

f = Cells(Rows.Count, 3).End(xlUp).Row
    
For i = 1 To f
    If objMail.Subject = "RE: " & Range("c" & i) Or objMail.Subject = Range("c" & i) Then
    Range("c" & i).Interior.Color = RGB(255, 255, 0)

    End If

Next i

    
    
    
'    If objMail.SenderName = "Пичманова Ольга Александровна" Or objMail.SenderName = "Байрамгулова Ирина Игоревна" Then
'        If objMail.ReceivedTime > "01.03.2021" Then
'            If Application.CountIf(Range("G:G"), IdMail) = 0 Then
'                iRow = iRow + 1: iCount = iCount + 1
'                Cells(iRow, 1) = iCount
'                Cells(iRow, 2) = objMail.SenderName
'                Cells(iRow, 3) = objMail.ReceivedTime
'                'Cells(iRow, 3) = objMail.SenderEmailAddress
'                Cells(iRow, 4) = objMail.Subject
'                'Cells(iRow, 5) = objMail.CreationTime
'                Cells(iRow, 5) = Left(objMail.Body, 100)
'                'Cells(iRow, 7) = IdMail '"'" & IdMail
'                'MsgBox (objMail.CreationTime)
'
'            End If
'        End If
'    End If
    Next
    
    objOutlook.Quit
    
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton84_Click()
    f = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To f
        Range("k" & i).FormulaArray = _
            "=INDEX([Table.xlsx]отправления!C11,MATCH(RC[-4]&RC[-2],[Table.xlsx]отправления!C7&[Table.xlsx]отправления!C9,0))"
    Next i

End Sub

Private Sub CommandButton85_Click()
    Call Shell("C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\Реестры v2.bat")

End Sub

Private Sub CommandButton86_Click()
    Call Shell("C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\Склеиваем и архивируем.bat")
End Sub

Private Sub CommandButton87_Click()

f = Cells(Rows.Count, 1).End(xlUp).Row

For i = 1 To f
    If Range("o" & i) = "" Then
        Range("o" & i) = "б/н"
    End If
    
    If Range("l" & i) = "" Or Range("l" & i) = "упаковано в пакет Pony" Then
        Range("l" & i) = "норма"
    End If
    
    Range("m" & i).Rows.Clear
Next i


End Sub

Private Sub CommandButton88_Click()


For ypak = 2 To 5

    For i = 1 To 30
        If Range("b" & i) = Sheets("Наименования").Range("a" & ypak) Then
            Range("d" & i) = Range("c" & i) * Sheets("Наименования").Range("b" & ypak)
        End If
    Next i
    
Next ypak


End Sub

Private Sub CommandButton89_Click()
    Dim d As Date
    
    d = "01.04.21"
    
    For pochta = 1 To 30
        X = 0
        
        For i = 1 To 25000
        If Workbooks("Table.xlsx").Sheets("отправления").Range("f" & i) > d Then
            If Workbooks("Table.xlsx").Sheets("отправления").Range("v" & i) = "Почта" Then
            
                If Workbooks("Table.xlsx").Sheets("отправления").Range("a" & i) = Range("a" & pochta) Then
                    X = X + 1
                End If
            End If
        End If
        
        Next i
        
        
    If Range("b" & pochta) = Sheets("Наименования").Range("a2") Then
        Range("e" & pochta) = X
    End If
    
    
    Next pochta


End Sub

Private Sub CommandButton90_Click()
    For i = 2 To 30
        Range("f" & i) = Range("d" & i) - Range("e" & i)
    Next i
    
End Sub

Private Sub CommandButton91_Click()

X = "112,0*80,0*19,0"
If X > "112,0*80,0*18,0" Then
    MsgBox ("ok")
End If

End Sub

Private Sub CommandButton92_Click()

f = Cells(Rows.Count, 3).End(xlUp).Row
For i = 1 To f
    If IsEmpty(Range("b" & i)) = True Then
         Range("b" & i).FormulaR1C1 = _
        "=INDEX([Table.xlsx]отправления!C7,MATCH(RC[1],[Table.xlsx]отправления!C8,0))"
        Range("b" & i).Copy
        Range("b" & i).PasteSpecial Paste:=xlPasteValues
    End If
Next i


End Sub

Private Sub CommandButton93_Click()
f = Cells(Rows.Count, 2).End(xlUp).Row
For i = 1 To f
    If IsEmpty(Range("c" & i)) = True Then
        Range("c" & i).FormulaR1C1 = "=VLOOKUP(RC[-1],[Table.xlsx]отправления!C7:C8,2,0)"
        Range("c" & i).Copy
        Range("c" & i).PasteSpecial Paste:=xlPasteValues
         
    End If
Next i

Range("c:c").NumberFormat = "#"

End Sub

Private Sub CommandButton94_Click()
    
    
    f = Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    For i = 1 To f
        Cells.Find(What:="не отвечает", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    ActiveCell = ("Клиент не отвечает по телефону. Уточнить актуальность.")
    
    Next i
    
    
    
    

    f = Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    For i = 1 To f
        Cells.Find(What:="недоступен", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate

    ActiveCell = ("Клиент не отвечает по телефону. Уточнить актуальность.")

    Next i

    
    
    
    f = Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    For i = 1 To f
        Cells.Find(What:="сбрасывает", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    ActiveCell = ("Клиент не отвечает по телефону. Уточнить актуальность.")
    
    Next i
    
   
    
    f = Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    For i = 1 To f
        Cells.Find(What:="актуальность", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    ActiveCell = ("Клиент не отвечает по телефону. Уточнить актуальность.")
    
    Next i
    
    
    
    
    f = Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    For i = 1 To f
        Cells.Find(What:="заблокирован", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    ActiveCell = ("Клиент не отвечает по телефону. Уточнить актуальность.")
    
    Next i
    
   
    
    f = Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    For i = 1 To f
        Cells.Find(What:="занят", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    ActiveCell = ("Клиент не отвечает по телефону. Уточнить актуальность.")
    
    Next i
    
   
    
    f = Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    For i = 1 To f
        Cells.Find(What:="автоответчик", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    ActiveCell = ("Клиент не отвечает по телефону. Уточнить актуальность.")
    
    Next i
    
    
    
    
    f = Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    For i = 1 To f
        Cells.Find(What:="не выходит на связь", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    ActiveCell = ("Клиент не отвечает по телефону. Уточнить актуальность.")
    
    Next i
    
        f = Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    For i = 1 To f
        Cells.Find(What:="срабатывает на сброс гудков", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    ActiveCell = ("Клиент не отвечает по телефону. Уточнить актуальность.")
    
    Next i
    
            f = Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    For i = 1 To f
        Cells.Find(What:="сбрасывают звонок", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    ActiveCell = ("Клиент не отвечает по телефону. Уточнить актуальность.")
    
    Next i
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    f = Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    For i = 1 To f
        Cells.Find(What:="не зарегистрирован", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    ActiveCell = ("Номер клиента в сети не зарегестрирован.")
    
    Next i
    
    
    
    f = Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    For i = 1 To f
        Cells.Find(What:="неверный", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    ActiveCell = ("Неверный номер")
    
    Next i
    
    
    
    
    f = Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    For i = 1 To f
        Cells.Find(What:="неверный", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    ActiveCell = ("Неверный номер")
    
    Next i
    
        f = Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    For i = 1 To f
        Cells.Find(What:="Просьба уточнить телефон", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    ActiveCell = ("Неверный номер")
    
    Next i


        f = Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    For i = 1 To f
        Cells.Find(What:="уточнить номер", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    ActiveCell = ("Неверный номер")
    
    Next i
    
            f = Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    For i = 1 To f
        Cells.Find(What:="не обслуживается", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    ActiveCell = ("Неверный номер")
    
    Next i
    
End Sub

Private Sub CommandButton95_Click()
Dim d As Date
    
    d = "01.01.21"
    
    For pochta = 1 To 30
        X = 0
        
        For i = 1 To 16000
        If Workbooks("Table.xlsx").Sheets("отправления").Range("f" & i) > d Then
            If Workbooks("Table.xlsx").Sheets("отправления").Range("w" & i) = "Почта" Then
            
                If Workbooks("Table.xlsx").Sheets("отправления").Range("a" & i) = Range("a" & pochta) Then
                    For cena = 1 To 200
                        If Workbooks("Table.xlsx").Sheets("отправления").Range("i" & i) = Workbooks("Table.xlsx").Sheets("цены_наименования").Range("b" & cena) Then
                            If Workbooks("Table.xlsx").Sheets("цены_наименования").Range("g" & cena) > 5000 Then
                                X = X + 1

                            End If
                        End If
                    Next cena
                
  
                End If
            End If
        End If
        
        Next i
        
        
    If Range("b" & pochta) = Sheets("Наименования").Range("a2") Then
        Range("e" & pochta) = X
    End If
    
    
    Next pochta
End Sub

Private Function GetFile(ByVal FileName As String, ByVal inFolder As Object) As Object
On Error GoTo errHandle
Set GetFile = inFolder.ParseName(FileName)
Exit Function
errHandle:
Set GetFile = Nothing
End Function


Private Sub CommandButton96_Click()
'    f = Cells(Rows.Count, 1).End(xlUp).Row
'     For i = 1 To f
'      If Range("g" & i) = 0 Then
'         If Dir("C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & city & "\Чеки\" & Range("c" & i) & ".pdf") = Range("c" & i) & ".pdf" Then
'         Kill "C:\Users\ShapkaMY\Desktop\2021\10 Октябрь\" & Trsdate & "\" & city & "\Чеки\" & Range("c" & i) & ".pdf"
'         End If
'    End If
'   Next i



    Dim pShell As Object, pFolder As Object, pFile As Object
    Set pShell = CreateObject("Shell.Application")
    Set pFolder = pShell.Namespace("C:\Users\ShapkaMY\Desktop\test3\Чеки.zip")
    Set pFile = GetFile("273573.pdf", pFolder)
    If Not pFile Is Nothing Then pFile.InvokeVerb ("Delete")
    pFolder.CopyHere "C:\Users\ShapkaMY\Desktop\test3\273573.pdf", 16

End Sub

Private Sub CommandButton97_Click()

    Dim sFileName As String, sNewFileName As String
 
    sFileName = "C:\Users\ShapkaMY\Desktop\test3\273583.pdf"    'имя исходного файла
    sNewFileName = "C:\Users\ShapkaMY\Desktop\test3\b2b_273583.pdf"    'имя файла для переименования
    If Dir(sFileName, 16) = "" Then
        MsgBox "Нет такого файла", vbCritical, "www.excel-vba.ru"
        Exit Sub
    End If
 
    Name sFileName As sNewFileName 'переименовываем файл
 
    MsgBox "Файл переименован", vbInformation, "www.excel-vba.ru"

End Sub

Private Sub CommandButton98_Click()

Dim s7zipPath$, sArcPath$, sArcFile$, sDestPath$, sDelim$, CmdLine$
    
    ' путь к архиватору 7zip
    s7zipPath = "C:\Program Files\7-Zip\7z.exe"
 
    ' путь к архиву (полный или относительный)
    sArcPath = "C:\Users\ShapkaMY\Desktop\test3\Чеки.zip"
    
    ' имя файла в архиве, который нужно распаковать
    sArcFile = "273583.pdf"
 
    ' путь к папке, куда распаковать файлы (полный или относительный)
    sDestPath = "C:\Users\ShapkaMY\Desktop\test3\"
    
    CmdLine = """" & s7zipPath & """" & " x " & """" & sArcPath & """" & " " & """" & sArcFile & """" & " -o" & """" & sDestPath & """" & " -y"
    
    ' асинхронный запуск
    'Shell CmdLine
    
    ' синхронный запуск
    CreateObject("WScript.Shell").Run CmdLine, 1, True
End Sub

Function ZIPOneFile(sZIPFileName As String, sFileToZIP As String)
    Dim objShell As Object
    Dim lcnt As Long
 
    Set objShell = CreateObject("Shell.Application")
    '??????? ?????? ZIP-?????, ???? ??? ??? ???
    If Dir(sZIPFileName, 16) = "" Then
        CreateNewZip (sZIPFileName)
    End If
    lcnt = objShell.Namespace((sZIPFileName)).Items.Count
    '???????? ????? ?? ????? ? ?????
    objShell.Namespace((sZIPFileName)).CopyHere CStr(sFileToZIP)
    '?????????? ????????? ?????????
    Do Until objShell.Namespace((sZIPFileName)).Items.Count = lcnt + 1
        DoEvents
    Loop
End Function
Private Sub CommandButton99_Click()
    Call ZIPOneFile("C:\Users\ShapkaMY\Desktop\test3\Чеки.zip", "C:\Users\ShapkaMY\Desktop\test3\b2b_273583")

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub TextBox23_Change()

End Sub


























