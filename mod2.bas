Attribute VB_Name = "Module2"
Public Sub labelsTamano(frm As Object)
    frm.lblDesc.Visible = True
    frm.lblCons.Visible = False
    frm.lblStock1.Visible = True
    frm.lblStock2.Visible = True
    frm.lblStock3.Visible = True
    
    
    
    frm.ListBox1.ColumnWidths = frm.ListBox1.Width * 0.35 & ";" & _
                                frm.ListBox1.Width * 0.39 & ";" & _
                                frm.ListBox1.Width * 0.095 & ";" & _
                                frm.ListBox1.Width * 0.08

    frm.lblArt.Width = frm.ListBox1.Width * 0.35
    frm.lblDesc.Width = frm.ListBox1.Width * 0.39
    frm.lblDesc.left = frm.ListBox1.left + frm.lblArt.Width
    
    With frm.lblStock1
        .Height = frm.Height * 0.041
        .Width = frm.ListBox1.Width * 0.095
        .top = frm.Height * 0.16
        .left = frm.lblDesc.left + frm.lblDesc.Width
        .Font.Name = "Bierstadt"
        .Font.Size = frm.Height * 0.024
        .Caption = "Reform."
    End With

    With frm.lblStock2
        .Height = frm.Height * 0.041
        .Width = frm.ListBox1.Width * 0.16
        .top = frm.Height * 0.16
        .left = frm.lblStock1.left + frm.lblStock1.Width
        .Font.Name = "Bierstadt"
        .Font.Size = frm.Height * 0.024
        .Caption = "Nuevo"
    End With

    With frm.lblStock3
        .Height = frm.Height * 0.041
        .Width = frm.ListBox1.Width * 0.08
        .top = frm.Height * 0.16
        .left = frm.lblStock2.left + frm.lblStock2.Width
        .Font.Name = "Bierstadt"
        .Font.Size = frm.Height * 0.024
        .Visible = False
    End With
End Sub

Public Sub elementosPageDetalles(frm As Object)
    Dim bloque1x As Long, bloque2x As Long, bloque3x As Long
    Dim bloque1y As Long, bloque2y As Long, bloque3y As Long
    Dim bloque1yy As Long, bloque2yy As Long, bloque3yy As Long
    Dim altura As Long, ancho As Long, altura2 As Long, ancho2 As Long
    Dim letra As Long
    
    bloque1y = frm.Height * 0.05
    bloque2y = frm.Height * 0.16
    bloque3y = frm.Height * 0.27
    bloque1yy = frm.Height * 0.085
    bloque2yy = frm.Height * 0.195
    bloque3yy = frm.Height * 0.302
    
    bloque1x = frm.Width * 0.065
    bloque2x = frm.Width * 0.315
    bloque3x = frm.Width * 0.56
    
    altura = frm.Height * 0.033
    ancho = frm.Width * 0.1
    
    altura2 = altura * 1.5
    ancho2 = ancho * 1.7
    
    letra = frm.Height * 0.024
    letraCaja = frm.Height * 0.022

    With frm.Label11
        .Height = altura
        .Width = ancho
        .top = bloque1y
        .left = bloque1x
        .Font.Size = letra
    End With
    
    With frm.Label12
        .Height = altura
        .Width = ancho
        .top = bloque2y * 1.37
        .left = bloque1x
        .Font.Size = letra
    End With
    
    With frm.Label13
        .Height = altura
        .Width = ancho
        '.Top = bloque3y * 1.43
        .top = bloque1y
        .left = bloque2x
        .Font.Size = letra
    End With
    
    With frm.Label15
        .Height = altura
        .Width = ancho
        .top = bloque2y * 1.2
        .left = bloque2x
        .Font.Size = letra
    End With
    
    With frm.Label27
        .Height = altura
        .Width = ancho
        '.Top = bloque3y * 1.05
        .top = frm.Label12.top + frm.Label12.Height + (altura2 * 2.4) - altura2 - altura
        .left = bloque2x
        .Font.Size = letra
    End With
    
    With frm.Label14
        .Height = altura
        .Width = ancho
        .top = bloque1y
        .left = bloque3x
        .Font.Size = letra
    End With
    
    With frm.Label17
        .Height = altura
        .Width = ancho
        '.Top = bloque2y
        '.Top = frm.Label15.Top
        .top = frm.Label11.top + ((frm.Label27.top - bloque1y) / 2) '+ (.Height + altura2) / 16
        .left = bloque3x
        .Font.Size = letra
    End With
    
    With frm.Label16
        .Height = altura
        .Width = ancho
        '.Top = bloque3y
        .top = frm.Label27.top
        .left = bloque3x
        .Font.Size = letra
    End With
    
    With frm.Label18
        .Height = altura
        .Width = ancho
        .top = bloque3y
        .left = bloque3x
        .Font.Size = letra
        .Visible = False
    End With
    
    With frm.TextBox1
        .Height = altura2 * 2.4
        .Width = ancho2 * 1.35
        .top = frm.Label11.top + frm.Label11.Height
        .left = bloque1x
        .Font.Size = letraCaja
        .Enabled = True
        .WordWrap = True
        .Multiline = True
        .Locked = True
    End With
    
    With frm.TextBox2
        .Height = altura2 * 2.4
        .Width = ancho2 * 1.35
        .top = frm.Label12.top + frm.Label12.Height
        .left = bloque1x
        .Font.Size = letraCaja
        .Enabled = True
        .WordWrap = True
        .Multiline = True
        .Locked = True
    End With
    
    With frm.TextBox3
        .Height = altura2 * 2
        .Width = ancho2 * 1.35
        .top = frm.Label13.top + frm.Label13.Height
        .left = frm.Label13.left
        .Font.Size = letraCaja
        .Enabled = True
        .WordWrap = True
        .Multiline = True
        .Locked = True
    End With
    
    With frm.TextBox4
        .Height = altura2
        .Width = frm.Label16.Width * 0.8
        .top = frm.Label14.top + frm.Label14.Height
        .left = frm.Label14.left
        .Font.Size = letraCaja
        .Enabled = True
        .WordWrap = True
        .Multiline = True
        .Locked = True
    End With
    
    With frm.TextBox5
        .Height = altura2
        .Width = ancho2
        .top = frm.Label15.top + frm.Label15.Height
        .left = bloque2x
        .Font.Size = letraCaja
        .Enabled = True
        .WordWrap = True
        .Multiline = True
        .Locked = True
    End With
    
    With frm.TextBox6
        .Height = altura2
        .Width = frm.Label16.Width * 0.8
        .top = frm.Label16.top + frm.Label16.Height
        .left = bloque3x
        .Font.Size = letraCaja
        .Enabled = True
        .WordWrap = True
        .Multiline = True
        .Locked = True
    End With
    
    With frm.TextBox7
        .Height = altura2
        .Width = frm.Label16.Width * 0.8
        .top = frm.Label17.top + frm.Label17.Height
        .left = bloque3x
        .Font.Size = letraCaja
        .Enabled = True
        .WordWrap = True
        .Multiline = True
        .Locked = True
    End With
    
    With frm.TextBox8
        .Height = altura2
        .Width = frm.Label16.Width * 0.8
        .top = frm.Label18.top + frm.Label18.Height
        .left = bloque3x
        .Font.Size = letraCaja
        .Enabled = True
        .WordWrap = True
        .Multiline = True
        .Locked = True
    End With
    
    With frm.TextBox9
        .Height = altura2
        .Width = ancho2
        .top = frm.Label27.top + frm.Label27.Height
        .left = bloque2x
        .Font.Size = letraCaja
        .Enabled = True
        .WordWrap = True
        .Multiline = True
        .Locked = True
    End With
    
    With frm.btVolver789
        .Height = frm.MultiPage1.Height
        .Width = frm.Width * 0.025
        .top = 0
        .left = 0
        .Font.Size = letra
    End With
   
    With frm.ListBox2
        .Height = frm.Height * 0.35
        .Width = frm.MultiPage1.Width * 0.6
        .top = frm.TextBox2.top + frm.TextBox2.Height * 1.7
        .left = frm.TextBox2.left
        .ColumnCount = 5
        .Font.Name = "Bierstadt"
        .Font.Size = frm.Height * 0.022
        .ColumnWidths = frm.ListBox2.Width * 0.25 & ";" & _
                        frm.ListBox2.Width * 0.36 & ";" & _
                        frm.ListBox2.Width * 0.15 & ";" & _
                        frm.ListBox2.Width * 0.13 & ";" & _
                        frm.ListBox2.Width * 0.1
        .TextAlign = fmTextAlignLeft
    End With
    
    With frm.Label28
        .Height = frm.Height * 0.041
        .Width = frm.ListBox2.Width * 0.25
        .top = frm.ListBox2.top - .Height
        .left = frm.ListBox2.left
        .Caption = "Fecha"
        .Font.Size = frm.Height * 0.024
    End With
    
    With frm.Label34
        .Height = frm.Height * 0.041
        .Width = frm.ListBox2.Width * 0.36
        .top = frm.Label28.top
        .left = frm.ListBox2.left + frm.Label28.Width
        .Caption = "Usuario"
        .Font.Size = frm.Height * 0.024
    End With
    
    With frm.Label29
        .Height = frm.Height * 0.041
        .Width = frm.ListBox2.Width * 0.16
        .top = frm.Label28.top
        .left = frm.Label34.left + frm.Label34.Width
        .Caption = "Condición"
        .Font.Size = frm.Height * 0.024
    End With
    
    With frm.Label30
        .Height = frm.Height * 0.041
        .Width = frm.ListBox2.Width * 0.12
        .top = frm.Label28.top
        .left = frm.Label29.left + frm.Label29.Width
        .Caption = "Código"
        .Font.Size = frm.Height * 0.024
    End With
    
    With frm.Label31
        .Height = frm.Height * 0.041
        .Width = frm.ListBox2.Width * 0.11
        .top = frm.Label28.top
        .left = frm.Label30.left + frm.Label30.Width
        .Caption = "QYT"
        .Font.Size = frm.Height * 0.024
    End With

    With frm.ListBox3
        .Height = frm.Height * 0.25
        .Width = frm.Width * 0.5
        .top = frm.Height * 0.45
        .left = bloque1x
        .ColumnCount = 5
    End With
End Sub

Public Sub plantilla(frm As Object)
    Dim vector() As Variant
    Dim indice As Long
    Dim item As String
    Dim delta As Long
    Dim fila As Range, ultFila As Long
    Dim fila2 As Range, ultfila2 As Long
    Dim alpha As Long
    Dim primero As String
    Dim contador As Long
    
    indice = frm.ListBox1.ListIndex
    If indice <> -1 Then
        item = frm.ListBox1.List(indice, 0)
        ultFila = Exist.ListObjects("tabExistencias").ListRows.Count + 2
        
        If item = "" Then item = frm.ListBox1.List(indice - 1, 0)
        
        Set fila = Exist.Range("D1:D" & ultFila).Find(item, LookIn:=xlValues)
        delta = Trim(Replace(fila.Address, "$D$", Delete))
        
        frm.TextBox1.Value = Exist.Range("D" & delta).Value
        frm.TextBox2.Value = Exist.Range("E" & delta).Value
        frm.TextBox3.Value = Exist.Range("F" & delta).Value
        frm.TextBox4.Value = Exist.Range("M" & delta).Value
        frm.TextBox5.Value = Exist.Range("N" & delta).Value
        frm.TextBox6.Value = Exist.Range("I" & delta).Value
        frm.TextBox7.Value = Exist.Range("H" & delta).Value
        frm.TextBox8.Value = Exist.Range("J" & delta).Value
        frm.TextBox9.Value = Exist.Range("O" & delta).Value
        
        frm.ListBox2.Clear
        ultfila2 = consumo.ListObjects("DailyConsumption").ListRows.Count + 2
        Set fila2 = consumo.Range("A2:A" & ultfila2).Find(item, LookIn:=xlValues, SearchDirection:=xlPrevious)
        
        If Not fila2 Is Nothing Then
            alpha = Trim(Replace(fila2.Address, "$A$", Delete))
            primero = fila2.Address
            
            ReDim vector(1 To ultFila, 1 To 5)
            contador = 1
            
            vector(1, 1) = consumo.Range("E" & alpha).Value
            vector(1, 2) = consumo.Range("F" & alpha).Value
            vector(1, 3) = consumo.Range("B" & alpha).Value
            vector(1, 4) = consumo.Range("C" & alpha).Value
            vector(1, 5) = "    " & consumo.Range("D" & alpha).Value
            
            Do
                Set fila2 = consumo.Range("A2:A" & ultfila2).FindPrevious(fila2)
                If fila2.Address <> primero Then
                    alpha = Trim(Replace(fila2.Address, "$A$", Delete))
                    contador = contador + 1
                    vector(contador, 1) = consumo.Range("E" & alpha).Value
                    vector(contador, 2) = consumo.Range("F" & alpha).Value
                    vector(contador, 3) = consumo.Range("B" & alpha).Value
                    vector(contador, 4) = consumo.Range("C" & alpha).Value
                    If Not consumo.Range("D" & alpha).Value Like "-*" Then
                        vector(contador, 5) = "     " & consumo.Range("D" & alpha).Value
                    Else
                        vector(contador, 5) = "    " & consumo.Range("D" & alpha).Value
                    End If
                End If
            Loop While fila2.Address <> primero
            
            vector = Application.Transpose(vector)
            ReDim Preserve vector(1 To 5, 1 To contador + 1)
            vector = Application.Transpose(vector)
            frm.ListBox2.List = vector
        Else
            frm.ListBox2.AddItem ("Sin consumos registrados...")
        End If
        
        frm.MultiPage1.Pages.item(2).Visible = True
        frm.MultiPage1.Pages.item(0).Visible = False
        frm.MultiPage1.Pages.item(1).Visible = False
    Else
        MsgBox "Selecciona un artículo.", vbExclamation, "Error"
    End If
End Sub



'kalo24junMod2
