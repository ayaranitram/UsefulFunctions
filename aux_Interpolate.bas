Attribute VB_Name = "aux_Interpolate"
' THIS IS THE SECOND VERSION OF MY INTERPOLATE FUNCTION, now the search is not done row by row.
' This version has been revisited to make variable names easier to read



Public Function INTERPOLATE(Optional ByVal data_range As Variant = "", Optional ByVal input_value As Variant = "", Optional ByVal input_column As Integer = 0, Optional ByVal output_column As Integer = 0, Optional ByVal alternative_sheet As String = "this:sheet")
    '
If TypeName(data_range) <> "Range" Then
    If data_range = "" And input_value = "" And input_column = 0 And OutputColum = 0 And alternative_sheet = "this:sheet" Then
        INTERPOLATE = "syntax: INTERPOLATE(data_range, input_value to interpolate, [input_column], [output_column], [alternative_sheet])"
        Exit Function
    ElseIf input_value = "" Then
        INTERPOLATE = "missing second argument, value to interpolate or look for"
        Exit Function
    End If
ElseIf input_value = "" Then
    INTERPOLATE = "missing second argument, value to interpolate or look for"
    Exit Function
End If

' si el nombre de la hoja esta vacio, el valor por defecto sera la hoja activa
If alternative_sheet = "this:sheet" Then
    alternative_sheet = ""
End If

    INTERPOLATE = INTERPOLAR(data_range, input_value, input_column, output_column, alternative_sheet)
End Function


' funcion interpolar a partir de datos en una tabla
' el primer argumento es el rango de celdas, escrito como string "A1:J35" o como rango
' el segundo argumento es el valor de entrada que hay que interpolar
' el tercer argumento es la posicion de la columna de los valores de entrada, es pcional y por defecto es 1
' el cuarto argumento es la posiciãn de la culumna de los valores de salida, es opcional y por defecto es la siguiente columna

Public Function INTERPOLAR(ByVal rango_datos As Variant, ByVal valor_entrada As Variant, Optional ByVal columna_entrada As Integer = 0, Optional ByVal columna_salida As Integer = 0, Optional ByVal hoja_alternativa As String = "esta:hoja")

Dim InRow, InCol, InTop, InBot, InLeft, InRight, InPos, InRows, PosShift As Long
Dim r, l As Long
Dim InOut, TestVar As Variant


' verificar si rango_datos es de tipo Range o String
If TypeName(rango_datos) = "Range" Then
' es un rango, miramos los extremos del rango:
    InTop = rango_datos.Row
    InBot = rango_datos.Rows(rango_datos.Rows.Count).Row
    InRows = rango_datos.Rows.Count
    InLeft = rango_datos.Column
    InRight = rango_datos.Columns(rango_datos.Columns.Count).Column

    ' si el nombre de la hoja estˆ vacio, el valor por defecto serˆ el indicado en el rango
    If hoja_alternativa = "esta:hoja" Then
        hoja_alternativa = rango_datos.Worksheet.Name
    ElseIf hoja_alternativa = "" Then
        hoja_alternativa = rango_datos.Worksheet.Name
    End If


Else
' no es un rango, ha de ser texto que debemos interpretar:

    ' si el nombre de la hoja esta vacio, el valor por defecto sera la hoja activa
    If hoja_alternativa = "esta:hoja" Then
        hoja_alternativa = ActiveSheet.Name
    End If

    ' encontrar las primeras letras dentro de rango_datos, que representan la primera columna de la tabla
    l = 1
    Do Until IsNumeric(Mid(rango_datos, l, 1)) = True Or l = Len(rango_datos)
        l = l + 1
    Loop
    InLeft = ABCtoNUM(Left(rango_datos, l - 1))

    ' encontrar los primeros numeros dentro de rango_datos, que representan la primera fila de la tabla
    r = l
    Do Until IsNumeric(Mid(rango_datos, r, 1)) = False Or r = Len(rango_datos)
        r = r + 1
    Loop
    InTop = Val(Mid(rango_datos, l, r - l))
    
    ' encontrar las ultimas letras dentro de rango_datos, que representan la ultima columna de la tabla
    r = r + 1
    l = r
    Do Until IsNumeric(Mid(rango_datos, r, 1)) = True Or r = Len(rango_datos)
        r = r + 1
    Loop
    InRight = ABCtoNUM(Mid(rango_datos, l, r - l))

    ' encontrar los ultimos numeros dentro de rango_datos, que representan la ultima fila de la tabla
    InBot = Val(Right(rango_datos, Len(rango_datos) - r + 1))

    InRows = InBot - InTop + 1

End If

' definir columna_entrada por defecto
If columna_entrada = 0 Then
    columna_entrada = 1
End If
' definir columna_salida por defecto
If (columna_salida = 0) Then
    ' sera la proxima columna si estaria dentro de la tabla
    If (columna_entrada + 1) <= (InRight - InLeft + 1) Then
        columna_salida = columna_entrada + 1
    ' sera la columna anterior si estaria dentro de la tabla
    ElseIf (columna_entrada - 1) >= 1 Then
        columna_salida = columna_entrada - 1
    ' sera la igual a columna_entrada si no queda otra opcion
    Else
        columna_salida = columna_entrada
    End If
End If


' redefinir el valor de columna_entrada y columna_salida para que no sea relativo al rango_datos
columna_entrada = InLeft + columna_entrada - 1
columna_salida = InLeft + columna_salida - 1



' si Valor es numerico se podria interpolar, asumimos que en columna_entrada tambien hay valores numericos
If (IsNumeric(valor_entrada) = True Or IsDate(valor_entrada) = True) And columna_entrada <> columna_salida Then
    ' buscar los dos valores entre los cuales interpolar
    Dim InXPos, InX0, InY0, InX1, InY1, InXTop, InXBot, InYTop, InYBot, InXBefore, InYPos As Double
    Dim InPosBefore, InTemporal, InTopNew, InBotNew, nLoop As Long
    
    nLoop = 0

    ' buscar el inicio de la tabla no vacia en columna_salida
    InPos = InTop
    While IsEmpty(Sheets(hoja_alternativa).Cells(InPos, columna_salida)) = True And InPos + 1 < InBot
        InPos = InPos + 1
    Wend
    If InPos < InBot And InPos > InTop Then InTop = InPos
    
    InPos = InBot
    While IsEmpty(Sheets(hoja_alternativa).Cells(InPos, columna_salida)) = True And InPos - 1 > InTop
        InPos = InPos - 1
    Wend
    If InPos > InTop And InPos < InBot Then InBot = InPos

    ' definir los valores inciales y topes entre los cuales interpolar
    InXTop = Sheets(hoja_alternativa).Cells(InTop, columna_entrada)
    InXBot = Sheets(hoja_alternativa).Cells(InBot, columna_entrada)
    
   
    InPos = Int(InBot - (InBot - InTop) / (InXBot - InXTop) * (InXBot - valor_entrada))
    If InPos > InBot Then InPos = InBot
    If InPos < InTop Then InPos = InTop
    While IsEmpty(Sheets(hoja_alternativa).Cells(InPos, columna_salida)) = True And InPos - 1 >= InTop
        InPos = InPos - 1
    Wend
    While IsEmpty(Sheets(hoja_alternativa).Cells(InPos, columna_salida)) = True And InPos + 1 < InBot
        InPos = InPos + 1
    Wend
    PosShift = 1
    While IsEmpty(Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_salida)) = True And InPos + PosShift <= InBot
        PosShift = PosShift + 1
    Wend
    
    InTopNew = InTop
    InBotNew = InBot
    
    InX0 = Sheets(hoja_alternativa).Cells(InPos, columna_entrada)
    InX1 = Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_entrada)
    
    If (InXTop - valor_entrada) * (InXBot - valor_entrada) > 0 Then
    ' valor_entrada fuera del rango de la tabla, hay que extrapolar
    
        If InXTop < InXBot Then
            If valor_entrada < InXTop Then
                InPos = InTop
                While IsEmpty(Sheets(hoja_alternativa).Cells(InPos, columna_salida)) = True And InPos + 1 < InBot
                    InPos = InPos + 1
                Wend
                PosShift = 1
                While IsEmpty(Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_salida)) = True And InPos + PosShift + 1 <= InBot
                    PosShift = PosShift + 1
                Wend
                
                InX0 = Sheets(hoja_alternativa).Cells(InPos, columna_entrada)
                InX1 = Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_entrada)
                InTopNew = InPos
            ElseIf valor_entrada > InXBot Then
                InPos = InBot
                While IsEmpty(Sheets(hoja_alternativa).Cells(InPos, columna_salida)) = True And InPos - 1 > InTop
                    InPos = InPos - 1
                Wend
                PosShift = 1
                While IsEmpty(Sheets(hoja_alternativa).Cells(InPos - PosShift, columna_salida)) = True And InPos - PosShift - 1 >= InTop
                    PosShift = PosShift + 1
                Wend
                InPos = InPos - PosShift
                
                InX0 = Sheets(hoja_alternativa).Cells(InPos, columna_entrada)
                InX1 = Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_entrada)
                InBotNew = InPos + PosShift
            Else
                INTERPOLAR = "ERROR ni mayor ni menor"
                Exit Function
            End If
        Else
            If valor_entrada > InXTop Then
                InPos = InTop
                While IsEmpty(Sheets(hoja_alternativa).Cells(InPos, columna_salida)) = True And InPos + 1 < InBot
                    InPos = InPos + 1
                Wend
                PosShift = 1
                While IsEmpty(Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_salida)) = True And InPos + PosShift + 1 <= InBot
                    PosShift = PosShift + 1
                Wend
                
                InX0 = Sheets(hoja_alternativa).Cells(InPos, columna_entrada)
                InX1 = Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_entrada)
                InTopNew = InPos
            ElseIf valor_entrada < InXBot Then
                InPos = InBot
                While IsEmpty(Sheets(hoja_alternativa).Cells(InPos, columna_salida)) = True And InPos - 1 >= InTop
                    InPos = InPos - 1
                Wend
                PosShift = 1
                While IsEmpty(Sheets(hoja_alternativa).Cells(InPos - PosShift, columna_salida)) = True And InPos - PosShift - 1 > InTop
                    PosShift = PosShift + 1
                Wend
                InPos = InPos - PosShift
                
                InX0 = Sheets(hoja_alternativa).Cells(InPos, columna_entrada)
                InX1 = Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_entrada)
                InBotNew = InPos + PosShift
            Else
                INTERPOLAR = "ERROR ni mayor ni menor"
                Exit Function
            End If
        End If
    
    
    Else
    ' valor_entrada dentro del rago de la tabla, buscar valores para interpolar
    
    Do Until (InX0 - valor_entrada) * (InX1 - valor_entrada) <= 0 Or (InPos + 1) = InBot
        
        nLoop = nLoop + 1
        
        InXPos = Sheets(hoja_alternativa).Cells(InPos, columna_entrada)
        InTemporal = InPos
        
        If InXTop < InXBot Then
            If valor_entrada < InXPos Then
                InBotNew = InPos
            Else
                InTopNew = InPos
            End If
        Else
            If valor_entrada > InXPos Then
                InBotNew = InPos
            Else
                InTopNew = InPos
            End If
        End If
            
        InXTopNew = Sheets(hoja_alternativa).Cells(InTopNew, columna_entrada)
        InXBotNew = Sheets(hoja_alternativa).Cells(InBotNew, columna_entrada)
        
        InPos = Int(InBotNew - (InBotNew - InTopNew) / (InXBotNew - InXTopNew) * (InXBotNew - valor_entrada))
            
        
' experimental, para buscar en tablas con celdas vacias:
        
        'TestVar = IsEmpty(Sheets(hoja_alternativa).Cells(InPos, columna_salida))
        'InYPos = Sheets(hoja_alternativa).Cells(InPos, columna_salida)
        While IsEmpty(Sheets(hoja_alternativa).Cells(InPos, columna_salida)) = True And InPos - 1 >= InTopNew
            InPos = InPos - 1
            'TestVar = IsEmpty(Sheets(hoja_alternativa).Cells(InPos, columna_salida))
            'InYPos = Sheets(hoja_alternativa).Cells(InPos, columna_salida)
        Wend
        While IsEmpty(Sheets(hoja_alternativa).Cells(InPos, columna_salida)) = True And InPos < InBotNew
            InPos = InPos + 1
            'TestVar = IsEmpty(Sheets(hoja_alternativa).Cells(InPos, columna_salida))
            'InYPos = Sheets(hoja_alternativa).Cells(InPos, columna_salida)
        Wend
        
        PosShift = 1 ' en caso de anular el experimento hay que dejar esta linea
        'TestVar = IsEmpty(Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_salida))
        'InYPos = Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_salida)
        While IsEmpty(Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_salida)) = True And (InPos + PosShift + 1) <= InBotNew
            PosShift = PosShift + 1
            'TestVar = IsEmpty(Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_salida))
            'InYPos = Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_salida)
        Wend
        
        
' fin de experimento
            
        
        ' evita bucles infinitos forzando avazar la busqueda cuando quedamos repitiendo los mismos valores de entrada
        If InPos = InTemporal And InPos + 1 < InBot Then
            InPos = InPos + 1
            While IsEmpty(Sheets(hoja_alternativa).Cells(InPos, columna_salida)) = True And InPos < InBotNew
                InPos = InPos + 1
            Wend
            PosShift = 1 ' en caso de anular el experimento hay que dejar esta linea
            While IsEmpty(Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_salida)) = True And (InPos + PosShift + 1) <= InBotNew
                PosShift = PosShift + 1
            Wend
        End If
            
        InX0 = Sheets(hoja_alternativa).Cells(InPos, columna_entrada)
        InX1 = Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_entrada)

    Loop
    
    End If ' If (InXTop - valor_entrada) * (InXBot - valor_entrada) > 0 Then

' si valor_entrada no es numerico no se puede interpolar y entonces haremos busqueda de texto
Else

    Dim InXStr As String

    ' definir los valores inciales para empezar a buscar
    InPos = InTop
    InXStr = Sheets(hoja_alternativa).Cells(InPos, columna_entrada)

    Do Until (InXStr = valor_entrada) Or (InPos = InBot)

        InPos = InPos + 1
        InXStr = Sheets(hoja_alternativa).Cells(InPos, columna_entrada)

    Loop

    If InXStr = valor_entrada Then
        InOut = Sheets(hoja_alternativa).Cells(InPos, columna_salida)
    Else
        InOut = "Not Found"
    End If

    INTERPOLAR = InOut

    Exit Function

End If




If valor_entrada = InX0 Then
' el valor a interpolar esta explicito en la tabla
    InOut = Sheets(hoja_alternativa).Cells(InPos, columna_salida)

ElseIf valor_entrada = InX1 Then
' el valor a interpolar esta explicito en la tabla
    InOut = Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_salida)

ElseIf (InX0 - valor_entrada) * (InX1 - valor_entrada) < 0 Then
' interpolando entre dos valores de la tabla que acotan a valor_entrada
    InY0 = Sheets(hoja_alternativa).Cells(InPos, columna_salida)
    InY1 = Sheets(hoja_alternativa).Cells(InPos + PosShift, columna_salida)
        
    If IsNumeric(InY0) = True And IsNumeric(InY1) = True Then
        InOut = InY0 + (valor_entrada - InX0) / (InX1 - InX0) * (InY1 - InY0)
    Else
        InOut = "between " & InY0 & " and " & InY1
    End If
    
ElseIf InPos + PosShift = InBot Or InPos = InTop Then

    If Abs(InXTop - valor_entrada) <= Abs(InXBot - valor_entrada) Then
' extrapolando por el tope de la tabla
        InY0 = Sheets(hoja_alternativa).Cells(InTopNew, columna_salida)
        InY1 = Sheets(hoja_alternativa).Cells(InTopNew + PosShift, columna_salida)
        InX0 = Sheets(hoja_alternativa).Cells(InTopNew, columna_entrada)
        InX1 = Sheets(hoja_alternativa).Cells(InTopNew + PosShift, columna_entrada)

        InOut = InY0 + (valor_entrada - InX0) / (InX1 - InX0) * (InY1 - InY0)
        
    Else
' extrapolando por la base de la tabla
        InY0 = Sheets(hoja_alternativa).Cells(InBotNew - PosShift, columna_salida)
        InY1 = Sheets(hoja_alternativa).Cells(InBotNew, columna_salida)
        InX0 = Sheets(hoja_alternativa).Cells(InBotNew - PosShift, columna_entrada)
        InX1 = Sheets(hoja_alternativa).Cells(InBotNew, columna_entrada)

        InOut = InY0 + (valor_entrada - InX0) / (InX1 - InX0) * (InY1 - InY0)
        
    End If
Else
' NPI
   InOut = "error"

End If


'INTERPOLAR = InOut & " in " & nLoop & " loops" ' escribir numero de bucles no tiene valor para el usuario final
INTERPOLAR = InOut

End Function

