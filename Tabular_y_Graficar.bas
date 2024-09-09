Attribute VB_Name = "Tabular_y_Graficar"
Sub TabularFuncion()
Attribute TabularFuncion.VB_Description = "Tabular y Graficar"
Attribute TabularFuncion.VB_ProcData.VB_Invoke_Func = "A\n14"
    Dim variable As String
    Dim funcion As String
    Dim rangoInicioStr As String, rangoFinStr As String, intervaloStr As String
    Dim rangoInicio As Double, rangoFin As Double, intervalo As Double
    Dim filaActual As Long
    Dim valorVariable As Double
    Dim resultado As Variant
    Dim formula As String
    Dim columnaVariable As Long, columnaFuncion As Long
    Dim celdaInicial As Range
    Dim i As Integer
    Dim tablaRango As Range
    Dim tabla As ListObject
    Dim grafico As ChartObject
    Dim graficoRango As Range
    
    
    
    Set celdaInicial = Selection.Cells(1, 1)
    filaActual = celdaInicial.Row
    columnaVariable = celdaInicial.Column
    columnaFuncion = columnaVariable + 1
    

    variable = InputBox("Ingrese el nombre de la variable:", "Variable")
    If variable = "" Then Exit Sub
    

    funcion = InputBox("Ingrese la función (en términos de '" & variable & "'):", "Función")
    If funcion = "" Then Exit Sub
    
    
    
 
    rangoInicioStr = InputBox("Ingrese el inicio del rango:", "Inicio del Rango")
    rangoFinStr = InputBox("Ingrese el fin del rango:", "Fin del Rango")
    If rangoInicioStr = "" Or rangoFinStr = "" Then Exit Sub
    

    intervaloStr = InputBox("Ingrese el intervalo para la tabulación:", "Intervalo")
    If intervaloStr = "" Then Exit Sub
    

    On Error Resume Next
    rangoInicio = CDbl(rangoInicioStr)
    rangoFin = CDbl(rangoFinStr)
    intervalo = CDbl(intervaloStr)
    If Err.Number <> 0 Then
        MsgBox "El rango o intervalo ingresado no es válido."
        Exit Sub
    End If
    On Error GoTo 0

    With Cells(filaActual, columnaVariable)
        .Value = variable
        .Font.Size = 14
        .Font.Bold = True
    End With
    
    With Cells(filaActual, columnaFuncion)
        .Value = "F(" & variable & ") = " & funcion
        .Font.Size = 14
        .Font.Bold = True
    End With

    funcion = AgregarMultiplicacionImplicita(funcion)

    filaActual = filaActual + 1
    For i = rangoInicio To rangoFin Step intervalo
        valorVariable = i
        formula = Replace(funcion, variable, valorVariable)
        
       
        On Error Resume Next
        resultado = Evaluate(formula)
        If Err.Number <> 0 Then
            resultado = "Error"
            Err.Clear
        End If
        On Error GoTo 0
        
        Cells(filaActual, columnaVariable).Value = valorVariable
        Cells(filaActual, columnaFuncion).Value = resultado
        
        filaActual = filaActual + 1
    Next i
    
     Set tablaRango = Range(Cells(celdaInicial.Row, columnaVariable), Cells(filaActual - 1, columnaFuncion))
    

    On Error Resume Next
    Set tabla = ActiveSheet.ListObjects.Add(xlSrcRange, tablaRango, , xlYes)
    On Error GoTo 0
    

    If Not tabla Is Nothing Then
        tabla.TableStyle = "TableStyleMedium3"
    End If
    
    Set graficoRango = Range(Cells(celdaInicial.Row + 1, columnaVariable + 1), Cells(filaActual - 1, columnaFuncion))
    Set grafico = ActiveSheet.ChartObjects.Add(Left:=Cells(celdaInicial.Row, columnaFuncion + 2).Left, _
                                               Width:=300, _
                                               Top:=Cells(celdaInicial.Row, columnaFuncion).Top, _
                                               Height:=300)
    
    With grafico.Chart
        .ChartType = xlXYScatterLines
        .HasTitle = True
        .ChartTitle.Text = "Gráfico de Dispersión de " & variable
        .HasLegend = False
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "F(" & variable & ")"
        
        With .Axes(xlCategory)
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = RGB(0, 0, 255)
            .Format.Line.Weight = 1
        End With
        
         With .SeriesCollection.NewSeries
            .XValues = Range(Cells(celdaInicial.Row + 1, columnaVariable), Cells(filaActual - 1, columnaVariable)) ' Valores X
            .Values = Range(Cells(celdaInicial.Row + 1, columnaFuncion), Cells(filaActual - 1, columnaFuncion)) ' Valores Y
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 5
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            .Format.Line.Weight = 1
        End With
    End With
    
    MsgBox "La tabulación se completó con éxito."
End Sub

Function AgregarMultiplicacionImplicita(ByVal funcion As String) As String
    Dim i As Integer
    Dim resultado As String
    Dim caracterActual As String, caracterSiguiente As String
    
    resultado = ""
    
    For i = 1 To Len(funcion)
        caracterActual = Mid(funcion, i, 1)
        If i < Len(funcion) Then
            caracterSiguiente = Mid(funcion, i + 1, 1)
            If IsNumeric(caracterActual) And (caracterSiguiente Like "[A-Za-z]") Then
                resultado = resultado & caracterActual & "*"
            Else
                resultado = resultado & caracterActual
            End If
        Else
            resultado = resultado & caracterActual
        End If
    Next i
    
    AgregarMultiplicacionImplicita = resultado
End Function







