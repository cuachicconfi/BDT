Imports System.IO
Imports System.Text
Public Class clsBDT
  Private objBD As clsBD
  Private dtYield As DataTable
  Private dtVol As DataTable
  Private dtOpciones As DataTable
  Private dtPrueba As DataTable
  Private dtSuma As DataTable
  Private dtCurva As DataTable
  Private dtError As DataTable
  Private dblShortRateLat As Double(,)
  Private dblValorCupon As Double(,)
  Private dblEntramadoBDT As Double(,)
  Private dblArbolPrecios As Double(,)
  Private enumFrecuencia As enumFrecuenciaDePago
  Private enumCurva As enumCurvas
  Private enumDura As enumDuracion

  Public Enum enumFrecuenciaDePago
    Mensual = 1 / 12
    Bimestral = 1 / 6
    Trimestral = 1 / 4
    Semestral = 1 / 2
    Anual = 1
  End Enum
  Public Enum enumCurvas
    bonosm = 0
    irs = 1
    libor = 2
  End Enum
  Public Enum enumDuracion
    unAño = 1
    dosAños = 2
    tresAños = 3
    cuatroAños = 4
    cincoAños = 5
    seisAños = 6
    sieteAños = 7
    ochoAños = 8
    nueveAños = 9
    diezAños = 10
    quinceAños = 11
    veinteAños = 12
    treintaAños = 13
  End Enum

  ''' <summary>
  ''' Constructor de la clase BDT
  ''' </summary>
  ''' <param name="_dtInicial">Valores Iniciales de la curva. En base 1 (no porcentual)</param>
  ''' <param name="_enumFrecuenciadePago">Frecuencia con la que se va a pagar</param>
  ''' <param name="_dtVolatilidad">Vector con la volatilidad</param>
  ''' <remarks></remarks>
  Public Sub New(ByVal _dtInicial As DataTable, _
                 ByVal _enumFrecuenciadePago As enumFrecuenciaDePago, _
                 Optional ByVal _dtVolatilidad As DataTable = Nothing) 'Valores en Base 1 (No porcentual)
    dtYield = _dtInicial.Copy
    enumFrecuencia = _enumFrecuenciadePago
    If Not _dtVolatilidad Is Nothing Then
      dtVol = _dtVolatilidad
    End If

  End Sub

  ''' <summary>
  ''' Constructor de la clase BDT - Consigue los valores de las curvas
  ''' </summary>
  ''' <param name="_enumFrecuenciadePago">Frecuencia con la que se va a pagar</param>
  ''' <remarks></remarks>
  Public Sub New(ByVal _enumCurvas As enumCurvas, _
                 ByVal _enumDuracion As enumDuracion, _
                 ByVal _enumFrecuenciadePago As enumFrecuenciaDePago)

    Me.objBD = New clsBD(VAR.Configuracion.LeeSetting("CadenaConexionBD"))
    enumCurva = _enumCurvas
    enumDura = _enumDuracion
    enumFrecuencia = _enumFrecuenciadePago

    consigueDatosCurva()

    'dtYield = _dtInicial.Copy

    'If Not _dtVolatilidad Is Nothing Then
    'dtVol = _dtVolatilidad
    'End If

  End Sub

  ''' <summary>
  ''' Regresa la calibración del vector _dtInicial, el cual da el menor error cuadrado dado el método BDT
  ''' </summary>
  ''' <returns>Calibración dado el vector _dtInicial</returns>
  ''' <remarks></remarks>
  Public Function calibracion() As Double()
    Dim dblXnSig() As Double
    Dim dblXnAnt() As Double
    Dim dblXn() As Double
    Dim dblFXnAnt() As Double
    Dim dblFXn() As Double

    Dim dblSumaErrorCuadrado As Double
    Try
      'Dimensionamos los arreglos
      ReDim dblXn(Me.dtYield.Rows.Count - 1)
      ReDim dblXnSig(Me.dtYield.Rows.Count - 1)
      ReDim dblXnAnt(Me.dtYield.Rows.Count - 1)
      ReDim dblFXnAnt(Me.dtYield.Rows.Count - 1)
      ReDim dblFXn(Me.dtYield.Rows.Count - 1)

      For j As Integer = 0 To dblXn.Length - 1
        dblXn(j) = 0.99
        dblXnAnt(j) = 0.00001
        dblXnSig(j) = 0.99
      Next

      'Empezamos la iteración con valor máximo de 1000 para que no continúe por siempre
      For iPosicion As Integer = 0 To Me.dtYield.Rows.Count - 1
        For iTeracion As Integer = 0 To 100
          dblSumaErrorCuadrado = 0
          'Al principio le asignamos valores
          If iTeracion = 0 Then

            For j As Integer = iPosicion To dblXn.Length - 1
              dblXn(j) = 0.99
              dblXnAnt(j) = 0.00001
              dblXnSig(j) = 0.99
            Next

            'Calculamos el valor de FXn, i.e. Error Cuadrado de nuestra B.D.T.
            calculaBDT(dblXnAnt)

            For j As Integer = 0 To Me.dtError.Rows.Count - 1
              dblFXnAnt(j) = Me.dtError(j)(0)
            Next

          Else
            Array.Copy(dblXn, dblXnAnt, dblXn.Length)
            Array.Copy(dblXnSig, dblXn, dblXnSig.Length)
          End If

          'Valor de FXn
          calculaBDT(dblXn)

          If iTeracion <> 0 Then
            Array.Copy(dblFXn, dblFXnAnt, dblFXn.Length)
          End If

          For j As Integer = 0 To Me.dtError.Rows.Count - 1
            dblFXn(j) = Me.dtError(j)(0)
          Next

          dblXnSig(iPosicion) = dblXn(iPosicion) - ((dblXn(iPosicion) - dblXnAnt(iPosicion)) / _
                                                    (dblFXn(iPosicion) - dblFXnAnt(iPosicion)) * _
                                                    dblFXn(iPosicion))

          If dblFXn(iPosicion) < 0.000000001 Then
            Exit For
          End If
        Next
      Next

      For iSuma As Integer = 0 To dblXnSig.Length - 1
        dblSumaErrorCuadrado += dblFXn(iSuma)
      Next

      calculaBDT(dblXnSig)
      Me.imprimePrimeraMatriz()
      Me.imprimeSegundaMatriz()
      Me.imprimeTerceraMatriz()
      Me.imprimeCalibracion(dblXnSig, dblSumaErrorCuadrado)
      Return dblXnSig
    Catch ex As Exception
      Throw New Exception("Error en clsBDT.calibracion1por1: " & vbCrLf & _
                          ex.Message & vbCrLf)
    End Try
  End Function

  ''' <summary>
  ''' Calcula el árbol de precios y regresa el valor del cupón.
  ''' </summary>
  ''' <returns>El valor del cupón</returns>
  ''' <remarks></remarks>
  ''' 
  Public Function calculaPrecio() As Decimal
    Dim dblValorCalibracion As Double()
    Try
      dblValorCalibracion = Me.calibracion()
      Me.calculaBDT(dblValorCalibracion)

      Me.creaTablaOpciones()

      Me.CalculaArboldePrecios()
      Return Me.dblArbolPrecios(0, 0)
    Catch ex As Exception
      Throw New Exception("Error en clsBDT.calculaPrecio:" & vbCrLf & _
                 ex.Message & vbCrLf)
    End Try
  End Function

#Region "BDT"
  Private Sub calculaBDT(ByVal dblPrueba As Double())
    Try
      calculaVolatilidad()
      primerPaso(dblPrueba)
      segundoPaso()
      tercerPaso()
      Precio()
      Curva()
      ErrorCuadrado()
    Catch ex As Exception
      Throw New Exception("Error en clsBDT.calculaBDT: " & vbCrLf & _
                          ex.Message & vbCrLf)
    End Try
  End Sub

  Private Sub primerPaso(ByVal dblPrueba() As Double) 'Enrejado de Corto Plazo
    Dim intTamaño As Integer
    Dim j() As Double
    Try
      ReDim Me.dblShortRateLat(dblPrueba.Length - 1, dblPrueba.Length - 1)
      ReDim j(dblPrueba.Length - 1)
      'Tenemos que calcular los valores (j) con los cuáles se va a elevar e
      'La fórmula es la siguiente: dblPrueba(i)*e^(j*volatilidad(i))
      intTamaño = dblPrueba.Length - 1

      j(0) = 0
      For i As Integer = 1 To intTamaño
        j(i) = j(i - 1) + 1 / enumFrecuencia
      Next

      For iRow As Integer = 0 To intTamaño
        For iColumn As Integer = 0 To iRow
          If iRow = 0 And iColumn = 0 Then
            Me.dblShortRateLat(iRow, iColumn) = dblPrueba(iRow)
          Else
            Me.dblShortRateLat(iRow, iColumn) = CDbl(dblPrueba(iRow) * Math.Exp(j(iColumn) * Me.dtVol.Rows(iRow).Item(0)))
          End If

        Next
      Next
    Catch ex As Exception
      Throw New Exception("Error en clsBDT.primerPaso: " & vbCrLf & _
                          ex.Message & vbCrLf)
    End Try
  End Sub
  Private Sub segundoPaso() 'Valor Cupón
    Try
      ReDim Me.dblValorCupon(Me.dtYield.Rows.Count - 1, dtYield.Rows.Count - 1)
      For iRow As Integer = 0 To dtYield.Rows.Count - 1
        For iColumn As Integer = 0 To iRow
          Me.dblValorCupon(iRow, iColumn) = CDbl(1 / (1 + Me.dblShortRateLat(iRow, iColumn)))
        Next
      Next
    Catch ex As Exception
      Throw New Exception("Error en clsBDT.segundoPaso: " & vbCrLf & _
                          ex.Message & vbCrLf)
    End Try
  End Sub
  Private Sub tercerPaso()
    Try
      ReDim Me.dblEntramadoBDT(Me.dtYield.Rows.Count, dtYield.Rows.Count)
      For iRow As Integer = 0 To dtYield.Rows.Count
        If iRow = 0 Then
          Me.dblEntramadoBDT(0, 0) = 1
        Else
          For iColumn As Integer = 0 To iRow
            If iColumn = 0 Then
              Me.dblEntramadoBDT(iRow, iColumn) = CDbl(0.5 * Me.dblEntramadoBDT(iRow - 1, iColumn) * Me.dblValorCupon(iRow - 1, iColumn))
            ElseIf iColumn = iRow Then
              Me.dblEntramadoBDT(iRow, iColumn) = CDbl(0.5 * Me.dblEntramadoBDT(iRow - 1, iColumn - 1) * Me.dblValorCupon(iRow - 1, iColumn - 1))
            Else
              Me.dblEntramadoBDT(iRow, iColumn) = CDbl(0.5 * (Me.dblEntramadoBDT(iRow - 1, iColumn) * Me.dblValorCupon(iRow - 1, iColumn) + _
                                                            Me.dblEntramadoBDT(iRow - 1, iColumn - 1) * Me.dblValorCupon(iRow - 1, iColumn - 1)))
            End If
          Next
        End If
      Next
    Catch ex As Exception
      Throw New Exception("Error en clsBDT.tercerPaso: " & vbCrLf & _
                          ex.Message & vbCrLf)
    End Try
  End Sub
  Private Sub Precio()
    Dim dblSuma As Double
    Dim dr As DataRow
    Try
      Me.dtSuma = New DataTable
      Me.dtSuma.Columns.Add("Valor", Type.GetType("System.Double"))

      For iRow As Integer = 1 To Me.dblEntramadoBDT.GetLength(0) - 1
        dblSuma = 0
        For iColumn As Integer = 0 To Me.dblEntramadoBDT.GetLength(1) - 1
          dblSuma += dblEntramadoBDT(iRow, iColumn)
        Next
        dr = Me.dtSuma.NewRow
        dr.Item(0) = dblSuma * 100
        Me.dtSuma.Rows.Add(dr)
      Next
    Catch ex As Exception
      Throw New Exception("Error en clsBDT.Precio: " & vbCrLf & _
                          ex.Message & vbCrLf)
    End Try
  End Sub
  Private Sub Curva()
    Dim dr As DataRow
    Dim j() As Double
    Try
      ReDim j(Me.dtSuma.Rows.Count - 1)
      dtCurva = New DataTable
      dtCurva.Columns.Add("Valor", Type.GetType("System.Double"))

      j(0) = 1 / enumFrecuencia
      For i As Integer = 1 To Me.dtSuma.Rows.Count - 1
        j(i) = j(i - 1) + 1 / enumFrecuencia
      Next

      For iRow As Integer = 0 To Me.dtSuma.Rows.Count - 1
        dr = dtCurva.NewRow
        dr.Item(0) = 1 / (j(iRow))
        dr.Item(0) = Math.Pow(100 / Me.dtSuma(iRow)(0), dr(0))
        'dr.Item(0) = Math.Pow(Math.Abs(Me.dtSuma.Rows(iRow).Item(0)), dr.Item(0))
        dr.Item(0) = CDbl(dr.Item(0) - 1)
        Me.dtCurva.Rows.Add(dr)
      Next

    Catch ex As Exception
      Throw New Exception("Error en clsBDT.Curva: " & vbCrLf & ex.Message & vbCrLf)
    End Try
  End Sub
  Private Sub ErrorCuadrado()
    Dim dr As DataRow
    Try
      dtError = New DataTable
      dtError.Columns.Add("Valor", Type.GetType("System.Double"))

      For iRow As Integer = 0 To Me.dtCurva.Rows.Count - 1
        dr = dtError.NewRow
        dr.Item(0) = CDbl(Math.Pow(Me.dtCurva.Rows(iRow).Item(0) - Me.dtYield.Rows(iRow).Item(0), 2))
        Me.dtError.Rows.Add(dr)
      Next

    Catch ex As Exception
      Throw New Exception("Error en clsBDT.ErrorCuadrado: " & vbCrLf & ex.Message & vbCrLf)
    End Try
  End Sub
  Private Sub CalculaArboldePrecios()
    Dim dblValorPrimerCupon As Decimal
    Dim dblCurva As Double()
    Dim intMaxCol As Integer
    Dim intMaxFil As Integer
    Try
      intMaxFil = Me.dblShortRateLat.GetLength(0)
      intMaxCol = Me.dblShortRateLat.GetLength(1)
      ReDim Me.dblArbolPrecios(intMaxFil, intMaxCol)
      ReDim dblCurva(Me.dtCurva.Rows.Count - 1)

      For iRow As Integer = 0 To Me.dtCurva.Rows.Count - 1
        dblCurva(iRow) = Me.dtCurva(iRow)(0)
      Next

      dblValorPrimerCupon = CDbl(Me.dtYield(0)(0) / 12 + 0.001 / 12) * 100

      For iColumna As Integer = 0 To intMaxCol
        Me.dblArbolPrecios(intMaxFil, iColumna) = 100 + dblValorPrimerCupon
      Next

      For iFila As Integer = intMaxFil - 1 To 0 Step -1
        For iColumna As Integer = iFila To 0 Step -1

          If Me.dtOpciones Is Nothing Then
            Me.dblArbolPrecios(iFila, iColumna) = 0.5 * (Me.dblArbolPrecios(iFila + 1, iColumna) + dblValorPrimerCupon) / (1 + (Me.dblShortRateLat(iFila, iColumna))) + _
                                                    0.5 * (Me.dblArbolPrecios(iFila + 1, iColumna + 1) + dblValorPrimerCupon) / (1 + (Me.dblShortRateLat(iFila, iColumna)))
          Else
            If Me.dtOpciones(iFila)(0) = 1 And iFila <> 0 Then 'Sí se puede cambiar.
              Me.dblArbolPrecios(iFila, iColumna) = Math.Min(0.5 * (Me.dblArbolPrecios(iFila + 1, iColumna) + dblValorPrimerCupon) / (1 + (Me.dblShortRateLat(iFila, iColumna))) + _
                                                    0.5 * (Me.dblArbolPrecios(iFila + 1, iColumna + 1) + dblValorPrimerCupon) / (1 + (Me.dblShortRateLat(iFila, iColumna))), 100 + dblValorPrimerCupon)
            Else 'No hay cambio
              Me.dblArbolPrecios(iFila, iColumna) = 0.5 * (Me.dblArbolPrecios(iFila + 1, iColumna) + dblValorPrimerCupon) / (1 + (Me.dblShortRateLat(iFila, iColumna))) + _
                                                    0.5 * (Me.dblArbolPrecios(iFila + 1, iColumna + 1) + dblValorPrimerCupon) / (1 + (Me.dblShortRateLat(iFila, iColumna)))
            End If
          End If
        Next
      Next

      Me.imprimeArboldePrecios()
    Catch ex As Exception
      Throw New Exception("Error en clsBDT.CalculaArboldePrecios: " & vbCrLf & _
                          ex.Message & vbCrLf)
    End Try
  End Sub
#End Region

#Region "Aux"
  Private Sub creaTablaOpciones()
    Dim intTamano As Integer
    Dim dr As DataRow
    Dim intVez As Integer
    Dim j() As Double
    Try
      ReDim j(Me.dtYield.Rows.Count)
      intTamano = Me.dtYield.Rows.Count
      Me.dtOpciones = New DataTable("Opciones")
      Me.dtOpciones.Columns.Add("Valor", Type.GetType("System.Int16"))


      intVez = 1
      For iRow As Integer = 0 To intTamano - 1
        dr = Me.dtOpciones.NewRow
        If iRow = 0 Then
          j(iRow) = 1 / 12
        Else
          j(iRow) = j(iRow - 1) + 1 / 12
        End If

        If Math.Round(j(iRow), 6) = Me.enumFrecuencia * intVez Then
          intVez += 1
          dr(0) = 1
        Else
          dr(0) = 0
        End If
        Me.dtOpciones.Rows.Add(dr)
      Next
    Catch ex As Exception
      Throw New Exception(" en clsBDT.creaTablaOpciones: " & vbCrLf & ex.Message & vbCrLf)
    End Try
  End Sub
  Private Sub calculaVolatilidad()
    Dim dblValores As Double()
    Dim dblVola As Double
    Dim dr As DataRow
    Try
      If Me.dtVol Is Nothing Then
        ReDim dblValores(Me.dtYield.Rows.Count - 1)

        For iRow As Integer = 0 To Me.dtYield.Rows.Count - 1
          dblValores(iRow) = CDbl(Me.dtYield.Rows(iRow).Item(0))
        Next

        'dblVola = Me.Varianza(dblValores)
        dblVola = 0.08
        dtVol = New DataTable
        dtVol.Columns.Add("Valores", Type.GetType("System.Double"))

        For iRow As Integer = 0 To Me.dtYield.Rows.Count - 1
          dr = Me.dtVol.NewRow
          dr.Item(0) = dblVola
          Me.dtVol.Rows.Add(dr)
        Next
      End If
    Catch ex As Exception
      Throw New Exception("Error en clsBD.calculaVolatilidad: " & ex.Message)
    End Try

  End Sub
  Private Function Varianza(ByVal Datos() As Double) As Double
    Dim i As Integer
    Dim n As Integer

    Dim nSumXQuad As Double
    Dim nSumX As Double
    Try
      nSumXQuad = 0
      nSumX = 0

      n = UBound(Datos)

      For i = 1 To n
        nSumXQuad = nSumXQuad + Datos(i) ^ 2
        nSumX = nSumX + Datos(i)
      Next i

      Varianza = ((n * nSumXQuad) - (nSumX) ^ 2) / (n * (n - 1))
    Catch ex As Exception
      Throw New Exception("Error en clsBDT.Varianza: " & ex.Message)
    End Try
  End Function
  Private Sub consigueDatosCurva()
    Dim strQuery As String
    Dim dtValores As DataTable
    Dim dtValoresInterpolados As DataTable
    Try
      strQuery = " SELECT 	VALOR, PLAZO  " & _
                    " FROM 	CURVAS " & _
                    " WHERE 	NOMBRE = '{0}' " & _
                    " 	AND	FECHA = '" & Format(DateTime.Today, VAR.Configuracion.LeeSetting("FormatoFecha")) & "' " & _
                    " 	ORDER BY 2 "

      Select Case Me.enumCurva
        Case enumCurvas.bonosm
          strQuery = String.Format(strQuery, "BONOSM")
        Case enumCurvas.irs
          strQuery = String.Format(strQuery, "IRS_Y")
        Case enumCurvas.libor
          strQuery = String.Format(strQuery, "LIBOR")
        Case Else
          Throw New Exception("Curva Incorrecta ")
      End Select

      dtValores = objBD.dsFromSQL(strQuery).Tables(0)
      dtValoresInterpolados = Me.interpolaValores(dtValores)

      Me.dtYield = New DataTable("Valores")
      Me.dtYield.Columns.Add("Valores", Type.GetType("System.Double"))

      For iValor As Integer = 0 To dtValoresInterpolados.Rows.Count - 1
        Me.dtYield(iValor)(0) = dtValoresInterpolados(iValor)(0)
      Next
    Catch ex As Exception
      Throw New Exception(" en clsBDT.consigueDatosCurva: " & vbCrLf & ex.Message & vbCrLf)
    End Try
  End Sub
  Private Function interpolaValores(ByVal dtValores As DataTable) As DataTable
    Dim intNumeroMeses As Integer
    Dim dtValoresInterpol As DataTable
    Dim dr As DataRow
    Dim intDuracionMes As Integer
    Dim dblValorUno As Double
    Dim dblValorDos As Double
    Dim intPlazoUno As Integer
    Dim intPlazoDos As Integer
    Dim iDatos As Integer
    Try

      If enumCurva = enumCurvas.irs Then
        intNumeroMeses = enumDura * 13
        intDuracionMes = 28
      Else
        intNumeroMeses = enumDura * 12
        intDuracionMes = 30
      End If

      'Borra los valores de la tabla que sean menores a 30 días o que tengan valor 0(No nos sirven)
      For iPlazo As Integer = 0 To dtValores.Rows.Count - 1
        If dtValores(iPlazo)(1) < 30 Or dtValores(iPlazo)(0) = 0 Then
          dtValores(iPlazo).Delete() 'No sirve
        End If
      Next

      dtValoresInterpol = New DataTable("Valores Interpolados")
      dtValoresInterpol.Columns.Add("Valores", Type.GetType("System.Double"))
      dtValoresInterpol.Columns.Add("Plazo", Type.GetType("System.Int16"))

      'Llenamos los plazos
      For iMeses As Integer = 0 To intNumeroMeses - 1
        dr = dtValoresInterpol.NewRow
        dr(1) = (iMeses + 1) * intDuracionMes
        dtValoresInterpol.Rows.Add(dr)
      Next

      'Metemos en la tabla los datos que ya tenemos
      For iMeses As Integer = 0 To intNumeroMeses - 1
        For iDatos = 0 To dtValores.Rows.Count - 1
          If dtValores(iDatos)(1) = dtValoresInterpol(iMeses)(1) Then
            dtValoresInterpol(iMeses)(1) = dtValores(iDatos)(0)
            Exit For
          End If
        Next
      Next

      'Interpolamos
      'Primero asignamos nuestros primeros Marcadores (Valor1, 2; Plazo1, 2)
      dblValorUno = dtValores(0)(0)
      dblValorDos = dtValores(1)(0)
      intPlazoUno = dtValores(0)(1)
      intPlazoDos = dtValores(1)(1)
      For iMeses As Integer = 0 To intNumeroMeses - 1
        iDatos = 1
        If dtValoresInterpol(iMeses)(0) = 0 Then
          dtValoresInterpol(iMeses)(0) = dblValorUno + ((dtValoresInterpol(iMeses)(1) - intPlazoUno) / _
                                                        (intPlazoDos - intPlazoUno)) * _
                                                        (dblValorDos - dblValorUno)
        Else
          'Cambiamos los "marcadores"
          If dtValoresInterpol(iMeses)(1) = dtValores(iDatos)(1) Then
            dblValorUno = dtValores(iDatos)(0)
            dblValorDos = dtValores(iDatos + 1)(0)
            intPlazoUno = dtValores(iDatos)(1)
            intPlazoDos = dtValores(iDatos + 1)(1)
            iDatos += 1
          End If
        End If
      Next

      Return dtValoresInterpol
    Catch ex As Exception
      Throw New Exception(" en clsBDT.interpolaValores: " & vbCrLf & _
                          ex.Message & vbCrLf)
    End Try
  End Function
#End Region

#Region "Imprimir"
  Private Sub imprimePrimeraMatriz()
    Dim strPath As String
    Dim sb As StringBuilder
    Try
      sb = New StringBuilder()
      strPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
      sb.AppendLine("Primera Matriz")
      sb.AppendLine("")
      For iColumna As Integer = 0 To Me.dblShortRateLat.GetLength(0) - 1
        For iFila As Integer = 0 To Me.dblShortRateLat.GetLength(1) - 1
          sb.Append(Me.dblShortRateLat(iColumna, iFila) & ", ")
        Next
        sb.AppendLine()
      Next

      Using outfile As New StreamWriter(strPath & "\Var\BDT\Resultados\PrimeraMatriz.csv")
        outfile.Write(sb.ToString())
      End Using

    Catch ex As Exception
      Throw New Exception(" en clsBDT.imprimePrimeraMatriz: " & vbCrLf & ex.Message & vbCrLf)
    End Try
  End Sub
  Private Sub imprimeSegundaMatriz()
    Dim strPath As String
    Dim sb As StringBuilder
    Try
      sb = New StringBuilder()
      strPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
      sb.AppendLine("Segunda Matriz ")
      sb.AppendLine("")
      For iColumna As Integer = 0 To Me.dblValorCupon.GetLength(0) - 1
        For iFila As Integer = 0 To Me.dblValorCupon.GetLength(1) - 1
          sb.Append(Me.dblValorCupon(iColumna, iFila) & ", ")
        Next
        sb.AppendLine()
      Next

      Using outfile As New StreamWriter(strPath & "\Var\BDT\Resultados\SegundaMatriz.csv")
        outfile.Write(sb.ToString())
      End Using

    Catch ex As Exception
      Throw New Exception(" en clsBDT.imprimeSegundaMatriz: " & vbCrLf & ex.Message & vbCrLf)
    End Try
  End Sub
  Private Sub imprimeTerceraMatriz()
    Dim strPath As String
    Dim sb As StringBuilder
    Try
      sb = New StringBuilder()
      strPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
      sb.AppendLine("Tercera Matriz ")
      sb.AppendLine("")
      For iColumna As Integer = 0 To Me.dblEntramadoBDT.GetLength(0) - 1
        For iFila As Integer = 0 To Me.dblEntramadoBDT.GetLength(1) - 1
          sb.Append(Me.dblEntramadoBDT(iColumna, iFila) & ", ")
        Next
        sb.AppendLine()
      Next

      Using outfile As New StreamWriter(strPath & "\Var\BDT\Resultados\TerceraMatriz.csv")
        outfile.Write(sb.ToString())
      End Using

    Catch ex As Exception
      Throw New Exception(" en clsBDT.imprimeTerceraMatriz: " & vbCrLf & ex.Message & vbCrLf)
    End Try
  End Sub
  Private Sub imprimeCalibracion(ByVal dblResultado As Double(), ByVal dblSumaErrorCuadrado As Double)
    Dim strPath As String
    Dim sb As StringBuilder
    Try
      sb = New StringBuilder()
      strPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
      sb.AppendLine("Vector Resultado de la Calibracion: ")
      sb.AppendLine("")
      For Each dblValor As Double In dblResultado
        sb.AppendLine(dblValor)
      Next

      sb.AppendLine("Suma del Error Cuadrado: " & vbCrLf & dblSumaErrorCuadrado)

      Using outfile As New StreamWriter(strPath & "\Var\BDT\Resultados\Calibracion.txt")
        outfile.Write(sb.ToString())
      End Using

    Catch ex As Exception
      Throw New Exception(" en clsBDT.imprimeCalibracion: " & vbCrLf & ex.Message & vbCrLf)
    End Try
  End Sub
  Private Sub imprimeArboldePrecios()
    Dim strPath As String
    Dim sb As StringBuilder
    Try
      sb = New StringBuilder()
      strPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
      sb.AppendLine("Árbol de Precios")
      sb.AppendLine("")
      For iColumna As Integer = 0 To Me.dblArbolPrecios.GetLength(0) - 1
        For iFila As Integer = 0 To Me.dblArbolPrecios.GetLength(1) - 1
          sb.Append(Me.dblArbolPrecios(iColumna, iFila) & ", ")
        Next
        sb.AppendLine()
      Next

      Using outfile As New StreamWriter(strPath & "\Var\BDT\Resultados\ArboldePrecios.csv")
        outfile.Write(sb.ToString())
      End Using

    Catch ex As Exception
      Throw New Exception(" en clsBDT.imprimeArboldePrecios: " & vbCrLf & ex.Message & vbCrLf)
    End Try
  End Sub
#End Region
End Class
