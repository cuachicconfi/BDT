Public Class Form1
  Private objBDT As clsBDT
  Private objBD As clsBD

  Private Sub cmdPrueba60_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrueba60.Click

    Dim dtDatos As DataTable
    'Dim dtCondiciones As DataTable
    Dim dtVolatilidad As DataTable
    Dim dr As DataRow
    Dim dblDatos As Double()
    'Dim dblCondiciones As Double()
    Dim dblVolatilidad As Double()
    Dim dblRes As Double()
    Try
      dblDatos = New Double(59) {4.3001, 4.2894, 4.2809, 4.296, 4.299, 4.2973, 4.2887, 4.2873, 4.3771, 4.4726, 4.5069, 4.5412, 4.6898, 4.7623, 4.8172, 4.8719, 4.9267, 4.9815, 5.0362, 5.0909, 5.1456, 5.2002, 5.2549, 5.3095, 5.4803, 5.5925, 5.66, 5.7275, 5.7949, 5.8624, 5.9297, 5.9971, 6.0644, 6.1317, 6.199, 6.2662, 6.4246, 6.567, 6.6427, 6.7183, 6.7939, 6.8695, 6.945, 7.0205, 7.0959, 7.1713, 7.2467, 7.322, 7.3876, 7.4493, 7.5236, 7.5979, 7.6721, 7.7463, 7.8205, 7.9687, 8.0427, 8.1167, 8.1907, 8.118}
      ' dblDatos = New Double(59) {0.043001, 0.042894, 0.042809, 0.04296, 0.04299, 0.042973, 0.042887, 0.042873, 0.043771, 0.044726, 0.045069, 0.045412, 0.046898, 0.047623, 0.048172, 0.048719, 0.049267, 0.049815, 0.050362, 0.050909, 0.051456, 0.052002, 0.052549, 0.053095, 0.054803, 0.055925, 0.0566, 0.057275, 0.057949, 0.058624, 0.059297, 0.059971, 0.060644, 0.061317, 0.06199, 0.062662, 0.064246, 0.06567, 0.066427, 0.067183, 0.067939, 0.068695, 0.06945, 0.070205, 0.070959, 0.071713, 0.072467, 0.07322, 0.073876, 0.074493, 0.075236, 0.075979, 0.076721, 0.077463, 0.078205, 0.079687, 0.080427, 0.081167, 0.081907, 0.08118}
      'dblCondiciones = New Double(59) {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1}
      dblVolatilidad = New Double(59) {0.09, 0.092827298, 0.095752089, 0.09867688, 0.101601671, 0.104526462, 0.107451253, 0.110376045, 0.113300836, 0.116225627, 0.119150418, 0.122075209, 0.125, 0.128333333, 0.131666667, 0.135, 0.138333333, 0.141666667, 0.145, 0.148333333, 0.151666667, 0.155, 0.158333333, 0.161666667, 0.165, 0.167916667, 0.170833333, 0.17375, 0.176666667, 0.179583333, 0.1825, 0.185416667, 0.188333333, 0.19125, 0.194166667, 0.197083333, 0.2, 0.201666667, 0.203333333, 0.205, 0.206666667, 0.208333333, 0.21, 0.211666667, 0.213333333, 0.215, 0.216666667, 0.218333333, 0.22, 0.221666667, 0.223333333, 0.225, 0.226666667, 0.228333333, 0.23, 0.233333333, 0.235, 0.236666667, 0.238333333, 0.24}
      dtDatos = New DataTable
      'dtCondiciones = New DataTable
      dtVolatilidad = New DataTable
      dtDatos.Columns.Add("Valores", Type.GetType("System.Double"))
      'dtCondiciones.Columns.Add("Valores", Type.GetType("System.Int16"))
      dtVolatilidad.Columns.Add("Valores", Type.GetType("System.Double"))
      For iMatriz As Integer = 0 To dblDatos.Length - 1
        dr = dtDatos.NewRow
        dr(0) = dblDatos(iMatriz)
        dtDatos.Rows.Add(dr)

        'dr = dtCondiciones.NewRow
        'dr(0) = dblCondiciones(iMatriz)
        'dtCondiciones.Rows.Add(dr)

        dr = dtVolatilidad.NewRow
        dr(0) = dblVolatilidad(iMatriz)
        dtVolatilidad.Rows.Add(dr)
      Next

      objBDT = New clsBDT(dtDatos, clsBDT.enumFrecuenciaDePago.Mensual, dtVolatilidad)

      MsgBox(objBDT.calculaPrecio())
    Catch ex As Exception
      MsgBox("Error en BDT.cmdPrueba3_Click: " & vbCrLf & _
             ex.Message & vbCrLf)
    End Try
  End Sub

  Private Sub cmdPrueba30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrueba30.Click
    Dim dtDatos As DataTable
    'Dim dtCondiciones As DataTable
    Dim dtVolatilidad As DataTable
    Dim dr As DataRow
    Dim dblDatos As Double()
    'Dim dblCondiciones As Double()
    Dim dblVolatilidad As Double()
    Dim dblRes As Double()
    Try
      dblDatos = New Double(29) {0.043001, 0.042894, 0.042809, 0.04296, 0.04299, 0.042973, 0.042887, 0.042873, 0.043771, 0.044726, 0.045069, 0.045412, 0.046898, 0.047623, 0.048172, 0.048719, 0.049267, 0.049815, 0.050362, 0.050909, 0.051456, 0.052002, 0.052549, 0.053095, 0.054803, 0.055925, 0.0566, 0.057275, 0.057949, 0.058624}
      'dblCondiciones = New Double(29) {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0}
      dblVolatilidad = New Double(29) {9, 9.282729805, 9.575208914, 9.867688022, 10.16016713, 10.45264624, 10.74512535, 11.03760446, 11.33008357, 11.62256267, 11.91504178, 12.20752089, 12.5, 12.83333333, 13.16666667, 13.5, 13.83333333, 14.16666667, 14.5, 14.83333333, 15.16666667, 15.5, 15.83333333, 16.16666667, 16.5, 16.79166667, 17.08333333, 17.375, 17.66666667, 17.95833333}

      dtDatos = New DataTable
      'dtCondiciones = New DataTable
      dtVolatilidad = New DataTable
      dtDatos.Columns.Add("Valores", Type.GetType("System.Double"))
      'dtCondiciones.Columns.Add("Valores", Type.GetType("System.Int16"))
      dtVolatilidad.Columns.Add("Valores", Type.GetType("System.Double"))
      For iMatriz As Integer = 0 To dblDatos.Length - 1
        dr = dtDatos.NewRow
        dr(0) = dblDatos(iMatriz)
        dtDatos.Rows.Add(dr)

        'dr = dtCondiciones.NewRow
        'dr(0) = dblCondiciones(iMatriz)
        'dtCondiciones.Rows.Add(dr)

        dr = dtVolatilidad.NewRow
        dr(0) = dblVolatilidad(iMatriz)
        dtVolatilidad.Rows.Add(dr)
      Next

      objBDT = New clsBDT(dtDatos, clsBDT.enumFrecuenciaDePago.Mensual, dtVolatilidad)
      dblRes = objBDT.calibracion()
    Catch ex As Exception
      MsgBox("Error en BDT.cmdPrueba3_Click: " & vbCrLf & _
             ex.Message & vbCrLf)
    End Try
  End Sub

  Private Sub cmdPruebaSinDatos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPruebaSinDatos.Click

    'Dim dblRes As Double()
    Try
      
      objBDT = New clsBDT(clsBDT.enumCurvas.bonosm, clsBDT.enumDuracion.cincoAños, clsBDT.enumFrecuenciaDePago.Mensual)

      MsgBox(objBDT.calculaPrecio())
    Catch ex As Exception
      MsgBox("Error en BDT.cmdPruebaSinDatos_Click: " & vbCrLf & _
             ex.Message & vbCrLf)
    End Try
  End Sub
End Class
