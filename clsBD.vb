Imports System.Data.OleDb
Public Class clsBD
  Private strConnectionString As String
  Private oledbDA As OleDbDataAdapter
  Private oledbConn As OleDbConnection  
  Public Sub New(ByVal strConnectionString As String)
    Me.strConnectionString = strConnectionString
    Me.oledbConn = New OleDbConnection(Me.strConnectionString)
  End Sub
  Public Function dsFromSQL(ByVal strQuery As String) As DataSet
    Dim dsResultado As DataSet
    Try
      dsResultado = New DataSet
      Me.oledbDA = New OleDb.OleDbDataAdapter(strQuery, Me.strConnectionString)
      oledbDA.Fill(dsResultado, strQuery)
      Return dsResultado
    Catch ex As Exception
      Throw New Exception("Error en clsBD.dsFromSQL:" & vbCrLf & ex.Message & vbCrLf)
    End Try
  End Function
  Public Function dtWithSchemmaFromSQL(ByVal strQuery As String) As DataTable
    Dim dtResultado As DataTable
    Try
      Me.oledbDA = New OleDb.OleDbDataAdapter(strQuery, Me.strConnectionString)
      dtResultado = dsFromSQL(strQuery).Tables(0)
      oledbDA.FillSchema(dtResultado, SchemaType.Mapped)
      Return dtResultado
    Catch ex As Exception
      Throw New Exception("Error en clsBD.dtWithSchemmaFromSQL:" & vbCrLf & ex.Message & vbCrLf)
    End Try
  End Function
  Public Function distinctRows(ByVal dt As DataTable, ByVal strColumna As String) As SortedList
    Dim slAux As SortedList
    Try
      slAux = New SortedList
      If dt.Columns.Contains(strColumna) = False Then
        Throw New Exception("Error en Interfaces.BD.distinctRows: " & _
                            "La tabla no contiene la columna ('" & strColumna & "')" & vbCrLf)
      End If
      For iRow As Integer = 0 To dt.Rows.Count - 1
        If slAux.ContainsKey(dt.Rows(iRow).Item(strColumna)) = False Then
          slAux.Add(dt.Rows(iRow).Item(strColumna), dt.Rows(iRow).Item(strColumna))
        End If
      Next
      Return slAux
    Catch ex As Exception
      Throw New Exception("Error en countDistinctRows:" & vbCrLf & ex.Message & vbCrLf)
    End Try
  End Function
  Public Function distinctRows(ByVal dv As DataView, ByVal strColumna As String) As SortedList
    Try
      Return distinctRows(Me.dataViewToDataTable(dv), strColumna)
    Catch ex As Exception
      Throw New Exception("Error en countDistinctRows:" & vbCrLf & ex.Message & vbCrLf)
    End Try
  End Function
  Public Function executeLotOfQuerys(ByVal alQuerys As ArrayList) As Boolean
    Dim cmd As OleDbCommand
    Dim bool As Boolean
    Dim transac As OleDbTransaction
    Dim strInsertQueryOra As String
    Dim strBigQuery As String
    Dim strBuilder As System.Text.StringBuilder
    Try
      If Me.oledbConn.State = ConnectionState.Closed Then
        Me.oledbConn.Open()
      End If
      transac = Me.oledbConn.BeginTransaction(IsolationLevel.Serializable)
      cmd = New OleDbCommand
      strBuilder = New System.Text.StringBuilder
      cmd.Connection = Me.oledbConn
      cmd.Transaction = transac
      If Not VAR.Configuracion.LeeSetting("FormatoFechaDBMS").Contains("{0}") Then
        For i As Integer = 0 To alQuerys.Count - 1
          strBuilder.Append(alQuerys.Item(i) & ";" & vbCrLf)
        Next
        strBigQuery = strBuilder.ToString
        cmd.CommandText = strBigQuery
        cmd.ExecuteNonQuery()
      Else
        For i As Integer = 0 To alQuerys.Count - 1
          strInsertQueryOra = alQuerys.Item(i)
          cmd.CommandText = strInsertQueryOra
          cmd.ExecuteNonQuery()
        Next
      End If
      transac.Commit()
      bool = True
    Catch ex As OleDbException
      If Not transac Is Nothing Then
        transac.Rollback()
      End If
      Throw New Exception("Error en executeLotOfQuerys" & vbCrLf & ex.Message)
      bool = False
    Finally
      Me.oledbConn.Close()
    End Try
    Return bool
  End Function
  Public Function regresaFechaHoraServidorSQL() As DateTime
    Dim strSQLFechaHora As String
    Dim dtFechaHora As DataTable
    Dim datResp As DateTime
    Dim strSrvFechaHoraDBMS As String
    Try
      strSQLFechaHora = "SELECT GETDATE() as fechahora"
      strSrvFechaHoraDBMS = VAR.Configuracion.LeeSetting("FormatoFechaServidorDBMS")
      If strSrvFechaHoraDBMS.Contains("{0}") Then
        strSrvFechaHoraDBMS = String.Format(strSrvFechaHoraDBMS, "fechahora")
        strSQLFechaHora = "SELECT " & strSrvFechaHoraDBMS
      End If
      dtFechaHora = dsFromSQL(strSQLFechaHora).Tables(0)
      datResp = CDate(dtFechaHora.Rows(0).Item("fechahora"))
      Return datResp
    Catch ex As Exception
      Throw New Exception("Error en regresaFechaHoraServidorSQL: " & vbCrLf & ex.Message & vbCrLf)
    End Try
  End Function
  Public Function dataViewToDataTable(ByVal dv As DataView) As DataTable
    Dim dtResultado As DataTable
    Dim dr As DataRow
    Try
      dtResultado = dv.Table.Clone
      If dv.Count > 0 Then
        For iRow As Integer = 0 To dv.Count - 1
          dr = dtResultado.NewRow
          For iCol As Integer = 0 To dtResultado.Columns.Count - 1
            dr.Item(iCol) = dv.Item(iRow).Item(iCol)
          Next
          dtResultado.Rows.Add(dr)
        Next
      End If
      Return dtResultado
    Catch ex As Exception
      Throw New Exception("Error en Interfaces.BD.dataViewToDataTable:" & vbCrLf & ex.Message & vbCrLf)
    End Try
  End Function
End Class

