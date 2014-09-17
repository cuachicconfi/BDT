<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
    Me.cmdPrueba60 = New System.Windows.Forms.Button
    Me.cmdPrueba30 = New System.Windows.Forms.Button
    Me.cmdPruebaSinDatos = New System.Windows.Forms.Button
    Me.SuspendLayout()
    '
    'cmdPrueba60
    '
    Me.cmdPrueba60.Location = New System.Drawing.Point(102, 34)
    Me.cmdPrueba60.Name = "cmdPrueba60"
    Me.cmdPrueba60.Size = New System.Drawing.Size(75, 23)
    Me.cmdPrueba60.TabIndex = 2
    Me.cmdPrueba60.Text = "Prueba 60"
    Me.cmdPrueba60.UseVisualStyleBackColor = True
    '
    'cmdPrueba30
    '
    Me.cmdPrueba30.Location = New System.Drawing.Point(221, 34)
    Me.cmdPrueba30.Name = "cmdPrueba30"
    Me.cmdPrueba30.Size = New System.Drawing.Size(75, 23)
    Me.cmdPrueba30.TabIndex = 3
    Me.cmdPrueba30.Text = "Prueba 30"
    Me.cmdPrueba30.UseVisualStyleBackColor = True
    '
    'cmdPruebaSinDatos
    '
    Me.cmdPruebaSinDatos.Location = New System.Drawing.Point(138, 76)
    Me.cmdPruebaSinDatos.Name = "cmdPruebaSinDatos"
    Me.cmdPruebaSinDatos.Size = New System.Drawing.Size(122, 23)
    Me.cmdPruebaSinDatos.TabIndex = 4
    Me.cmdPruebaSinDatos.Text = "Prueba sin Datos"
    Me.cmdPruebaSinDatos.UseVisualStyleBackColor = True
    '
    'Form1
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(399, 133)
    Me.Controls.Add(Me.cmdPruebaSinDatos)
    Me.Controls.Add(Me.cmdPrueba30)
    Me.Controls.Add(Me.cmdPrueba60)
    Me.Name = "Form1"
    Me.Text = "Ejemplo BDT"
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents cmdPrueba60 As System.Windows.Forms.Button
  Friend WithEvents cmdPrueba30 As System.Windows.Forms.Button
  Friend WithEvents cmdPruebaSinDatos As System.Windows.Forms.Button

End Class
