<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Bienvenida
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Bienvenida))
        Me.lblSaludo = New System.Windows.Forms.Label()
        Me.lstArchivos = New System.Windows.Forms.ListView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.btnAbrirArchivo = New System.Windows.Forms.Button()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.BtnLibro = New System.Windows.Forms.Button()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblSaludo
        '
        Me.lblSaludo.AutoSize = True
        Me.lblSaludo.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSaludo.Location = New System.Drawing.Point(185, 18)
        Me.lblSaludo.Name = "lblSaludo"
        Me.lblSaludo.Size = New System.Drawing.Size(16, 24)
        Me.lblSaludo.TabIndex = 2
        Me.lblSaludo.Text = "."
        '
        'lstArchivos
        '
        Me.lstArchivos.HideSelection = False
        Me.lstArchivos.Location = New System.Drawing.Point(121, 287)
        Me.lstArchivos.Margin = New System.Windows.Forms.Padding(2)
        Me.lstArchivos.Name = "lstArchivos"
        Me.lstArchivos.Size = New System.Drawing.Size(939, 282)
        Me.lstArchivos.TabIndex = 4
        Me.lstArchivos.UseCompatibleStateImageBehavior = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(124, 261)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(218, 24)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Recientes     Favoritos"
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.White
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.Image = Global.Proyecto_Clase.My.Resources.Resources.fomulac1
        Me.Button4.Location = New System.Drawing.Point(419, 34)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(224, 211)
        Me.Button4.TabIndex = 10
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.White
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Image = Global.Proyecto_Clase.My.Resources.Resources.Libro_Blanco
        Me.Button1.Location = New System.Drawing.Point(142, 34)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(243, 211)
        Me.Button1.TabIndex = 7
        Me.Button1.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.Proyecto_Clase.My.Resources.Resources.panelb
        Me.PictureBox1.Location = New System.Drawing.Point(-7, 346)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(124, 687)
        Me.PictureBox1.TabIndex = 6
        Me.PictureBox1.TabStop = False
        '
        'btnAbrirArchivo
        '
        Me.btnAbrirArchivo.BackColor = System.Drawing.Color.White
        Me.btnAbrirArchivo.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAbrirArchivo.Image = Global.Proyecto_Clase.My.Resources.Resources.openb
        Me.btnAbrirArchivo.Location = New System.Drawing.Point(-6, 114)
        Me.btnAbrirArchivo.Name = "btnAbrirArchivo"
        Me.btnAbrirArchivo.Size = New System.Drawing.Size(124, 118)
        Me.btnAbrirArchivo.TabIndex = 5
        Me.btnAbrirArchivo.UseVisualStyleBackColor = False
        '
        'btnSalir
        '
        Me.btnSalir.BackColor = System.Drawing.Color.White
        Me.btnSalir.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSalir.Image = Global.Proyecto_Clase.My.Resources.Resources.closeB
        Me.btnSalir.Location = New System.Drawing.Point(-6, 232)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(124, 116)
        Me.btnSalir.TabIndex = 3
        Me.btnSalir.Text = "Cerrar"
        Me.btnSalir.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnSalir.UseVisualStyleBackColor = False
        '
        'BtnLibro
        '
        Me.BtnLibro.BackColor = System.Drawing.Color.White
        Me.BtnLibro.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnLibro.Image = Global.Proyecto_Clase.My.Resources.Resources.newb
        Me.BtnLibro.Location = New System.Drawing.Point(-6, -1)
        Me.BtnLibro.Name = "BtnLibro"
        Me.BtnLibro.Size = New System.Drawing.Size(124, 118)
        Me.BtnLibro.TabIndex = 0
        Me.BtnLibro.UseVisualStyleBackColor = False
        '
        'Bienvenida
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SlateGray
        Me.ClientSize = New System.Drawing.Size(1071, 692)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.btnAbrirArchivo)
        Me.Controls.Add(Me.lstArchivos)
        Me.Controls.Add(Me.btnSalir)
        Me.Controls.Add(Me.lblSaludo)
        Me.Controls.Add(Me.BtnLibro)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Bienvenida"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Bienvenida"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BtnLibro As Button
    Friend WithEvents lblSaludo As Label
    Friend WithEvents btnSalir As Button
    Friend WithEvents lstArchivos As ListView
    Friend WithEvents btnAbrirArchivo As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Label1 As Label
End Class
