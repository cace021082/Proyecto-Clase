Public Class frmprueba2

    Private Async Sub frmprueba2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Aquí puedes agregar cualquier código adicional que necesites para configurar tu ventana de carga.

        ' Esperar 1 segundos
        Await Task.Delay(1000)

        ' Abrir el formulario Bienvenida
        Dim bienvenida As New Bienvenida()
        bienvenida.Show()

        ' Cerrar la ventana de carga
        Me.Hide()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub
End Class