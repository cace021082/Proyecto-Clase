Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Bienvenida
    Private Sub BtnLibro_Click(sender As Object, e As EventArgs) Handles BtnLibro.Click
        Dim form1 As New Form1()
        form1.Show()

        ' Oculta el formulario Bienvenida en lugar de cerrarlo
        Me.Hide()
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Application.Exit()
    End Sub

    Private Sub lblNombre_Click(sender As Object, e As EventArgs) Handles lblSaludo.Click

    End Sub

    Private Sub Bienvenida_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblSaludo.Text = ObtenerSaludo()
        ' Configurar ListView
        lstArchivos.View = View.Details
        lstArchivos.FullRowSelect = True
        lstArchivos.GridLines = False
        lstArchivos.Columns.Add("Nombre", 300, HorizontalAlignment.Left)
        lstArchivos.Columns.Add("Ubicación", 300, HorizontalAlignment.Left)
        lstArchivos.Columns.Add("Fecha Modificación", 150, HorizontalAlignment.Left)

        ' Cargar historial de archivos abiertos
        CargarHistorial()
    End Sub

    Private Function ObtenerSaludo() As String
        Dim hora As Integer = DateTime.Now.Hour
        If hora < 12 Then
            Return "Buenos dias"
        ElseIf hora < 18 Then
            Return "Buenas Tardes"
        Else
            Return "Buenas Noches"
        End If
    End Function

    Private Sub CargarArchivosRecientes()
        ' Simulación de archivos recientes
        Dim archivos As New List(Of String) From {
            "Informe Grupo4 - lunes. at 18:12",
            "PRUEBA DE PROYECTO - 21 de julio",
            "sistema-frances - 21 de julio"
        }

        For Each archivo In archivos
            lstArchivos.Items.Add(archivo)
        Next
    End Sub

    Private Sub btnAbrirArchivo_Click(sender As Object, e As EventArgs) Handles btnAbrirArchivo.Click
        ' Abrir un cuadro de diálogo para seleccionar un archivo
        Dim ofd As New OpenFileDialog()
        ofd.Filter = "Archivos Excel|*.xlsx;*.xlsm;*.xlsb;*.xls"
        ofd.Title = "Seleccionar un archivo de Excel"

        If ofd.ShowDialog() = DialogResult.OK Then
            ' Verificar si Form1 ya está abierto
            Dim frm As Form1 = Nothing

            For Each frmAbierto As Form In Application.OpenForms
                If TypeOf frmAbierto Is Form1 Then
                    frm = CType(frmAbierto, Form1)
                    Exit For
                End If
            Next

            ' Si Form1 no está abierto, crearlo
            If frm Is Nothing Then
                frm = New Form1()
                frm.Show()
                Me.Hide()
            End If

            ' Llamar al método público para abrir el archivo
            frm.AbrirArchivoDesdeOtroFormulario(ofd.FileName)
            AgregarArchivoALista(ofd.FileName)
            GuardarHistorial(ofd.FileName)
        End If
    End Sub

    Public Sub AgregarArchivoALista(rutaArchivo As String)
        Dim nombre As String = Path.GetFileName(rutaArchivo)
        Dim ubicacion As String = Path.GetDirectoryName(rutaArchivo)
        Dim fechaMod As String = File.GetLastWriteTime(rutaArchivo).ToString("dd/MM/yyyy HH:mm")

        ' Agregar al ListView
        Dim listItem As New ListViewItem(nombre)
        listItem.SubItems.Add(ubicacion)
        listItem.SubItems.Add(fechaMod)
        lstArchivos.Items.Add(listItem)
    End Sub

    Public Sub GuardarHistorial(rutaArchivo As String)
        Dim historialPath As String = "historial.txt"

        ' Leer todas las líneas del historial si existe
        Dim lineas As New List(Of String)
        If File.Exists(historialPath) Then
            lineas = File.ReadAllLines(historialPath).ToList()
        End If

        ' Obtener el nombre del archivo (sin ruta)
        Dim nombreArchivo As String = Path.GetFileName(rutaArchivo)

        ' Eliminar cualquier entrada anterior con el mismo nombre
        lineas = lineas.Where(Function(linea) Path.GetFileName(linea) <> nombreArchivo).ToList()

        ' Agregar el nuevo archivo al inicio de la lista (para que el más reciente quede primero)
        lineas.Insert(0, rutaArchivo)

        ' Guardar la lista actualizada en el archivo
        File.WriteAllLines(historialPath, lineas)
    End Sub

    Private Sub CargarHistorial()
        Dim historialPath As String = "historial.txt"

        ' Limpiar la lista antes de cargar
        lstArchivos.Items.Clear()

        If File.Exists(historialPath) Then
            Dim lineas As String() = File.ReadAllLines(historialPath)

            ' Cargar en orden de arriba hacia abajo (el más reciente primero)
            For Each rutaArchivo In lineas
                AgregarArchivoALista(rutaArchivo)
            Next
        End If
    End Sub

    Private Sub lstArchivos_DoubleClick(sender As Object, e As EventArgs) Handles lstArchivos.DoubleClick
        ' Verificar si hay un elemento seleccionado en el ListView
        If lstArchivos.SelectedItems.Count = 0 Then
            MessageBox.Show("Selecciona un archivo del historial.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        ' Obtener la ruta del archivo desde el ListView
        Dim selectedItem As ListViewItem = lstArchivos.SelectedItems(0)
        Dim filePath As String = selectedItem.SubItems(1).Text & "\" & selectedItem.SubItems(0).Text ' Asegúrate de que la ruta está en la columna correcta

        ' Verificar si el archivo existe
        If Not File.Exists(filePath) Then
            MessageBox.Show("El archivo no existe.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' Ocultar Bienvenida y abrir Form1 pasando la ruta
        Dim frm As New Form1()
        frm.rutaArchivo = filePath ' Pasar la ruta del archivo
        frm.Show()
        AgregarArchivoALista(filePath)
        GuardarHistorial(filePath)
        Me.Hide()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim form1 As New Form1()
        form1.Show()

        ' Oculta el formulario Bienvenida en lugar de cerrarlo
        Me.Hide()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim pdfPath As String = "C:\Users\PJL\Desktop\II Trimestre 2025\Teoria de la Computacion\Proyecto\Manual de usuario.pdf"
        If System.IO.File.Exists(pdfPath) Then
            Try
                Process.Start(pdfPath)
            Catch ex As Exception
                MessageBox.Show("No se pudo abrir el archivo PDF. " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            MessageBox.Show("El archivo PDF no se encontró.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub lstArchivos_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstArchivos.SelectedIndexChanged

    End Sub
End Class