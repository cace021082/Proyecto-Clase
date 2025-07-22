Imports System.Data.OleDb
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.Logging
Imports OfficeOpenXml
Imports OfficeOpenXml.FormulaParsing


Public Class Form1
    Private currentFilePath As String

    Private formulas(,) As String ' Arreglo para almacenar fórmulas

    Private editingCell As DataGridViewCell = Nothing

    Public rutaArchivo As String ' Variable para recibir la ruta
    Public Sub columnas()

        Dim letras As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

        For i As Integer = 0 To 25
            Dim columnaNombre As String = letras(i)
            dgvhoja1.Columns.Add(columnaNombre, columnaNombre)
        Next


    End Sub
    Public Sub filas()

        ' Añadir filas y actualizar el índice de cada una
        For i As Integer = 0 To 100
            Dim rowIndex As Integer = dgvhoja1.Rows.Add()

        Next


    End Sub


    'LAS OPERACIONES ARITMETICAS'

    Public Sub Operaciones()
        Dim operacion As String
        operacion = cbxOperaciones.Text

        Dim resultado As Double = If(operacion = "Multiplicación", 1, 0)
        Dim valorCelda As Double
        Dim seleccionadas As String = "Celdas seleccionadas:" & vbCrLf
        Dim contador As Integer = 0

        For Each celda As DataGridViewCell In dgvhoja1.SelectedCells
            seleccionadas &= "Fila: " & celda.RowIndex & " Columna: " & celda.ColumnIndex & vbCrLf

            ' Intentar convertir el valor de la celda a un número
            If Double.TryParse(celda.Value.ToString(), valorCelda) Then
                contador += 1
                Select Case operacion
                    Case "Suma"
                        resultado += valorCelda
                    Case "Resta"
                        If contador = 1 Then
                            resultado = valorCelda
                        Else
                            resultado -= valorCelda
                        End If
                    Case "Multiplicación"
                        resultado *= valorCelda
                    Case "Promedio"
                        resultado += valorCelda

                End Select
            Else
                MessageBox.Show("La celda en Fila: " & celda.RowIndex & " Columna: " & celda.ColumnIndex & " no contiene un valor numérico.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
        Next



        If operacion = "Promedio" And contador > 0 Then
            resultado /= contador

        End If

        If operacion = "Divicion" And contador > 0 Then

            If dgvhoja1.SelectedCells.Count <> 2 Then
                MessageBox.Show("Por favor, selecciona exactamente dos celdas para realizar la división.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim valor1 As Double
            Dim valor2 As Double
            Dim celda1 As DataGridViewCell = dgvhoja1.SelectedCells(0)
            Dim celda2 As DataGridViewCell = dgvhoja1.SelectedCells(1)

            seleccionadas &= "Fila: " & celda1.RowIndex & " Columna: " & celda1.ColumnIndex & vbCrLf
            seleccionadas &= "Fila: " & celda2.RowIndex & " Columna: " & celda2.ColumnIndex & vbCrLf


            ' Intentar convertir el valor de las celdas a números


            If Not Double.TryParse(celda1.Value.ToString(), valor1) Then
                MessageBox.Show("La celda en Fila: " & celda1.RowIndex & " Columna: " & celda1.ColumnIndex & " no contiene un valor numérico.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If Not Double.TryParse(celda2.Value.ToString(), valor2) Then
                MessageBox.Show("La celda en Fila: " & celda2.RowIndex & " Columna: " & celda2.ColumnIndex & " no contiene un valor numérico.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If valor2 = 0 Then
                MessageBox.Show("La división por cero no está permitida.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim division As Double = valor1 / valor2
            resultado = division
        End If

        seleccionadas &= vbCrLf & "Resultado de " & operacion & ": " & resultado
        MessageBox.Show(seleccionadas)

    End Sub

    Private Function ReplaceCellReferences(match As Match) As String
        Dim cellName As String = match.Value
        Dim cellValue As Double = GetCellValue(cellName)
        Return cellValue.ToString()
    End Function

    Private Function GetCellValue(cellName As String) As Double

        ' Convierte el nombre de la celda (e.g., "A1") en índices de fila y columna


        Try
            Dim columnName As String = cellName.Substring(0, 1)
            Dim rowIndex As Integer = Integer.Parse(cellName.Substring(1)) - 1

            Dim columnIndex As Integer = -1
            For Each column As DataGridViewColumn In dgvhoja1.Columns
                If column.Name = columnName Then
                    columnIndex = column.Index
                    Exit For
                End If
            Next

            If columnIndex = -1 OrElse rowIndex < 0 OrElse rowIndex >= dgvhoja1.RowCount Then
                Throw New Exception($"Referencia de celda no válida: {cellName}.")
            End If

            Dim cell As DataGridViewCell = dgvhoja1.Rows(rowIndex).Cells(columnIndex)

            If cell.Value IsNot Nothing Then
                Dim result As Double
                If Double.TryParse(cell.Value.ToString(), result) Then
                    Return result
                Else
                    Throw New Exception($"Valor no válido en la celda {cellName}: {cell.Value}.")
                End If
            Else
                Throw New Exception($"La celda {cellName} está vacía.")
            End If
        Catch ex As Exception
            MessageBox.Show("Error al obtener el valor de la celda: " & ex.Message)
            Return 0
        End Try
    End Function


    ' ´FUNCION PARA EVALUAR LA FORMULA SUMA

    Private Function EvaluateSumFormula(formula As String) As Double
        Dim pattern As String = "=SUMA\(([A-Z]+\d+):([A-Z]+\d+)\)"
        Dim match As Match = Regex.Match(formula, pattern, RegexOptions.IgnoreCase)
        Dim dt As New DataTable()
        Return dt.Compute(formula, String.Empty)

        If match.Success Then
            Dim startCell As String = match.Groups(1).Value
            Dim endCell As String = match.Groups(2).Value

            Dim startColumn As Integer = ColumnLetterToIndex(Regex.Match(startCell, "[A-Z]+").Value)
            Dim startRow As Integer = Integer.Parse(Regex.Match(startCell, "\d+").Value) - 1
            Dim endColumn As Integer = ColumnLetterToIndex(Regex.Match(endCell, "[A-Z]+").Value)
            Dim endRow As Integer = Integer.Parse(Regex.Match(endCell, "\d+").Value) - 1

            Dim sum As Double = 0

            For rowIndex As Integer = startRow To endRow
                For colIndex As Integer = startColumn To endColumn
                    Dim cellValue As Object = dgvhoja1.Rows(rowIndex).Cells(colIndex).Value
                    Dim cellNumber As Double

                    If Double.TryParse(cellValue.ToString(), cellNumber) Then
                        sum += cellNumber
                    End If
                Next
            Next

            Return sum
        Else
            Throw New ArgumentException("Formato de fórmula no válido.")
        End If
    End Function


    Private Function EvaluateAverageFormula(formula As String) As Double


        ' ´FUNCION PARA EVALUAR LA FORMULA PROMEDIO

        Dim pattern As String = "=PROMEDIO\(([A-Z]+\d+):([A-Z]+\d+)\)"
        Dim match As Match = Regex.Match(formula, pattern, RegexOptions.IgnoreCase)

        If match.Success Then
            Dim startCell As String = match.Groups(1).Value
            Dim endCell As String = match.Groups(2).Value

            Dim startColumn As Integer = ColumnLetterToIndex(Regex.Match(startCell, "[A-Z]+").Value)
            Dim startRow As Integer = Integer.Parse(Regex.Match(startCell, "\d+").Value) - 1
            Dim endColumn As Integer = ColumnLetterToIndex(Regex.Match(endCell, "[A-Z]+").Value)
            Dim endRow As Integer = Integer.Parse(Regex.Match(endCell, "\d+").Value) - 1

            Dim sum As Double = 0
            Dim count As Integer = 0

            For rowIndex As Integer = startRow To endRow
                For colIndex As Integer = startColumn To endColumn
                    Dim cellValue As Object = dgvhoja1.Rows(rowIndex).Cells(colIndex).Value
                    Dim cellNumber As Double

                    If Double.TryParse(cellValue.ToString(), cellNumber) Then
                        sum += cellNumber
                        count += 1
                    End If
                Next
            Next

            If count > 0 Then
                Return sum / count
            Else
                Return 0
            End If
        Else
            Throw New ArgumentException("Formato de fórmula no válido.")
        End If
    End Function

    Private Function ColumnLetterToIndex(columnLetter As String) As Integer
        Dim sum As Integer = 0

        For Each ch As Char In columnLetter
            sum *= 26
            sum += (Asc(ch) - Asc("A"c) + 1)
        Next

        Return sum - 1
    End Function


    Private Sub ButtonGuardar_Click(sender As Object, e As EventArgs)

        ' CONFIGURAR EL CONTEXTO DE LA LICENCIA

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        ' CREAR UN NUEVO PAQUETE DE EXCEL

        Dim package As New ExcelPackage()

        ' AGREGAR UNA NUEVA HOJA DE TRABAJO

        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Datos")

        'AGREGAR LOS ENCABEZADOS DE COLUMNA

        For col As Integer = 1 To dgvhoja1.Columns.Count
            worksheet.Cells(1, col).Value = dgvhoja1.Columns(col - 1).HeaderText
        Next

        'AGREGAR LOS DATOS DEL DATAGRIDVIEW

        For row As Integer = 1 To dgvhoja1.Rows.Count
            For col As Integer = 1 To dgvhoja1.Columns.Count
                worksheet.Cells(row + 1, col).Value = dgvhoja1.Rows(row - 1).Cells(col - 1).Value
            Next
        Next

        'GUARDAR EL ARCHIVO EXCEL EN UNA UBICACIÓN ESPECÍFICA

        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Excel Files|*.xlsx"
        saveFileDialog.Title = "Guardar archivo Excel"
        saveFileDialog.FileName = "Datos.xlsx"

        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Dim fi As New FileInfo(saveFileDialog.FileName)
            package.SaveAs(fi)
            MessageBox.Show("Datos exportados correctamente.")
        End If

    End Sub

    Private Sub cambiar_colorTexto()
        Dim Caso As String
        Caso = cbxcolortexto.Text

        For Each cell As DataGridViewCell In dgvhoja1.SelectedCells
            Select Case Caso
                Case "Rojo"
                    cell.Style.ForeColor = Color.Red
                Case "Azul"
                    cell.Style.ForeColor = Color.Blue
                Case "Verde"
                    cell.Style.ForeColor = Color.Green
                Case "Amarillo"
                    cell.Style.ForeColor = Color.Yellow
                Case "Naranja"
                    cell.Style.ForeColor = Color.Orange
                Case "Morado"
                    cell.Style.ForeColor = Color.Purple
                Case "Cafe"
                    cell.Style.ForeColor = Color.Brown
                Case "Rosado"
                    cell.Style.ForeColor = Color.Pink
                Case "Blanco"
                    cell.Style.ForeColor = Color.White
                Case Else
                    cell.Style.ForeColor = Color.Black ' Color por defecto
            End Select

        Next

    End Sub

    Private Sub cambiar_colorFondo()
        Dim Caso As String
        Caso = cbxcolorfondo.Text

        For Each cell As DataGridViewCell In dgvhoja1.SelectedCells
            Select Case Caso
                Case "Rojo"
                    cell.Style.BackColor = Color.Red
                Case "Azul"
                    cell.Style.BackColor = Color.Blue
                Case "Verde"
                    cell.Style.BackColor = Color.Green
                Case "Amarillo"
                    cell.Style.BackColor = Color.Yellow
                Case "Naranja"
                    cell.Style.BackColor = Color.Orange
                Case "Morado"
                    cell.Style.BackColor = Color.Purple
                Case "Cafe"
                    cell.Style.BackColor = Color.Brown
                Case "Rosado"
                    cell.Style.BackColor = Color.Pink
                Case "Blanco"
                    cell.Style.BackColor = Color.White
                Case Else
                    cell.Style.BackColor = Color.Black ' Color por defecto
            End Select

        Next

    End Sub

    Private Sub tipo_letra()
        Dim letra As String
        Dim numero As Int16
        letra = cbxtipoletra.Text
        numero = Val(txtnumerofuente.Text)

        Dim selectedCells As DataGridViewSelectedCellCollection = dgvhoja1.SelectedCells

        For Each cell As DataGridViewCell In selectedCells
            cell.Style.Font = New Font(letra, numero)
        Next

    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        columnas()
        filas()
        txtnumerofuente.Text = 10
        AddHandler dgvhoja1.RowPostPaint, AddressOf dgvhoja1_RowPostPaint

        'INICIALIZAR EL ARREGLO DE FÓRMULAS DESPUÉS DE QUE DGVHOJA1 TENGA FILAS Y COLUMNAS

        formulas = New String(dgvhoja1.RowCount - 1, dgvhoja1.ColumnCount - 1) {}

        ' VERIFICAR SI HAY UNA RUTA VÁLIDA ANTES DE CARGAR EL ARCHIVO

        If Not String.IsNullOrEmpty(rutaArchivo) Then
            CargarArchivo(rutaArchivo)
        End If

    End Sub

    Private Sub dgvhoja1_RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs)
        Using b As SolidBrush = New SolidBrush(dgvhoja1.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + 4)
        End Using
    End Sub



    Private Sub dgvhoja1_SelectionChanged(sender As Object, e As EventArgs) Handles dgvhoja1.SelectionChanged
        Dim dataGridView As DataGridView = DirectCast(sender, DataGridView)

        If dataGridView.SelectedCells.Count > 0 Then
            Dim selectedCell As DataGridViewCell = dataGridView.SelectedCells(0)
            Dim columnName As String = dataGridView.Columns(selectedCell.ColumnIndex).Name
            Dim rowIndex As Integer = selectedCell.RowIndex


            ' ACTUALIZA LBLCELDA CON LA REFERENCIA DE LA CELDA


            lblcelda.Text = columnName & (rowIndex + 1).ToString()


            ' VERIFICA SI FORMULAS ESTÁ INICIALIZADO ANTES DE ACCEDER A ÉL


            If formulas IsNot Nothing Then

                ' MUESTRA LA FÓRMULA SI EXISTE, DE LO CONTRARIO, MUESTRA EL VALOR

                If formulas(rowIndex, selectedCell.ColumnIndex) IsNot Nothing Then
                    lblnumero.Text = formulas(rowIndex, selectedCell.ColumnIndex)
                Else
                    If selectedCell.Value IsNot Nothing Then
                        lblnumero.Text = selectedCell.Value.ToString()
                    Else
                        lblnumero.Text = "" ' O CUALQUIER OTRO VALOR POR DEFECTO SI LA CELDA ESTÁ VACÍA
                    End If
                End If
            Else


                ' SI FORMULAS NO ESTÁ INICIALIZADO, MUESTRA EL VALOR DE LA CELDA


                If selectedCell.Value IsNot Nothing Then
                    lblnumero.Text = selectedCell.Value.ToString()
                Else
                    lblnumero.Text = ""
                End If
            End If
        Else

            ' Si no hay celdas seleccionadas, limpia lblnumero

            lblnumero.Text = ""
        End If
    End Sub

    Private Sub dgvhoja1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvhoja1.CellDoubleClick
        Dim rowIndex As Integer = e.RowIndex
        Dim columnIndex As Integer = e.ColumnIndex

        If rowIndex >= 0 AndAlso columnIndex >= 0 Then


            ' Asigna el valor de lblnumero a la celda

            dgvhoja1.Rows(rowIndex).Cells(columnIndex).Value = lblnumero.Text


            ' Actualiza el arreglo de fórmulas si es necesario

            formulas(rowIndex, columnIndex) = Nothing ' Limpia la fórmula si existe


            ' Actualiza lblnumero para reflejar el nuevo valor de la celda

            lblnumero.Text = dgvhoja1.Rows(rowIndex).Cells(columnIndex).Value.ToString()
        End If
    End Sub

    Private Sub dgvhoja1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvhoja1.CellClick
        Dim rowIndex As Integer = e.RowIndex
        Dim columnIndex As Integer = e.ColumnIndex

        If rowIndex >= 0 AndAlso columnIndex >= 0 Then

            ' Asigna el valor de lblnumero a la celda

            dgvhoja1.Rows(rowIndex).Cells(columnIndex).Value = lblnumero.Text


            ' Actualiza el arreglo de fórmulas si es necesario

            formulas(rowIndex, columnIndex) = Nothing ' Limpia la fórmula si existe


            ' Actualiza lblnumero para reflejar el nuevo valor de la celda

            lblnumero.Text = dgvhoja1.Rows(rowIndex).Cells(columnIndex).Value.ToString()
        End If

        ' Verificamos si estamos en el modo de copiar formato

        If isCopyingFormat AndAlso e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then

            ' Si hemos seleccionado una celda válida, aplicamos el formato

            Dim destinationCell As DataGridViewCell = dgvhoja1.Rows(e.RowIndex).Cells(e.ColumnIndex)


            ' Copiamos el formato de la celda de origen a la celda de destino

            destinationCell.Style = sourceCell.Style

            ' Desactivamos el modo de copiar formato

            isCopyingFormat = False

            MessageBox.Show("Formato aplicado correctamente.", "Formato Aplicado", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Operaciones()
    End Sub

    Private Sub cbxcolortexto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxcolortexto.SelectedIndexChanged
        cambiar_colorTexto()
    End Sub

    Private Sub cbxcolorfondo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxcolorfondo.SelectedIndexChanged
        cambiar_colorFondo()
    End Sub

    Private Sub cbxtipoletra_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxtipoletra.SelectedIndexChanged
        tipo_letra()
    End Sub

    Private Sub txtnumerofuente_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtnumerofuente.KeyPress
        If Asc(e.KeyChar) = 13 Then
            tipo_letra()
        End If
    End Sub

    ' Evento que detecta cuando se edita una celda y evalúa las fórmulas ingresadas
    Private Sub dgvhoja1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvhoja1.CellEndEdit
        Dim cell As DataGridViewCell = dgvhoja1.Rows(e.RowIndex).Cells(e.ColumnIndex)
        Dim cellValue As String = If(cell.Value IsNot Nothing AndAlso Not IsDBNull(cell.Value), cell.Value.ToString(), "")

        If cellValue.StartsWith("=") Then
            Dim expression As String = cellValue.Substring(1).Trim()
            Try
                Dim functionName As String = expression.Split("("c)(0).ToUpper()
                Dim args As String = ""
                If expression.Length > functionName.Length + 1 Then
                    args = expression.Substring(functionName.Length + 1).TrimEnd(")"c)
                End If
                Select Case functionName
                    Case "SUMA" : cell.Value = EvaluateSum(args)
                    Case "PROMEDIO" : cell.Value = EvaluateAverage(args)
                    Case "CONCATENAR" : cell.Value = EvaluateConcatenate(args)
                    Case "CONTAR" : cell.Value = EvaluateCount(args)
                    Case "CONTARA" : cell.Value = EvaluateCountA(args)
                    Case "MAX" : cell.Value = EvaluateMax(args)
                    Case "MIN" : cell.Value = EvaluateMin(args)
                    Case "SI" : cell.Value = EvaluateIf(args)
                    Case "BUSCARV" : cell.Value = EvaluateVLookup(args)
                    Case "DIA" : cell.Value = EvaluateDay(args)
                    Case "MES" : cell.Value = EvaluateMonth(args)
                    Case "AÑO" : cell.Value = EvaluateYear(args)
                    Case "HOY" : cell.Value = EvaluateToday()
                    Case "AHORA" : cell.Value = EvaluateNow()
                    Case "REDONDEAR" : cell.Value = EvaluateRound(args)
                    Case "SI.ERROR" : cell.Value = EvaluateIfError(args)
                    Case "EXTRAER" : cell.Value = EvaluateExtract(args)
                    Case "LAPSO" : cell.Value = EvaluateElapsedTime(args)
                    Case "MAYU" : cell.Value = EvaluateUpperCase(args)
                    Case "MINU" : cell.Value = EvaluateLowerCase(args)
                    Case Else : cell.Value = EvaluateExpression(expression)
                End Select
                formulas(e.RowIndex, e.ColumnIndex) = cellValue
            Catch ex As Exception
                MessageBox.Show("Error en la expresión: " & ex.Message)
            End Try
        Else
            formulas(e.RowIndex, e.ColumnIndex) = Nothing
        End If
    End Sub
    Private Function EvaluateUpperCase(args As String) As String

        ' Evaluar el texto y convertirlo a mayúsculas

        Return UCase(EvaluateText(args.Trim()))
    End Function

    Private Function EvaluateLowerCase(args As String) As String

        ' Evaluar el texto y convertirlo a minúsculas

        Return LCase(EvaluateText(args.Trim()))
    End Function
    Private Function EvaluateExpression(expression As String) As Boolean

        If expression.StartsWith("PROMEDIO") Then

            ' Extraer el rango de celdas de la expresión

            Dim range As String = expression.Substring(expression.IndexOf("("c) + 1, expression.IndexOf(")"c) - expression.IndexOf("("c) - 1)

            ' Llamar a EvaluateAverage para obtener el promedio

            Dim avg As Double = EvaluateAverage(range)

            ' Evaluar la condición (si el promedio es mayor o igual a 70)

            Return avg >= 70
        End If

        ' Aquí podrías agregar lógica para otros tipos de evaluaciones si es necesario.

        Return False
    End Function
    Private Function ReplaceCellReferences(expression As String) As String
        Dim result As String = expression
        Dim matches = System.Text.RegularExpressions.Regex.Matches(expression, "[A-Z][0-9]+")

        For Each match As System.Text.RegularExpressions.Match In matches
            Dim cellReference As String = match.Value
            Dim cellValue As String = EvaluateText(cellReference)
            result = result.Replace(cellReference, cellValue)
        Next

        Return result
    End Function
    Private Function EvaluateSum(args As String) As Double
        Dim sum As Double = 0.0
        Dim ranges As String() = args.Split(","c)
        For Each range As String In ranges
            sum += EvaluateRange(range.Trim())
        Next
        Return sum
    End Function

    Private Function EvaluateAverage(args As String) As Double
        Try
            ' Asumir que el argumento es un rango como "A1:A5"

            Dim cells As List(Of DataGridViewCell) = GetCellsFromRange(args) ' Obtener todas las celdas del rango
            Dim sum As Double = 0
            Dim count As Integer = 0

            ' Sumar todos los valores de las celdas y contar cuántas hay

            For Each cell As DataGridViewCell In cells
                If cell.Value IsNot Nothing Then
                    sum += Convert.ToDouble(cell.Value)
                    count += 1
                End If
            Next

            If count > 0 Then
                Return sum / count ' Devuelve el promedio
            End If
        Catch ex As Exception
            Console.WriteLine("Error en EvaluateAverage: " & ex.Message)
        End Try
        Return 0
    End Function

    Private Function EvaluateConcatenate(args As String) As String
        Dim result As String = ""
        Dim parts As String() = args.Split(","c)

        For Each part As String In parts

            ' Evaluar cada parte y agregarla a la cadena resultante

            result &= EvaluateText(part.Trim()) & " "
        Next


        ' Eliminar el último espacio en blanco si es necesario

        If result.EndsWith(" ") Then
            result = result.Substring(0, result.Length - 1)
        End If

        Return result
    End Function
    Private Function EvaluateRange(range As String) As Double
        Dim sum As Double = 0.0

        ' Asumiendo que el rango está en el formato A1:B2

        Dim cells As List(Of DataGridViewCell) = GetCellsFromRange(range)
        For Each cell As DataGridViewCell In cells
            Dim cellValue As Double
            If Double.TryParse(cell.Value.ToString(), cellValue) Then
                sum += cellValue
            End If
        Next
        Return sum
    End Function

    Private Function EvaluateRangeWithCount(range As String) As (Double, Integer)
        Dim sum As Double = 0.0
        Dim count As Integer = 0


        ' Asumiendo que el rango está en el formato A1:B2

        Dim cells As List(Of DataGridViewCell) = GetCellsFromRange(range)
        For Each cell As DataGridViewCell In cells
            Dim cellValue As Double
            If Double.TryParse(cell.Value.ToString(), cellValue) Then
                sum += cellValue
                count += 1
            End If
        Next
        Return (sum, count)
    End Function

    Private Function EvaluateText(text As String) As String


        ' Intentar obtener la celda a partir del texto

        Dim cell As DataGridViewCell = Nothing
        Try
            Dim cellReference As (Integer, Integer) = ParseCell(text)
            cell = dgvhoja1.Rows(cellReference.Item1).Cells(cellReference.Item2)
        Catch ex As Exception

            ' No hacer nada si la referencia de celda no es válida

        End Try


        ' Retorna el valor de la celda si es texto, o el texto mismo si es una cadena

        If cell IsNot Nothing AndAlso cell.Value IsNot Nothing Then
            Return cell.Value.ToString()
        End If
        Return text
    End Function

    Private Function GetCellsFromRange(range As String) As List(Of DataGridViewCell)
        Dim cells As New List(Of DataGridViewCell)()
        Dim parts As String() = range.Split(":"c)
        Dim startCell As (Integer, Integer) = ParseCell(parts(0))
        Dim endCell As (Integer, Integer) = ParseCell(parts(1))

        For rowIndex As Integer = startCell.Item1 To endCell.Item1
            For colIndex As Integer = startCell.Item2 To endCell.Item2
                cells.Add(dgvhoja1.Rows(rowIndex).Cells(colIndex))
            Next
        Next
        Return cells
    End Function

    Private Function ParseCell(cell As String) As (Integer, Integer)

        ' Asumiendo que las celdas están en formato "A1", "B2", etc.

        Dim col As Integer = Asc(cell.Substring(0, 1).ToUpper()) - Asc("A")
        Dim row As Integer = Integer.Parse(cell.Substring(1)) - 1
        Return (row, col)
    End Function


    Private Sub Btnabrir_Click(sender As Object, e As EventArgs)

        ' Mostrar el cuadro de diálogo para seleccionar el archivo Excel

        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Excel Files|*.xlsx"
        openFileDialog.Title = "Abrir archivo Excel"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            currentFilePath = openFileDialog.FileName
            Dim fi As New FileInfo(currentFilePath)


            ' Configurar el contexto de la licencia

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial


            ' Cargar el archivo Excel

            Using package As New ExcelPackage(fi)
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)


                ' Limpiar el DataGridView antes de cargar nuevos datos

                dgvhoja1.Rows.Clear()
                dgvhoja1.Columns.Clear()


                ' Leer los encabezados de las columnas

                For col As Integer = 1 To worksheet.Dimension.End.Column
                    dgvhoja1.Columns.Add(worksheet.Cells(1, col).Text, worksheet.Cells(1, col).Text)
                Next


                ' Leer los datos de las filas

                For row As Integer = 2 To worksheet.Dimension.End.Row
                    Dim rowData As New List(Of String)()
                    For col As Integer = 1 To worksheet.Dimension.End.Column
                        rowData.Add(worksheet.Cells(row, col).Text)
                    Next
                    dgvhoja1.Rows.Add(rowData.ToArray())
                Next
            End Using

            MessageBox.Show("Datos cargados correctamente.")
        End If
        filas()
    End Sub

    Private Sub btnguardar_Click(sender As Object, e As EventArgs)
        If String.IsNullOrEmpty(currentFilePath) Then
            MessageBox.Show("Primero debe abrir un archivo Excel.")
            Return
        End If

        Dim fi As New FileInfo(currentFilePath)


        ' Configurar el contexto de la licencia

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial


        ' Cargar el archivo Excel

        Using package As New ExcelPackage(fi)
            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)


            ' Limpiar la hoja de trabajo antes de guardar los nuevos datos

            worksheet.Cells.Clear()


            ' Guardar los encabezados de las columnas

            For col As Integer = 1 To dgvhoja1.Columns.Count
                worksheet.Cells(1, col).Value = dgvhoja1.Columns(col - 1).HeaderText
            Next


            ' Guardar los datos del DataGridView

            For row As Integer = 1 To dgvhoja1.Rows.Count
                For col As Integer = 1 To dgvhoja1.Columns.Count
                    worksheet.Cells(row + 1, col).Value = dgvhoja1.Rows(row - 1).Cells(col - 1).Value
                Next
            Next



            ' Guardar el archivo Excel

            package.Save()
            MessageBox.Show("Datos guardados correctamente.")
        End Using
    End Sub

    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs) Handles GroupBox4.Enter

    End Sub

    ' Mostrar el cuadro de diálogo para seleccionar el archivo Excel
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click



        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Excel Files|*.xlsx"
        openFileDialog.Title = "Abrir archivo Excel"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            currentFilePath = openFileDialog.FileName
            Dim fi As New FileInfo(currentFilePath)


            ' Configurar el contexto de la licencia

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial


            ' Cargar el archivo Excel

            Using package As New ExcelPackage(fi)
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)


                ' Limpiar el DataGridView antes de cargar nuevos datos

                dgvhoja1.Rows.Clear()
                dgvhoja1.Columns.Clear()


                ' *** Volver a generar las columnas con letras ***

                columnas() ' Llama al método que genera las columnas de A-Z, AA-ZZ


                ' Leer los encabezados de las columnas

                For col As Integer = 1 To worksheet.Dimension.End.Column
                    dgvhoja1.Columns.Add(worksheet.Cells(1, col).Text, worksheet.Cells(1, col).Text)
                Next


                ' Leer los datos de las filas

                For row As Integer = 2 To worksheet.Dimension.End.Row
                    Dim rowData As New List(Of String)()
                    For col As Integer = 1 To worksheet.Dimension.End.Column
                        rowData.Add(worksheet.Cells(row, col).Text)
                    Next
                    dgvhoja1.Rows.Add(rowData.ToArray())
                Next
            End Using

            MessageBox.Show("Datos cargados correctamente.")
            Bienvenida.AgregarArchivoALista(openFileDialog.FileName)
            Bienvenida.GuardarHistorial(openFileDialog.FileName)
        End If
        filas()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        If String.IsNullOrEmpty(currentFilePath) Then
            MessageBox.Show("Primero debe abrir un archivo Excel.")
            Return
        End If

        Dim fi As New FileInfo(currentFilePath)

        ' Configurar el contexto de la licencia

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial


        ' Cargar el archivo Excel

        Using package As New ExcelPackage(fi)
            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)


            ' Limpiar la hoja de trabajo antes de guardar los nuevos datos

            worksheet.Cells.Clear()




            ' Guardar los encabezados de las columnas

            For col As Integer = 1 To dgvhoja1.Columns.Count
                worksheet.Cells(1, col).Value = dgvhoja1.Columns(col - 1).HeaderText
            Next


            ' Guardar los datos del DataGridView

            For row As Integer = 1 To dgvhoja1.Rows.Count
                For col As Integer = 1 To dgvhoja1.Columns.Count
                    worksheet.Cells(row + 1, col).Value = dgvhoja1.Rows(row - 1).Cells(col - 1).Value
                Next
            Next


            ' Guardar el archivo Excel

            package.Save()
            MessageBox.Show("Datos guardados correctamente.")
        End Using
    End Sub

    Private Sub PBGuardarComo_Click(sender As Object, e As EventArgs) Handles PBGuardarComo.Click


        ' Configurar el contexto de la licencia

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial


        ' Crear un nuevo paquete de Excel

        Dim package As New ExcelPackage()



        ' Agregar una nueva hoja de trabajo

        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Datos")


        ' Agregar los datos del DataGridView SIN encabezados

        For row As Integer = 0 To dgvhoja1.Rows.Count - 1
            For col As Integer = 0 To dgvhoja1.Columns.Count - 1
                worksheet.Cells(row + 1, col + 1).Value = dgvhoja1.Rows(row).Cells(col).Value
            Next
        Next

        ' Guardar el archivo Excel en una ubicación específica

        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Excel Files|*.xlsx"
        saveFileDialog.Title = "Guardar archivo Excel"
        saveFileDialog.FileName = "Datos.xlsx"

        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Dim fi As New FileInfo(saveFileDialog.FileName)
            package.SaveAs(fi)
            MessageBox.Show("Datos exportados correctamente.")
        End If
        Bienvenida.AgregarArchivoALista(saveFileDialog.FileName)
        Bienvenida.GuardarHistorial(saveFileDialog.FileName)
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub txtnumerofuente_TextChanged(sender As Object, e As EventArgs) Handles txtnumerofuente.TextChanged

    End Sub

    Private Sub btnNegrita_Click(sender As Object, e As EventArgs) Handles btnNegrita.Click
        For Each celda As DataGridViewCell In dgvhoja1.SelectedCells
            If celda.Style.Font Is Nothing Then
                celda.Style.Font = New Font(dgvhoja1.Font, FontStyle.Bold)
            Else
                Dim currentFont As Font = celda.Style.Font
                celda.Style.Font = New Font(currentFont.FontFamily, currentFont.Size, currentFont.Style Xor FontStyle.Bold)
            End If
        Next
    End Sub

    Private Sub BtnCursiva_Click(sender As Object, e As EventArgs) Handles BtnCursiva.Click
        For Each celda As DataGridViewCell In dgvhoja1.SelectedCells
            If celda.Style.Font Is Nothing Then
                celda.Style.Font = New Font(dgvhoja1.Font, FontStyle.Italic)
            Else
                Dim currentFont As Font = celda.Style.Font
                celda.Style.Font = New Font(currentFont.FontFamily, currentFont.Size, currentFont.Style Xor FontStyle.Italic)
            End If
        Next
    End Sub

    Private Sub btnSubrayar_Click(sender As Object, e As EventArgs) Handles btnSubrayar.Click
        For Each celda As DataGridViewCell In dgvhoja1.SelectedCells
            If celda.Style.Font Is Nothing Then
                celda.Style.Font = New Font(dgvhoja1.Font, FontStyle.Underline)
            Else
                Dim currentFont As Font = celda.Style.Font
                celda.Style.Font = New Font(currentFont.FontFamily, currentFont.Size, currentFont.Style Xor FontStyle.Underline)
            End If
        Next
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        For Each celda As DataGridViewCell In dgvhoja1.SelectedCells
            celda.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
        Next
    End Sub

    Private Sub PBCentrar_Click(sender As Object, e As EventArgs) Handles PBCentrar.Click
        For Each celda As DataGridViewCell In dgvhoja1.SelectedCells
            celda.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next
    End Sub

    Private Sub PBDerecha_Click(sender As Object, e As EventArgs) Handles PBDerecha.Click
        For Each celda As DataGridViewCell In dgvhoja1.SelectedCells
            celda.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        Next
    End Sub

    Private Sub dgvhoja1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvhoja1.CellContentClick

    End Sub

    Private Sub lblcelda_Click(sender As Object, e As EventArgs) Handles lblcelda.Click

    End Sub

    Private Sub lblnumero_Click(sender As Object, e As EventArgs) Handles lblnumero.Click

    End Sub

    Private Sub GroupBox5_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub PBAyuda_Click(sender As Object, e As EventArgs) Handles PBAyuda.Click
        Dim pdfPath As String = "C:\Users\PL\Desktop\Teoria de la computacion\Teoria de la computacion\Proyecto Clase\Proyecto Clase\Resources\ManualusoExcelito.pdf"
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

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub BTNCopiar_Click(sender As Object, e As EventArgs) Handles BTNCopiar.Click
        If dgvhoja1.SelectedCells.Count > 0 Then
            Dim textoCopiado As New System.Text.StringBuilder()

            ' Obtener la selección ordenada por fila y columna
            Dim filasOrdenadas = dgvhoja1.SelectedCells.Cast(Of DataGridViewCell)() _
                            .OrderBy(Function(c) c.RowIndex) _
                            .ThenBy(Function(c) c.ColumnIndex) _
                            .GroupBy(Function(c) c.RowIndex)

            ' Recorrer cada fila seleccionada
            For Each fila In filasOrdenadas
                Dim filaTexto As New List(Of String)

                ' Recorrer celdas dentro de la fila
                For Each celda In fila.OrderBy(Function(c) c.ColumnIndex)
                    If celda.Value IsNot Nothing Then
                        filaTexto.Add(celda.Value.ToString())
                    Else
                        filaTexto.Add("") ' Espacio vacío si la celda está vacía
                    End If
                Next

                ' Agregar fila al texto copiado, separando con tabulaciones
                textoCopiado.AppendLine(String.Join(vbTab, filaTexto))
            Next

            ' Copiar al portapapeles
            Clipboard.SetText(textoCopiado.ToString())
        Else
            MessageBox.Show("Seleccione al menos una celda para copiar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub BTNPegar_Click(sender As Object, e As EventArgs) Handles BTNPegar.Click
        If dgvhoja1.SelectedCells.Count = 0 Then
            MessageBox.Show("Seleccione una celda para pegar los datos.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' Obtener el contenido del portapapeles
        Dim datosPortapapeles As String = Clipboard.GetText()
        If String.IsNullOrEmpty(datosPortapapeles) Then Exit Sub

        ' Dividir los datos en filas y columnas
        Dim filas As String() = datosPortapapeles.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)

        ' Obtener la celda de inicio (la primera seleccionada)
        Dim filaInicio As Integer = dgvhoja1.SelectedCells(0).RowIndex
        Dim columnaInicio As Integer = dgvhoja1.SelectedCells(0).ColumnIndex

        ' Recorrer las filas copiadas
        For i As Integer = 0 To filas.Length - 1
            Dim columnas As String() = filas(i).Split(vbTab) ' Separar columnas por tabulación

            For j As Integer = 0 To columnas.Length - 1
                Dim filaDestino As Integer = filaInicio + i
                Dim columnaDestino As Integer = columnaInicio + j

                ' Verificar que la celda esté dentro del DataGridView
                If filaDestino < dgvhoja1.RowCount AndAlso columnaDestino < dgvhoja1.ColumnCount Then
                    dgvhoja1.Rows(filaDestino).Cells(columnaDestino).Value = columnas(j)
                End If
            Next
        Next
    End Sub

    Private Sub BTNCortar_Click(sender As Object, e As EventArgs) Handles BTNCortar.Click
        If dgvhoja1.SelectedCells.Count = 0 Then
            MessageBox.Show("Seleccione al menos una celda para cortar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim textoCopiado As New System.Text.StringBuilder()

        ' Ordenar las celdas por fila y columna
        Dim filasOrdenadas = dgvhoja1.SelectedCells.Cast(Of DataGridViewCell)() _
                        .OrderBy(Function(c) c.RowIndex) _
                        .ThenBy(Function(c) c.ColumnIndex) _
                        .GroupBy(Function(c) c.RowIndex)

        ' Recorrer cada fila seleccionada
        For Each fila In filasOrdenadas
            Dim filaTexto As New List(Of String)

            ' Recorrer celdas dentro de la fila
            For Each celda In fila.OrderBy(Function(c) c.ColumnIndex)
                If celda.Value IsNot Nothing Then
                    filaTexto.Add(celda.Value.ToString())
                    celda.Value = Nothing ' Borra el contenido de la celda
                Else
                    filaTexto.Add("") ' Espacio vacío si la celda está vacía
                End If
            Next

            ' Agregar fila al texto copiado, separando con tabulaciones
            textoCopiado.AppendLine(String.Join(vbTab, filaTexto))
        Next

        ' Copiar al portapapeles
        Clipboard.SetText(textoCopiado.ToString())
    End Sub

    Private Sub BTNAutoSuma_Click(sender As Object, e As EventArgs) Handles BTNAutoSuma.Click
        If dgvhoja1.SelectedCells.Count = 0 Then
            MessageBox.Show("Seleccione al menos una celda con valores numéricos para sumar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim filasSeleccionadas As New HashSet(Of Integer)
        Dim columnasSeleccionadas As New HashSet(Of Integer)
        Dim sumaFilas As New Dictionary(Of Integer, Double)
        Dim sumaColumnas As New Dictionary(Of Integer, Double)
        Dim totalSuma As Double = 0

        ' Recorrer las celdas seleccionadas
        For Each celda As DataGridViewCell In dgvhoja1.SelectedCells
            Dim fila As Integer = celda.RowIndex
            Dim columna As Integer = celda.ColumnIndex

            filasSeleccionadas.Add(fila)
            columnasSeleccionadas.Add(columna)

            If IsNumeric(celda.Value) Then
                Dim valor As Double = Convert.ToDouble(celda.Value)

                ' Sumar por filas
                If Not sumaFilas.ContainsKey(fila) Then sumaFilas(fila) = 0
                sumaFilas(fila) += valor

                ' Sumar por columnas
                If Not sumaColumnas.ContainsKey(columna) Then sumaColumnas(columna) = 0
                sumaColumnas(columna) += valor

                ' Sumar el total
                totalSuma += valor
            End If
        Next

        ' Obtener los límites de la selección
        Dim filaMax As Integer = filasSeleccionadas.Max()
        Dim columnaMax As Integer = columnasSeleccionadas.Max()

        ' Determinar la fila y columna de suma (si están vacías, usar la última seleccionada; si no, agregar una nueva)
        Dim filaSuma As Integer = If(IsCellEmpty(dgvhoja1, filaMax + 1, columnasSeleccionadas.First()), filaMax + 1, filaMax + 2)
        Dim columnaSuma As Integer = If(IsCellEmpty(dgvhoja1, filasSeleccionadas.First(), columnaMax + 1), columnaMax + 1, columnaMax + 2)

        ' Si la fila suma no existe, agregar una nueva
        While filaSuma >= dgvhoja1.RowCount
            dgvhoja1.Rows.Add()
        End While

        ' Si la columna suma no existe, agregar una nueva
        While columnaSuma >= dgvhoja1.ColumnCount
            dgvhoja1.Columns.Add("Columna" & columnaSuma, "Columna " & columnaSuma)
        End While

        ' Insertar la suma de cada fila en la última columna vacía de la selección
        For Each fila In sumaFilas.Keys
            Dim colDestino = If(IsCellEmpty(dgvhoja1, fila, columnaMax + 1), columnaMax + 1, columnaMax + 2)
            dgvhoja1.Rows(fila).Cells(colDestino).Value = sumaFilas(fila)
        Next

        ' Insertar la suma de cada columna en la última fila vacía de la selección
        For Each columna In sumaColumnas.Keys
            Dim filaDestino = If(IsCellEmpty(dgvhoja1, filaMax + 1, columna), filaMax + 1, filaMax + 2)
            dgvhoja1.Rows(filaDestino).Cells(columna).Value = sumaColumnas(columna)
        Next

        ' Insertar la suma total en la celda vacía en la esquina inferior derecha
        dgvhoja1.Rows(filaSuma).Cells(columnaSuma).Value = totalSuma
    End Sub

    ' Función para verificar si una celda está vacía
    Private Function IsCellEmpty(dgv As DataGridView, fila As Integer, columna As Integer) As Boolean
        If fila < dgv.RowCount AndAlso columna < dgv.ColumnCount Then
            Dim cellValue = dgv.Rows(fila).Cells(columna).Value
            Return cellValue Is Nothing OrElse cellValue.ToString().Trim() = ""
        End If
        Return True
    End Function

    Private Sub BTNAutoSumaS_Click(sender As Object, e As EventArgs) Handles BTNAutoSumaS.Click
        If dgvhoja1.SelectedCells.Count = 0 Then
            MessageBox.Show("Seleccione al menos una celda con valores numéricos para sumar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim suma As Double = 0
        Dim filasSeleccionadas As New HashSet(Of Integer)
        Dim columnasSeleccionadas As New HashSet(Of Integer)

        ' Recorrer las celdas seleccionadas
        For Each celda As DataGridViewCell In dgvhoja1.SelectedCells
            If IsNumeric(celda.Value) Then
                suma += Convert.ToDouble(celda.Value)
            End If
            filasSeleccionadas.Add(celda.RowIndex)
            columnasSeleccionadas.Add(celda.ColumnIndex)
        Next

        ' Verificar si se encontraron valores numéricos
        If suma = 0 Then
            MessageBox.Show("No se encontraron valores numéricos en la selección.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim filaDestino As Integer
        Dim columnaDestino As Integer

        ' Determinar si la selección está en una sola fila o en una sola columna
        If filasSeleccionadas.Count = 1 Then
            ' Selección en una sola fila → colocar el resultado en la siguiente columna
            filaDestino = filasSeleccionadas.First()
            columnaDestino = columnasSeleccionadas.Max() + 1

            ' Si la columna destino no existe, agregar una nueva
            If columnaDestino >= dgvhoja1.ColumnCount Then
                dgvhoja1.Columns.Add("Columna" & columnaDestino, "Columna " & columnaDestino)
            End If
        ElseIf columnasSeleccionadas.Count = 1 Then
            ' Selección en una sola columna → colocar el resultado en la siguiente fila
            filaDestino = filasSeleccionadas.Max() + 1
            columnaDestino = columnasSeleccionadas.First()

            ' Si la fila destino no existe, agregar una nueva
            If filaDestino >= dgvhoja1.RowCount Then
                dgvhoja1.Rows.Add()
            End If
        Else
            MessageBox.Show("Seleccione valores en una única fila o columna.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' Colocar la suma en la celda destino
        dgvhoja1.Rows(filaDestino).Cells(columnaDestino).Value = suma
    End Sub

    Private Sub BTNCopiarFormato_Enter(sender As Object, e As EventArgs) Handles BTNCopiarFormato.Enter

    End Sub

    Dim sourceCell As DataGridViewCell = Nothing ' Celda de origen
    Dim isCopyingFormat As Boolean = False ' Bander
    Private Sub BTNCopiarFormato_Click(sender As Object, e As EventArgs) Handles BTNCopiarFormato.Click
        ' Si no hemos seleccionado una celda, no hacemos nada
        If dgvhoja1.SelectedCells.Count > 0 Then
            ' Guardamos la celda seleccionada como origen
            sourceCell = dgvhoja1.SelectedCells(0)
            isCopyingFormat = True ' Activamos el modo de copiar formato
            MessageBox.Show("Formato copiado. Ahora, seleccione la celda de destino para pegar el formato.", "Copiar Formato", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Por favor, seleccione una celda de origen para copiar el formato.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub CargarArchivo(filePath As String)
        Try
            Dim fi As New FileInfo(filePath)

            ' Configurar la licencia de EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial

            Using package As New ExcelPackage(fi)
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)

                ' Limpiar el DataGridView antes de cargar nuevos datos
                dgvhoja1.Rows.Clear()
                dgvhoja1.Columns.Clear()

                ' *** Volver a generar las columnas con letras ***
                columnas() ' Llama al método que genera las columnas de A-Z, AA-ZZ


                ' Leer los encabezados de las columnas
                For col As Integer = 1 To worksheet.Dimension.End.Column
                    dgvhoja1.Columns.Add(worksheet.Cells(1, col).Text, worksheet.Cells(1, col).Text)
                Next

                ' Leer los datos de las filas SIN modificar las cabeceras
                For row As Integer = 1 To worksheet.Dimension.End.Row
                    Dim rowData As New List(Of String)()
                    For col As Integer = 1 To worksheet.Dimension.End.Column
                        rowData.Add(worksheet.Cells(row, col).Text)
                    Next
                    dgvhoja1.Rows.Add(rowData.ToArray())
                Next

                ' *** Asegurar que haya al menos 100 filas ***
                While dgvhoja1.Rows.Count < 100
                    dgvhoja1.Rows.Add()
                End While
            End Using

            MessageBox.Show("Datos cargados correctamente.")
        Catch ex As Exception
            MessageBox.Show("Error al abrir el archivo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub AbrirArchivoDesdeOtroFormulario(ByVal filePath As String)
        Try
            Dim fi As New FileInfo(filePath)

            ' Configurar la licencia de EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial

            ' Cargar el archivo Excel
            Using package As New ExcelPackage(fi)
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)

                ' Limpiar el DataGridView antes de cargar nuevos datos
                dgvhoja1.Rows.Clear()
                dgvhoja1.Columns.Clear()

                ' *** Volver a generar las columnas con letras ***
                columnas() ' Asegura que las columnas sean A-Z, AA-ZZ

                ' Leer los datos de las filas SIN modificar las cabeceras
                For row As Integer = 1 To worksheet.Dimension.End.Row
                    Dim rowData As New List(Of String)()
                    For col As Integer = 1 To worksheet.Dimension.End.Column
                        rowData.Add(worksheet.Cells(row, col).Text)
                    Next
                    dgvhoja1.Rows.Add(rowData.ToArray())
                Next
                ' *** Asegurar que haya al menos 100 filas ***
                While dgvhoja1.Rows.Count < 100
                    dgvhoja1.Rows.Add()
                End While
            End Using

            ' Mostrar Form1 en caso de que esté minimizado o no visible
            Me.Show()
            Me.BringToFront()

            MessageBox.Show("Datos cargados correctamente en Form1.")

        Catch ex As Exception
            MessageBox.Show("Error al abrir el archivo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Función que cuenta el número de celdas que contienen números en un rango
    Private Function EvaluateCount(args As String) As Integer
        Dim count As Integer = 0
        Dim cells As List(Of DataGridViewCell) = GetCellsFromRange(args)
        For Each cell As DataGridViewCell In cells
            If IsNumeric(cell.Value) Then count += 1
        Next
        Return count
    End Function

    ' Función que cuenta el número de celdas no vacías en un rango
    Private Function EvaluateCountA(args As String) As Integer
        Dim count As Integer = 0
        Dim cells As List(Of DataGridViewCell) = GetCellsFromRange(args)
        For Each cell As DataGridViewCell In cells
            If Not String.IsNullOrEmpty(cell.Value?.ToString()) Then count += 1
        Next
        Return count
    End Function

    ' Función que devuelve el valor máximo dentro de un rango de celdas
    Private Function EvaluateMax(args As String) As Double
        Dim maxVal As Double = Double.MinValue
        Dim cells As List(Of DataGridViewCell) = GetCellsFromRange(args)
        For Each cell As DataGridViewCell In cells
            Dim value As Double
            If Double.TryParse(cell.Value?.ToString(), value) Then
                If value > maxVal Then maxVal = value
            End If
        Next
        Return If(maxVal = Double.MinValue, 0, maxVal)
    End Function

    ' Función que devuelve el valor mínimo dentro de un rango de celdas
    Private Function EvaluateMin(args As String) As Double
        Dim minVal As Double = Double.MaxValue
        Dim cells As List(Of DataGridViewCell) = GetCellsFromRange(args)
        For Each cell As DataGridViewCell In cells
            Dim value As Double
            If Double.TryParse(cell.Value?.ToString(), value) Then
                If value < minVal Then minVal = value
            End If
        Next
        Return If(minVal = Double.MaxValue, 0, minVal)
    End Function

    ' Función que evalúa una condición lógica y devuelve un valor si es verdadera y otro si es falsa
    Private Function EvaluateIf(args As String) As String
        Try
            ' Dividir el argumento en condición, valor verdadero y valor falso
            Dim parts As String() = args.Split(","c)
            If parts.Length <> 3 Then Throw New ArgumentException("El formato de la fórmula es incorrecto.")

            ' Evaluar la condición (por ejemplo, PROMEDIO(A1:A5) >= 70)
            Dim condition As Boolean = EvaluateExpression(parts(0)) ' Evaluamos la expresión de la condición

            ' Dependiendo de la condición, retornar uno u otro valor
            If condition Then
                Return parts(1).Trim() ' Valor si la condición es verdadera
            Else
                Return parts(2).Trim() ' Valor si la condición es falsa
            End If
        Catch ex As Exception
            Console.WriteLine("Error en EvaluateIf: " & ex.Message)
        End Try
        Return "Error"
    End Function

    ' Función que busca un valor en la primera columna de un rango de datos y devuelve el valor de otra columna en la misma fila
    Private Function EvaluateVLookup(args As String) As String
        Try
            ' Separar los argumentos
            Dim parts As String() = args.Split(","c)
            If parts.Length <> 3 Then Throw New ArgumentException("Formato incorrecto en BUSCARV")

            ' Extraer valores
            Dim searchValue As String = parts(0).Trim()
            Dim tableStart As String = parts(1).Trim().Split(":")(0) ' Primera celda del rango (ej. A1)
            Dim tableEnd As String = parts(1).Trim().Split(":")(1) ' Última celda del rango (ej. D5)
            Dim columnIndex As Integer = Integer.Parse(parts(2).Trim()) - 1 ' Convertir a índice base 0

            ' Verificar que el índice de columna sea válido
            If columnIndex < 0 Then Throw New ArgumentException("Número de columna fuera de rango")

            ' Obtener la primera celda de la tabla
            Dim startCell As DataGridViewCell = ObtenerCelda(tableStart)
            Dim endCell As DataGridViewCell = ObtenerCelda(tableEnd)

            ' Verificar que las celdas existen
            If startCell Is Nothing OrElse endCell Is Nothing Then
                Throw New ArgumentException("Rango no válido")
            End If

            ' Determinar los límites del rango
            Dim startRow As Integer = startCell.RowIndex
            Dim endRow As Integer = endCell.RowIndex
            Dim startCol As Integer = startCell.ColumnIndex

            ' Recorrer la primera columna del rango para buscar el valor
            For i As Integer = startRow To endRow
                Dim cell As DataGridViewCell = dgvhoja1.Rows(i).Cells(startCol)
                If cell IsNot Nothing AndAlso cell.Value IsNot Nothing AndAlso cell.Value.ToString().Trim() = searchValue Then
                    ' Verificar si la columna a devolver está dentro del rango
                    If startCol + columnIndex < dgvhoja1.Columns.Count Then
                        Return dgvhoja1.Rows(i).Cells(startCol + columnIndex).Value.ToString()
                    Else
                        Return "Error: Índice de columna fuera de rango"
                    End If
                End If
            Next

        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try

        Return "No encontrado"
    End Function
    Private Function ObtenerCelda(referencia As String) As DataGridViewCell
        Try
            ' Verificar que la referencia no sea nula o vacía
            If String.IsNullOrEmpty(referencia) Then
                Console.WriteLine("Error: Referencia vacía")
                Return Nothing
            End If

            ' Separar letras (columna) y números (fila)
            Dim columnaLetra As String = ""
            Dim filaNumero As String = ""

            For Each ch As Char In referencia
                If Char.IsLetter(ch) Then
                    columnaLetra &= ch
                ElseIf Char.IsDigit(ch) Then
                    filaNumero &= ch
                End If
            Next

            ' Convertir la parte de la fila a número
            Dim filaIndex As Integer
            If Not Integer.TryParse(filaNumero, filaIndex) Then
                Console.WriteLine("Error: No se pudo extraer la fila de la referencia " & referencia)
                Return Nothing
            End If

            ' Convertir columna de letra a índice numérico (A=0, B=1, ..., Z=25, AA=26, AB=27, ...)
            Dim columnaIndex As Integer = 0
            For i As Integer = 0 To columnaLetra.Length - 1
                columnaIndex *= 26
                columnaIndex += Asc(columnaLetra(i)) - Asc("A") + 1
            Next
            columnaIndex -= 1 ' Ajustar a índice basado en 0

            ' Ajustar fila a índice basado en 0
            filaIndex -= 1

            ' Verificar si los índices están dentro de los límites del DataGridView
            If filaIndex >= 0 AndAlso filaIndex < dgvhoja1.Rows.Count AndAlso
           columnaIndex >= 0 AndAlso columnaIndex < dgvhoja1.Columns.Count Then

                Return dgvhoja1.Rows(filaIndex).Cells(columnaIndex)
            Else
                Console.WriteLine("Error: La referencia está fuera de los límites del DataGridView")
            End If

        Catch ex As Exception
            Console.WriteLine("Error en ObtenerCelda: " & ex.Message)
        End Try

        Return Nothing
    End Function
    ' Funciones para extraer el día, mes o año de una fecha proporcionada
    Private Function EvaluateDay(args As String) As Integer
        Try
            ' Convertir el argumento (como "A14") a las coordenadas de fila y columna
            Dim column As Integer = ConvertColumnToIndex(args.Substring(0, 1)) ' Convierte "A" a índice de columna
            Dim row As Integer = Convert.ToInt32(args.Substring(1)) - 1 ' Convierte "14" a índice de fila (empieza desde 0)

            ' Acceder a la celda utilizando las coordenadas obtenidas
            Dim cell As DataGridViewCell = dgvhoja1.Rows(row).Cells(column)

            If cell IsNot Nothing AndAlso cell.Value IsNot Nothing Then
                Dim dateValue As DateTime
                If DateTime.TryParse(cell.Value.ToString(), dateValue) Then
                    Return dateValue.Day
                End If
            End If

        Catch ex As Exception
            ' Imprimir la excepción para entender mejor el error
            Console.WriteLine("Error en EvaluateDay: " & ex.Message)
            Console.WriteLine("Pila de llamadas: " & ex.StackTrace)
        End Try

        Return 0 ' Devuelve 0 si no se encuentra una fecha válida
    End Function

    ' Función auxiliar para convertir una letra de columna en un índice de columna (por ejemplo, "A" -> 0, "B" -> 1, etc.)
    Private Function ConvertColumnToIndex(col As String) As Integer
        ' Convierte la letra de la columna a su índice (Ejemplo: "A" -> 0, "B" -> 1, "C" -> 2)
        Return Asc(col.ToUpper()) - Asc("A"c)
    End Function

    Private Function EvaluateMonth(args As String) As Integer
        Try
            ' Convertir el argumento (como "A14") a las coordenadas de fila y columna
            Dim column As Integer = ConvertColumnToIndex(args.Substring(0, 1)) ' Convierte "A" a índice de columna
            Dim row As Integer = Convert.ToInt32(args.Substring(1)) - 1 ' Convierte "14" a índice de fila (empieza desde 0)

            ' Acceder a la celda utilizando las coordenadas obtenidas
            Dim cell As DataGridViewCell = dgvhoja1.Rows(row).Cells(column)

            If cell IsNot Nothing AndAlso cell.Value IsNot Nothing Then
                Dim dateValue As DateTime
                If DateTime.TryParse(cell.Value.ToString(), dateValue) Then
                    Return dateValue.Month
                End If
            End If

        Catch ex As Exception
            ' Imprimir la excepción para entender mejor el error
            Console.WriteLine("Error en EvaluateDay: " & ex.Message)
            Console.WriteLine("Pila de llamadas: " & ex.StackTrace)
        End Try

        Return 0 ' Devuelve 0 si no se encuentra una fecha válida
    End Function

    Private Function EvaluateYear(args As String) As Integer
        Try
            ' Convertir el argumento (como "A14") a las coordenadas de fila y columna
            Dim column As Integer = ConvertColumnToIndex(args.Substring(0, 1)) ' Convierte "A" a índice de columna
            Dim row As Integer = Convert.ToInt32(args.Substring(1)) - 1 ' Convierte "14" a índice de fila (empieza desde 0)

            ' Acceder a la celda utilizando las coordenadas obtenidas
            Dim cell As DataGridViewCell = dgvhoja1.Rows(row).Cells(column)

            If cell IsNot Nothing AndAlso cell.Value IsNot Nothing Then
                Dim dateValue As DateTime
                If DateTime.TryParse(cell.Value.ToString(), dateValue) Then
                    Return dateValue.Year
                End If
            End If

        Catch ex As Exception
            ' Imprimir la excepción para entender mejor el error
            Console.WriteLine("Error en EvaluateDay: " & ex.Message)
            Console.WriteLine("Pila de llamadas: " & ex.StackTrace)
        End Try

        Return 0 ' Devuelve 0 si no se encuentra una fecha válida
    End Function
    ' Función que devuelve la fecha actual
    Private Function EvaluateToday() As Date
        Return DateTime.Today
    End Function

    ' Función que devuelve la fecha y hora actuales
    Private Function EvaluateNow() As Date
        Return DateTime.Now
    End Function

    ' Función que redondea un número a un número específico de dígitos
    Private Function EvaluateRound(args As String) As Double
        Try
            ' Dividir los argumentos para obtener la celda y el número de decimales
            Dim parts As String() = args.Split(","c)

            ' Verificar que tenemos dos parámetros (celda y número de decimales)
            If parts.Length <> 2 Then
                Throw New ArgumentException("Formato incorrecto en REDONDEAR. Debes pasar una celda y el número de decimales.")
            End If

            ' Obtener la referencia de la celda (por ejemplo "A1")
            Dim cellReference As String = parts(0).Trim()

            ' Obtener el número de decimales
            Dim decimals As Integer
            If Not Integer.TryParse(parts(1).Trim(), decimals) Then
                Throw New ArgumentException("El número de decimales no es válido.")
            End If

            ' Obtener las filas y columnas correspondientes a la celda (por ejemplo "A1")
            ' Suponemos que "A1" es la celda en la primera fila y columna
            ' Convertir la referencia de la celda (por ejemplo "A1") a índices de fila y columna
            Dim column As Integer = Asc(Char.ToUpper(cellReference(0))) - Asc("A"c) ' Columna de la celda
            Dim row As Integer = Integer.Parse(cellReference.Substring(1)) - 1 ' Fila de la celda (considerando que las filas en DataGridView empiezan desde 0)

            ' Acceder a la celda usando los índices obtenidos
            Dim cell As DataGridViewCell = dgvhoja1.Rows(row).Cells(column)

            ' Verificar si la celda tiene un valor
            If cell IsNot Nothing AndAlso cell.Value IsNot Nothing Then
                ' Convertir el valor de la celda a un número
                Dim value As Double
                If Double.TryParse(cell.Value.ToString(), value) Then
                    ' Redondear el valor
                    Return Math.Round(value, decimals)
                End If
            End If

        Catch ex As Exception
            ' Si ocurre algún error, mostrarlo
            Console.WriteLine("Error en EvaluateRound: " & ex.Message)
        End Try

        ' Si ocurre algún error o no hay valor válido, devolver 0
        Return 0
    End Function

    Private Function EvaluateExtract(args As String) As String
        Try
            ' Dividir los argumentos para obtener la celda y el número de la posición
            Dim parts As String() = args.Split(","c)

            ' Verificar que tenemos dos parámetros (celda y número de posición)
            If parts.Length <> 2 Then
                Throw New ArgumentException("Formato incorrecto en EXTRAER. Debes pasar una celda y la posición del carácter.")
            End If

            ' Obtener la referencia de la celda (por ejemplo "A1")
            Dim cellReference As String = parts(0).Trim()

            ' Obtener el número de la posición
            Dim position As Integer
            If Not Integer.TryParse(parts(1).Trim(), position) Then
                Throw New ArgumentException("La posición no es válida.")
            End If

            ' Convertir la referencia de la celda (por ejemplo "A1") a índices de fila y columna
            Dim column As Integer = Asc(Char.ToUpper(cellReference(0))) - Asc("A"c) ' Columna de la celda
            Dim row As Integer = Integer.Parse(cellReference.Substring(1)) - 1 ' Fila de la celda (considerando que las filas en DataGridView empiezan desde 0)

            ' Acceder a la celda usando los índices obtenidos
            Dim cell As DataGridViewCell = dgvhoja1.Rows(row).Cells(column)

            ' Verificar si la celda tiene un valor
            If cell IsNot Nothing AndAlso cell.Value IsNot Nothing Then
                ' Obtener el valor de la celda como texto
                Dim text As String = cell.Value.ToString()

                ' Verificar que la posición sea válida dentro de la longitud del texto
                If position > 0 AndAlso position <= text.Length Then
                    ' Extraer el carácter correspondiente a la posición (ajustando el índice ya que es 1-based)
                    Return text.Substring(position - 1, 1)
                End If
            End If

        Catch ex As Exception
            ' Si ocurre algún error, mostrarlo
            Console.WriteLine("Error en EvaluateExtract: " & ex.Message)
        End Try

        ' Si ocurre algún error o no hay valor válido, devolver una cadena vacía
        Return ""
    End Function

    ' Función que devuelve un valor especificado si la evaluación de una expresión genera un error
    Private Function EvaluateIfError(args As String) As String
        Dim parts As String() = args.Split(","c)
        If parts.Length <> 2 Then Throw New ArgumentException("Formato incorrecto en SI.ERROR")
        Try
            Return EvaluateExpression(parts(0)).ToString()
        Catch
            Return parts(1)
        End Try
    End Function

    Private Function EvaluateElapsedTime(args As String) As Integer
        Try
            ' Dividir los argumentos para obtener las dos celdas con las fechas
            Dim parts As String() = args.Split(","c)

            ' Verificar que tenemos dos parámetros (referencia de celda 1 y referencia de celda 2)
            If parts.Length <> 2 Then
                Throw New ArgumentException("Formato incorrecto en TRANSCURSO. Debes pasar dos celdas con fechas.")
            End If

            ' Obtener las referencias de las celdas
            Dim cellReference1 As String = parts(0).Trim()
            Dim cellReference2 As String = parts(1).Trim()

            ' Convertir las referencias de las celdas (por ejemplo "A1" y "A2") a índices de fila y columna
            Dim column1 As Integer = Asc(Char.ToUpper(cellReference1(0))) - Asc("A"c)
            Dim row1 As Integer = Integer.Parse(cellReference1.Substring(1)) - 1

            Dim column2 As Integer = Asc(Char.ToUpper(cellReference2(0))) - Asc("A"c)
            Dim row2 As Integer = Integer.Parse(cellReference2.Substring(1)) - 1

            ' Acceder a las celdas usando los índices obtenidos
            Dim cell1 As DataGridViewCell = dgvhoja1.Rows(row1).Cells(column1)
            Dim cell2 As DataGridViewCell = dgvhoja1.Rows(row2).Cells(column2)

            ' Verificar si las celdas tienen valores
            If cell1 IsNot Nothing AndAlso cell1.Value IsNot Nothing AndAlso cell2 IsNot Nothing AndAlso cell2.Value IsNot Nothing Then
                ' Intentar convertir los valores de las celdas en fechas
                Dim dateValue1 As DateTime
                Dim dateValue2 As DateTime

                If DateTime.TryParse(cell1.Value.ToString(), dateValue1) AndAlso DateTime.TryParse(cell2.Value.ToString(), dateValue2) Then
                    ' Calcular la diferencia en días
                    Dim timeDifference As TimeSpan = dateValue2 - dateValue1
                    Return timeDifference.Days
                End If
            End If

        Catch ex As Exception
            ' Si ocurre algún error, mostrarlo
            Console.WriteLine("Error en EvaluateElapsedTime: " & ex.Message)
        End Try

        ' Si ocurre algún error o no hay valores válidos, devolver 0
        Return 0
    End Function

    Private Sub dgvhoja1_KeyDown(sender As Object, e As KeyEventArgs) Handles dgvhoja1.KeyDown
        If e.Control AndAlso e.KeyCode = Keys.Z Then
            btnUndo.PerformClick() ' Simula el clic en el botón de Deshacer
        ElseIf e.Control AndAlso e.KeyCode = Keys.Y Then
            btnRedo.PerformClick() ' Simula el clic en el botón de Rehacer
        End If
    End Sub

    Dim undoStack As New Stack(Of Action)()
    Dim redoStack As New Stack(Of Action)()

    Private Sub GuardarAccion(accion As Action)
        undoStack.Push(accion)
        redoStack.Clear() ' Al hacer una nueva acción, vacía el redoStack
    End Sub

    Private Sub btnUndo_Click(sender As Object, e As EventArgs) Handles btnUndo.Click
        If undoStack.Count > 0 Then
            Dim accion As Action = undoStack.Pop()
            redoStack.Push(accion) ' Guarda en redo antes de deshacer
            accion.Invoke() ' Ejecuta la acción de deshacer
        End If
    End Sub

    Private Sub btnRedo_Click(sender As Object, e As EventArgs) Handles btnRedo.Click
        If redoStack.Count > 0 Then
            Dim accion As Action = redoStack.Pop()
            undoStack.Push(accion) ' Guarda en undo antes de rehacer
            accion.Invoke() ' Ejecuta la acción de rehacer
        End If
    End Sub

    Private Sub dgvhoja1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvhoja1.CellValueChanged
        Dim fila As Integer = e.RowIndex
        Dim columna As Integer = e.ColumnIndex
        Dim valorAnterior As Object = dgvhoja1.Rows(fila).Cells(columna).Tag
        Dim valorNuevo As Object = dgvhoja1.Rows(fila).Cells(columna).Value

        ' Guardar la acción en la pila de deshacer
        GuardarAccion(Sub()
                          dgvhoja1.Rows(fila).Cells(columna).Value = valorAnterior
                      End Sub)

        ' Guardar el nuevo valor en Tag para futuros cambios
        dgvhoja1.Rows(fila).Cells(columna).Tag = valorNuevo
    End Sub

    Private Sub cbxOperaciones_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxOperaciones.SelectedIndexChanged

    End Sub

    Private Sub GroupBox3_Enter(sender As Object, e As EventArgs) Handles GroupBox3.Enter

    End Sub
End Class
