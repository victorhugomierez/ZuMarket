

Private Sub BT_buscar_Click()
Dim unfila As Variant
Dim fila As Variant

Me.Lista.RowSource = Clear

unfila = Sheets("BaseInventario").Range("B" & Rows.Count).End(xlUp).Row
For fila = 6 To unfila

If Range("b" & fila) Like Me.TextBuscar.Value Then

Me.Lista.AddItem Cells(fila, 1)

Me.Lista.List(Lista.ListCount - 1, 0) = Cells(fila, 2)
Me.Lista.List(Lista.ListCount - 1, 1) = Cells(fila, 3)
Me.Lista.List(Lista.ListCount - 1, 2) = Cells(fila, 4)
Me.Lista.List(Lista.ListCount - 1, 3) = Cells(fila, 5)
Me.Lista.List(Lista.ListCount - 1, 4) = Cells(fila, 6)
Me.Lista.List(Lista.ListCount - 1, 5) = Cells(fila, 7)
Me.Lista.List(Lista.ListCount - 1, 6) = Cells(fila, 8)

End If
Next fila


End Sub



Private Sub BT_Eliminar_Click()

Dim fila As Object
Dim linea As Variant
Dim IdBuscado As Variant

If Me.Lista.ListIndex = -1 Then
MsgBox "¡Amor! PRIMERO SELECCIONE EL PRODUCTO A ELIMINAR"
Else
Me.TextCodigoProducto.Value = Me.Lista

IdBuscar = Me.TextCodigoProducto

Set fila = Sheets("BaseInventario").Range("B:B").Find(IdBuscar, lookat:=xlWhole)
linea = fila.Row
Range("B" & linea).EntireRow.Delete

End If
MsgBox "¡Amor! ELIMINASTE EL PRODUCTO"
Me.Lista.RowSource = "BaseInventario"
Me.Lista.ColumnCount = 7


End Sub

Private Sub BT_Guardar_Click()
 Sheets("BaseInventario").Range("a7").EntireRow.Insert , copyorigin:=xlFormatFromRightOrBelow
 
 Range("b7").Value = Me.TextCodigoProducto.Value
 Range("c7").Value = Me.TextNombreProducto.Value
 Range("d7").Value = Me.TextCantidadStock.Value
 Range("e7").Value = Me.TextPrecioUnidad.Value
 Range("f7").Value = Me.TextFechaAdquisicion.Value
 Range("g7").Value = Me.TextFechaCaducidad.Value
 Range("h7").Value = Me.TextValorFinal.Value
 
 Me.Lista.RowSource = "BaseInventario"
 Me.Lista.ColumnCount = 7
 
 Me.TextCodigoProducto = Empty
Me.TextNombreProducto = Empty
Me.TextCantidadStock = Empty
Me.TextPrecioUnidad = Empty
Me.TextFechaAdquisicion = Empty
Me.TextFechaCaducidad = Empty
Me.TextValorFinal = Empty
 
 
End Sub

Private Sub BT_IrEditar_Click()
Dim fila As Object
Dim linea As Variant
Dim ValorBuscado As Variant


If Me.Lista.ListIndex = -1 Then
MsgBox "¡AMOR! TE FALTO SELECCIONAR EL PRODUCTO ANTES DE EDITAR"
Else

Me.TextCodigoProducto = Me.Lista
ValorBuscado = Me.TextCodigoProducto
Set fila = Sheets("BaseInventario").Range("B:B").Find(ValorBuscado, lookat:=xlWhole)
Lines = fila.Row
MsgBox (Lines)
Me.TextNombreProducto.Value = Range("c" & Lines).Value
Me.TextCantidadStock.Value = Range("d" & Lines).Value
Me.TextPrecioUnidad.Value = Range("e" & Lines).Value
Me.TextFechaAdquisicion.Value = Range("f" & Lines).Value
Me.TextFechaCaducidad.Value = Range("g" & Lines).Value
Me.TextValorFinal.Value = Range("h" & Lines).Value


End If
End Sub

Private Sub BT_MostrarCampos_Click()

Me.Height = 489.75

End Sub

Private Sub BT_OcultarCampos_Click()

Me.Height = 320.25

End Sub

Private Sub PlanillaExcel_Click()

Application.Visible = True
Unload Me

End Sub

Private Sub TextGuardarCambios_Click()
Dim fila As Object
Dim linea As Variant
Dim ValorBuscado As Variant
Me.TextCodigoProducto = Me.Lista
ValorBuscado = Me.TextCodigoProducto
Set fila = Sheets("BaseInventario").Range("B:B").Find(ValorBuscado, lookat:=xlWhole)
Lines = fila.Row

Range("c" & Lines).Value = Me.TextNombreProducto.Value
Range("d" & Lines).Value = Me.TextCantidadStock.Value
Range("e" & Lines).Value = Me.TextPrecioUnidad.Value
Range("f" & Lines).Value = Me.TextFechaAdquisicion.Value
Range("g" & Lines).Value = Me.TextFechaCaducidad.Value
Range("h" & Lines).Value = Me.TextValorFinal.Value

Me.TextCodigoProducto = Empty
Me.TextNombreProducto = Empty
Me.TextCantidadStock = Empty
Me.TextPrecioUnidad = Empty
Me.TextFechaAdquisicion = Empty
Me.TextFechaCaducidad = Empty
Me.TextValorFinal = Empty

MsgBox "LO HICISTE BIEN AMOR, SE MODIFICO CON EXITO"
Me.Lista.RowSource = "BaseInventario"
Me.Lista.ColumnCount = 7


End Sub

Private Sub UserForm_Initialize()
Me.Lista.RowSource = "BaseInventario"
Me.Lista.ColumnCount = 7
Me.Lista.ColumnHeads = True
Me.Lista.ColumnWidths = "150;150;150;150;150;150;190"

End Sub
