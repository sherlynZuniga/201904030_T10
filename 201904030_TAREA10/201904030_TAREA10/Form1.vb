Public Class Form1
    Private Sub OperarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OperarToolStripMenuItem.Click

        'operar
        'asignando valores a vectores
        dpi(fila) = Val(TextBox1.Text)
        nombre(fila) = TextBox2.Text
        art(fila) = Val(TextBox3.Text)
        meses(fila) = Val(TextBox4.Text)

        'operando
        Select Case (Val(TextBox4.Text))
            Case < 12
                interes(fila) = (Val(TextBox4.Text) * 3 / 100)
            Case 12 To 24
                interes(fila) = (Val(TextBox4.Text) * 4.5 / 100)
            Case 25 To 36
                interes(fila) = (Val(TextBox4.Text) * 5.5 / 100)
            Case > 37
                interes(fila) = (Val(TextBox4.Text) * 6.2 / 100)
        End Select

        'descuento 
        If (Val(art(fila)) > 8000) And (Val(meses(fila) <= 12)) Then
            interes(fila) = (Val(TextBox4.Text) * 2 / 100)
        End If

        'valor total 
        total(fila) = Val(art(fila)) + Val(interes(fila))


    End Sub

    Private Sub MostrarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MostrarToolStripMenuItem.Click

        'mostrar
        DataGridView1.Rows.Add(dpi(fila), nombre(fila), art(fila), meses(fila), interes(fila), total(fila))


    End Sub

    Private Sub OrdenarDescToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OrdenarDescToolStripMenuItem.Click

        'ordenar desc
        Dim i As Byte
        Dim j As Byte
        Dim temp1
        Dim temp2
        Dim temp3
        Dim temp4
        Dim temp5
        Dim temp6

        For i = 0 To 5
            For j = i + 1 To 6
                If (dpi(j) <> Nothing) Then
                    If (dpi(i) < dpi(j)) Then
                        temp1 = dpi(i)
                        dpi(i) = dpi(j)
                        dpi(j) = temp1

                        temp2 = nombre(i)
                        nombre(i) = nombre(j)
                        nombre(j) = temp2

                        temp3 = art(i)
                        art(i) = art(j)
                        art(j) = temp3

                        temp4 = meses(i)
                        meses(i) = meses(j)
                        meses(j) = temp4

                        temp5 = interes(i)
                        interes(i) = interes(j)
                        interes(j) = temp5

                        temp6 = total(i)
                        total(i) = total(j)
                        total(j) = temp6
                    End If
                Else
                    Exit For
                End If
            Next j
        Next i


    End Sub

    Private Sub LimpiarEntradasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LimpiarEntradasToolStripMenuItem.Click

        'limpiar entradas
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        DataGridView1.Rows.Clear()


    End Sub

    Private Sub LimpiarVectoresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LimpiarVectoresToolStripMenuItem.Click

        'limpiar vectores
        fila = 0
        dpi(fila) = 0
        nombre(fila) = 0
        art(fila) = 0
        meses(fila) = 0
        interes(fila) = 0
        total(fila) = 0



    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'salir
        Form2.Show()


    End Sub
End Class
