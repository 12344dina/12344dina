- 👋 Hi, I’m @12344dina
- 👀 I’m interested in ...
- 🌱 I’m currently learning ...
- 💞️ I’m looking to collaborate on ...
- 📫 How to reach me ...

<!---
12344dina/12344dina is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->
Public Class Form1
    Dim codigo, Codigo_VT, Codigo_C, Codigo_V, Fecha_venta, IGV_VT, Total_venta As String
    Dim numero, mayor, bandera As Integer
    Dim respuesta As MsgBoxResult
    Sub llenarGrid()
        Dim ds As New DataSet
        Dim adp As New OleDb.OleDbDataAdapter("SELECT * FROM Venta", cad)

        ds.Tables.Add("tabla")
        adp.Fill(ds.Tables("tabla"))

        Me.dgvVenta.DataSource = ds.Tables("tabla")

    End Sub
    Sub generarCodigo()
        mayor = 0
        cmd.Connection = cad
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select codigo from Venta"
        Try
            drVenta = cmd.ExecuteReader
            While drVenta.Read
                codigo = drVenta(0).ToString
                numero = Int(Microsoft.VisualBasic.Right(codigo, 3))
                If numero > mayor Then
                    mayor = numero
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        drVenta.Close()
        codigo = Microsoft.VisualBasic.Left("VT000", 5 - Len(Trim(Str(mayor + 1)))) & Trim(Str(mayor + 1))
        txtCodigo_VT.Text = codigo
    End Sub
    Sub generarCodigo()
        mayor = 0
        cmd.Connection = cad
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select codigo from Venta"
        Try
            drVenta = cmd.ExecuteReader
            While drVenta.Read
                codigo = drVenta(0).ToString
                numero = Int(codigo)
                If numero > mayor Then
                    mayor = numero
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        drVenta.Close()
        codigo = Microsoft.VisualBasic.Left("000000", 6 - Len(Trim(Str(mayor + 1)))) & Trim(Str(mayor + 1))
        txtCodigo_VT.Text = codigo
    End Sub
    Sub generarCodigo()
        mayor = 0
        cmd.Connection = cad
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select codigo from Venta"
        Try
            drVenta = cmd.ExecuteReader
            While drVenta.Read
                codigo = drVenta(0).ToString
                numero = Int(Microsoft.VisualBasic.Right(codigo, 3))
                If numero > mayor Then
                    mayor = numero
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        drVenta.Close()
        codigo = Microsoft.VisualBasic.Left("C000", 4 - Len(Trim(Str(mayor + 1)))) & Trim(Str(mayor + 1))
        txtCodigo_C.Text = codigo
    End Sub
    Sub generarCodigo()
        mayor = 0
        cmd.Connection = cad
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select codigo from Venta"
        Try
            drVenta = cmd.ExecuteReader
            While drVenta.Read
                codigo = drVenta(0).ToString
                numero = Int(codigo)
                If numero > mayor Then
                    mayor = numero
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        drVenta.Close()
        codigo = Microsoft.VisualBasic.Left("000000", 6 - Len(Trim(Str(mayor + 1)))) & Trim(Str(mayor + 1))
        txtCodigo_C.Text = codigo
    End Sub
    Sub generarCodigo()
        mayor = 0
        cmd.Connection = cad
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select codigo from Venta"
        Try
            drVenta = cmd.ExecuteReader
            While drVenta.Read
                codigo = drVenta(0).ToString
                numero = Int(Microsoft.VisualBasic.Right(codigo, 3))
                If numero > mayor Then
                    mayor = numero
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        drVenta.Close()
        codigo = Microsoft.VisualBasic.Left("V000", 4 - Len(Trim(Str(mayor + 1)))) & Trim(Str(mayor + 1))
        txtCodigo_V.Text = codigo
    End Sub
    Sub generarCodigo1()
        mayor = 0
        cmd.Connection = cad
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select codigo from Venta"
        Try
            drVenta = cmd.ExecuteReader
            While drVenta.Read
                codigo = drVenta(0).ToString
                numero = Int(codigo)
                If numero > mayor Then
                    mayor = numero
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        drVenta.Close()
        codigo = Microsoft.VisualBasic.Left("000000", 6 - Len(Trim(Str(mayor + 1)))) & Trim(Str(mayor + 1))
        txtCodigo_V.Text = codigo
    End Sub
    Sub activar_botones()
        btnBuscar.Enabled = True
        btnEliminar.Enabled = True
        btnNuevo.Enabled = True
        btnModificar.Enabled = True

        btnAceptar.Enabled = False
        btnCancelar.Enabled = False
    End Sub

    Sub desactivar_botones()
        btnBuscar.Enabled = False
        btnEliminar.Enabled = False
        btnNuevo.Enabled = False
        btnModificar.Enabled = False

        btnAceptar.Enabled = True
        btnCancelar.Enabled = True
    End Sub
    Sub activar_cuadros()
        txtCodigo_VT.ReadOnly = False
        txtCodigo_C.ReadOnly = False
        txtCodigo_V.ReadOnly = False
        dtpfi.Enabled = False
        txt_IGV_VT.Enabled = False
        txtTotal_venta.ReadOnly = True
    End Sub
    Sub desactivar_cuadros()
        txtCodigo_VT.ReadOnly = True
        txtCodigo_C.ReadOnly = True
        txtCodigo_V.ReadOnly = True
        dtpfi.Enabled = True
        txt_IGV_VT.Enabled = True
        txtTotal_venta.ReadOnly = False
    End Sub
    Sub mostrar()
        cmd.Connection = cad
        cmd.CommandType = CommandType.Text
        SQL = "select * from venta"
        cmd.CommandText = SQL
        Try
            drVenta = cmd.ExecuteReader
            If drVenta.HasRows Then
                While drVenta.Read
                    txtCodigo_VT.Text = drVenta(0).ToString
                    txtCodigo_C.Text = drVenta(1).ToString
                    txtCodigo_V.Text = drVenta(2).ToString
                    dtpfi.Value = drVenta(3).ToString
                    IGV_VT.Value = drVenta(4).ToString
                    txtTotal_venta.Text = drVenta(5).ToString
                End While
            Else
                MsgBox("la tabla esta vacia")
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        drVenta.Close()
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        activar_botones()
        desactivar_cuadros()
        conectar()
        mostrar()
        llenarGrid()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        desactivar_botones()
        activar_cuadros()
        generarCodigo()
        bandera = 1
        txtCodigo_C.Text = ""
        txtCodigo_V.Text = ""
        txtCodigo_C.Focus()
    End Sub

    Private Sub btnModificar_Click(sender As Object, e As EventArgs) Handles btnModificar.Click
        desactivar_botones()
        activar_cuadros()
        bandera = 2
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click

        If bandera = 1 Then
            'boton NUEVO
            Codigo_VT = txtCodigo_VT.Text
            Codigo_C = txtCodigo_C.Text
            Codigo_V = txtCodigo_V.Text
            Fecha_venta = dtpfi.Value
            IGV_VT = txt_IGV_VT.Text
            Total_venta = txtTotal_venta.Text
            If Codigo_C = "" Then
                MsgBox("Debe ingresar CODIGO DE CLIENTE")
                txtCodigo_C.Focus()
                Exit Sub
            End If
            If Codigo_V = "" Then
                MsgBox("Debe ingresar codigo de venta")
                txtCodigo_V.Focus()
                Exit Sub
            Else
                If Len(Codigo_V) < 6 Then
                    MsgBox("la clave debe contener 5 caracteres")
                    txtCodigo_V.Focus()
                    Exit Sub
                End If
            End If
            cmd.Connection = cad
            cmd.CommandType = CommandType.Text
            SQL = "insert into venta values('" & Codigo_VT & "','" & Codigo_C & "','" & Codigo_V & "','" & Fecha_venta & "','" & IGV_VT & "','" & Total_venta & "')"
            cmd.CommandText = SQL
            Try
                cmd.ExecuteNonQuery()
                MsgBox("Registro insertado con exito")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            activar_botones()
            desactivar_cuadros()
            llenarGrid()
        Else
            'boton MODIFICAR
            Codigo_VT = txtCodigo_VT.Text
            Codigo_C = txtCodigo_C.Text
            Codigo_V = txtCodigo_V.Text
            Fecha_venta = dtpfi.Value
            IGV_VT = txt_IGV_VT.Text
            Total_venta = txtTotal_venta.Text
            If Codigo_C = "" Then
                MsgBox("Debe ingresar CODIGO DE CLIENTE")
                txtCodigo_C.Focus()
                Exit Sub
            End If
            If Codigo_V = "" Then
                MsgBox("Debe ingresar codigo de venta")
                txtCodigo_V.Focus()
                Exit Sub
            Else
                If Len(Codigo_V) < 6 Then
                    MsgBox("la clave debe contener 5 caracteres")
                    txtCodigo_V.Focus()
                    Exit Sub
                End If
            End If
            cmd.Connection = cad
            cmd.CommandType = CommandType.Text
            SQL = "update venta set Fecha_venta='" & Fecha_venta
            SQL = SQL + "', IGV_VT ='" & IGV_VT & "'Where codigo_VT='" & Codigo_VT & "','" & Codigo_C & "','" & Codigo_V & "'"
            cmd.CommandText = SQL
            Try
                cmd.ExecuteNonQuery()
                MsgBox("Registro modificado con exito")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            activar_botones()
            desactivar_cuadros()
            llenarGrid()
        End If
    End Sub

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        activar_botones()
        desactivar_cuadros()
        mostrar()
        llenarGrid()
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        Codigo_V = txtCodigo_VT.Text
        cmd.Connection = cad
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "delete from venta where codigo_VT='" & Codigo_VT & "','" & Codigo_C & "','" & Codigo_V & "'"
            respuesta = MsgBox("Esta seguro de eliminar el Registro??", MsgBoxStyle.YesNo, "Advertencia")
            If respuesta = MsgBoxResult.Yes Then
                cmd.ExecuteNonQuery()
                MsgBox("Registro eliminado")
                mostrar()
            Else
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        llenarGrid()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Codigo_V = InputBox("Ingrese el codigo de venta", "Buscar Registro")
        cmd.Connection = cad
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select * from venta where codigo_VT='" & Codigo_VT & "','" & Codigo_C & "','" & Codigo_V & "'"
        Try
            drVenta = cmd.ExecuteReader
            If drVenta.HasRows Then
                While drVenta.Read
                    txtCodigo_VT.Text = drVenta(0).ToString
                    txtCodigo_C.Text = drVenta(1).ToString
                    txtCodigo_V.Text = drVenta(2).ToString
                    dtpfi.Value = drVenta(3).ToString
                    IGV_VT.Value = drVenta(4).ToString
                    txtTotal_venta.Text = drVenta(5).ToString
                End While
            Else
                MsgBox("No se encontro el registro con la clave " & Codigo_VT ,  Codigo_C , Codigo_V &
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        drVenta.Close()
    End Sub

End Class
corrigue este codigo
