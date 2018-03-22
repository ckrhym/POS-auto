Imports MySql.Data.MySqlClient
Module database
    Public attempts
    Public qty As Integer
    Dim strs = ""
    Dim ident As Integer
    Dim rd As MySqlDataReader
    Private myconn As MySqlConnection
    Private command As MySqlCommand
    Public datatable As DataTable
    Private adapter As MySqlDataAdapter
    Private reader As MySqlDataReader
    Private strconn = "server=localhost;user=root;database=dbmbaj"
    Public name
    Public price = 0.00
    Public description = ""
    Public code = ""
    Dim strquery = ""
    Dim remainingstock = 0
    Public Sub connect()
        myconn = New MySqlConnection(strconn)
        myconn.Open()
        myconn.Close()
    End Sub
    Public Function getstatus(ByVal sql As String) As String
        Dim this = ""
        'myconn.Open()
        'command = New MySqlCommand(sql, myconn)
        'rd = command.ExecuteReader
        Try
            rd.Read()
            this = rd.GetValue(8)
        Catch ex As Exception

        End Try
        ' myconn.Close()
        rd.Close()
        Return this
    End Function
    Public Sub sqlManager(ByVal query As String, ByVal msg As String)
        myconn.Open()
        command = New MySqlCommand(query, myconn)
        With command
            .CommandType = CommandType.Text
            .ExecuteNonQuery()
        End With
        MessageBox.Show(msg, "MBAJ", MessageBoxButtons.OK, MessageBoxIcon.Information)
        myconn.Close()
    End Sub
    Public Sub sqlManager2(ByVal query As String)
        myconn.Open()
        command = New MySqlCommand(query, myconn)
        With command
            .CommandType = CommandType.Text
            .ExecuteNonQuery()
        End With
        ' MessageBox.Show(msg, "MBAJ")
        myconn.Close()
    End Sub
    Public Sub display(ByVal query As String, ByVal datagrid As DataGridView)
        myconn.Open()
        adapter = New MySqlDataAdapter(query, myconn)
        datatable = New DataTable
        adapter.Fill(datatable)
        datagrid.DataSource = datatable
        myconn.Close()
    End Sub
    Public Sub closeconn()
        myconn.Close()
    End Sub
    Public Sub openconn()
        myconn.Open()
    End Sub
    Public Function returnDT(ByVal sql As String) As DataTable
        myconn.Open()
        adapter = New MySqlDataAdapter(sql, myconn)
        datatable = New DataTable
        adapter.Fill(datatable)
        myconn.Close()
        Return datatable
    End Function

    Public Sub login(ByVal txtuser As TextBox, ByVal txtpass As TextBox)
        myconn.Open()
        Dim str = "select * from users where username='" & txtuser.Text & "' and password='" & frmUsers.Encrypt(txtpass.Text, "StarterPack") & "'"
        command = New MySqlCommand(str, myconn)
        rd = command.ExecuteReader
        Dim count = 0
        Dim id = ""
        Dim st
        Try
            Dim status = False
            ' Dim name = ""
            Dim type = 0
            While rd.Read
                count = count + 1
                id = rd.GetValue(0).ToString
                frmLogin.cashierID = rd.GetValue(0).ToString
                name = rd.GetValue(1) + " " + rd.GetValue(3)
                type = rd.GetValue(9)
                st = rd.GetValue(10)
                ' Form1.Lbluser.Text = rd.GetValue(3)
            End While
            If count = 1 Then
                If type = 1 Then
                    attempts = 0
                    If st = 1 Then
                        frmMain.Show()
                        frmLogin.Hide()
                        txtpass.Text = ""
                        txtuser.Text = ""
                        frmCashier.lblname.Text = name.ToString
                        frmMain.admin.Text = name.ToString
                        frmCashier.lblid.Text = id
                    Else
                        MsgBox("The account is inactive", MessageBoxIcon.Information, "MBAJ")
                        txtpass.Text = ""
                        txtuser.Text = ""
                    End If
                ElseIf type = 2
                    If st = 1 Then
                        frmCashier.lblname.Text = name.ToString
                        frmCashier.Show()
                        frmLogin.Hide()
                        txtpass.Text = ""
                        txtuser.Text = ""
                        frmCashier.btnmenucash.Visible = False
                        frmCashier.lblid.Text = id
                    Else
                        MsgBox("The account is inactive", MessageBoxIcon.Information, "MBAJ")
                        txtpass.Text = ""
                        txtuser.Text = ""
                    End If
                Else
                    MsgBox("Invalid Account", MessageBoxIcon.Information, "MBAJ")
                End If
            Else
                attempts += 1
                If attempts < 3 Then
                    MsgBox("username or password is incorrect", MessageBoxIcon.Information, "MBAJ")
                ElseIf attempts = 3
                    MsgBox("Too many attempts", MessageBoxIcon.Error, "MBAJ")
                    frmLogin.Dispose()
                End If
                txtpass.Text = ""
                txtuser.Text = ""
                txtuser.Select()

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        myconn.Close()
    End Sub

    Public Sub trycode(ByVal txtcode As TextBox)
        myconn.Open()
        Dim stocks
        Dim str = "select * from viewcashier where id=" & txtcode.Text & " and stat=1"
        command = New MySqlCommand(str, myconn)
        rd = command.ExecuteReader
        Dim count = 0

        Try
            Dim num As Integer = 0
            While rd.Read
                stocks = rd.GetValue(3)
                count = count + 1
                code = rd.GetValue(0) ' code
                price = rd.GetValue(4) ' price
                description = rd.GetValue(1).ToString & " " + rd.GetValue(2).ToString & " " & rd.GetValue(5).ToString ' name + descripton + brand
            End While

            If count = 1 Then
                If stocks = 0 Then
                    MessageBox.Show("Out of Stock", "MBAJ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    frmCashier.txtcode.Text = ""
                Else
                    frmQTY.ShowDialog()
                End If
            Else
                MessageBox.Show("Invalid product code", "MBAJ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                frmCashier.txtcode.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        myconn.Close()
    End Sub
    Public Sub qtyTry(ByVal txtcode As TextBox, ByVal qty As TextBox)
        myconn.Close()
        getcountstock(txtcode.Text)
        myconn.Open()
        Try
            If CInt(qty.Text) > remainingstock Then
                strs = "no"
            Else
                strs = "yes"
            End If
        Catch ex As Exception
            strs = "no"
        End Try
        If strs = "no" Then
            If CInt(qty.Text) > remainingstock Then
                MessageBox.Show("Not Enough Stock", "MBAJ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                MessageBox.Show("Invalid Quantity", "MBAJ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            Dim productid1 As String
            Dim found = False
            Dim rowcount = 0
            For i As Integer = 0 To frmCashier.dtgtrans.Rows.Count - 1
                productid1 = frmCashier.dtgtrans.Rows(i).Cells(0).Value.ToString()
                If txtcode.Text = productid1 Then
                    found = True
                    rowcount = i
                End If
            Next
            If found = True Then
                myconn.Close()
                getcountstock(txtcode.Text)
                myconn.Open()
                If remainingstock < frmCashier.dtgtrans.Rows(rowcount).Cells("qty").Value + CInt(Val(qty.Text)) Then
                    MessageBox.Show("Not Enough stock", "MBAJ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    frmCashier.dtgtrans.Rows(rowcount).Cells("qty").Value += Val(CInt(qty.Text))
                    frmCashier.dtgtrans.Rows(rowcount).Cells("subtotal").Value = Format(CDbl(frmCashier.dtgtrans.Rows(rowcount).Cells("price").Value.ToString) * CDbl(frmCashier.dtgtrans.Rows(rowcount).Cells("qty").Value.ToString), "0.00")
                    frmCashier.txtcode.Text = ""
                End If

            Else
                frmCashier.dtgtrans.Rows.Add(code, Format(price, "0.00"), description, CInt(qty.Text), Format((price * CInt(qty.Text)), "0.00"))
                frmCashier.txtcode.Text = ""
            End If
        End If
        frmQTY.txtqty.Text = "1"
        frmQTY.Close()
    End Sub
    Public Sub productCode(ByVal txtcode As TextBox)
        myconn.Open()
        Dim str = "select * from product where id=" & txtcode.Text & ""
        command = New MySqlCommand(str, myconn)
        rd = command.ExecuteReader
        Dim count = 0
        Try
            While rd.Read
                count = count + 1
            End While
            If count = 1 Then
                Dim row = 0
                Dim found = False
                For i As Integer = 0 To frmaddstocks.dtgStock.Rows.Count - 1
                    Dim productid1 As String = frmaddstocks.dtgStock.Rows(i).Cells("id").Value.ToString()
                    If productid1 = frmaddstocks.txtcode.Text And frmaddstocks.dtgStock.Rows.Count <> 0 Then
                        row = i
                        found = True
                    End If
                Next
                If found = True Then
                    frmInfo.price.Text = frmaddstocks.dtgStock.Rows(row).Cells(3).Value
                    frmInfo.price.Enabled = False
                    frmInfo.Show()
                Else
                    frmInfo.Show()
                End If
            Else
                MsgBox("Product code no match")
                frmaddstocks.txtcode.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        myconn.Close()
    End Sub
    Public Sub toGrid(ByVal dg As DataGridView, ByVal query As String)
        myconn.Open()
        command = New MySqlCommand(query, myconn)
        rd = command.ExecuteReader
        Dim count = 0
        Dim name = ""
        Try
            Dim num As Integer = 0
            While rd.Read
                count = count + 1
                code = rd.GetValue(0) ' code
                name = rd.GetValue(1) ' name
                description = rd.GetValue(2) ' description
            End While
            Dim strs = ""
            If count = 1 Then
                Try
                Catch ex As Exception
                    strs = "no"
                End Try
                If strs = "no" Then
                    MsgBox("invalid")
                    frmaddstocks.txtcode.Text = ""
                Else
                    Dim row = 0
                    Dim found = False
                    For i As Integer = 0 To dg.Rows.Count - 1
                        Dim productid1 As String = dg.Rows(i).Cells("id").Value.ToString()
                        If productid1 = frmaddstocks.txtcode.Text Then
                            row = i
                            found = True
                        End If
                    Next
                    If found = True Then
                        dg.Rows(row).Cells("qty").Value += Val(frmInfo.qty.Text)
                        frmaddstocks.dtgStock.Rows(row).Cells(5).Value = CDbl(frmaddstocks.dtgStock.Rows(row).Cells(4).Value) * CDbl(frmaddstocks.dtgStock.Rows(row).Cells(3).Value.ToString)
                        frmaddstocks.txtcode.Text = ""
                        frmInfo.Close()
                    Else
                        frmInfo.price.Enabled = True
                        dg.Rows.Add(code, name, description, frmInfo.price.Text, frmInfo.qty.Text, CDbl(frmInfo.price.Text) * CDbl(frmInfo.qty.Text))
                        frmaddstocks.txtcode.Text = ""
                        frmInfo.Close()
                    End If
                End If
            Else
                MessageBox.Show("Invalid product code", "MBAJ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                frmaddstocks.txtcode.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        myconn.Close()
    End Sub
    Public Sub trap(ByVal e As KeyPressEventArgs)
        '--------------Trap
        Dim Sn As String = "0123456789"
        Dim So As String = Chr(8) & Chr(13) & Chr(1) & Chr(3) & Chr(22)
        If Not (Sn & So).Contains(e.KeyChar) Then
            e.Handled = True
        End If
        '--------------
    End Sub
    Public Sub trapdouble(ByVal e As KeyPressEventArgs)
        '--------------Trap
        Dim Sn As String = "0123456789."
        Dim So As String = Chr(8) & Chr(13) & Chr(1) & Chr(3) & Chr(22)
        If Not (Sn & So).Contains(e.KeyChar) Then
            e.Handled = True
        End If
        '--------------
    End Sub
    Public Sub stockin(ByVal dt As DataGridView, ByVal code As TextBox, ByVal supid As ComboBox, ByVal totals As TextBox)
        Try
            strquery = "insert into stockin (code, supplier_id, Date, total) values (" & code.Text & ", " & supid.SelectedValue & ",'" & Format(Now, "yyyy-MM-dd") & "'," & totals.Text & ")"
            sqlManager2(strquery)
            For Each dr As DataGridViewRow In dt.Rows
                strquery = "insert into stockin_details (prod_id,stockin_code,qty,price,subtotal) values (" & dr.Cells(0).Value.ToString & "," & code.Text & ", " & dr.Cells(4).Value & ", " & dr.Cells(3).Value & ", " & CDbl(dr.Cells(3).Value.ToString) * CDbl(dr.Cells(4).Value.ToString) & ")"
                sqlManager2(strquery)
                updatestock(dr.Cells(0).Value.ToString, dr.Cells(4).Value.ToString)
            Next
            MessageBox.Show("Stockin Complete", "MBAJ", MessageBoxButtons.OK, MessageBoxIcon.Information)
            display("select * from viewproducts", frmProduct.dtgproduct)
            frmaddstocks.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        myconn.Close()
    End Sub
    Public Sub updatestock(ByVal id As String, ByVal qty As String)
        strquery = "update product set stocks=stocks+" & qty & " where id=" & id & ""
        sqlManager2(strquery)
    End Sub
    Public Sub stockout(ByVal id As String, ByVal qty As String)
        strquery = "update product set stocks=stocks-" & qty & " where id=" & id & ""
        sqlManager2(strquery)
    End Sub
    Public Sub endtrans(ByVal dt As DataGridView, ByVal id As Label, ByVal total As Label, ByVal userid As Label)
        Try
            strquery = "insert into transaction(id,date,total,user_id,time) values (" & id.Text & ", '" & Format(Now, "yyyy-MM-dd") & "'," & total.Text & "," & userid.Text & ",'" & Format(Now, "hh:mm tt") & "')"
            sqlManager2(strquery)
            For Each dr As DataGridViewRow In dt.Rows
                strquery = "insert into transaction_details (trans_id,prod_id,price,qty) values (" & id.Text & "," & dr.Cells(0).Value & ", " & dr.Cells(1).Value & ", " & dr.Cells(3).Value & ")"
                sqlManager2(strquery)
                stockout(dr.Cells(0).Value.ToString, dr.Cells(3).Value.ToString)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        myconn.Close()
    End Sub
    Public Sub getCount(ByVal cnt As Label)
        myconn.Open()
        command = New MySqlCommand("select count(id) as trans from transaction", myconn)
        rd = command.ExecuteReader
        Dim count = 0
        Try
            Dim num As Integer = 0
            While rd.Read
                count = rd.GetValue(0) + 1
            End While
        Catch ex As Exception

        End Try
        cnt.Text = count
        myconn.Close()
    End Sub
    Public Sub newtrans(ByVal dt As DataGridView, cnt As Label)
        frmReceipt.ShowDialog()
        dt.Rows.Clear()
        getCount(cnt)
    End Sub
    Public Sub printReciept(ByVal DT As DataGridView)
        Dim total As Double = 0.00
        Dim str = "MBAJ AUTO PARTS,MOTOR PARTS and ACCESSORIES" & vbNewLine & vbNewLine & "                                 OFFICIAL RECEIPT" & vbNewLine &
            "-----------------------------------------------------------------" & vbNewLine & "Code     Description        Price     Qty     Subtotal" & vbNewLine
        For Each dr As DataGridViewRow In DT.Rows
            str = str & dr.Cells(0).Value.ToString & "     " & dr.Cells(2).Value.ToString & "     " & dr.Cells(1).Value.ToString & "     " & dr.Cells(3).Value.ToString & "     " & dr.Cells(4).Value.ToString & vbNewLine
            total += CDbl(dr.Cells(4).Value.ToString)
        Next
        Dim amount As Double = frmCashier.lblamount.Text
        MsgBox(str & vbNewLine & "Total:" & total.ToString("C2") & "              Amount:" & amount.ToString("C2") & "          Change:" & CDbl(frmCashier.lblchange.Text).ToString("C2") & vbNewLine & "-----------------------------------------------------------------")
    End Sub
    Public Sub getcountstock(ByVal id As String)
        myconn.Open()
        Dim count = 0
        Try
            command = New MySqlCommand("select * from product where id=" & id & "", myconn)
            rd = command.ExecuteReader
            Dim num As Integer = 0
            While rd.Read
                count = rd.GetValue(3)
            End While
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        frmCashier.stocks = count
        remainingstock = count
        myconn.Close()
    End Sub
    Public Sub getcountstock2(ByVal query As String)
        myconn.Open()
        Dim count = 0
        Try
            command = New MySqlCommand(query, myconn)
            rd = command.ExecuteReader
            Dim num As Integer = 0
            While rd.Read
                count = rd.GetValue(3)
            End While
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        frmCashier.stocks = count
        remainingstock = count
        myconn.Close()
    End Sub
    Public Sub authenticateCancel(ByVal pass As String)
        myconn.Open()
        Dim str = "select * from users where password='" & frmUsers.Encrypt(pass, "StarterPack") & "' and type_id=1"
        command = New MySqlCommand(str, myconn)
        rd = command.ExecuteReader
        Dim count = 0
        Try
            While rd.Read
                count = count + 1
            End While
            If count = 1 Then
                frmCashier.dtgtrans.Rows.Clear()
                frmCashier.reset()
                frmAuth.Close()
                frmAuth.txtpass.Text = ""
            Else
                MessageBox.Show("Invalid", "MBAJ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                frmAuth.Close()
                frmAuth.txtpass.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        myconn.Close()
    End Sub
    Public Sub authenticateVoid(ByVal pass As String)
        myconn.Open()
        Dim str = "select * from users where password='" & frmUsers.Encrypt(pass, "StarterPack") & "' and type_id=1"
        command = New MySqlCommand(str, myconn)
        rd = command.ExecuteReader
        Dim count = 0
        Try
            While rd.Read
                count = count + 1
            End While
            If count = 1 Then
                frmCashier.dtgtrans.Rows.RemoveAt(frmCashier.index)
                frmCashier.getTotal()
                frmAuth.Close()
                frmAuth.txtpass.Text = ""
            Else
                frmAuth.Close()
                MessageBox.Show("Invalid", "MBAJ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                frmAuth.txtpass.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        myconn.Close()
    End Sub
    Public Sub authenticateReturn(ByVal pass As String)
        myconn.Open()
        Dim str = "select * from users where password='" & frmUsers.Encrypt(pass, "StarterPack") & "' and type_id=1"
        command = New MySqlCommand(str, myconn)
        rd = command.ExecuteReader
        Dim count = 0
        Try
            While rd.Read
                count = count + 1
            End While
            If count = 1 Then
                frmReturn.ShowDialog()
                frmAuth.Close()
                frmAuth.txtpass.Text = ""
            Else
                frmAuth.Close()
                MessageBox.Show("Invalid", "MBAJ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                frmAuth.txtpass.Text = ""
            End If
        Catch ex As Exception
        End Try
        myconn.Close()
    End Sub
    Public Sub returnCode(ByVal code As TextBox, ByVal lblname As Label, ByVal lbldes As Label, ByVal price As Label)
        myconn.Open()
        Dim str = "select name,description,SRP from product where id=" & code.Text & ""
        command = New MySqlCommand(str, myconn)
        rd = command.ExecuteReader
        Dim count = 0
        Dim name = ""
        Dim description = ""
        Dim prices As Double
        Try
            While rd.Read
                count = count + 1
                name = rd.GetValue(0)
                description = rd.GetValue(1)
                prices = rd.GetValue(2)
            End While
            If count = 1 Then
                lblname.Text = name
                lbldes.Text = description
                price.Text = prices.ToString("C2")
            Else
                lblname.Text = "No items"
                lbldes.Text = "Found"
                price.Text = "."
            End If
        Catch ex As Exception
        End Try
        myconn.Close()
    End Sub
    Public Sub createPO(ByVal dt As DataGridView, ByVal cmb As ComboBox, ByVal name As Label, ByVal transID As Label)
        Try
            strquery = "insert into purchaseorder (Preparedby,Supplier_ID,Date,remarks) values ('" & name.Text & "', " & cmb.SelectedValue & ",'" & Format(Now, "yyyy-MM-dd") & "','" & "Pending" & "')"
            sqlManager2(strquery)
            For Each dr As DataGridViewRow In dt.Rows
                strquery = "insert into po_details (trans_id,prod_id,qty) values (" & transID.Text & ", " & dr.Cells(0).Value & "," & dr.Cells(3).Value & ")"
                sqlManager2(strquery)
            Next
            MessageBox.Show("Creating Purchase Order Complete", "MBAJ", MessageBoxButtons.OK, MessageBoxIcon.Information)
            frmPO.Dispose()
            'frmPOSubmit.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub countPO(ByVal cnt As Label)
        myconn.Open()
        command = New MySqlCommand("select count(id) as trans from purchaseorder", myconn)
        rd = command.ExecuteReader
        Dim count = 0
        Try
            Dim num As Integer = 0
            While rd.Read
                count = rd.GetValue(0) + 1
            End While
        Catch ex As Exception
        End Try
        cnt.Text = count
        myconn.Close()
    End Sub
    Public Sub getCountPO(ByVal cnt As Label, ByVal query As String)
        myconn.Open()
        command = New MySqlCommand(query, myconn)
        rd = command.ExecuteReader
        Dim count = 0
        Try
            Dim num As Integer = 0
            While rd.Read
                count = rd.GetValue(0) + 1
            End While
        Catch ex As Exception

        End Try
        cnt.Text = count
        myconn.Close()
    End Sub
    Public Sub getSuppID(ByVal query As String, ByVal cnt As Label)
        myconn.Open()
        command = New MySqlCommand(query, myconn)
        rd = command.ExecuteReader
        Dim count = 0
        Try
            Dim num As Integer = 0
            While rd.Read
                count = rd.GetValue(0)
            End While
        Catch ex As Exception
        End Try
        cnt.Text = count
        myconn.Close()
    End Sub
    Public Sub POgrid(ByVal query As String, ByVal dt As DataGridView)
        myconn.Open()
        command = New MySqlCommand(query, myconn)
        rd = command.ExecuteReader
        Dim count = 0
        Try
            Dim num As Integer = 0
            While rd.Read
                dt.Rows.Add(rd.GetValue(0), rd.GetValue(1), rd.GetValue(2), 10)
            End While
        Catch ex As Exception
        End Try
        myconn.Close()
    End Sub
End Module