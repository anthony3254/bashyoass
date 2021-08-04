Imports System.Data
Imports System.Data.SqlClient
Imports FirebirdSql.Data
Imports FirebirdSql.Data.FirebirdClient

Public Class f_EMOS_list_order

    Private Sub f_EMOS_list_order_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim da As New FbDataAdapter
        Dim dt As New DataTable
        'mengambil 4 hari sebelum sebagai tanggal mulai
        DateTimePicker1.Value = Format(DateAdd("d", -5, CDate(Now())), "yyyy-MM-dd")
        DateTimePicker2.Value = Format(CDate(Now()), "yyyy-MM-dd")

        'mengambil kode subdist
        Dim kodesubdist As String
        If f_menu.ToolStripKodeSubdistAwal.Text = "-" Then
            kodesubdist = f_menu.ToolstripKodesubdist.Text
        Else
            kodesubdist = f_menu.ToolStripKodeSubdistAwal.Text
        End If

        'lihat apakah masih ada faktur EMOS yang masih belum terkirim
        BukaKoneksi()
        da = New FbDataAdapter("select count(distinct invoiceno) CT from ARINVDET d left outer join ARINV a on a.ARINVOICEID = d.ARINVOICEID left outer join sodet sd on sd.soid = d.soid and d.soseq = sd.SEQ left outer join so sod on sod.soid = d.soid left outer join extended e1 on e1.EXTENDEDID = sd.EXTENDEDID left outer join extended eA on eA.EXTENDEDID = a.EXTENDEDID where a.invoicedate between cast (dateadd(month, -1, cast('" & DateTimePicker2.Value & "' as timestamp)) as date) and '" & DateTimePicker2.Value & "' and (ea.CUSTOMFIELD1 = 0 or ea.CUSTOMFIELD1 is null) and D.IS_PROMO = 0 and ( LEFT(sod.sono,4) = 'EMOS' and ( e1.customfield9 is not null) )", conn)
        dt = New DataTable
        da.Fill(dt)
        If dt.Rows(0).Item("CT") > 0 Then
            RichTextBox1.Clear()
            RichTextBox1.AppendText("Anda memiliki " & dt.Rows(0).Item("CT") & " invoice EMOS yang masih pending. " & vbNewLine & "  Untuk proses po pending silakan jalankan menu FET-TOOLS-SYNC EMOS terlebih dahulu.")
            Button1.Enabled = False

        Else
            RichTextBox1.Clear()
            RichTextBox1.AppendText("Tidak ada faktur EMOS yang pending.")
            tampil()
        End If
        conn.Close()
    End Sub
    Private Sub f_EMOS_list_order_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Dim da As New FbDataAdapter
        Dim dt As New DataTable
        'mengambil 4 hari sebelum sebagai tanggal mulai
        DateTimePicker1.Value = Format(DateAdd("d", -5, CDate(Now())), "yyyy-MM-dd")
        DateTimePicker2.Value = Format(CDate(Now()), "yyyy-MM-dd")

        'mengambil kode subdist
        Dim kodesubdist As String
        If f_menu.ToolStripKodeSubdistAwal.Text = "-" Then
            kodesubdist = f_menu.ToolstripKodesubdist.Text
        Else
            kodesubdist = f_menu.ToolStripKodeSubdistAwal.Text
        End If

        'lihat apakah masih ada faktur EMOS yang masih belum terkirim
        BukaKoneksi()
        da = New FbDataAdapter("select count(distinct invoiceno) CT from ARINVDET d left outer join ARINV a on a.ARINVOICEID = d.ARINVOICEID left outer join sodet sd on sd.soid = d.soid and d.soseq = sd.SEQ left outer join so sod on sod.soid = d.soid left outer join extended e1 on e1.EXTENDEDID = sd.EXTENDEDID left outer join extended eA on eA.EXTENDEDID = a.EXTENDEDID where a.invoicedate between cast (dateadd(month, -1, cast('" & DateTimePicker2.Value & "' as timestamp)) as date) and '" & DateTimePicker2.Value & "' and (ea.CUSTOMFIELD1 = 0 or ea.CUSTOMFIELD1 is null) and D.IS_PROMO = 0 and ( LEFT(sod.sono,4) = 'EMOS' and ( e1.customfield9 is not null) )", conn)
        dt = New DataTable
        da.Fill(dt)
        If dt.Rows(0).Item("CT") > 0 Then
            RichTextBox1.Clear()
            RichTextBox1.AppendText("Anda memiliki " & dt.Rows(0).Item("CT") & " invoice EMOS yang masih pending. " & vbNewLine & "  Untuk proses po pending silakan jalankan menu FET-TOOLS-SYNC EMOS terlebih dahulu.")
            Button1.Enabled = False

        Else
            RichTextBox1.Clear()
            RichTextBox1.AppendText("Tidak ada faktur EMOS yang pending.")
            tampil()
        End If
        conn.Close()
    End Sub
    Private Sub tampil()
        Try
            DataGridView1.Rows.Clear()
            ds = New DataSet

            If f_menu.ToolStripKodeSubdistAwal.Text = "-" Then
                ds = GetDataPrice.GetOrderEMOSNew_NEW2(f_menu.ToolstripKodesubdist.Text, Format(DateTimePicker1.Value, "yyyy-MM-dd"), Format(DateTimePicker2.Value, "yyyy-MM-dd"), txnodok.Text)
            Else
                ds = GetDataPrice.GetOrderEMOSNew_NEW2(f_menu.ToolStripKodeSubdistAwal.Text, Format(DateTimePicker1.Value, "yyyy-MM-dd"), Format(DateTimePicker2.Value, "yyyy-MM-dd"), txnodok.Text)
            End If

            ds.Tables(0).Columns.Add(New DataColumn("SYNC", GetType(String)))

            DataGridView1.AutoGenerateColumns = False
            Dim i As Integer
            Dim j As Integer
            'Dim kodebp As String
            'Dim CustPONum As String
            'Dim custnumbp As String
            'Dim dtablesync As New DataTable
            'Dim syncall As String = Nothing
            'Dim FlagSync As String = ""

            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim namacustomer As String = Nothing
                Dim WARNA As Integer

                '<5 Juli 2021, Anthony. Penambahan column di tampilan depan (column Status Upload)>
                'kodebp = ds.Tables(0).Rows(i).Item("kodebp")
                'CustPONum = ds.Tables(0).Rows(i).Item("ORIGCUSTOMER_PO_NUMBER")
                'custnumbp = ds.Tables(0).Rows(i).Item("custnumbp")

                'dtablesync = GetDataPrice.GetFlagSyncSubdist(kodebp, custnumbp, CustPONum)

                'If dtablesync.Rows.Count < 0 Or dtablesync.Rows.Count = 0 Then
                '    FlagSync = "Not yet Synced"
                'End If


                'If dtablesync.Rows.Count > 0 Then
                '    For j = 0 To dtablesync.Rows.Count - 1
                '        syncall = syncall + dtablesync.Rows(j).Item("flag_sync")
                '    Next



                '    If syncall Like "%Y%" Then
                '        FlagSync = "Synching"
                '    ElseIf syncall Like "%E%" Then
                '        FlagSync = "Error"
                '    ElseIf syncall Like "%C%" Then
                '        FlagSync = "Checking"
                '    ElseIf syncall Like "%N%" Then
                '        FlagSync = "Not Synced"
                '    Else
                '        FlagSync = "Synced"
                '    End If

                '    'If dtablesync.Rows(0).Item("flag_sync") = "S" Then
                '    '    FlagSync = "Synced"
                '    'ElseIf dtablesync.Rows(0).Item("flag_sync") = "Y" Then
                '    '    FlagSync = "Synching"
                '    'ElseIf dtablesync.Rows(0).Item("flag_sync") = "N" Then
                '    '    FlagSync = "Not Synced"
                '    'ElseIf dtablesync.Rows(0).Item("flag_sync") = "C" Then
                '    '    FlagSync = "Checking"

                '    'End If
                '    'ElseIf dtablesync.Rows.Count > 1 Then
                '    '    FlagSync = "Incomplete"
                'End If
                ''</5 Juli 2021, Anthony. Penambahan column di tampilan depan (column Status Upload)>

                BukaKoneksi()
                da = New FbDataAdapter("select name from persondata where personno ='" & ds.Tables(0).Rows(i).Item("custnumbp") & "'", conn)
                dt = New DataTable
                da.Fill(dt)
                If dt.Rows.Count > 0 Then
                    namacustomer = dt.Rows(0).Item(0).ToString
                    WARNA = 1
                Else
                    namacustomer = "Customer Isn't Mapping"
                    WARNA = 0
                    'DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                End If
                conn.Close()
                Dim row As String() = New String() {ds.Tables(0).Rows(i).Item("ORIGCUSTOMER_PO_NUMBER").ToString, Format(CDate(ds.Tables(0).Rows(i).Item("po_date")), "dd MMM yyyy"), ds.Tables(0).Rows(i).Item("paymenttype").ToString, ds.Tables(0).Rows(i).Item("custnumbp").ToString, namacustomer, ds.Tables(0).Rows(i).Item("status").ToString, ds.Tables(0).Rows(i).Item("cancelreason").ToString, ds.Tables(0).Rows(i).Item("PERCENTAGE_PROSES"), ds.Tables(0).Rows(i).Item("PERCENTAGE_DIPENUHI"), ds.Tables(0).Rows(i).Item("Note")}
                DataGridView1.Rows.Add(row)
                DataGridView1.Refresh()
                If WARNA = 0 Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Orange
                End If
                If Integer.Parse(ds.Tables(0).Rows(i).Item("PERCENTAGE_PROSES").ToString) > 0 And Integer.Parse(ds.Tables(0).Rows(i).Item("PERCENTAGE_PROSES").ToString) < 100 Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                End If
                If ds.Tables(0).Rows(i).Item("status").ToString = "Canceled" Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Crimson
                End If
                If Not ds.Tables(0).Rows(i).Item("Note").Equals("-") Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        tampil()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        My.Forms.f_EMOS_det_order.MdiParent = f_menu
        My.Forms.f_EMOS_det_order.txSO_No.Text = DataGridView1.SelectedRows.Item(0).Cells(0).Value
        My.Forms.f_EMOS_det_order.txSO_Date.Text = DataGridView1.SelectedRows.Item(0).Cells(1).Value
        My.Forms.f_EMOS_det_order.txtop.Text = DataGridView1.SelectedRows.Item(0).Cells(2).Value
        My.Forms.f_EMOS_det_order.txstatus.Text = DataGridView1.SelectedRows.Item(0).Cells(5).Value
        My.Forms.f_EMOS_det_order.txcustno.Text = DataGridView1.SelectedRows.Item(0).Cells(3).Value.ToString & ";" & DataGridView1.SelectedRows.Item(0).Cells(4).Value.ToString
        My.Forms.f_EMOS_det_order.Show()
        My.Forms.f_EMOS_det_order.Focus()

    End Sub
    Public Shared MouseX As Integer
    Public Shared MouseY As Integer
    Private Sub DataGridView1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseDown
        MouseX = e.X
        MouseY = e.Y
    End Sub
    Private Sub f_EMOS_list_order_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Dim da As New FbDataAdapter
        Dim dt As New DataTable
        'mengambil 4 hari sebelum sebagai tanggal mulai
        DateTimePicker1.Value = Format(DateAdd("d", -5, CDate(Now())), "yyyy-MM-dd")
        DateTimePicker2.Value = Format(CDate(Now()), "yyyy-MM-dd")

        'mengambil kode subdist
        Dim kodesubdist As String
        If f_menu.ToolStripKodeSubdistAwal.Text = "-" Then
            kodesubdist = f_menu.ToolstripKodesubdist.Text
        Else
            kodesubdist = f_menu.ToolStripKodeSubdistAwal.Text
        End If

        'lihat apakah masih ada faktur EMOS yang masih belum terkirim
        BukaKoneksi()
        da = New FbDataAdapter("select count(distinct invoiceno) CT from ARINVDET d left outer join ARINV a on a.ARINVOICEID = d.ARINVOICEID left outer join sodet sd on sd.soid = d.soid and d.soseq = sd.SEQ left outer join so sod on sod.soid = d.soid left outer join extended e1 on e1.EXTENDEDID = sd.EXTENDEDID left outer join extended eA on eA.EXTENDEDID = a.EXTENDEDID where a.invoicedate between cast (dateadd(month, -1, cast('" & DateTimePicker2.Value & "' as timestamp)) as date) and '" & DateTimePicker2.Value & "' and (ea.CUSTOMFIELD1 = 0 or ea.CUSTOMFIELD1 is null) and D.IS_PROMO = 0 and ( LEFT(sod.sono,4) = 'EMOS' and ( e1.customfield9 is not null) )", conn)
        dt = New DataTable
        da.Fill(dt)
        If dt.Rows(0).Item("CT") > 0 Then
            RichTextBox1.Clear()
            RichTextBox1.AppendText("Anda memiliki " & dt.Rows(0).Item("CT") & " invoice EMOS yang masih pending. " & vbNewLine & "  Untuk proses po pending silakan jalankan menu FET-TOOLS-SYNC EMOS terlebih dahulu.")
            Button1.Enabled = False

        Else
            RichTextBox1.Clear()
            RichTextBox1.AppendText("Tidak ada faktur EMOS yang pending.")
            tampil()
        End If
        conn.Close()
    End Sub

    '<5 juli 2021, Anthony. Penambahan fungsi klik kanan pada table datanya untuk Sync EMOS>
    Private Sub SyncEMOSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SyncEMOSToolStripMenuItem.Click
        Try
            'Buat ambil invoiceno, itemno
            BukaKoneksi()
            da = New FbDataAdapter("SELECT X.ORIGCUSTOMER_PO_NUMBER,CASE WHEN c.RESERVED8 IS NULL THEN c.ADDRESSLINE3 ELSE c.RESERVED8 END KODEBP, X.CUSTNUMBP,X.ITEMNO, cast(SUM(X.QUANTITY) as int) QTY_DIPENUHI, cast(round(SUM(X.DISC),2) as money) TOTAL_DISC,cast( round(SUM(X.HARGA_INC_DISC),2) as money) TOTAL_HARGA,x.customfield1 FLAGPROSES,c.addressline3 KODEBPSPLIT, invoiceno, extract (YEAR from (cast (tgl_berangkat as TIMESTAMP_1))) || '-'|| lpad ( extract (month from (cast (tgl_berangkat as TIMESTAMP_1))),2,0) || '-'|| lpad ( extract (day from (cast (tgl_berangkat as TIMESTAMP_1))),2,0) || ' '|| lpad ( extract (hour from (cast (tgl_berangkat as TIMESTAMP_1))),2,0) || ':'|| lpad ( extract (minute from (cast (tgl_berangkat as TIMESTAMP_1))),2,0) || ':'|| '00' || '.'||'000'TGL_BERANGKAT, PENGEMUDI, extract (YEAR from (cast (tgl_nota as TIMESTAMP_1))) || '-'|| lpad ( extract (month from (cast (tgl_nota as TIMESTAMP_1))),2,0) || '-'|| lpad ( extract (day from (cast (tgl_nota as TIMESTAMP_1))),2,0) || ' '|| lpad ( extract (hour from (cast (tgl_nota as TIMESTAMP_1))),2,0) || ':'|| lpad ( extract (minute from (cast (tgl_nota as TIMESTAMP_1))),2,0) || ':'|| '00' || '.'||'000' tgl_nota, sono,extract (YEAR from (cast (sodate as TIMESTAMP_1))) || '-'|| lpad ( extract (month from (cast (sodate as TIMESTAMP_1))),2,0) || '-'|| lpad ( extract (day from (cast (sodate as TIMESTAMP_1))),2,0)  sodate, invoiceamount, cashdiscount, tax1amount, X.extendedid FROM( select e1.customfield9 ORIGCUSTOMER_PO_NUMBER,trim(trailing ' ' from p.personno) CUSTNUMBP,d.ITEMNO, d.QUANTITY, (d.BRUTOUNITPRICE - d.ITEMCOST) * d.QUANTITY as disc,d.QUANTITY * d.ITEMCOST as harga_inc_disc,a.invoiceno, ea.customfield10 tgl_nota,ea.CUSTOMFIELD2 || ' ' || ea.CUSTOMFIELD3 TGL_BERANGKAT,ea.CUSTOMFIELD4 PENGEMUDI,a.invoicedate tgl_nota2,sod.sono,sod.sodate,a.INVOICEAMOUNT,a.CASHDISCOUNT,a.CASHDISCPC,d.TAXABLEAMOUNT1 TAX1AMOUNT, eA.CUSTOMFIELD1, eA.EXTENDEDID from ARINVDET d left outer join ARINV a on a.ARINVOICEID = d.ARINVOICEID left outer join sodet sd on sd.soid = d.soid and d.soseq = sd.SEQ left outer join so sod on sod.soid = d.soid left outer join PERSONDATA p on sod.CUSTOMERID = p.ID left outer join extended e ON sod.EXTENDEDID = e.EXTENDEDID left outer join extended e1 on e1.EXTENDEDID = sd.EXTENDEDID left outer join extended eA on eA.EXTENDEDID = a.EXTENDEDID where (ea.CUSTOMFIELD1 = 0 or ea.CUSTOMFIELD1 is null) and D.IS_PROMO = 0 and ( LEFT(sod.sono,4) = 'EMOS' and ( e1.customfield9 is not null)) ) X cross join company c where x.ORIGCUSTOMER_PO_NUMBER ='" & DataGridView1.SelectedRows(0).Cells(0).ToString & "' AND (c.RESERVED8 = '" & f_menu.ToolstripKodesubdist.Text & "' or c.ADDRESSLINE3 = '" & f_menu.ToolstripKodesubdist.Text & "' or c.RESERVED8 = '" & f_menu.ToolstripKodesubdist.Text & "') AND x.CUSTNUMBP = '" & DataGridView1.SelectedRows(0).Cells(3).ToString & "' GROUP BY X.ITEMNO,X.ORIGCUSTOMER_PO_NUMBER,c.ADDRESSLINE3,c.RESERVED8,X.CUSTNUMBP,invoiceno,tgl_berangkat,Pengemudi,tgl_nota,sono,sodate,invoiceamount,cashdiscount,tax1amount, x.customfield1, x.EXTENDEDID ORDER BY ORIGCUSTOMER_PO_NUMBER,ITEMNO", conn)
            dt = New DataTable
            da.Fill(dt)

            Dim itemno As String = dt.Rows(0).Item("itemno")
            Dim invoiceno As String = dt.Rows(0).Item("invoiceno")

            Dim checkSubd As New DataTable
            checkSubd = GetDataPrice.CheckDataEmosB2B(DataGridView1.SelectedRows(0).Cells(0).ToString, f_menu.ToolstripKodesubdist.Text, DataGridView1.SelectedRows(0).Cells(3).ToString, invoiceno, itemno)
            'checkSubd = GetDataPrice.CheckDataEmosB2B(dtBelumProses.Rows(i).Item("ORIGCUSTOMER_PO_NUMBER").ToString, dtBelumProses.Rows(i).Item("KODEBP").ToString, dtBelumProses.Rows(i).Item("CUSTNUMBP").ToString, dtBelumProses.Rows(i).Item("INVOICENO").ToString, dtBelumProses.Rows(i).Item("ITEMNO").ToString)
            If checkSubd.Rows.Count < 0 Then
                MsgBox("Data tidak dapat di Sync")
                'bikin alert error no data found
            End If

            Dim cmd_FB As New FbCommand
            Dim tgl_berangkat
            If dt.Rows(0).Item("TGL_BERANGKAT").ToString = "" Then
                tgl_berangkat = DBNull.Value.ToString
            Else
                tgl_berangkat = dt.Rows(0).Item("TGL_BERANGKAT").ToString
            End If

            If checkSubd.Rows.Count > 0 Then
                'update
                Try
                    GetDataPrice.SyncDataEmosToB2B(dt.Rows(0).Item("ORIGCUSTOMER_PO_NUMBER").ToString, dt.Rows(0).Item("KODEBP").ToString, dt.Rows(0).Item("CUSTNUMBP").ToString, dt.Rows(0).Item("ITEMNO").ToString, dt.Rows(0).Item("QTY_DIPENUHI").ToString, dt.Rows(0).Item("TOTAL_DISC").ToString, dt.Rows(0).Item("TOTAL_HARGA").ToString, dt.Rows(0).Item("FLAGPROSES").ToString, dt.Rows(0).Item("KODEBPSPLIT").ToString, dt.Rows(0).Item("INVOICENO").ToString, tgl_berangkat, dt.Rows(0).Item("PENGEMUDI").ToString, dt.Rows(0).Item("TGL_NOTA").ToString, dt.Rows(0).Item("SONO").ToString, dt.Rows(0).Item("SODATE").ToString, dt.Rows(0).Item("INVOICEAMOUNT").ToString, dt.Rows(0).Item("CASHDISCOUNT").ToString, dt.Rows(0).Item("TAX1AMOUNT").ToString, dt.Rows(0).Item("EXTENDEDID").ToString, "C")

                    cmd_FB = New FbCommand("update extended set customfield1='1' where extendedid='" & dt.Rows(0).Item("EXTENDEDID").ToString & "'", conn)
                    cmd_FB.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox("Update Sync EMOS error")
                End Try
            Else
                'insert
                Try
                    GetDataPrice.SyncDataEmosToB2B(dt.Rows(0).Item("ORIGCUSTOMER_PO_NUMBER").ToString, dt.Rows(0).Item("KODEBP").ToString, dt.Rows(0).Item("CUSTNUMBP").ToString, dt.Rows(0).Item("ITEMNO").ToString, dt.Rows(0).Item("QTY_DIPENUHI").ToString, dt.Rows(0).Item("TOTAL_DISC").ToString, dt.Rows(0).Item("TOTAL_HARGA").ToString, dt.Rows(0).Item("FLAGPROSES").ToString, dt.Rows(0).Item("KODEBPSPLIT").ToString, dt.Rows(0).Item("INVOICENO").ToString, tgl_berangkat, dt.Rows(0).Item("PENGEMUDI").ToString, dt.Rows(0).Item("TGL_NOTA").ToString, dt.Rows(0).Item("SONO").ToString, dt.Rows(0).Item("SODATE").ToString, dt.Rows(0).Item("INVOICEAMOUNT").ToString, dt.Rows(0).Item("CASHDISCOUNT").ToString, dt.Rows(0).Item("TAX1AMOUNT").ToString, dt.Rows(0).Item("EXTENDEDID").ToString, "N")

                    cmd_FB = New FbCommand("update extended set customfield1='1' where extendedid='" & dt.Rows(0).Item("EXTENDEDID").ToString & "'", conn)
                    cmd_FB.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox("insert sync emos error")
                End Try

            End If
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    '</5 juli 2021, Anthony. Penambahan fungsi klik kanan pada table datanya untuk Sync EMOS>

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        Dim rowClicked As DataGridView.HitTestInfo = DataGridView1.HitTest(e.X, e.Y)

        'Select Right Clicked Row if its not the header row
        If e.Button = System.Windows.Forms.MouseButtons.Right AndAlso e.RowIndex > -1 Then
            'Clear any currently sellected rows
            DataGridView1.ClearSelection()
            Me.DataGridView1.Rows(e.RowIndex).Selected = True
            ContextMenuStrip1.Show(DataGridView1, New System.Drawing.Point(MouseX, MouseY))
        End If
    End Sub

    Private Sub CancelOrderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CancelOrderToolStripMenuItem.Click

        If DataGridView1.SelectedRows(0).Cells(5).Value = "Canceled" Then
            MsgBox("PO sudah dicancel !", vbCritical)
            Exit Sub
        ElseIf DataGridView1.SelectedRows(0).Cells(5).Value <> "Completed" Then
            BukaKoneksi()
            ' da = New FbDataAdapter("select distinct arinvoiceid from ARINVDET where SOID in(select SOID from SO s left outer join EXTENDED e on e.EXTENDEDID=s.EXTENDEDID where e.CUSTOMFIELD1='1' and s.PONO='" & DataGridView1.SelectedRows.Item(0).Cells(0).Value & "')", conn)
            da = New FbDataAdapter("select distinct f.ARINVOICEID from so s left outer join sodet d on d.SOID = s.SOID left outer join extended ed on ed.EXTENDEDID = d.EXTENDEDID left outer join persondata pd on pd.id = s.CUSTOMERID left outer join arinvdet t on t.soid = d.SOID left outer join arinv f on f.ARINVOICEID = t.ARINVOICEID where ed.customfield1= 1 and ed.customfield9 = '" & DataGridView1.SelectedRows.Item(0).Cells(0).Value & "' and trim(trailing from pd.personno) = '" & DataGridView1.SelectedRows.Item(0).Cells(3).Value & "' and ed.customfield9 is not null and f.ARINVOICEID is not null", conn)
            '25juni2021
            ' da = New FbDataAdapter("select distinct arinvoiceid from ARINVDET where SOID in(select SOID from SO s left outer join EXTENDED e on e.EXTENDEDID=s.EXTENDEDID where e.CUSTOMFIELD1='1' and e.customfield9 = '" & DataGridView1.SelectedRows.Item(0).Cells(0).Value & "')", conn)
            dt = New DataTable
            da.Fill(dt)
            If dt.Rows.Count = 0 Then
                Dim tanya As String
                Dim kodesubdist As String = f_menu.ToolstripKodesubdist.Text
                Dim kodesubdist_split As String = f_menu.ToolStripKodeSubdistAwal.Text

                tanya = MsgBox("Yakin Pesanan No. " & DataGridView1.SelectedRows(0).Cells(0).Value.ToString & " akan di Cancel ?", vbYesNo + vbQuestion, "Canceled")

                If tanya = vbYes Then
                    If DataGridView1.SelectedRows(0).Cells(4).Value.ToString = "Customer Isn't Mapping" Then

                        GetDataPrice.SendFlagProses_EMOS_Header_NEW2(DataGridView1.SelectedRows(0).Cells(0).Value.ToString, f_menu.ToolStripKodeSubdistAwal.Text, "C", "Salah Mapping", DataGridView1.SelectedRows(0).Cells(3).Value.ToString)

                        'update cancelreason pada table b2b
                        'GetDataPrice.UpdCancelReasonEmosB2B(DataGridView1.SelectedRows(0).Cells(0).Value.ToString, f_menu.ToolStripKodeSubdistAwal.Text, "C", "Salah Mapping", DataGridView1.SelectedRows(0).Cells(3).Value.ToString)

                        MsgBox("PO Sudah di Cancel", vbOKOnly)

                        tampil()
                    Else
                        My.Forms.frm_dialog_cancel.tx_PONO.Text = DataGridView1.SelectedRows.Item(0).Cells(0).Value
                        My.Forms.frm_dialog_cancel.txKodePelanggan.Text = DataGridView1.SelectedRows.Item(0).Cells(3).Value
                        My.Forms.frm_dialog_cancel.txtop.Text = DataGridView1.SelectedRows.Item(0).Cells(2).Value
                        My.Forms.frm_dialog_cancel.txSO_Date.Text = DataGridView1.SelectedRows.Item(0).Cells(1).Value
                        My.Forms.frm_dialog_cancel.Show()
                        Me.Close()
                        f_menu.Hide()
                    End If

                    'SendFlagCancel_SO_EMOS_NEW2(ByVal nodok As String, ByVal kodesubdist As String, ByVal alasan As String, ByVal kodecust_FINA As String, ByVal kodesubdist_split As String, ByVal FLAG_PROSES As String)

                    ' GetDataPrice.SendFlagCancel_SO_EMOS_NEW2(DataGridView1.SelectedRows.Item(0).Cells(0).Value, kodesubdist, DataGridView1.SelectedRows.Item(0).Cells(1).Value, DataGridView1.SelectedRows.Item(0).Cells(3).Value, kodesubdist_split, "C")

                    'GetDataPrice.SendFlagProses_EMOS_Header_NEW2(DataGridView1.SelectedRows.Item(0).Cells(0).Value, f_menu.ToolstripKodesubdist.Text, "C")
                    'cmd = New FbCommand("update SO set STATUS=1, CLOSED=1 where soid in (select distinct s.soid from so s left outer join sodet d on d.SOID = s.SOID left outer join extended ed on ed.EXTENDEDID = d.EXTENDEDID left outer join persondata pd on pd.id = s.CUSTOMERID where ed.customfield1= 1 and trim(trailing from pd.personno) ='" & DataGridView1.SelectedRows.Item(0).Cells(3).Value & "' and ed.customfield9 = '" & DataGridView1.SelectedRows.Item(0).Cells(0).Value & "'", conn)
                    ' cmd.ExecuteNonQuery()

                    'MsgBox("PO di Cancel", vbOKOnly)
                    'tampil()
                End If

            Else
                MsgBox("PO sudah menjadi Invoice di FINA, Tidak dapat di cancel !", vbCritical)
                Exit Sub
            End If
        Else
            MsgBox("PO sudah menjadi Invoice di FINA, Tidak dapat di cancel !", vbCritical)
            Exit Sub
        End If
    End Sub

    Private Sub ResendOrderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResendOrderToolStripMenuItem.Click
        Try
            Dim tanya As String
            Dim kodesubdist1 As String = f_menu.ToolstripKodesubdist.Text
            Dim kodesubdist2 As String = f_menu.ToolStripKodeSubdistAwal.Text
            Dim cmd_FB As New FbCommand

            tanya = MsgBox("Yakin Pesanan No. " & DataGridView1.SelectedRows(0).Cells(0).Value.ToString & " akan di kirim ulang ? Pastikan statusnya masih Proses dan telah menjadi faktur atau telah masuk ke fina sebagai SO. Data baru akan terupdate +- 5 s/d 10 menit", vbYesNo + vbQuestion, "Canceled")

            If tanya = vbYes Then
                If DataGridView1.SelectedRows(0).Cells(5).Value = "Canceled" Then
                    MsgBox("PO sudah di cancel, tidak dapat di resend !", vbCritical)
                    Exit Sub
                ElseIf DataGridView1.SelectedRows(0).Cells(5).Value = "Completed" Then
                    MsgBox("PO dinyatakan telah selesai, tidak dapat di resend !", vbCritical)
                    Exit Sub
                ElseIf DataGridView1.SelectedRows(0).Cells(5).Value = "PENDING" Then
                    MsgBox("PO masih pending, tidak dapat di resend !", vbCritical)
                    Exit Sub
                Else
                    BukaKoneksi()
                    'Dim tanggal_SO_EMOS As String
                    'Dim so_no As String
                    'Dim customer As String

                    Dim da_fb As New FbDataAdapter
                    Dim dt_fb As New DataTable

                    da_fb = New FbDataAdapter("select distinct ea.EXTENDEDID from ARINVDET d left outer join ARINV a on a.ARINVOICEID = d.ARINVOICEID left outer join sodet sd on sd.soid = d.soid and d.soseq = sd.SEQ left outer join so sod on sod.soid = d.soid left outer join extended e ON sod.EXTENDEDID = e.EXTENDEDID left outer join extended e1 on e1.EXTENDEDID = sd.EXTENDEDID left outer join extended eA on eA.EXTENDEDID = a.EXTENDEDID where e1.customfield9='" & DataGridView1.SelectedRows(0).Cells(0).Value.ToString & "' and D.IS_PROMO = 0 and ( LEFT(sod.sono,4) = 'EMOS' and ( e1.customfield9 is not null) ) and ea.customfield1 =1", conn)
                    dt_fb = New DataTable
                    da_fb.Fill(dt_fb)

                    If dt_fb.Rows.Count > 0 Then

                        'update extended
                        Dim i As Integer
                        Dim extendedid As String = "'" & dt_fb.Rows(0).Item("EXTENDEDID").ToString & "'"

                        If dt_fb.Rows.Count > 1 Then
                            For i = 1 To dt_fb.Rows.Count - 1
                                extendedid = extendedid & ",'" & dt_fb.Rows(i).Item("EXTENDEDID").ToString & "'"
                            Next
                        End If
                        
                        Try
                            cmd_FB = New FbCommand("update extended set customfield1 = 0 where extendedid in (" & extendedid & ")", conn)
                            cmd_FB.ExecuteNonQuery()

                        Catch ex As Exception
                            MsgBox(ex.Message.ToString, vbCritical)
                        End Try

                        MsgBox("SO berhasil di resend !", vbCritical)
                    Else
                        MsgBox("PO belum pernah diproses di fina, silahkan diproses terlebih dahulu !", vbCritical)
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString, vbCritical)
            Exit Sub
        End Try

        conn.Close()
    End Sub

    Private Sub ResendInvoiceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResendInvoiceToolStripMenuItem.Click
        Try
            BukaKoneksi()
            Dim da_inv As FbDataAdapter
            Dim dt_inv As DataTable
            Dim kodesubdist As String
            Dim no_inv_fina As String

            'da_inv = New FbDataAdapter("select distinct CASE WHEN a.purchaseorderno IS NULL THEN a.description ELSE a.purchaseorderno END purchaseorderno,a.invoiceno,a.invoiceamount,a.extendedid,p.personno,e.customfield10 from ARINV a left outer join persondata p on p.ID=a.customerid left outer join EXTENDED e ON a.EXTENDEDID=e.EXTENDEDID where a.EXTENDEDID in(select extendedid from EXTENDED where extendedid in(select distinct a.extendedid from ARINV a left outer join ARINVDET d on d.ARINVOICEID=a.ARINVOICEID left outer join so s on s.SOID=d.SOID left outer join extended e1 on e1.EXTENDEDID=s.EXTENDEDID where e1.CUSTOMFIELD1 = '1' ) and (a.purchaseorderno='" & DataGridView1.SelectedRows(0).Cells(0).Value & "' or a.description='" & DataGridView1.SelectedRows(0).Cells(0).Value & "'))", conn)
            da_inv = New FbDataAdapter("select distinct e1.customfield9 purchaseorderno, a.invoiceno, a.invoiceamount,a.extendedid,p.personno,e.customfield10 from ARINV a left outer join persondata p on p.ID=a.customerid left outer join EXTENDED e ON a.EXTENDEDID=e.EXTENDEDID left outer join ARINVDET b ON a.ARINVOICEID=b.ARINVOICEID left outer join SO s ON s.SOID=b.SOID left outer join extended e1 on e1.EXTENDEDID = s.EXTENDEDID WHERE e1.CUSTOMFIELD1='1' AND e1.customfield9 = '" & DataGridView1.SelectedRows(0).Cells(0).Value & "')", conn)
            dt_inv = New DataTable
            da_inv.Fill(dt_inv)

            If dt_inv.Rows.Count > 0 Then
                Dim i As Integer
                For i = 0 To dt_inv.Rows.Count - 1
                    If f_menu.ToolStripKodeSubdistAwal.Text = "-" Then
                        kodesubdist = f_menu.ToolstripKodesubdist.Text
                    Else
                        kodesubdist = f_menu.ToolStripKodeSubdistAwal.Text
                    End If

                    If GetDataPrice.SendFlagStatus_EMOS_Invoice2_NEW2(dt_inv.Rows(i).Item("purchaseorderno"), kodesubdist, dt_inv.Rows(i).Item("invoiceno"), dt_inv.Rows(i).Item("customfield10"), dt_inv.Rows(i).Item("invoiceamount"), dt_inv.Rows(i).Item("personno")) = 1 Then

                        no_inv_fina = dt_inv.Rows(i).Item("invoiceno")

                        Dim da_item_inv As New FbDataAdapter
                        Dim dt_item_inv As New DataTable

                        da_item_inv = New FbDataAdapter("SELECT X.ITEMNO, SUM(X.QUANTITY) QTY, SUM(X.DISC) DISC, SUM(X.HARGA_INC_DISC) TOTAL FROM ( select d.ITEMNO, d.QUANTITY, (d.BRUTOUNITPRICE - d.ITEMCOST) * d.QUANTITY as disc, d.QUANTITY * d.ITEMCOST as harga_inc_disc from ARINVDET d left outer join ARINV a on a.ARINVOICEID = d.ARINVOICEID left outer join so sod on sod.soid = d.soid left outer join extended e on sod.EXTENDEDID = e.EXTENDEDID where a.INVOICENO = '" & no_inv_fina & "' AND D.IS_PROMO = 0 and ( LEFT(sod.sono,4) = 'EMOS' and e.customfield9 is not null ) union select RIGHT(D.ITEMOVDESC, 5) ITEMNO, 0 QUANTITY, (d.QUANTITY * d.ITEMCOST) * -1 as disc, d.QUANTITY * d.ITEMCOST as harga_inc_disc from ARINVDET d left outer join ARINV a on a.ARINVOICEID = d.ARINVOICEID left outer join so sod on sod.soid = d.soid left outer join extended e on sod.EXTENDEDID = e.EXTENDEDID where a.INVOICENO = '" & no_inv_fina & "' AND D.IS_PROMO = 1 and ( LEFT(sod.sono,4) = 'EMOS' and e.customfield9 is not null ) ) X GROUP BY X.ITEMNO", conn)
                        dt_item_inv = New DataTable
                        da_item_inv.Fill(dt_item_inv)

                        'terakhir cek sampai sini

                        If dt_item_inv.Rows.Count > 0 Then
                            Dim x As Integer

                            For x = 0 To dt_item_inv.Rows.Count - 1
                                GetDataPrice.SendQtyStatus_EMOS_Invoice_Detail_NEW2(dt_inv.Rows(i).Item("purchaseorderno"), kodesubdist, dt_item_inv.Rows(x).Item(0).ToString, dt_item_inv.Rows(x).Item(1), dt_item_inv.Rows(x).Item(2), dt_item_inv.Rows(x).Item(3), dt_inv.Rows(i).Item("personno").ToString)
                            Next
                        End If

                        GetDataPrice.SendFlagProses_EMOS_Header_NEW2(dt_inv.Rows(i).Item("purchaseorderno"), kodesubdist, "I", "", dt_inv.Rows(i).Item("personno").ToString)

                        Dim cmd_FB As FbCommand

                        cmd_FB = New FbCommand("update extended set customfield1='1' where extendedid='" & dt_inv.Rows(i).Item("extendedid") & "'", conn)
                        cmd_FB.ExecuteNonQuery()

                        MsgBox("Invoice berhasil di resend !", vbCritical)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString, vbCritical)
            Exit Sub
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub ResendDeliveryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResendDeliveryToolStripMenuItem.Click
        Try
            BukaKoneksi()
            Dim da_deliv As FbDataAdapter
            Dim dt_deliv As DataTable
            Dim kodesubdist As String

            'da_deliv = New FbDataAdapter("Select distinct CASE WHEN x.DESCRIPTION IS NULL THEN x.PONO ELSE x.DESCRIPTION END PONO, a.INVOICENO, a.INVOICEDATE, s.LASTNAME salesman, p.PERSONNO, p.NAME, a.INVOICEAMOUNT, e.CUSTOMFIELD2 || ' ' || e.CUSTOMFIELD3 TGL_BERANGKAT,e.CUSTOMFIELD4 pengemudi from ARINV a Left outer join EXTENDED e on e.EXTENDEDID=a.EXTENDEDID left outer join SALESMAN s on s.SALESMANID=a.SALESMANID left outer join ARINVDET d on d.ARINVOICEID=a.ARINVOICEID left outer join SO x on x.SOID=d.SOID left outer join EXTENDED e2 on e2.EXTENDEDID=x.EXTENDEDID left outer join PERSONDATA p on p.ID=a.CUSTOMERID where e2.CUSTOMFIELD1='1' and (x.PONO ='" & DataGridView1.SelectedRows(0).Cells(0).Value & "' OR x.DESCRIPTION ='" & DataGridView1.SelectedRows(0).Cells(0).Value & "') order by a.ARINVOICEID desc", conn)
            da_deliv = New FbDataAdapter("Select distinct e2.customfield9 PONO, a.INVOICENO, a.INVOICEDATE, s.LASTNAME salesman, p.PERSONNO, p.NAME, a.INVOICEAMOUNT, e.CUSTOMFIELD2 || ' ' || e.CUSTOMFIELD3 TGL_BERANGKAT, e.CUSTOMFIELD4 pengemudi from ARINV a Left outer join EXTENDED e on e.EXTENDEDID = a.EXTENDEDID left outer join SALESMAN s on s.SALESMANID = a.SALESMANID left outer join ARINVDET d on d.ARINVOICEID = a.ARINVOICEID left outer join SO x on x.SOID = d.SOID left outer join EXTENDED e2 on e2.EXTENDEDID = x.EXTENDEDID left outer join PERSONDATA p on p.ID = a.CUSTOMERID where e2.CUSTOMFIELD1 = '1' and e2.customfield9 = '" & DataGridView1.SelectedRows(0).Cells(0).Value & "') order by a.ARINVOICEID desc", conn)
            dt_deliv = New DataTable
            da_deliv.Fill(dt_deliv)

            If dt_deliv.Rows.Count > 0 Then
                If f_menu.ToolStripKodeSubdistAwal.Text = "-" Then
                    kodesubdist = f_menu.ToolstripKodesubdist.Text
                Else
                    kodesubdist = f_menu.ToolStripKodeSubdistAwal.Text
                End If

                GetDataPrice.SendDataDelivery_Header_NEW2(dt_deliv.Rows(0).Item("PONO"), kodesubdist, dt_deliv.Rows(0).Item("INVOICENO"), dt_deliv.Rows(0).Item("TGL_BERANGKAT"), dt_deliv.Rows(0).Item("pengemudi"))

                MsgBox("Delivery Berhasil dikirim", vbCritical)
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString, vbCritical)
            Exit Sub
        End Try

        conn.Close()
    End Sub

    Private Sub RichTextBox3_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox3.TextChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class