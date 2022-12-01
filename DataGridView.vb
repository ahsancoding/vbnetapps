' Library ODBC Connector to MySQL 
Imports System.Data.Odbc

' Ambil data dari DB tanpa Parameter DENGAN HEADER KOLOM DATA GRID VIEW:
Sub RefreshDataView_barang()
    bukaConn()
    Dim SqlQuery As String = "select a.kd_barang,b.nama_barang,d.total_qty_terima, d.keterangan, a.stok from stok_tito a, barang as b, penerimaan_barang as c,detil_penerimaan_barang as d " & _
    " where a.kd_barang = b.kd_barang And b.kd_barang = d.kd_barang and c.no_penerimaan='" & TextBox1.Text & "' And c.no_penerimaan = d.no_penerimaan" & _
    " and a.kd_divwh='11' and a.kd_lokasi='" & frmlogin.txtkd_lokasi.Text & "' group by a.kd_barang "
    Dim SqlCommand As New OdbcCommand
    Dim sqlAdapter As New OdbcDataAdapter
    Dim TABLE As New DataTable
    With SqlCommand
        .CommandText = SqlQuery
        .Connection = MyCn
    End With
    With sqlAdapter
        .SelectCommand = SqlCommand
        .Fill(TABLE)
    End With
    DataGridView1.Rows.Clear()
    For i = 0 To TABLE.Rows.Count - 1
        With DataGridView1
            .Rows.Add(TABLE.Rows(i)("kd_barang"), TABLE.Rows(i)("nama_barang"), TABLE.Rows(i)("total_qty_terima"), TABLE.Rows(i)("stok"), TABLE.Rows(i)("keterangan"))
        End With
    Next
End Sub


' Pencarian Data dari DB with Parameter Pencarian DENGAN HEADER KOLOM DATA GRID VIEW:
Sub RefreshDataViewImport(ByVal xlokasi As String)
    Dim SqlQuery As String = "SELECT a.no_penerimaan,a.tgl_terima,a.kd_lokasi,b.nama_lokasi,a.no_ref,a.ket,ifnull(a.tgl_update_qc,'-') as tgl_update_qc,a.container" & _
    " FROM penerimaan_barang as a,lokasi as b,detil_penerimaan_barang as c,barang d" & _
    " WHERE a.kd_lokasi=b.kd_lokasi and a.no_penerimaan=c.no_penerimaan and c.kd_barang=d.kd_barang and d.jenis='sp' and a.container='" & cmbjenis_penerimaan.Text & "'" & _
    " AND DATE(a.tgl_terima) BETWEEN '" & Format(CDate(dtp_tgl1.Value), "yyyy-MM-dd") & "' and '" & Format(CDate(dtp_tgl2.Value), "yyyy-MM-dd") & "'" & _
    " and a.kd_lokasi='" & xlokasi & "' and a.no_penerimaan like '%" & Replace(txtno_lpb.Text, "'", "''") & "%' group by a.no_penerimaan order by a.tgl_terima"
    Dim SqlCommand As New OdbcCommand
    Dim sqlAdapter As New OdbcDataAdapter
    Dim TABLE As New DataTable
    With SqlCommand
        .CommandText = SqlQuery
        .Connection = MyCn
    End With
    With sqlAdapter
        .SelectCommand = SqlCommand
        .Fill(TABLE)
    End With
    dgvlaporan_penerimaan.Rows.Clear()
    For i = 0 To TABLE.Rows.Count - 1
        With dgvlaporan_penerimaan
            .Rows.Add(TABLE.Rows(i)("no_penerimaan"), TABLE.Rows(i)("tgl_terima"), TABLE.Rows(i)("kd_lokasi"), TABLE.Rows(i)("nama_lokasi"), TABLE.Rows(i)("no_ref"), TABLE.Rows(i)("ket"), TABLE.Rows(i)("tgl_update_qc"), TABLE.Rows(i)("container"))
            .Columns(1).DefaultCellStyle.Format = "dd-MM-yyyy"
        End With
    Next
End Sub


' Selection Row DAtagrid 
If dgvlaporan_penerimaan.CurrentRow.Cells(6).Value = "-" Then
    FBL_Terimaqc.txtno_lpb.Text = dgvlaporan_penerimaan.CurrentRow.Cells(0).Value
    FBL_Terimaqc.dtp_tgl.Value = dgvlaporan_penerimaan.CurrentRow.Cells(1).Value
    FBL_Terimaqc.txtno_bl.Text = dgvlaporan_penerimaan.CurrentRow.Cells(4).Value
    FBL_Terimaqc.txtketerangan.Text = dgvlaporan_penerimaan.CurrentRow.Cells(5).Value
    FBL_Terimaqc.txtjenis_penerimaan.Text = dgvlaporan_penerimaan.CurrentRow.Cells(7).Value
    FBL_Terimaqc.ShowDialog()
Else