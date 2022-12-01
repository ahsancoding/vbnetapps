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
