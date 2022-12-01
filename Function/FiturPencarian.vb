    'Function untuk mencari pengecekan QC berdasarkan No LPB
    Sub findPengecekanQC(ByVal nolpb As String)
        bukaConn()
        Dim SqlQuery As String = "select a.no_lpqc, a.tgl_lpqc, a.keterangan, a.no_lpb, a.tgl_approval from pengecekan_qc AS a where no_lpb like '%" & nolpb & "%'"
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
        DGPengecekanQC.Rows.Clear()
        For i = 0 To TABLE.Rows.Count - 1
            With DGPengecekanQC
                .Rows.Add(TABLE.Rows(i)("no_lpb"), TABLE.Rows(i)("tgl_lpqc"), TABLE.Rows(i)("no_lpqc"), TABLE.Rows(i)("tgl_approval"), TABLE.Rows(i)("keterangan"))
            End With
        Next
    End Sub