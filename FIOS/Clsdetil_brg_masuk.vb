Imports System.Data.Odbc
Public Class Clsdetil_brg_masuk
    Private Fno_penerimaan As String
    Private Fnopermintaan As String
    Private Fkd_barang As String
    Private Ftotal_qty_order As Integer
    Private Ftotal_qty_terima As Integer
    Private Falokasi As String
    Private FTgl_alokasi As Date
    Private FTgl_terima As Date
    Private FNama_barang As String
    Private FKeterangan As String
    Private FKd_lokasimasuk As String
    Private Fstatus As String
    Private Fkdcust As String
    Private Fkdstatus As String
    Private Fqtyjumlah As Integer
    'Public Property PNo_penerimaan() As String
    '    Get
    '        Return Fno_penerimaan
    '    End Get
    '    Set(ByVal value As String)
    '        Fno_penerimaan = value
    '    End Set
    'End Property
    'Public Property PKd_barang() As String
    '    Get
    '        Return Fkd_barang
    '    End Get
    '    Set(ByVal value As String)
    '        Fkd_barang = value
    '    End Set
    'End Property
    'Public Property PTotal_qty_order() As Integer
    '    Get
    '        Return Ftotal_qty_order
    '    End Get
    '    Set(ByVal value As Integer)
    '        Ftotal_qty_order = value
    '    End Set
    'End Property
    'Public Property PTotal_qty_terima() As Integer
    '    Get
    '        Return Ftotal_qty_terima
    '    End Get
    '    Set(ByVal value As Integer)
    '        Ftotal_qty_terima = value
    '    End Set
    'End Property
    'Public Property PAlokasi() As String
    '    Get
    '        Return Falokasi
    '    End Get
    '    Set(ByVal value As String)
    '        Falokasi = value
    '    End Set
    'End Property
    'Public Property PTgl_alokasi() As Date
    '    Get
    '        Return FTgl_alokasi
    '    End Get
    '    Set(ByVal value As Date)
    '        FTgl_alokasi = value
    '    End Set
    'End Property
    'Public Property PTgl_terima() As Date
    '    Get
    '        Return FTgl_terima
    '    End Get
    '    Set(ByVal value As Date)
    '        FTgl_terima = value
    '    End Set
    'End Property
    'Public Property PNama_Barang() As String
    '    Get
    '        Return FNama_barang
    '    End Get
    '    Set(ByVal value As String)
    '        FNama_barang = value
    '    End Set
    'End Property
    'Public Property PKeterangan() As String
    '    Get
    '        Return FKeterangan
    '    End Get
    '    Set(ByVal value As String)
    '        FKeterangan = value
    '    End Set
    'End Property
    'Public Property PKd_lokasimasuk() As String
    '    Get
    '        Return FKd_lokasimasuk
    '    End Get
    '    Set(ByVal value As String)
    '        FKd_lokasimasuk = value
    '    End Set
    'End Property
    'Public Property Pstatus() As String
    '    Get
    '        Return Fstatus
    '    End Get
    '    Set(ByVal value As String)
    '        Fstatus = value
    '    End Set
    'End Property
    'Public Property Pkdcust() As String
    '    Get
    '        Return Fkdcust
    '    End Get
    '    Set(ByVal value As String)
    '        Fkdcust = value
    '    End Set
    'End Property
    'Public Property Pqtyjumlah() As Integer
    '    Get
    '        Return Fqtyjumlah
    '    End Get
    '    Set(ByVal value As Integer)
    '        Fqtyjumlah = value
    '    End Set
    'End Property
    'Public Property Pnopermintaan() As String
    '    Get
    '        Return Fnopermintaan
    '    End Get
    '    Set(ByVal value As String)
    '        Fnopermintaan = value
    '    End Set
    'End Property
    'Public Property Pkdstatus() As String
    '    Get
    '        Return Fkdstatus
    '    End Get
    '    Set(ByVal value As String)
    '        Fkdstatus = value
    '    End Set
    'End Property
    Private Fqtyretur As Integer
    'Public Property Pqtyretur() As Integer
    '    Get
    '        Return Fqtyretur
    '    End Get
    '    Set(ByVal value As Integer)
    '        Fqtyretur = value
    '    End Set
    'End Property
    Private Fqtyalokasi As Integer
    'Public Property Pqtyalokasi() As Integer
    '    Get
    '        Return Fqtyalokasi
    '    End Get
    '    Set(ByVal value As Integer)
    '        Fqtyalokasi = value
    '    End Set
    'End Property
    Private Fketerangana As String
    'Public Property Pketerangana() As String
    '    Get
    '        Return Fketerangana
    '    End Get
    '    Set(ByVal value As String)
    '        Fketerangana = value
    '    End Set
    'End Property
    'Private Fno_penerimaan As String
    'Private Fnopermintaan As String
    'Private Fkd_barang As String
    'Private Ftotal_qty_order As Integer
    'Private Ftotal_qty_terima As Integer
    'Private Falokasi As String
    'Private FTgl_alokasi As Date
    'Private FTgl_terima As Date
    'Private FNama_barang As String
    'Private FKeterangan As String
    'Private FKd_lokasimasuk As String
    'Private Fstatus As String
    'Private Fkdcust As String
    'Private Fkdstatus As String
    'Private Fqtyjumlah As Integer

    Private FQty_bagus As Integer
    Private FQty_rusak As Integer
    Private FTgl_cek_qc As Date
    Private FKeterangan_qc As String
    Private FKd_user As String
    Private FNo_lpqc As String

    Private FId_login As String
    Private FNo_lpb As String


    Public Property PNo_penerimaan() As String
        Get
            Return Fno_penerimaan
        End Get
        Set(ByVal value As String)
            Fno_penerimaan = value
        End Set
    End Property
    Public Property PKd_barang() As String
        Get
            Return Fkd_barang
        End Get
        Set(ByVal value As String)
            Fkd_barang = value
        End Set
    End Property
    Public Property PTotal_qty_order() As Integer
        Get
            Return Ftotal_qty_order
        End Get
        Set(ByVal value As Integer)
            Ftotal_qty_order = value
        End Set
    End Property
    Public Property PTotal_qty_terima() As Integer
        Get
            Return Ftotal_qty_terima
        End Get
        Set(ByVal value As Integer)
            Ftotal_qty_terima = value
        End Set
    End Property
    Public Property PAlokasi() As String
        Get
            Return Falokasi
        End Get
        Set(ByVal value As String)
            Falokasi = value
        End Set
    End Property
    Public Property PTgl_alokasi() As Date
        Get
            Return FTgl_alokasi
        End Get
        Set(ByVal value As Date)
            FTgl_alokasi = value
        End Set
    End Property
    Public Property PTgl_terima() As Date
        Get
            Return FTgl_terima
        End Get
        Set(ByVal value As Date)
            FTgl_terima = value
        End Set
    End Property
    Public Property PNama_Barang() As String
        Get
            Return FNama_barang
        End Get
        Set(ByVal value As String)
            FNama_barang = value
        End Set
    End Property
    Public Property PKeterangan() As String
        Get
            Return FKeterangan
        End Get
        Set(ByVal value As String)
            FKeterangan = value
        End Set
    End Property
    Public Property PKd_lokasimasuk() As String
        Get
            Return FKd_lokasimasuk
        End Get
        Set(ByVal value As String)
            FKd_lokasimasuk = value
        End Set
    End Property
    Public Property Pstatus() As String
        Get
            Return Fstatus
        End Get
        Set(ByVal value As String)
            Fstatus = value
        End Set
    End Property
    Public Property Pkdcust() As String
        Get
            Return Fkdcust
        End Get
        Set(ByVal value As String)
            Fkdcust = value
        End Set
    End Property
    Public Property Pqtyjumlah() As Integer
        Get
            Return Fqtyjumlah
        End Get
        Set(ByVal value As Integer)
            Fqtyjumlah = value
        End Set
    End Property
    Public Property Pnopermintaan() As String
        Get
            Return Fnopermintaan
        End Get
        Set(ByVal value As String)
            Fnopermintaan = value
        End Set
    End Property
    Public Property Pkdstatus() As String
        Get
            Return Fkdstatus
        End Get
        Set(ByVal value As String)
            Fkdstatus = value
        End Set
    End Property
    ' Private Fqtyretur As Integer
    Public Property Pqtyretur() As Integer
        Get
            Return Fqtyretur
        End Get
        Set(ByVal value As Integer)
            Fqtyretur = value
        End Set
    End Property
    'Private Fqtyalokasi As Integer
    Public Property Pqtyalokasi() As Integer
        Get
            Return Fqtyalokasi
        End Get
        Set(ByVal value As Integer)
            Fqtyalokasi = value
        End Set
    End Property
    ' Private Fketerangana As String
    Public Property Pketerangana() As String
        Get
            Return Fketerangana
        End Get
        Set(ByVal value As String)
            Fketerangana = value
        End Set
    End Property
    Public Property PQty_bagus() As Integer
        Get
            Return FQty_bagus
        End Get
        Set(ByVal value As Integer)
            FQty_bagus = value
        End Set
    End Property
    Public Property PQty_rusak() As Integer
        Get
            Return FQty_rusak
        End Get
        Set(ByVal value As Integer)
            FQty_rusak = value
        End Set
    End Property
    Public Property PTgl_cek_qc() As Date
        Get
            Return FTgl_cek_qc
        End Get
        Set(ByVal value As Date)
            FTgl_cek_qc = value
        End Set
    End Property
    Public Property PKeterangan_qc() As String
        Get
            Return FKeterangan_qc
        End Get
        Set(ByVal value As String)
            FKeterangan_qc = value
        End Set
    End Property
    Public Property PKd_user() As String
        Get
            Return FKd_user
        End Get
        Set(ByVal value As String)
            FKd_user = value
        End Set
    End Property
    Public Property PNo_lpqc() As String
        Get
            Return FNo_lpqc
        End Get
        Set(ByVal value As String)
            FNo_lpqc = value
        End Set
    End Property
    Public Property PId_login() As String
        Get
            Return FId_login
        End Get
        Set(ByVal value As String)
            FId_login = value
        End Set
    End Property
    Public Property PNo_lpb() As String
        Get
            Return FNo_lpb
        End Get
        Set(ByVal value As String)
            FNo_lpb = value
        End Set
    End Property
    Public Function simpan() As Integer
        Dim xSimpan As String
        Dim X As Integer
        Dim myCmd As OdbcCommand
        xSimpan = "INSERT INTO detil_penerimaan_barang"
        xSimpan &= "(no_penerimaan,kd_barang,total_qty_order,total_qty_terima)"
        xSimpan &= "values (?,?,?,?)"
        myCmd = New OdbcCommand(xSimpan, MyCn)

        myCmd.Parameters.AddWithValue("no_penerimaan", Fno_penerimaan)
        myCmd.Parameters.AddWithValue("kd_barang", Fkd_barang)
        myCmd.Parameters.AddWithValue("total_qty_order", Ftotal_qty_order)
        myCmd.Parameters.AddWithValue("total_qty_terima", Ftotal_qty_terima)
        myCmd.Prepare()
        X = myCmd.ExecuteNonQuery()

        myCmd.Dispose()
        Return X
    End Function
    Public Function simpan_detil_returan() As Integer
        Dim xSimpan As String
        Dim X As Integer
        Dim myCmd As OdbcCommand
        xSimpan = "INSERT INTO t_detil_returan"
        xSimpan &= "(no_lpb,kd_barang,qty_sj,qty_terima,keterangan,kd_lokasimasuk,kd_status)"
        xSimpan &= "values (?,?,?,?,?,?,?)"
        myCmd = New OdbcCommand(xSimpan, MyCn)

        myCmd.Parameters.AddWithValue("no_lpb", Fno_penerimaan)
        myCmd.Parameters.AddWithValue("kd_barang", Fkd_barang)
        myCmd.Parameters.AddWithValue("qty_sj", Ftotal_qty_order)
        myCmd.Parameters.AddWithValue("qty_terima", Ftotal_qty_terima)
        myCmd.Parameters.AddWithValue("keterangan", FKeterangan)
        myCmd.Parameters.AddWithValue("kd_lokasimasuk", FKd_lokasimasuk)
        myCmd.Parameters.AddWithValue("kd_status", Fkdstatus)
        myCmd.Prepare()
        X = myCmd.ExecuteNonQuery()

        myCmd.Dispose()
        Return X
    End Function
    Public Function getcek_detil_barang_masukstok() As Boolean
        Dim query As String
        Dim xmyread As OdbcDataReader
        Dim xmycmd As OdbcCommand

        query = "SELECT * FROM t_detil_returan WHERE no_lpb=? and kd_barang=?"
        xmycmd = New OdbcCommand(query, MyCn)
        xmycmd.Parameters.AddWithValue("no_lpb", Fno_penerimaan)
        xmycmd.Parameters.AddWithValue("kd_barang", Fkd_barang)
        xmyread = xmycmd.ExecuteReader
        If xmyread.HasRows Then
            xmyread.Read()
            Ftotal_qty_terima = xmyread.Item("total_qty_terima").ToString
            xmyread.Close()
            Return True
        Else
            xmyread.Close()
            Return False
        End If

    End Function
    Public Function simpan_detil_barang_masukstok() As Integer
        Dim xSimpan As String
        Dim X As Integer
        Dim myCmd As OdbcCommand
        xSimpan = "INSERT INTO t_detil_returan"
        xSimpan &= "(no_lpb,kd_barang,qty_sj,qty_terima,status,kd_cust,qty_retur,qty_alokasi,keterangan)"
        xSimpan &= "values (?,?,?,?,?,?,?,?,?)"
        myCmd = New OdbcCommand(xSimpan, MyCn)

        myCmd.Parameters.AddWithValue("no_lpb", Fno_penerimaan)
        myCmd.Parameters.AddWithValue("kd_barang", Fkd_barang)
        myCmd.Parameters.AddWithValue("qty_sj", Ftotal_qty_order)
        myCmd.Parameters.AddWithValue("qty_terima", Ftotal_qty_terima)
        myCmd.Parameters.AddWithValue("status", Fstatus)
        myCmd.Parameters.AddWithValue("kd_cust", Fkdcust)
        myCmd.Parameters.AddWithValue("qty_retur", Fqtyretur)
        myCmd.Parameters.AddWithValue("qty_alokasi", Fqtyretur)
        myCmd.Parameters.AddWithValue("keterangan", Fketerangana)

        myCmd.Prepare()
        X = myCmd.ExecuteNonQuery()

        myCmd.Dispose()
        Return X
    End Function
    Public Function ubah_detil_barang_masukstok() As Integer
        Dim xUbah As String
        Dim X As Integer
        Dim xMyCmd As OdbcCommand

        xUbah = "UPDATE t_detil_returan SET qty_sj=?,qty_terima=?  WHERE no_lpb=? and kd_barang=?"
        xMyCmd = New OdbcCommand(xUbah, MyCn)

        xMyCmd.Parameters.AddWithValue("qty_sj", Ftotal_qty_terima)
        xMyCmd.Parameters.AddWithValue("qty_terima", Ftotal_qty_terima)
        xMyCmd.Parameters.AddWithValue("no_lpb", Fno_penerimaan)
        xMyCmd.Parameters.AddWithValue("kd_barang", Fkd_barang)

        xMyCmd.Prepare()
        X = xMyCmd.ExecuteNonQuery()
        xMyCmd.Dispose()
        Return X

    End Function
    Public Function ubah_detil_barang_masukstokcekupdate() As Integer
        Dim xUbah As String
        Dim X As Integer
        Dim xMyCmd As OdbcCommand

        xUbah = "UPDATE t_detil_returan SET qty_retur=?,qty_alokasi=?  WHERE no_lpb=? and kd_barang=?"
        xMyCmd = New OdbcCommand(xUbah, MyCn)

        xMyCmd.Parameters.AddWithValue("qty_retur", Ftotal_qty_terima)
        xMyCmd.Parameters.AddWithValue("qty_alokasi", Ftotal_qty_terima)
        xMyCmd.Parameters.AddWithValue("no_lpb", Fno_penerimaan)
        xMyCmd.Parameters.AddWithValue("kd_barang", Fkd_barang)

        xMyCmd.Prepare()
        X = xMyCmd.ExecuteNonQuery()
        xMyCmd.Dispose()
        Return X

    End Function
    Public Function ubah() As Integer
        Dim xUbah As String
        Dim X As Integer
        Dim xMyCmd As OdbcCommand

        xUbah = "UPDATE detil_penerimaan_barang SET alokasi=?,tgl_alokasi=?,kd_lokasimasuk=?  WHERE no_penerimaan=? and kd_barang=?"
        xMyCmd = New OdbcCommand(xUbah, MyCn)

        xMyCmd.Parameters.AddWithValue("alokasi", Falokasi)
        xMyCmd.Parameters.AddWithValue("tgl_alokasi", FTgl_alokasi)
        xMyCmd.Parameters.AddWithValue("kd_lokasimasuk", Fkd_lokasimasuk)
        xMyCmd.Parameters.AddWithValue("no_penerimaan", Fno_penerimaan)
        xMyCmd.Parameters.AddWithValue("kd_barang", Fkd_barang)

        xMyCmd.Prepare()
        X = xMyCmd.ExecuteNonQuery()
        xMyCmd.Dispose()
        Return X

    End Function
    Public Function ubah_detilpenerimaan_barang() As Integer
        Dim xUbah As String
        Dim X As Integer
        Dim xMyCmd As OdbcCommand

        xUbah = "UPDATE detil_penerimaan_barang SET qty_terima=?,alokasi=?,tgl_alokasi=?,kd_lokasimasuk=?  WHERE  no_penerimaan=? and kd_barang=? "
        xMyCmd = New OdbcCommand(xUbah, MyCn)

        xMyCmd.Parameters.AddWithValue("qty_terima", Fqtyretur)
        xMyCmd.Parameters.AddWithValue("alokasi", Falokasi)
        xMyCmd.Parameters.AddWithValue("tgl_alokasi", FTgl_alokasi)
        xMyCmd.Parameters.AddWithValue("kd_lokasimasuk", FKd_lokasimasuk)
        xMyCmd.Parameters.AddWithValue("no_penerimaan", Fno_penerimaan)
        xMyCmd.Parameters.AddWithValue("kd_barang", Fkd_barang)


        xMyCmd.Prepare()
        X = xMyCmd.ExecuteNonQuery()
        xMyCmd.Dispose()
        Return X

    End Function
    Public Function ubah_detil() As Integer
        Dim xUbah As String
        Dim X As Integer
        Dim xMyCmd As OdbcCommand

        xUbah = "UPDATE t_detil_returan SET qty_sj=?,qty_terima=?  WHERE no_lpb=? and kd_barang=?"
        xMyCmd = New OdbcCommand(xUbah, MyCn)

        xMyCmd.Parameters.AddWithValue("alokasi", Falokasi)
        xMyCmd.Parameters.AddWithValue("tgl_alokasi", FTgl_alokasi)
        xMyCmd.Parameters.AddWithValue("no_penerimaan", Fno_penerimaan)
        xMyCmd.Parameters.AddWithValue("kd_barang", Fkd_barang)

        xMyCmd.Prepare()
        X = xMyCmd.ExecuteNonQuery()
        xMyCmd.Dispose()
        Return X

    End Function
    Public Function ubah_detil_peminjaman() As Integer
        Dim xUbah As String
        Dim X As Integer
        Dim xMyCmd As OdbcCommand

        xUbah = "UPDATE t_detil_proses_permintaan SET qty_jumlah=? WHERE no_permintaan=? and kd_barang=?"
        xMyCmd = New OdbcCommand(xUbah, MyCn)
        xMyCmd.Parameters.AddWithValue("qty_jumlah", Fqtyjumlah)
        xMyCmd.Parameters.AddWithValue("no_permintaan", Fnopermintaan)
        xMyCmd.Parameters.AddWithValue("kd_barang", Fkd_barang)

        xMyCmd.Prepare()
        X = xMyCmd.ExecuteNonQuery()
        xMyCmd.Dispose()
        Return X

    End Function
    Public Function TampilkanDatakeList(ByVal nm_brg As String) As List(Of Clsdetil_brg_masuk)
        Dim Q As String
        Dim xmycmd As OdbcCommand
        Dim xmyread As OdbcDataReader
        Dim tmpBaca As New List(Of Clsdetil_brg_masuk)

        Q = "select a.no_penerimaan,b.tgl_terima,a.kd_barang,c.nama_barang,a.total_qty_order,a.total_qty_terima,(a.total_qty_order-a.total_qty_terima) as selisih " & _
            " from detil_penerimaan_barang as a, penerimaan_barang as b, barang as c" & _
            " where a.no_penerimaan = b.no_penerimaan and a.kd_barang = c.kd_barang and" & _
            " a.total_qty_order<>a.total_qty_terima and c.nama_barang like '%" & Replace(nm_brg, "'", "''") & "%'"
        xmycmd = New OdbcCommand(Q, MyCn)

        xmyread = xmycmd.ExecuteReader
        If xmyread.HasRows Then
            While xmyread.Read
                Dim objTemp As New Clsdetil_brg_masuk
                objTemp.Fno_penerimaan = xmyread.Item("no_penerimaan")
                objTemp.FTgl_terima = xmyread.Item("tgl_terima")
                objTemp.Fkd_barang = xmyread.Item("kd_barang")
                objTemp.FNama_barang = xmyread.Item("nama_barang")
                objTemp.PTotal_qty_order = xmyread.Item("total_qty_order")
                objTemp.Ftotal_qty_terima = xmyread.Item("total_qty_terima")
                ' objTemp.FKeterangan = xmyread.Item("keterangan")
                tmpBaca.Add(objTemp)
            End While
        End If
        xmyread.Close() : Return tmpBaca
    End Function

    Public Function simpan_LPQC() As Integer
        Dim xSimpan As String
        Dim X As Integer
        Dim myCmd As OdbcCommand
        xSimpan = "INSERT INTO pengecekan_qc"
        xSimpan &= "(no_lpqc,id_login,kd_user,keterangan,no_lpb)"
        xSimpan &= "values (?,?,?,?,?)"
        myCmd = New OdbcCommand(xSimpan, MyCn)

        myCmd.Parameters.AddWithValue("no_lpqc", FNo_lpqc)
        myCmd.Parameters.AddWithValue("id_login", FId_login)
        myCmd.Parameters.AddWithValue("kd_user", FKd_user)
        myCmd.Parameters.AddWithValue("keterangan", FKeterangan)
        myCmd.Parameters.AddWithValue("no_lpb", FNo_lpb)
        myCmd.Prepare()
        X = myCmd.ExecuteNonQuery()

        myCmd.Dispose()
        Return X
    End Function


    Public Function ubah_qty_terima_qc_header12() As Integer
        Dim xUbah As String
        Dim X As Integer
        Dim xMyCmd As OdbcCommand

        xUbah = "UPDATE penerimaan_barang SET tgl_update_qc=current_timestamp(),kd_update_qc=?" & _
        " WHERE no_penerimaan=?"
        xMyCmd = New OdbcCommand(xUbah, MyCn)

        xMyCmd.Parameters.AddWithValue("kd_update_qc", FKd_user)
        xMyCmd.Parameters.AddWithValue("no_penerimaan", Fno_penerimaan)

        xMyCmd.Prepare()
        X = xMyCmd.ExecuteNonQuery()
        xMyCmd.Dispose()
        Return X

    End Function
    Public Function ubah_qty_terima_qc_header() As Integer
        Dim xUbah As String
        Dim X As Integer
        Dim xMyCmd As OdbcCommand

        xUbah = "UPDATE penerimaan_barang SET tgl_update_qc=current_timestamp(),kd_update_qc=?" & _
        " WHERE no_penerimaan=?"
        xMyCmd = New OdbcCommand(xUbah, MyCn)

        xMyCmd.Parameters.AddWithValue("kd_update_qc", FKd_user)
        xMyCmd.Parameters.AddWithValue("no_penerimaan", Fno_penerimaan)

        xMyCmd.Prepare()
        X = xMyCmd.ExecuteNonQuery()
        xMyCmd.Dispose()
        Return X

    End Function

    Public Function ubah_qty_terima_qc_header13() As Integer
        Dim xUbah As String
        Dim X As Integer
        Dim xMyCmd As OdbcCommand

        xUbah = "UPDATE penerimaan_barang SET tgl_update_qc=current_timestamp(),kd_update_qc=?" & _
        " WHERE no_penerimaan=?"
        xMyCmd = New OdbcCommand(xUbah, MyCn)

        xMyCmd.Parameters.AddWithValue("kd_update_qc", FKd_user)
        xMyCmd.Parameters.AddWithValue("no_penerimaan", Fno_penerimaan)

        xMyCmd.Prepare()
        X = xMyCmd.ExecuteNonQuery()
        xMyCmd.Dispose()
        Return X

    End Function


    Public Function ubah_qty_terima_qc_header2() As Integer
        Dim xUbah As String
        Dim X As Integer
        Dim xMyCmd As OdbcCommand

        xUbah = "UPDATE t_returan SET tgl_update_qc=curent_timestamp(),kd_update_qc=?" & _
        " WHERE no_lpb=?"
        xMyCmd = New OdbcCommand(xUbah, MyCn)

        xMyCmd.Parameters.AddWithValue("kd_update_qc", FKd_user)
        xMyCmd.Parameters.AddWithValue("no_lpb", Fno_penerimaan)

        xMyCmd.Prepare()
        X = xMyCmd.ExecuteNonQuery()
        xMyCmd.Dispose()
        Return X

    End Function


    Public Function simpan_LPQCdetail() As Integer
        Dim xSimpan As String
        Dim X As Integer
        Dim myCmd As OdbcCommand
        xSimpan = "INSERT INTO pengecekan_qc_detail"
        xSimpan &= "(no_lpqc,kd_barang,qty_terima,qty_rusak,qty_ready,keterangan)"
        xSimpan &= "values (?,?,?,?,?,?)"
        myCmd = New OdbcCommand(xSimpan, MyCn)

        myCmd.Parameters.AddWithValue("no_lpqc", FNo_lpqc)
        myCmd.Parameters.AddWithValue("kd_barang", Fkd_barang)
        myCmd.Parameters.AddWithValue("qty_terima", Ftotal_qty_terima)
        myCmd.Parameters.AddWithValue("qty_rusak", FQty_rusak)
        myCmd.Parameters.AddWithValue("qty_ready", FQty_bagus)
        myCmd.Parameters.AddWithValue("keterangan", FKeterangan)
        myCmd.Prepare()
        X = myCmd.ExecuteNonQuery()

        myCmd.Dispose()
        Return X
    End Function


    Public Function autonumber(ByVal xkd_user As String) As String
        '15/ADM/GA/VIII/13
        Dim bulan As String
        Dim romawi As String
        Dim nootomatis As String
        Dim strtemp As String = ""
        Dim strvalue As String = ""
        Dim str As String
        Dim ord As OdbcDataReader
        Dim oc As OdbcCommand
        romawi = ""
        bulan = Format(DateTime.Now.Date, "MM")
        If bulan = "01" Then
            romawi = "I"
        ElseIf bulan = "02" Then
            romawi = "II"
        ElseIf bulan = "03" Then
            romawi = "III"
        ElseIf bulan = "04" Then
            romawi = "IV"
        ElseIf bulan = "05" Then
            romawi = "V"
        ElseIf bulan = "06" Then
            romawi = "VI"
        ElseIf bulan = "07" Then
            romawi = "VII"
        ElseIf bulan = "08" Then
            romawi = "VIII"
        ElseIf bulan = "09" Then
            romawi = "IX"
        ElseIf bulan = "10" Then
            romawi = "X"
        ElseIf bulan = "11" Then
            romawi = "XI"
        ElseIf bulan = "12" Then
            romawi = "XII"
        End If

        str = "select no_lpqc from pengecekan_qc where year(tgl_lpqc)=year(curdate()) and month(tgl_lpqc)=month(curdate()) and kd_user='" & xkd_user & "' order by no_lpqc desc"
        oc = New OdbcCommand(str, MyCn)
        ord = oc.ExecuteReader

        If ord.Read Then

            strtemp = Val(Mid(ord.Item("no_lpqc"), 1, 4)) + 1

            strtemp = Right("0000", 4 - strtemp.Length) & strtemp
        Else
            strtemp = "0001"
        End If

        strvalue = Val(strtemp) + 1

        If strvalue >= 10 Then

            nootomatis = strtemp & "/LPQC/" & romawi & "/" & bulan & "/" & Format(DateTime.Now.Date, "yy")
        Else
            nootomatis = strtemp & "/LPQC/" & romawi & "/" & bulan & "/" & Format(DateTime.Now.Date, "yy")
        End If
        ord.Dispose()
        Return nootomatis

    End Function
End Class
