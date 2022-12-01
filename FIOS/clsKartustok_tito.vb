Imports System.Data.Odbc
Public Class clsKartustok_tito
    Private Fno_transfer As String
    Private Fkd_barang As String
    Private FKd_lokasi As String
    Private Fsaldo_awal As Integer
    Private Ftransfer_in As Integer
    Private Ftransfer_out As Integer
    Private Fsaldo_akhir As Integer
    Private Fkd_barangakhir As String
    Private FKeterangan As String
    Private Fkd_divwh As String
    Private Ftgl_transaksi As Date
    Public Property Pno_transfer() As String
        Get
            Return Fno_transfer
        End Get
        Set(ByVal value As String)
            Fno_transfer = value
        End Set
    End Property
    Public Property Pkd_barang() As String
        Get
            Return Fkd_barang
        End Get
        Set(ByVal value As String)
            Fkd_barang = value
        End Set
    End Property
    Public Property PKd_lokasi() As String
        Get
            Return FKd_lokasi
        End Get
        Set(ByVal value As String)
            FKd_lokasi = value
        End Set
    End Property
    Public Property PSaldo_awal() As Integer
        Get
            Return Fsaldo_awal
        End Get
        Set(ByVal value As Integer)
            Fsaldo_awal = value
        End Set
    End Property
    Public Property PTransfer_in() As Integer
        Get
            Return Ftransfer_in
        End Get
        Set(ByVal value As Integer)
            Ftransfer_in = value
        End Set
    End Property
    Public Property PTransfer_out() As Integer
        Get
            Return Ftransfer_out
        End Get
        Set(ByVal value As Integer)
            Ftransfer_out = value
        End Set
    End Property
    Public Property Psaldo_akhir() As Integer
        Get
            Return Fsaldo_akhir
        End Get
        Set(ByVal value As Integer)
            Fsaldo_akhir = value
        End Set
    End Property
    Public Property PKd_barangakhir() As String
        Get
            Return Fkd_barangakhir
        End Get
        Set(ByVal value As String)
            Fkd_barangakhir = value
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
    Public Property PKd_divwh() As String
        Get
            Return Fkd_divwh
        End Get
        Set(ByVal value As String)
            Fkd_divwh = value
        End Set
    End Property
    Public Property Ptgl_transaksi() As Date
        Get
            Return Ftgl_transaksi
        End Get
        Set(ByVal value As Date)
            Ftgl_transaksi = value
        End Set
    End Property
    Public Function autonumber(ByVal kd_lokasi As String, ByVal kd_divwh As String) As String
        bukaConn()
        Dim bulan As String
        Dim tanggal As String
        Dim nootomatis As String
        Dim strtemp As String = ""
        Dim strvalue As String = ""
        Dim str As String
        Dim ord As OdbcDataReader
        Dim oc As OdbcCommand
        bulan = Format(DateTime.Now.Date, "MM")
        tanggal = Format(DateTime.Now.Date, "dd")


        str = " SELECT no_transfer FROM kartustok_tito where kd_lokasi='" & kd_lokasi & "' and kd_divwh='" & kd_divwh & "' and month(tgl_transaksi)=month(curdate()) AND year(tgl_transaksi)=year(curdate()) ORDER By no_transfer desc"
        oc = New OdbcCommand(str, MyCn)
        ord = oc.ExecuteReader

        If ord.Read Then

            strtemp = Val(Right(ord.Item("no_transfer"), 4)) + 1

            strtemp = Right("0000", 4 - strtemp.Length) & strtemp
        Else
            strtemp = "0001"
        End If

        strvalue = Val(strtemp) + 1

        If strvalue >= 10 Then

            nootomatis = tanggal & bulan & Format(DateTime.Now.Date, "yy") & "-" & strtemp
        Else
            nootomatis = tanggal & bulan & Format(DateTime.Now.Date, "yy") & "-" & strtemp
        End If
        ord.Dispose()
        Return nootomatis
        MyCn.Close()
    End Function
    Public Function simpan() As Integer
        Dim xSimpan As String
        Dim X As Integer
        Dim myCmd As OdbcCommand
        xSimpan = "INSERT INTO kartustok_tito"
        xSimpan &= "(no_transfer,kd_barang,kd_lokasi,saldo_awal,transfer_in,transfer_out,saldo_akhir,kd_barangakhir,keterangan,kd_divwh,tgl_transaksi)"
        xSimpan &= "values (?,?,?,?,?,?,?,?,?,?,?)"
        myCmd = New OdbcCommand(xSimpan, MyCn)

        myCmd.Parameters.AddWithValue("no_transfer", Fno_transfer)
        myCmd.Parameters.AddWithValue("kd_barang", Fkd_barang)
        myCmd.Parameters.AddWithValue("kd_lokasi", FKd_lokasi)
        myCmd.Parameters.AddWithValue("saldo_awal", Fsaldo_awal)
        myCmd.Parameters.AddWithValue("transfer_in", Ftransfer_in)
        myCmd.Parameters.AddWithValue("transfer_out", Ftransfer_out)
        myCmd.Parameters.AddWithValue("saldo_akhir", Fsaldo_akhir)
        myCmd.Parameters.AddWithValue("kd_barangakhir", Fkd_barangakhir)
        myCmd.Parameters.AddWithValue("keterangan", FKeterangan)
        myCmd.Parameters.AddWithValue("kd_divwh", Fkd_divwh)
        myCmd.Parameters.AddWithValue("tgl_transaksi", Ftgl_transaksi)

        myCmd.Prepare()
        X = myCmd.ExecuteNonQuery()

        myCmd.Dispose()
        Return X

    End Function
    'MANUAL
    Public Function simpan_manual() As Integer
        bukaCon4()
        Dim xSimpan As String
        Dim X As Integer
        Dim myCmd As OdbcCommand
        xSimpan = "INSERT INTO kartustok_tito_manual"
        xSimpan &= "(no_transfer,kd_barang,kd_lokasi,saldo_awal,transfer_in,transfer_out,saldo_akhir,kd_barangakhir,keterangan,kd_divwh,tgl_transaksi)"
        xSimpan &= "values (?,?,?,?,?,?,?,?,?,?,?)"
        myCmd = New OdbcCommand(xSimpan, MyCn)

        myCmd.Parameters.AddWithValue("no_transfer", Fno_transfer)
        myCmd.Parameters.AddWithValue("kd_barang", Fkd_barang)
        myCmd.Parameters.AddWithValue("kd_lokasi", FKd_lokasi)
        myCmd.Parameters.AddWithValue("saldo_awal", Fsaldo_awal)
        myCmd.Parameters.AddWithValue("transfer_in", Ftransfer_in)
        myCmd.Parameters.AddWithValue("transfer_out", Ftransfer_out)
        myCmd.Parameters.AddWithValue("saldo_akhir", Fsaldo_akhir)
        myCmd.Parameters.AddWithValue("kd_barangakhir", Fkd_barangakhir)
        myCmd.Parameters.AddWithValue("keterangan", FKeterangan)
        myCmd.Parameters.AddWithValue("kd_divwh", Fkd_divwh)
        myCmd.Parameters.AddWithValue("tgl_transaksi", Ftgl_transaksi)

        myCmd.Prepare()
        X = myCmd.ExecuteNonQuery()

        myCmd.Dispose()
        Return X

    End Function
  

    Public Function autonumber_manual(ByVal kd_lokasi As String, ByVal kd_divwh As String) As String
        bukaCon4()
        Dim bulan As String
        Dim tanggal As String
        Dim nootomatis As String
        Dim strtemp As String = ""
        Dim strvalue As String = ""
        Dim str As String
        Dim ord As OdbcDataReader
        Dim oc As OdbcCommand
        bulan = Format(DateTime.Now.Date, "MM")
        tanggal = Format(DateTime.Now.Date, "dd")


        str = " SELECT no_transfer FROM kartustok_tito_manual where kd_lokasi='" & kd_lokasi & "' and kd_divwh='" & kd_divwh & "' and month(tgl_transaksi)=month(curdate()) AND year(tgl_transaksi)=year(curdate()) ORDER By no_transfer desc"
        oc = New OdbcCommand(str, MyCn)
        ord = oc.ExecuteReader

        If ord.Read Then

            strtemp = Val(Right(ord.Item("no_transfer"), 4)) + 1

            strtemp = Right("0000", 4 - strtemp.Length) & strtemp
        Else
            strtemp = "0001"
        End If

        strvalue = Val(strtemp) + 1

        If strvalue >= 10 Then

            nootomatis = tanggal & bulan & Format(DateTime.Now.Date, "yy") & "-" & strtemp
        Else
            nootomatis = tanggal & bulan & Format(DateTime.Now.Date, "yy") & "-" & strtemp
        End If
        ord.Dispose()
        Return nootomatis
        MyCn.Close()
    End Function
End Class
