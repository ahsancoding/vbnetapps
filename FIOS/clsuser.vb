Imports System.Data.Odbc
Public Class clsuser
    Private FKd_user As String
    Private Fusername As String
    Private Fpassword As String
    Private Fhak_akses As String
    Private Fkd_lokasi As String
    Private Fjenis_barang As String
    Private FTotal_login As Integer
    Private Fkddiv As String
    Public Property Pkddiv() As String
        Get
            Return Fkddiv
        End Get
        Set(ByVal value As String)
            Fkddiv = value
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
    Public Property PJenis_barang() As String
        Get
            Return Fjenis_barang
        End Get
        Set(ByVal value As String)
            Fjenis_barang = value
        End Set
    End Property
    Public Property PKd_lokasi() As String
        Get
            Return Fkd_lokasi
        End Get
        Set(ByVal value As String)
            Fkd_lokasi = value
        End Set
    End Property
    Public Property Pusername() As String
        Get
            Return Fusername
        End Get
        Set(ByVal value As String)
            Fusername = value
        End Set
    End Property

    Public Property Ppassword() As String
        Get
            Return Fpassword
        End Get
        Set(ByVal value As String)
            Fpassword = value
        End Set
    End Property

    Public Property Phak_akses() As String
        Get
            Return Fhak_akses
        End Get
        Set(ByVal value As String)
            Fhak_akses = value
        End Set
    End Property
    Public Property PTotal_login() As Integer
        Get
            Return FTotal_login
        End Get
        Set(ByVal value As Integer)
            FTotal_login = value
        End Set
    End Property
    Private Falias As String
    Public Property Palias() As String
        Get
            Return Falias
        End Get
        Set(ByVal value As String)
            Falias = value
        End Set
    End Property
    Private Faliasdiv As String
    Public Property Paliasdiv() As String
        Get
            Return Faliasdiv
        End Get
        Set(ByVal value As String)
            Faliasdiv = value
        End Set
    End Property
    Public Function simpan() As Integer
        Dim xSimpan As String
        Dim X As Integer
        Dim myCmd As OdbcCommand
        xSimpan = "INSERT INTO user"
        xSimpan &= "(kd_user,username,password,hak_akses,kd_lokasi,jenis_barang)"
        xSimpan &= "values (?,?,?,?,?)"
        myCmd = New OdbcCommand(xSimpan, MyCn)

        myCmd.Parameters.AddWithValue("kd_user", FKd_user)
        myCmd.Parameters.AddWithValue("username", Fusername)
        myCmd.Parameters.AddWithValue("password", Fpassword)
        myCmd.Parameters.AddWithValue("hak_akses", Fhak_akses)
        myCmd.Parameters.AddWithValue("kd_lokasi", Fkd_lokasi)
        myCmd.Parameters.AddWithValue("jenis_barang", Fjenis_barang)
        myCmd.Prepare()
        X = myCmd.ExecuteNonQuery()

        myCmd.Dispose()
        Return X
    End Function

    Public Function ubah() As Integer
        Dim xUbah As String
        Dim X As Integer
        Dim xMyCmd As OdbcCommand

        xUbah = "UPDATE user SET password=?,hak_akses=?,kd_lokasi=?,jenis_barang=?"
        xUbah &= "WHERE username=?"
        xMyCmd = New OdbcCommand(xUbah, MyCn)

        xMyCmd.Parameters.AddWithValue("password", Fpassword)
        xMyCmd.Parameters.AddWithValue("hak_akses", Fhak_akses)
        xMyCmd.Parameters.AddWithValue("kd_lokasi", Fkd_lokasi)
        xMyCmd.Parameters.AddWithValue("jenis_barang", Fjenis_barang)
        xMyCmd.Parameters.AddWithValue("username", Fusername)

        xMyCmd.Prepare()
        X = xMyCmd.ExecuteNonQuery()
        xMyCmd.Dispose()
        Return X

    End Function

    Public Function ubah_password() As Integer
        Dim xUbah As String
        Dim X As Integer
        Dim xMyCmd As OdbcCommand

        xUbah = "UPDATE user SET passwordn=? WHERE kd_user=?"
        xMyCmd = New OdbcCommand(xUbah, MyCn)

        xMyCmd.Parameters.AddWithValue("passwordn", Fpassword)
        xMyCmd.Parameters.AddWithValue("kd_user", FKd_user)

        xMyCmd.Prepare()
        X = xMyCmd.ExecuteNonQuery()
        xMyCmd.Dispose()
        Return X

    End Function
    Public Function hapus() As Integer
        Dim xUbah As String
        Dim X As Integer
        Dim xMyCmd As OdbcCommand

        xUbah = "DELETE FROM user WHERE username=?"
        xMyCmd = New OdbcCommand(xUbah, MyCn)
        xMyCmd.Parameters.AddWithValue("username", Fusername)
        xMyCmd.Prepare()
        X = xMyCmd.ExecuteNonQuery()

        xMyCmd.Dispose()
        Return X
    End Function

    Public Function getUser() As Boolean
        bukaConn()
        Dim query As String
        Dim xmyread As OdbcDataReader
        Dim xmycmd As OdbcCommand

        query = "SELECT * FROM user as a,divisi as b WHERE a.kd_div=b.kd_divisi and username=?"
        xmycmd = New OdbcCommand(query, MyCn)
        xmycmd.Parameters.AddWithValue("username", Fusername)
        xmyread = xmycmd.ExecuteReader
        If xmyread.HasRows Then
            xmyread.Read()
            FKd_user = xmyread.Item("kd_user").ToString
            Fusername = xmyread.Item("username").ToString
            Fpassword = xmyread.Item("passwordn").ToString
            Fhak_akses = xmyread.Item("hak_akses").ToString
            Fkd_lokasi = xmyread.Item("kd_lokasi").ToString
            Fjenis_barang = xmyread.Item("jenis_barang").ToString
            FTotal_login = Val(xmyread.Item("total_login").ToString)
            Fkddiv = xmyread.Item("kd_div").ToString
            Falias = xmyread.Item("alias").ToString
            Faliasdiv = xmyread.Item("nm_divisi").ToString
            xmyread.Close()
            Return True
        Else
            xmyread.Close()
            Return False
        End If
    End Function
    Public Function getKd_user() As Boolean
        Dim query As String
        Dim xmyread As OdbcDataReader
        Dim xmycmd As OdbcCommand

        query = "SELECT kd_user FROM user WHERE name=?"
        xmycmd = New OdbcCommand(query, MyCn)
        xmycmd.Parameters.AddWithValue("name", Fusername)
        xmyread = xmycmd.ExecuteReader
        If xmyread.HasRows Then
            xmyread.Read()
            FKd_user = xmyread.Item("kd_user").ToString
            xmyread.Close()
            Return True
        Else
            xmyread.Close()
            Return False
        End If
    End Function

    Public Function tampilUser(ByVal xusername As String) As List(Of clsuser)
        Dim Q As String
        Dim xmycmd As OdbcCommand
        Dim xmyread As OdbcDataReader
        Dim tmpBaca As New List(Of clsuser)

        Q = "SELECT kd_user,username,password,hak_akses,kd_lokasi,jenis_barang FROM user WHERE username like '%" & Replace(xusername, "'", "''") & "%'"
        xmycmd = New OdbcCommand(Q, MyCn)

        xmyread = xmycmd.ExecuteReader
        If xmyread.HasRows Then
            While xmyread.Read
                Dim objTemp As New clsuser
                objTemp.FKd_user = xmyread.Item("kd_user")
                objTemp.Fusername = xmyread.Item("username")
                objTemp.Fpassword = xmyread.Item("password")
                objTemp.Fhak_akses = xmyread.Item("hak_akses")
                objTemp.Fkd_lokasi = xmyread.Item("kd_lokasi")
                objTemp.Fjenis_barang = xmyread.Item("jenis_barang")
              
                tmpBaca.Add(objTemp)
            End While
        End If
        xmyread.Close() : Return tmpBaca
    End Function
End Class
