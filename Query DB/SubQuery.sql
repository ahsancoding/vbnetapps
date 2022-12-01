
-- Sub query tabelBarang, Kartu Stok Tito dan Stok Tito
select a.keterangan,a.kd_barang,b.nama_barang,sum(transfer_in)-sum(transfer_out) as qty_sisa, c.stok, c.kd_divwh, c.kd_lokasi
from kartustok_tito as a, barang AS b, (select * from stok_tito where kd_divwh='11' and kd_lokasi='01') as c
where a.kd_barang = b.kd_barang and b.kd_barang=c.kd_barang
and a.kd_divwh='11' and a.kd_lokasi='01' and a.keterangan='032/RT/IV/22HQAK'
group by a.kd_barang,a.keterangan