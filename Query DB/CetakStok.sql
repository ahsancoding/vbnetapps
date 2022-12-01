
 SELECT `barang`.`nama_barang`, `t_detil_storage_pa_pp`.`kd_barang`, `t_storage_pa_pp`.`kd_lokasi`, `t_detil_storage_pa_pp`.`qtybad`, `t_detil_storage_pa_pp`.`qtygood`, `penerimaan_barang`.`no_penerimaan`, `t_storage_pa_pp`.`tgl_trans`, `t_detil_storage_pa_pp`.`ket`, `t_detil_storage_pa_pp`.`Tgl_Simpan`, `t_detil_storage_pa_pp`.`kirim`, `t_storage_pa_pp`.`no_storage`, `pengecekan_qc_detail`.`qty_terima`
 FROM   `firman_indonesia`.`pengecekan_qc_detail` `pengecekan_qc_detail` 
 INNER JOIN (((`firman_indonesia`.`t_detil_storage_pa_pp` `t_detil_storage_pa_pp` INNER JOIN `firman_indonesia`.`t_storage_pa_pp` `t_storage_pa_pp` ON `t_detil_storage_pa_pp`.`no_storage`=`t_storage_pa_pp`.`no_storage`) 
 INNER JOIN `firman_indonesia`.`barang` `barang` ON `t_detil_storage_pa_pp`.`kd_barang`=`barang`.`kd_barang`) 
 INNER JOIN `firman_indonesia`.`penerimaan_barang` `penerimaan_barang` ON `t_storage_pa_pp`.`no_penerimaan`=`penerimaan_barang`.`no_penerimaan`) 
 ON `pengecekan_qc_detail`.`kd_barang`=`barang`.`kd_barang`

