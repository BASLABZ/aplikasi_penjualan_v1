-- phpMyAdmin SQL Dump
-- version 3.3.9
-- http://www.phpmyadmin.net
--
-- Host: localhost
-- Waktu pembuatan: 15. Juni 2014 jam 06:49
-- Versi Server: 5.5.8
-- Versi PHP: 5.3.5

SET SQL_MODE="NO_AUTO_VALUE_ON_ZERO";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

--
-- Database: `tokoanggutjaya`
--

-- --------------------------------------------------------

--
-- Struktur dari tabel `barang`
--

CREATE TABLE IF NOT EXISTS `barang` (
  `IdBarang` char(4) NOT NULL,
  `IdJenisBarang` char(4) NOT NULL,
  `NamaBarang` varchar(40) NOT NULL,
  `HargaBeli` int(11) NOT NULL,
  `HargaJual` int(11) NOT NULL,
  `Stok` int(11) NOT NULL,
  PRIMARY KEY (`IdBarang`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `barang`
--

INSERT INTO `barang` (`IdBarang`, `IdJenisBarang`, `NamaBarang`, `HargaBeli`, `HargaJual`, `Stok`) VALUES
('B001', 'J002', 'Kalengan Cinta', 800, 1000, 40),
('B002', 'J001', 'oreo', 2000, 3000, 87);

-- --------------------------------------------------------

--
-- Struktur dari tabel `detailpembelian`
--

CREATE TABLE IF NOT EXISTS `detailpembelian` (
  `IdDetailPembelian` char(6) NOT NULL,
  `IdPembelian` char(4) NOT NULL,
  `IdBarang` char(4) NOT NULL,
  `HargaBeli` int(11) NOT NULL,
  `Jumlah` int(11) NOT NULL,
  PRIMARY KEY (`IdDetailPembelian`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `detailpembelian`
--

INSERT INTO `detailpembelian` (`IdDetailPembelian`, `IdPembelian`, `IdBarang`, `HargaBeli`, `Jumlah`) VALUES
('DP0001', 'P002', 'B001', 800, 20),
('DP0002', 'P003', 'B001', 800, 0),
('DP0003', 'P003', 'B002', 2000, 0),
('DP0004', 'P004', 'B002', 2000, 10),
('DP0005', 'P005', 'B002', 2000, 50),
('DP0006', 'P006', 'B002', 2000, 2);

-- --------------------------------------------------------

--
-- Struktur dari tabel `detailpenjualan`
--

CREATE TABLE IF NOT EXISTS `detailpenjualan` (
  `IdPenjualan` char(4) NOT NULL,
  `IdBarang` char(4) NOT NULL,
  `HargaJual` int(11) NOT NULL,
  `Jumlah` int(11) NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `detailpenjualan`
--

INSERT INTO `detailpenjualan` (`IdPenjualan`, `IdBarang`, `HargaJual`, `Jumlah`) VALUES
('P001', 'B002', 3000, 10),
('P002', 'B002', 3000, 25),
('P003', 'B002', 3000, 10),
('P003', 'B001', 1000, 5),
('P004', 'B001', 1000, 5),
('P005', 'B002', 3000, 15),
('P007', 'B002', 3000, 3),
('P008', 'B001', 1000, 3),
('P009', 'B002', 3000, 3),
('P010', 'B001', 1000, 3),
('P011', 'B002', 3000, 1),
('P012', 'B002', 3000, 9),
('P012', 'B001', 1000, 4);

-- --------------------------------------------------------

--
-- Struktur dari tabel `jenisbarang`
--

CREATE TABLE IF NOT EXISTS `jenisbarang` (
  `IdJenisBarang` char(4) NOT NULL,
  `NamaJenisBarang` varchar(40) NOT NULL,
  PRIMARY KEY (`IdJenisBarang`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `jenisbarang`
--

INSERT INTO `jenisbarang` (`IdJenisBarang`, `NamaJenisBarang`) VALUES
('J001', 'Kaleng'),
('J002', 'Susu'),
('J003', 'Kalengan');

-- --------------------------------------------------------

--
-- Struktur dari tabel `pembelian`
--

CREATE TABLE IF NOT EXISTS `pembelian` (
  `IdPembelian` char(4) NOT NULL,
  `IdUser` char(4) NOT NULL,
  `IdSupplier` char(4) NOT NULL,
  `TanggalBeli` date NOT NULL,
  `Total` int(4) NOT NULL,
  PRIMARY KEY (`IdPembelian`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `pembelian`
--

INSERT INTO `pembelian` (`IdPembelian`, `IdUser`, `IdSupplier`, `TanggalBeli`, `Total`) VALUES
('P001', 'User', 'S001', '2014-05-29', 104000),
('P002', 'User', 'S002', '2014-05-29', 104000),
('P003', 'User', 'S001', '2014-05-29', 0),
('P004', 'User', 'S002', '2014-05-30', 104000),
('P005', 'User', 'S002', '2014-06-02', 100000),
('P006', 'User', 'S002', '2014-06-04', 4000);

-- --------------------------------------------------------

--
-- Struktur dari tabel `penjualan`
--

CREATE TABLE IF NOT EXISTS `penjualan` (
  `IdPenjualan` char(4) NOT NULL,
  `idUser` char(4) NOT NULL,
  `TglJual` date NOT NULL,
  `Total` int(11) NOT NULL,
  PRIMARY KEY (`IdPenjualan`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `penjualan`
--

INSERT INTO `penjualan` (`IdPenjualan`, `idUser`, `TglJual`, `Total`) VALUES
('P001', 'User', '2014-05-29', 30000),
('P002', 'User', '2014-05-29', 75000),
('P003', 'User', '2014-05-29', 35000),
('P004', 'User', '2014-05-29', 5000),
('P005', 'User', '2014-06-02', 45000),
('P006', 'User', '2014-06-04', 0),
('P007', 'User', '2014-06-04', 9000),
('P008', 'User', '2014-06-04', 3000),
('P009', 'User', '2014-06-04', 9000),
('P010', 'User', '2014-06-04', 3000),
('P011', 'User', '2014-06-04', 3000),
('P012', 'User', '2014-06-04', 31000);

-- --------------------------------------------------------

--
-- Struktur dari tabel `supplier`
--

CREATE TABLE IF NOT EXISTS `supplier` (
  `IdSupplier` char(4) NOT NULL,
  `NamaSupplier` varchar(40) NOT NULL,
  `AlamatSupplier` varchar(40) NOT NULL,
  PRIMARY KEY (`IdSupplier`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `supplier`
--

INSERT INTO `supplier` (`IdSupplier`, `NamaSupplier`, `AlamatSupplier`) VALUES
('S001', 'Toko Aneka', 'Sleman'),
('S002', 'Toko Aneka Sejahtera', 'Sleman');

-- --------------------------------------------------------

--
-- Struktur dari tabel `user`
--

CREATE TABLE IF NOT EXISTS `user` (
  `IdUser` char(4) NOT NULL,
  `Nama` varchar(40) NOT NULL,
  `Password` varchar(6) NOT NULL,
  `Level` varchar(8) NOT NULL,
  PRIMARY KEY (`IdUser`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `user`
--

INSERT INTO `user` (`IdUser`, `Nama`, `Password`, `Level`) VALUES
('U001', 'admin', 'admin', 'Admin'),
('U002', 'Kasir', 'kasir', 'Kasir'),
('U003', 'Pemilik', 'pem', 'Pemilik');
