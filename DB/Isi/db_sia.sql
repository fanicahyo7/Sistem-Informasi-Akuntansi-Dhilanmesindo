-- phpMyAdmin SQL Dump
-- version 4.8.4
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: Mar 14, 2019 at 08:54 AM
-- Server version: 10.1.37-MariaDB
-- PHP Version: 7.3.0

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET AUTOCOMMIT = 0;
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `db_sia`
--

-- --------------------------------------------------------

--
-- Table structure for table `tb_barang`
--

CREATE TABLE `tb_barang` (
  `Kode_Barang` varchar(100) NOT NULL,
  `Nama_Barang` varchar(100) NOT NULL,
  `Satuan` varchar(100) NOT NULL,
  `Stok` int(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_barang`
--

INSERT INTO `tb_barang` (`Kode_Barang`, `Nama_Barang`, `Satuan`, `Stok`) VALUES
('BR0001', 'asd', 'asddddd', 93);

-- --------------------------------------------------------

--
-- Table structure for table `tb_bukubiaya`
--

CREATE TABLE `tb_bukubiaya` (
  `Kode_Transaksi` varchar(100) NOT NULL,
  `Tanggal` date NOT NULL,
  `Keterangan` text NOT NULL,
  `Biaya` double NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_bukubiaya`
--

INSERT INTO `tb_bukubiaya` (`Kode_Transaksi`, `Tanggal`, `Keterangan`, `Biaya`) VALUES
('TY0001', '2019-03-30', 'anu', 50000);

-- --------------------------------------------------------

--
-- Table structure for table `tb_bukuhutang`
--

CREATE TABLE `tb_bukuhutang` (
  `Kode_Hutang` varchar(100) NOT NULL,
  `Tanggal` date NOT NULL,
  `Keterangan` text NOT NULL,
  `Debet` double NOT NULL,
  `Kredit` double NOT NULL,
  `Saldo` double NOT NULL,
  `Status` varchar(50) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_bukuhutang`
--

INSERT INTO `tb_bukuhutang` (`Kode_Hutang`, `Tanggal`, `Keterangan`, `Debet`, `Kredit`, `Saldo`, `Status`) VALUES
('BH0001', '2019-03-12', 'gaji', 0, 250000, 250000, 'BELUM LUNAS');

-- --------------------------------------------------------

--
-- Table structure for table `tb_bukukas`
--

CREATE TABLE `tb_bukukas` (
  `ID` int(100) NOT NULL,
  `Kode_Transaksi` varchar(100) NOT NULL,
  `Tanggal` date NOT NULL,
  `Keterangan` text NOT NULL,
  `Debet` double NOT NULL,
  `Kredit` double NOT NULL,
  `Saldo` double NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_bukukas`
--

INSERT INTO `tb_bukukas` (`ID`, `Kode_Transaksi`, `Tanggal`, `Keterangan`, `Debet`, `Kredit`, `Saldo`) VALUES
(1, 'MD0001', '2019-03-12', 'Modal Awal', 12000000, 0, 12000000),
(2, 'BH0001', '2019-03-12', 'gaji', 250000, 0, 12250000),
(3, 'TJ0001', '2019-03-12', 'Penjualan Kredit', 0, 0, 12250000),
(9, 'TJ0002', '2019-03-13', 'Penjualan Tunai', 60000, 0, 12310000),
(20, 'BP0002', '2019-03-30', 'Pembayaran Piutang', 80000, 0, 12390000),
(30, '-', '2019-03-31', 'Saldo Akhir Kas', 12390000, 0, 12390000),
(31, '-', '2019-04-01', 'Saldo Awal Kas', 12390000, 0, 12390000),
(32, '-', '2019-04-01', 'Laba', 90000, 0, 12480000);

-- --------------------------------------------------------

--
-- Table structure for table `tb_bukupembelian`
--

CREATE TABLE `tb_bukupembelian` (
  `Kode_Transaksi` varchar(100) NOT NULL,
  `Tanggal` date NOT NULL,
  `Keterangan` text NOT NULL,
  `Debet` double NOT NULL,
  `Kredit` double NOT NULL,
  `Saldo` double NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `tb_bukupenjualan`
--

CREATE TABLE `tb_bukupenjualan` (
  `Kode_Transaksi` varchar(100) NOT NULL,
  `Tanggal` date NOT NULL,
  `Keterangan` text NOT NULL,
  `Debet` double NOT NULL,
  `Kredit` double NOT NULL,
  `Saldo` double NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_bukupenjualan`
--

INSERT INTO `tb_bukupenjualan` (`Kode_Transaksi`, `Tanggal`, `Keterangan`, `Debet`, `Kredit`, `Saldo`) VALUES
('TJ0001', '2019-03-12', 'Penjualan Kredit', 0, 100000, 100000),
('TJ0002', '2019-03-13', 'Penjualan Tunai', 0, 60000, 160000);

-- --------------------------------------------------------

--
-- Table structure for table `tb_bukupersediaanbarang`
--

CREATE TABLE `tb_bukupersediaanbarang` (
  `ID` int(100) NOT NULL,
  `Kode_Transaksi` varchar(100) NOT NULL,
  `Tanggal` date NOT NULL,
  `Kode_Barang` varchar(100) NOT NULL,
  `Nama_Barang` varchar(100) NOT NULL,
  `Satuan` varchar(100) NOT NULL,
  `Dibeli` double NOT NULL,
  `Dijual` double NOT NULL,
  `Retur` double NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_bukupersediaanbarang`
--

INSERT INTO `tb_bukupersediaanbarang` (`ID`, `Kode_Transaksi`, `Tanggal`, `Kode_Barang`, `Nama_Barang`, `Satuan`, `Dibeli`, `Dijual`, `Retur`) VALUES
(1, 'TJ0001', '2019-03-12', 'BR0001', 'asd', 'asddddd', 0, 5, 0),
(2, 'TJ0002', '2019-03-13', 'BR0001', 'asd', 'asddddd', 0, 2, 0);

-- --------------------------------------------------------

--
-- Table structure for table `tb_bukupiutang`
--

CREATE TABLE `tb_bukupiutang` (
  `ID` int(100) NOT NULL,
  `Kode_Transaksi_Pembayaran` varchar(100) NOT NULL,
  `Kode_Transaksi_Penjualan` varchar(100) NOT NULL,
  `Tanggal` date NOT NULL,
  `Keterangan` text NOT NULL,
  `Debet` double NOT NULL,
  `Kredit` double NOT NULL,
  `Saldo` double NOT NULL,
  `Status` varchar(50) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_bukupiutang`
--

INSERT INTO `tb_bukupiutang` (`ID`, `Kode_Transaksi_Pembayaran`, `Kode_Transaksi_Penjualan`, `Tanggal`, `Keterangan`, `Debet`, `Kredit`, `Saldo`, `Status`) VALUES
(1, '-', 'TJ0001', '2019-03-12', 'Penjualan Kredit', 100000, 0, 100000, 'LUNAS'),
(2, 'BP0002', 'TJ0001', '2019-03-30', 'Pembayaran Piutang', 0, 100000, 0, 'LUNAS');

-- --------------------------------------------------------

--
-- Table structure for table `tb_labarugi`
--

CREATE TABLE `tb_labarugi` (
  `ID` int(11) NOT NULL,
  `Tanggal` date NOT NULL,
  `Pendapatan` varchar(100) NOT NULL,
  `Saldo_Pendapatan` double NOT NULL,
  `Biaya` varchar(100) NOT NULL,
  `Saldo_Biaya` double NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_labarugi`
--

INSERT INTO `tb_labarugi` (`ID`, `Tanggal`, `Pendapatan`, `Saldo_Pendapatan`, `Biaya`, `Saldo_Biaya`) VALUES
(13, '2019-03-31', 'Penjualan', 160000, 'Biaya', 50000),
(14, '2019-03-31', '-', 0, 'Potongan', 20000),
(15, '2019-03-31', 'Laba', 90000, '-', 0);

-- --------------------------------------------------------

--
-- Table structure for table `tb_modal`
--

CREATE TABLE `tb_modal` (
  `ID` int(100) NOT NULL,
  `Tanggal` date NOT NULL,
  `Keterangan` varchar(100) NOT NULL,
  `Jumlah` double NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_modal`
--

INSERT INTO `tb_modal` (`ID`, `Tanggal`, `Keterangan`, `Jumlah`) VALUES
(1, '2019-03-01', 'Modal Awal', 12000000),
(7, '2019-03-31', 'Laba 90000', 12090000);

-- --------------------------------------------------------

--
-- Table structure for table `tb_neraca`
--

CREATE TABLE `tb_neraca` (
  `ID` int(100) NOT NULL,
  `Tanggal` date NOT NULL,
  `Aktiva` varchar(100) NOT NULL,
  `Saldo_Aktiva` double NOT NULL,
  `Pasiva` varchar(100) NOT NULL,
  `Saldo_Pasiva` double NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_neraca`
--

INSERT INTO `tb_neraca` (`ID`, `Tanggal`, `Aktiva`, `Saldo_Aktiva`, `Pasiva`, `Saldo_Pasiva`) VALUES
(43, '2019-03-31', 'Kas', 12390000, '-', 0),
(44, '2019-03-31', 'Piutang', 0, '-', 0),
(45, '2019-03-31', 'Potongan', 20000, '-', 0),
(46, '2019-03-31', '-', 0, 'Hutang', 250000),
(47, '2019-03-31', '-', 0, 'Penjualan', 160000),
(48, '2019-03-31', '-', 0, 'Modal', 12000000);

-- --------------------------------------------------------

--
-- Table structure for table `tb_potongan`
--

CREATE TABLE `tb_potongan` (
  `ID` int(100) NOT NULL,
  `Tanggal` date NOT NULL,
  `Keterangan` text NOT NULL,
  `Jumlah` double NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_potongan`
--

INSERT INTO `tb_potongan` (`ID`, `Tanggal`, `Keterangan`, `Jumlah`) VALUES
(2, '2019-03-30', 'Potongan Penjualan TJ0001', 20000);

-- --------------------------------------------------------

--
-- Table structure for table `tb_retur`
--

CREATE TABLE `tb_retur` (
  `Kode_Retur` varchar(100) NOT NULL,
  `Tanggal` date NOT NULL,
  `Kode_Transaksi` varchar(100) NOT NULL,
  `Debet` double NOT NULL,
  `Kredit` double NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `tb_returrinci`
--

CREATE TABLE `tb_returrinci` (
  `Kode_Barang` varchar(100) NOT NULL,
  `Nama_Barang` varchar(100) NOT NULL,
  `Satuan` varchar(100) NOT NULL,
  `Jumlah_Retur` double NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `tb_user`
--

CREATE TABLE `tb_user` (
  `Username` varchar(100) NOT NULL,
  `Password` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_user`
--

INSERT INTO `tb_user` (`Username`, `Password`) VALUES
('admin', 'admin');

--
-- Indexes for dumped tables
--

--
-- Indexes for table `tb_barang`
--
ALTER TABLE `tb_barang`
  ADD PRIMARY KEY (`Kode_Barang`);

--
-- Indexes for table `tb_bukubiaya`
--
ALTER TABLE `tb_bukubiaya`
  ADD PRIMARY KEY (`Kode_Transaksi`);

--
-- Indexes for table `tb_bukuhutang`
--
ALTER TABLE `tb_bukuhutang`
  ADD PRIMARY KEY (`Kode_Hutang`);

--
-- Indexes for table `tb_bukukas`
--
ALTER TABLE `tb_bukukas`
  ADD PRIMARY KEY (`ID`);

--
-- Indexes for table `tb_bukupembelian`
--
ALTER TABLE `tb_bukupembelian`
  ADD PRIMARY KEY (`Kode_Transaksi`);

--
-- Indexes for table `tb_bukupenjualan`
--
ALTER TABLE `tb_bukupenjualan`
  ADD PRIMARY KEY (`Kode_Transaksi`);

--
-- Indexes for table `tb_bukupersediaanbarang`
--
ALTER TABLE `tb_bukupersediaanbarang`
  ADD PRIMARY KEY (`ID`);

--
-- Indexes for table `tb_bukupiutang`
--
ALTER TABLE `tb_bukupiutang`
  ADD PRIMARY KEY (`ID`);

--
-- Indexes for table `tb_labarugi`
--
ALTER TABLE `tb_labarugi`
  ADD PRIMARY KEY (`ID`);

--
-- Indexes for table `tb_modal`
--
ALTER TABLE `tb_modal`
  ADD PRIMARY KEY (`ID`);

--
-- Indexes for table `tb_neraca`
--
ALTER TABLE `tb_neraca`
  ADD PRIMARY KEY (`ID`);

--
-- Indexes for table `tb_potongan`
--
ALTER TABLE `tb_potongan`
  ADD PRIMARY KEY (`ID`);

--
-- Indexes for table `tb_retur`
--
ALTER TABLE `tb_retur`
  ADD PRIMARY KEY (`Kode_Retur`);

--
-- Indexes for table `tb_returrinci`
--
ALTER TABLE `tb_returrinci`
  ADD PRIMARY KEY (`Kode_Barang`);

--
-- Indexes for table `tb_user`
--
ALTER TABLE `tb_user`
  ADD PRIMARY KEY (`Username`);

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `tb_bukukas`
--
ALTER TABLE `tb_bukukas`
  MODIFY `ID` int(100) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=33;

--
-- AUTO_INCREMENT for table `tb_bukupersediaanbarang`
--
ALTER TABLE `tb_bukupersediaanbarang`
  MODIFY `ID` int(100) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;

--
-- AUTO_INCREMENT for table `tb_bukupiutang`
--
ALTER TABLE `tb_bukupiutang`
  MODIFY `ID` int(100) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;

--
-- AUTO_INCREMENT for table `tb_labarugi`
--
ALTER TABLE `tb_labarugi`
  MODIFY `ID` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=16;

--
-- AUTO_INCREMENT for table `tb_modal`
--
ALTER TABLE `tb_modal`
  MODIFY `ID` int(100) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=8;

--
-- AUTO_INCREMENT for table `tb_neraca`
--
ALTER TABLE `tb_neraca`
  MODIFY `ID` int(100) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=49;

--
-- AUTO_INCREMENT for table `tb_potongan`
--
ALTER TABLE `tb_potongan`
  MODIFY `ID` int(100) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
