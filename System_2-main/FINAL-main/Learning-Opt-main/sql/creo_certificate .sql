-- phpMyAdmin SQL Dump
-- version 5.2.1
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: Jul 10, 2025 at 01:21 AM
-- Server version: 10.4.32-MariaDB
-- PHP Version: 8.2.12

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;


CREATE TABLE uploaded_files (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  file_name TEXT NOT NULL,
  upload_date TEXT NOT NULL,
  content TEXT -- JSON string of rows
);


--
-- Database: `creo_certificate`
--

-- --------------------------------------------------------

CREATE TABLE IF NOT EXISTS tesda_records (
  id INT AUTO_INCREMENT PRIMARY KEY,
  title VARCHAR(255) NOT NULL,
  created_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- Optional sample data
INSERT INTO tesda_records (title) VALUES 
  ('TESDA NC II: Programming'),
  ('TESDA NC II: Animation'),
  ('TESDA NC II: Electrical Installation');

--
-- Table structure for table `credential_tbl`
--

CREATE TABLE `credential_tbl` (
  `credential_id` int(11) NOT NULL,
  `credential_username` varchar(20) DEFAULT NULL,
  `credential_password` varchar(255) NOT NULL,
  `credential_email` varchar(50) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `credential_tbl`
--

INSERT INTO `credential_tbl` (`credential_id`, `credential_username`, `credential_password`, `credential_email`) VALUES
(1, 'creoapp25', 'creotec123', '');

--
-- Indexes for dumped tables
--

--
-- Indexes for table `credential_tbl`
--
ALTER TABLE `credential_tbl`
  ADD PRIMARY KEY (`credential_id`);

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `credential_tbl`
--
ALTER TABLE `credential_tbl`
  MODIFY `credential_id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
