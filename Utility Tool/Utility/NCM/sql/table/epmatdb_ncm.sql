-- MySQL dump 10.13  Distrib 8.0.29, for Win64 (x86_64)
--
-- Host: localhost    Database: epmatdb
-- ------------------------------------------------------
-- Server version	8.0.13

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!50503 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `ncm`
--

DROP TABLE IF EXISTS `ncm`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `ncm` (
  `idqc` int(11) NOT NULL AUTO_INCREMENT,
  `id_report` int(11) NOT NULL,
  `id_sk_process` int(11) NOT NULL,
  `guid` varchar(90) NOT NULL,
  `guid_type` varchar(28) NOT NULL,
  `mod_user_id` int(11) NOT NULL,
  `mod_rmk` longtext,
  `mod_dt` bigint(20) DEFAULT NULL,
  `obser` longtext,
  `chk_user_id` int(11) DEFAULT NULL,
  `chk_rmk` longtext,
  `chk_dt` bigint(20) DEFAULT NULL,
  `bd_isagree` tinyint(4) DEFAULT NULL,
  `bd_ismodelupdated` int(11) DEFAULT NULL,
  `id_cat` int(11) DEFAULT NULL,
  `bd_user_id` int(11) DEFAULT NULL,
  `bd_rmk` longtext,
  `bd_dt` bigint(20) DEFAULT NULL,
  `backcheck` tinyint(4) DEFAULT NULL,
  `backcheck_user_id` int(11) DEFAULT NULL,
  `backcheck_rmk` longtext,
  `backcheck_dt` bigint(20) DEFAULT NULL,
  `id_sev` int(11) DEFAULT NULL,
  `sev_user_id` int(11) DEFAULT NULL,
  `sev_dt` bigint(20) DEFAULT NULL,
  `pm_id` int(11) DEFAULT NULL,
  `pm_rmk` longtext,
  `pm_dt` bigint(20) DEFAULT NULL,
  `qa_rmk` longtext,
  `status` longtext,
  `status_user_id` int(11) DEFAULT NULL,
  `status_dt` bigint(20) DEFAULT NULL,
  `status_sys` varchar(28) DEFAULT NULL,
  `fut_id_1` longtext,
  `fut_id_2` longtext,
  PRIMARY KEY (`idqc`),
  UNIQUE KEY `qc_id_UNIQUE` (`idqc`)
) ENGINE=InnoDB AUTO_INCREMENT=90157 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;
/*!40101 SET character_set_client = @saved_cs_client */;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2022-05-18 18:39:21
