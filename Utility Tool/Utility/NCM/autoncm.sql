CREATE TABLE `qms` (
  `idqms` varchar(50) NOT NULL AUTO_INCREMENT,
  `idprojectmast` varchar(45) NOT NULL,
  `chkby` varchar(45) NOT NULL,
  `guid` varchar(45) NOT NULL,
  `modby` varchar(45) DEFAULT NULL,
  `errordesc` longtext,
  `idseveritymast` varchar(45) DEFAULT NULL,
  `category` varchar(45) DEFAULT NULL,
  `accepted` tinyint DEFAULT NULL,
  `leadby` varchar(45) DEFAULT NULL,
  `modremark` longtext,
  `chkremark` longtext,
  `bkchkby` varchar(45) DEFAULT NULL,
  `leadremark` longtext,
  PRIMARY KEY (`idqms`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;
