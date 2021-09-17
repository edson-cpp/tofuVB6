GRANT ALL PRIVILEGES ON `tofu`.* TO 'eds'@'localhost';
GRANT ALL PRIVILEGES ON `tofu`.* TO 'eds'@'%';
UPDATE mysql.user SET password='359a15221ff9a9b7' WHERE User = 'eds';
CREATE DATABASE IF NOT EXISTS `tofu`;
USE tofu;
