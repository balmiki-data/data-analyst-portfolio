/* ==========================================================
   Quick-Commerce Sales Dashboard  |  SQL Server version (T-SQL)
   Works on SQL Server 2016 or newer (SSMS / Azure Data Studio).
   Same sample data as the Excel dashboard (300 orders, Jan-Jun 2026).
   Sample data is randomly generated for practice, not real company data.

   How to use:
   1) Run PART 1 completely (select all and Execute)
   2) Run PART 2 one query at a time: select ONE query, then Execute.
      Each query = one dashboard tile.
   Note: "GO" is a batch separator used by SSMS. Keep it on its own line.
   ========================================================== */

-- ================= PART 1: SETUP =================
IF DB_ID('sales_dashboard') IS NULL
    CREATE DATABASE sales_dashboard;
GO

USE sales_dashboard;
GO

DROP VIEW  IF EXISTS vw_orders;
DROP TABLE IF EXISTS orders;
GO

CREATE TABLE orders (
  order_id        VARCHAR(10)   NOT NULL PRIMARY KEY,
  order_date      DATE          NOT NULL,
  city            VARCHAR(30)   NOT NULL,
  category        VARCHAR(30)   NOT NULL,
  product         VARCHAR(40)   NOT NULL,
  quantity        INT           NOT NULL,
  unit_price      DECIMAL(10,2) NOT NULL,
  cost_price      DECIMAL(10,2) NOT NULL,
  delivery_mins   INT           NOT NULL
);
GO

INSERT INTO orders
  (order_id, order_date, city, category, product, quantity, unit_price, cost_price, delivery_mins)
VALUES
  ('ORD1001','2026-01-01','Ahmedabad','Dairy','Butter 100g',3,58,50,8),
  ('ORD1002','2026-01-01','Gandhinagar','Snacks','Chocolate',2,80,64,22),
  ('ORD1003','2026-01-02','Ahmedabad','Fruits & Veg','Banana',5,60,48,13),
  ('ORD1004','2026-01-02','Ahmedabad','Fruits & Veg','Apple',5,180,140,12),
  ('ORD1005','2026-01-03','Gandhinagar','Snacks','Potato Chips',5,20,16,20),
  ('ORD1006','2026-01-04','Surat','Beverages','Packaged Juice',4,120,96,17),
  ('ORD1007','2026-01-05','Ahmedabad','Household','Dishwash Liquid',5,135,105,10),
  ('ORD1008','2026-01-06','Surat','Snacks','Namkeen',3,50,40,14),
  ('ORD1009','2026-01-07','Ahmedabad','Beverages','Packaged Juice',3,120,96,10),
  ('ORD1010','2026-01-07','Ahmedabad','Dairy','Milk 1L',1,68,62,16),
  ('ORD1011','2026-01-08','Surat','Household','Detergent 1kg',1,210,170,13),
  ('ORD1012','2026-01-09','Ahmedabad','Dairy','Milk 1L',2,68,62,13),
  ('ORD1013','2026-01-09','Ahmedabad','Fruits & Veg','Onion',5,45,34,19),
  ('ORD1014','2026-01-09','Surat','Beverages','Tea 250g',1,130,105,16),
  ('ORD1015','2026-01-12','Gandhinagar','Fruits & Veg','Banana',1,60,48,18),
  ('ORD1016','2026-01-12','Ahmedabad','Fruits & Veg','Onion',2,45,34,12),
  ('ORD1017','2026-01-13','Vadodara','Personal Care','Face Wash',3,220,165,20),
  ('ORD1018','2026-01-13','Ahmedabad','Dairy','Curd 400g',1,45,37,10),
  ('ORD1019','2026-01-14','Ahmedabad','Personal Care','Toothpaste',2,95,72,13),
  ('ORD1020','2026-01-15','Surat','Fruits & Veg','Banana',2,60,48,13),
  ('ORD1021','2026-01-15','Surat','Personal Care','Soap 4-pack',4,150,118,18),
  ('ORD1022','2026-01-16','Gandhinagar','Personal Care','Shampoo',3,180,140,20),
  ('ORD1023','2026-01-16','Ahmedabad','Dairy','Milk 1L',3,68,62,15),
  ('ORD1024','2026-01-16','Surat','Beverages','Packaged Juice',2,120,96,10),
  ('ORD1025','2026-01-17','Gandhinagar','Household','Dishwash Liquid',4,135,105,23),
  ('ORD1026','2026-01-17','Ahmedabad','Household','Dishwash Liquid',3,135,105,14),
  ('ORD1027','2026-01-17','Ahmedabad','Beverages','Coffee 100g',1,280,225,15),
  ('ORD1028','2026-01-18','Vadodara','Snacks','Namkeen',2,50,40,14),
  ('ORD1029','2026-01-18','Ahmedabad','Beverages','Coffee 100g',3,280,225,13),
  ('ORD1030','2026-01-18','Vadodara','Dairy','Butter 100g',5,58,50,17),
  ('ORD1031','2026-01-19','Ahmedabad','Fruits & Veg','Onion',2,45,34,8),
  ('ORD1032','2026-01-19','Ahmedabad','Snacks','Namkeen',3,50,40,9),
  ('ORD1033','2026-01-19','Surat','Household','Detergent 1kg',4,210,170,18),
  ('ORD1034','2026-01-20','Ahmedabad','Dairy','Paneer 200g',3,95,78,12),
  ('ORD1035','2026-01-21','Surat','Household','Floor Cleaner',5,160,125,14),
  ('ORD1036','2026-01-21','Vadodara','Dairy','Curd 400g',4,45,37,21),
  ('ORD1037','2026-01-21','Ahmedabad','Beverages','Packaged Juice',2,120,96,8),
  ('ORD1038','2026-01-22','Vadodara','Household','Floor Cleaner',4,160,125,19),
  ('ORD1039','2026-01-23','Surat','Fruits & Veg','Onion',2,45,34,13),
  ('ORD1040','2026-01-24','Vadodara','Snacks','Chocolate',2,80,64,17),
  ('ORD1041','2026-01-24','Surat','Personal Care','Soap 4-pack',1,150,118,16),
  ('ORD1042','2026-01-24','Surat','Fruits & Veg','Tomato',3,40,30,21),
  ('ORD1043','2026-01-24','Ahmedabad','Fruits & Veg','Banana',1,60,48,14),
  ('ORD1044','2026-01-25','Vadodara','Fruits & Veg','Apple',5,180,140,19),
  ('ORD1045','2026-01-25','Ahmedabad','Beverages','Coffee 100g',4,280,225,14),
  ('ORD1046','2026-01-25','Ahmedabad','Beverages','Cold Drink 2L',4,95,80,18),
  ('ORD1047','2026-01-26','Ahmedabad','Personal Care','Face Wash',1,220,165,8),
  ('ORD1048','2026-01-26','Ahmedabad','Fruits & Veg','Apple',5,180,140,11),
  ('ORD1049','2026-01-26','Gandhinagar','Snacks','Chocolate',5,80,64,21),
  ('ORD1050','2026-01-27','Surat','Household','Tissue Roll',2,90,68,11),
  ('ORD1051','2026-01-27','Vadodara','Snacks','Namkeen',2,50,40,17),
  ('ORD1052','2026-01-28','Vadodara','Snacks','Biscuits',3,30,24,13),
  ('ORD1053','2026-01-28','Surat','Beverages','Packaged Juice',3,120,96,13),
  ('ORD1054','2026-01-29','Ahmedabad','Fruits & Veg','Apple',2,180,140,19),
  ('ORD1055','2026-01-29','Vadodara','Fruits & Veg','Apple',3,180,140,13),
  ('ORD1056','2026-01-29','Ahmedabad','Dairy','Milk 1L',5,68,62,13),
  ('ORD1057','2026-01-29','Ahmedabad','Fruits & Veg','Onion',3,45,34,14),
  ('ORD1058','2026-01-30','Gandhinagar','Dairy','Paneer 200g',4,95,78,13),
  ('ORD1059','2026-01-31','Ahmedabad','Dairy','Milk 1L',4,68,62,11),
  ('ORD1060','2026-02-01','Surat','Fruits & Veg','Apple',2,180,140,20),
  ('ORD1061','2026-02-02','Gandhinagar','Snacks','Chocolate',5,80,64,18),
  ('ORD1062','2026-02-02','Surat','Fruits & Veg','Apple',2,180,140,16),
  ('ORD1063','2026-02-03','Ahmedabad','Dairy','Paneer 200g',4,95,78,11),
  ('ORD1064','2026-02-05','Ahmedabad','Beverages','Coffee 100g',3,280,225,18),
  ('ORD1065','2026-02-05','Vadodara','Snacks','Potato Chips',3,20,16,16),
  ('ORD1066','2026-02-05','Gandhinagar','Fruits & Veg','Banana',2,60,48,20),
  ('ORD1067','2026-02-05','Ahmedabad','Fruits & Veg','Tomato',2,40,30,17),
  ('ORD1068','2026-02-06','Surat','Fruits & Veg','Onion',3,45,34,20),
  ('ORD1069','2026-02-09','Ahmedabad','Snacks','Potato Chips',2,20,16,15),
  ('ORD1070','2026-02-09','Ahmedabad','Dairy','Butter 100g',4,58,50,13),
  ('ORD1071','2026-02-09','Gandhinagar','Fruits & Veg','Tomato',1,40,30,18),
  ('ORD1072','2026-02-10','Vadodara','Personal Care','Shampoo',3,180,140,13),
  ('ORD1073','2026-02-10','Surat','Beverages','Cold Drink 2L',5,95,80,16),
  ('ORD1074','2026-02-10','Surat','Snacks','Chocolate',3,80,64,16),
  ('ORD1075','2026-02-11','Surat','Personal Care','Soap 4-pack',4,150,118,18),
  ('ORD1076','2026-02-11','Vadodara','Snacks','Chocolate',4,80,64,19),
  ('ORD1077','2026-02-11','Surat','Beverages','Tea 250g',3,130,105,20),
  ('ORD1078','2026-02-12','Ahmedabad','Household','Dishwash Liquid',4,135,105,9),
  ('ORD1079','2026-02-12','Surat','Beverages','Tea 250g',1,130,105,14),
  ('ORD1080','2026-02-12','Surat','Dairy','Curd 400g',3,45,37,17),
  ('ORD1081','2026-02-13','Ahmedabad','Snacks','Namkeen',5,50,40,21),
  ('ORD1082','2026-02-15','Ahmedabad','Beverages','Coffee 100g',5,280,225,14),
  ('ORD1083','2026-02-16','Vadodara','Beverages','Coffee 100g',4,280,225,16),
  ('ORD1084','2026-02-18','Ahmedabad','Beverages','Tea 250g',5,130,105,12),
  ('ORD1085','2026-02-18','Ahmedabad','Snacks','Namkeen',4,50,40,17),
  ('ORD1086','2026-02-19','Ahmedabad','Snacks','Namkeen',3,50,40,13),
  ('ORD1087','2026-02-19','Ahmedabad','Beverages','Tea 250g',1,130,105,13),
  ('ORD1088','2026-02-20','Vadodara','Fruits & Veg','Apple',3,180,140,8),
  ('ORD1089','2026-02-20','Vadodara','Household','Floor Cleaner',1,160,125,17),
  ('ORD1090','2026-02-21','Ahmedabad','Dairy','Paneer 200g',3,95,78,11),
  ('ORD1091','2026-02-23','Ahmedabad','Fruits & Veg','Onion',2,45,34,10),
  ('ORD1092','2026-02-24','Ahmedabad','Fruits & Veg','Onion',4,45,34,10),
  ('ORD1093','2026-02-24','Surat','Dairy','Milk 1L',3,68,62,20),
  ('ORD1094','2026-02-24','Vadodara','Dairy','Milk 1L',5,68,62,14),
  ('ORD1095','2026-02-24','Surat','Fruits & Veg','Tomato',5,40,30,18),
  ('ORD1096','2026-02-24','Ahmedabad','Snacks','Chocolate',5,80,64,13),
  ('ORD1097','2026-02-25','Surat','Snacks','Chocolate',4,80,64,19),
  ('ORD1098','2026-02-25','Surat','Dairy','Milk 1L',5,68,62,15),
  ('ORD1099','2026-02-26','Gandhinagar','Fruits & Veg','Tomato',2,40,30,18),
  ('ORD1100','2026-02-26','Ahmedabad','Snacks','Biscuits',1,30,24,15),
  ('ORD1101','2026-02-26','Surat','Beverages','Coffee 100g',3,280,225,19),
  ('ORD1102','2026-02-26','Ahmedabad','Personal Care','Shampoo',2,180,140,14),
  ('ORD1103','2026-02-27','Gandhinagar','Personal Care','Soap 4-pack',4,150,118,20),
  ('ORD1104','2026-02-27','Ahmedabad','Dairy','Paneer 200g',1,95,78,15),
  ('ORD1105','2026-02-28','Ahmedabad','Personal Care','Toothpaste',4,95,72,17),
  ('ORD1106','2026-02-28','Vadodara','Household','Tissue Roll',4,90,68,19),
  ('ORD1107','2026-02-28','Ahmedabad','Fruits & Veg','Apple',3,180,140,13),
  ('ORD1108','2026-02-28','Ahmedabad','Beverages','Cold Drink 2L',3,95,80,13),
  ('ORD1109','2026-03-01','Surat','Personal Care','Soap 4-pack',4,150,118,13),
  ('ORD1110','2026-03-01','Ahmedabad','Household','Tissue Roll',4,90,68,13),
  ('ORD1111','2026-03-02','Gandhinagar','Snacks','Chocolate',3,80,64,18),
  ('ORD1112','2026-03-02','Ahmedabad','Dairy','Butter 100g',3,58,50,15),
  ('ORD1113','2026-03-03','Vadodara','Dairy','Milk 1L',4,68,62,17),
  ('ORD1114','2026-03-03','Ahmedabad','Personal Care','Shampoo',5,180,140,15),
  ('ORD1115','2026-03-04','Surat','Fruits & Veg','Onion',5,45,34,17),
  ('ORD1116','2026-03-04','Vadodara','Fruits & Veg','Tomato',3,40,30,11),
  ('ORD1117','2026-03-04','Surat','Beverages','Cold Drink 2L',1,95,80,21),
  ('ORD1118','2026-03-05','Ahmedabad','Fruits & Veg','Banana',5,60,48,13),
  ('ORD1119','2026-03-05','Surat','Dairy','Milk 1L',4,68,62,21),
  ('ORD1120','2026-03-05','Ahmedabad','Household','Floor Cleaner',5,160,125,12),
  ('ORD1121','2026-03-05','Vadodara','Dairy','Butter 100g',4,58,50,16),
  ('ORD1122','2026-03-06','Ahmedabad','Household','Dishwash Liquid',1,135,105,17),
  ('ORD1123','2026-03-09','Ahmedabad','Personal Care','Face Wash',3,220,165,17),
  ('ORD1124','2026-03-09','Vadodara','Dairy','Curd 400g',5,45,37,14),
  ('ORD1125','2026-03-09','Surat','Beverages','Packaged Juice',4,120,96,14),
  ('ORD1126','2026-03-09','Surat','Snacks','Chocolate',3,80,64,16),
  ('ORD1127','2026-03-09','Gandhinagar','Household','Dishwash Liquid',5,135,105,14),
  ('ORD1128','2026-03-10','Surat','Dairy','Butter 100g',4,58,50,20),
  ('ORD1129','2026-03-10','Ahmedabad','Beverages','Packaged Juice',4,120,96,15),
  ('ORD1130','2026-03-10','Gandhinagar','Dairy','Milk 1L',4,68,62,18),
  ('ORD1131','2026-03-11','Ahmedabad','Household','Detergent 1kg',2,210,170,10),
  ('ORD1132','2026-03-12','Ahmedabad','Snacks','Namkeen',3,50,40,16),
  ('ORD1133','2026-03-13','Surat','Dairy','Milk 1L',3,68,62,20),
  ('ORD1134','2026-03-13','Ahmedabad','Dairy','Butter 100g',5,58,50,9),
  ('ORD1135','2026-03-13','Ahmedabad','Fruits & Veg','Onion',2,45,34,13),
  ('ORD1136','2026-03-13','Ahmedabad','Beverages','Cold Drink 2L',4,95,80,10),
  ('ORD1137','2026-03-13','Ahmedabad','Dairy','Milk 1L',1,68,62,13),
  ('ORD1138','2026-03-16','Ahmedabad','Dairy','Paneer 200g',5,95,78,14),
  ('ORD1139','2026-03-17','Surat','Beverages','Packaged Juice',2,120,96,18),
  ('ORD1140','2026-03-17','Ahmedabad','Beverages','Coffee 100g',5,280,225,17),
  ('ORD1141','2026-03-18','Ahmedabad','Snacks','Namkeen',4,50,40,13),
  ('ORD1142','2026-03-20','Vadodara','Household','Dishwash Liquid',1,135,105,16),
  ('ORD1143','2026-03-21','Surat','Dairy','Curd 400g',4,45,37,15),
  ('ORD1144','2026-03-22','Ahmedabad','Household','Tissue Roll',1,90,68,10),
  ('ORD1145','2026-03-22','Ahmedabad','Snacks','Namkeen',3,50,40,19),
  ('ORD1146','2026-03-24','Ahmedabad','Fruits & Veg','Banana',5,60,48,14),
  ('ORD1147','2026-03-25','Vadodara','Personal Care','Toothpaste',2,95,72,13),
  ('ORD1148','2026-03-26','Surat','Beverages','Packaged Juice',1,120,96,15),
  ('ORD1149','2026-03-28','Gandhinagar','Beverages','Tea 250g',4,130,105,21),
  ('ORD1150','2026-03-28','Ahmedabad','Household','Floor Cleaner',4,160,125,14),
  ('ORD1151','2026-03-28','Ahmedabad','Fruits & Veg','Onion',5,45,34,11),
  ('ORD1152','2026-03-29','Ahmedabad','Household','Tissue Roll',3,90,68,13),
  ('ORD1153','2026-03-29','Vadodara','Dairy','Milk 1L',4,68,62,17),
  ('ORD1154','2026-03-30','Gandhinagar','Household','Tissue Roll',5,90,68,18),
  ('ORD1155','2026-04-01','Ahmedabad','Personal Care','Shampoo',3,180,140,10),
  ('ORD1156','2026-04-01','Vadodara','Beverages','Coffee 100g',3,280,225,15),
  ('ORD1157','2026-04-02','Ahmedabad','Dairy','Milk 1L',5,68,62,16),
  ('ORD1158','2026-04-03','Gandhinagar','Dairy','Milk 1L',4,68,62,16),
  ('ORD1159','2026-04-03','Gandhinagar','Dairy','Butter 100g',1,58,50,23),
  ('ORD1160','2026-04-03','Ahmedabad','Snacks','Potato Chips',2,20,16,19),
  ('ORD1161','2026-04-04','Gandhinagar','Fruits & Veg','Apple',3,180,140,16),
  ('ORD1162','2026-04-05','Gandhinagar','Dairy','Paneer 200g',4,95,78,14),
  ('ORD1163','2026-04-06','Vadodara','Personal Care','Face Wash',3,220,165,18),
  ('ORD1164','2026-04-06','Surat','Household','Floor Cleaner',1,160,125,17),
  ('ORD1165','2026-04-07','Gandhinagar','Dairy','Curd 400g',1,45,37,17),
  ('ORD1166','2026-04-08','Ahmedabad','Personal Care','Face Wash',1,220,165,15),
  ('ORD1167','2026-04-08','Vadodara','Beverages','Cold Drink 2L',1,95,80,16),
  ('ORD1168','2026-04-08','Ahmedabad','Snacks','Potato Chips',5,20,16,11),
  ('ORD1169','2026-04-08','Ahmedabad','Fruits & Veg','Tomato',3,40,30,14),
  ('ORD1170','2026-04-08','Surat','Dairy','Paneer 200g',2,95,78,11),
  ('ORD1171','2026-04-09','Surat','Personal Care','Shampoo',1,180,140,18),
  ('ORD1172','2026-04-12','Vadodara','Snacks','Potato Chips',2,20,16,18),
  ('ORD1173','2026-04-13','Ahmedabad','Fruits & Veg','Apple',5,180,140,14),
  ('ORD1174','2026-04-13','Surat','Dairy','Milk 1L',4,68,62,17),
  ('ORD1175','2026-04-13','Surat','Household','Detergent 1kg',5,210,170,15),
  ('ORD1176','2026-04-13','Ahmedabad','Household','Tissue Roll',4,90,68,18),
  ('ORD1177','2026-04-14','Vadodara','Snacks','Chocolate',1,80,64,18),
  ('ORD1178','2026-04-15','Ahmedabad','Fruits & Veg','Onion',4,45,34,12),
  ('ORD1179','2026-04-15','Gandhinagar','Snacks','Chocolate',5,80,64,17),
  ('ORD1180','2026-04-16','Vadodara','Personal Care','Soap 4-pack',4,150,118,15),
  ('ORD1181','2026-04-18','Surat','Dairy','Curd 400g',4,45,37,13),
  ('ORD1182','2026-04-19','Ahmedabad','Dairy','Curd 400g',3,45,37,16),
  ('ORD1183','2026-04-19','Gandhinagar','Beverages','Coffee 100g',1,280,225,20),
  ('ORD1184','2026-04-19','Surat','Beverages','Tea 250g',2,130,105,17),
  ('ORD1185','2026-04-19','Surat','Snacks','Namkeen',1,50,40,11),
  ('ORD1186','2026-04-19','Gandhinagar','Dairy','Milk 1L',3,68,62,22),
  ('ORD1187','2026-04-19','Surat','Dairy','Paneer 200g',2,95,78,15),
  ('ORD1188','2026-04-20','Ahmedabad','Personal Care','Toothpaste',1,95,72,12),
  ('ORD1189','2026-04-21','Vadodara','Snacks','Chocolate',5,80,64,21),
  ('ORD1190','2026-04-22','Surat','Beverages','Cold Drink 2L',5,95,80,16),
  ('ORD1191','2026-04-23','Gandhinagar','Personal Care','Toothpaste',1,95,72,20),
  ('ORD1192','2026-04-24','Ahmedabad','Dairy','Milk 1L',2,68,62,18),
  ('ORD1193','2026-04-25','Gandhinagar','Snacks','Chocolate',1,80,64,21),
  ('ORD1194','2026-04-25','Ahmedabad','Beverages','Packaged Juice',4,120,96,17),
  ('ORD1195','2026-04-26','Surat','Dairy','Curd 400g',3,45,37,16),
  ('ORD1196','2026-04-27','Surat','Fruits & Veg','Banana',2,60,48,11),
  ('ORD1197','2026-04-27','Ahmedabad','Snacks','Potato Chips',3,20,16,11),
  ('ORD1198','2026-04-28','Surat','Dairy','Paneer 200g',5,95,78,18),
  ('ORD1199','2026-04-28','Ahmedabad','Dairy','Curd 400g',5,45,37,9),
  ('ORD1200','2026-04-29','Ahmedabad','Dairy','Milk 1L',1,68,62,14),
  ('ORD1201','2026-04-29','Ahmedabad','Household','Dishwash Liquid',3,135,105,13),
  ('ORD1202','2026-04-30','Surat','Dairy','Paneer 200g',4,95,78,18),
  ('ORD1203','2026-04-30','Surat','Dairy','Curd 400g',1,45,37,13),
  ('ORD1204','2026-05-02','Surat','Snacks','Potato Chips',3,20,16,19),
  ('ORD1205','2026-05-02','Surat','Snacks','Chocolate',3,80,64,17),
  ('ORD1206','2026-05-02','Ahmedabad','Personal Care','Face Wash',3,220,165,14),
  ('ORD1207','2026-05-04','Ahmedabad','Personal Care','Soap 4-pack',4,150,118,10),
  ('ORD1208','2026-05-05','Ahmedabad','Dairy','Curd 400g',2,45,37,12),
  ('ORD1209','2026-05-05','Ahmedabad','Fruits & Veg','Banana',5,60,48,8),
  ('ORD1210','2026-05-05','Ahmedabad','Beverages','Packaged Juice',2,120,96,13),
  ('ORD1211','2026-05-06','Ahmedabad','Dairy','Paneer 200g',5,95,78,16),
  ('ORD1212','2026-05-07','Ahmedabad','Household','Tissue Roll',4,90,68,18),
  ('ORD1213','2026-05-08','Vadodara','Beverages','Coffee 100g',5,280,225,10),
  ('ORD1214','2026-05-09','Ahmedabad','Household','Floor Cleaner',2,160,125,8),
  ('ORD1215','2026-05-10','Surat','Snacks','Namkeen',3,50,40,14),
  ('ORD1216','2026-05-10','Ahmedabad','Snacks','Potato Chips',1,20,16,12),
  ('ORD1217','2026-05-10','Ahmedabad','Dairy','Milk 1L',5,68,62,11),
  ('ORD1218','2026-05-11','Surat','Household','Detergent 1kg',4,210,170,15),
  ('ORD1219','2026-05-12','Gandhinagar','Snacks','Biscuits',3,30,24,16),
  ('ORD1220','2026-05-13','Gandhinagar','Snacks','Namkeen',1,50,40,16),
  ('ORD1221','2026-05-16','Surat','Fruits & Veg','Tomato',4,40,30,20),
  ('ORD1222','2026-05-16','Surat','Dairy','Curd 400g',4,45,37,16),
  ('ORD1223','2026-05-16','Surat','Dairy','Milk 1L',3,68,62,18),
  ('ORD1224','2026-05-17','Vadodara','Dairy','Paneer 200g',3,95,78,16),
  ('ORD1225','2026-05-17','Ahmedabad','Beverages','Tea 250g',3,130,105,18),
  ('ORD1226','2026-05-18','Ahmedabad','Snacks','Potato Chips',3,20,16,13),
  ('ORD1227','2026-05-18','Surat','Beverages','Tea 250g',2,130,105,14),
  ('ORD1228','2026-05-18','Vadodara','Beverages','Cold Drink 2L',4,95,80,11),
  ('ORD1229','2026-05-18','Surat','Fruits & Veg','Onion',2,45,34,19),
  ('ORD1230','2026-05-19','Vadodara','Fruits & Veg','Tomato',2,40,30,15),
  ('ORD1231','2026-05-19','Ahmedabad','Personal Care','Shampoo',5,180,140,14),
  ('ORD1232','2026-05-19','Gandhinagar','Fruits & Veg','Onion',5,45,34,13),
  ('ORD1233','2026-05-19','Ahmedabad','Dairy','Paneer 200g',2,95,78,16),
  ('ORD1234','2026-05-20','Vadodara','Fruits & Veg','Apple',1,180,140,11),
  ('ORD1235','2026-05-20','Surat','Beverages','Packaged Juice',4,120,96,13),
  ('ORD1236','2026-05-21','Vadodara','Dairy','Paneer 200g',3,95,78,12),
  ('ORD1237','2026-05-21','Gandhinagar','Personal Care','Soap 4-pack',3,150,118,17),
  ('ORD1238','2026-05-22','Surat','Personal Care','Face Wash',5,220,165,13),
  ('ORD1239','2026-05-22','Surat','Household','Dishwash Liquid',5,135,105,18),
  ('ORD1240','2026-05-22','Surat','Household','Detergent 1kg',5,210,170,18),
  ('ORD1241','2026-05-23','Gandhinagar','Dairy','Paneer 200g',3,95,78,15),
  ('ORD1242','2026-05-24','Vadodara','Fruits & Veg','Tomato',4,40,30,17),
  ('ORD1243','2026-05-24','Gandhinagar','Household','Tissue Roll',1,90,68,16),
  ('ORD1244','2026-05-26','Ahmedabad','Fruits & Veg','Tomato',1,40,30,15),
  ('ORD1245','2026-05-26','Ahmedabad','Fruits & Veg','Onion',2,45,34,15),
  ('ORD1246','2026-05-27','Vadodara','Beverages','Packaged Juice',4,120,96,17),
  ('ORD1247','2026-05-28','Surat','Fruits & Veg','Banana',1,60,48,16),
  ('ORD1248','2026-05-28','Surat','Beverages','Coffee 100g',1,280,225,22),
  ('ORD1249','2026-05-30','Vadodara','Dairy','Milk 1L',1,68,62,16),
  ('ORD1250','2026-05-30','Surat','Fruits & Veg','Apple',5,180,140,21),
  ('ORD1251','2026-05-31','Surat','Beverages','Packaged Juice',2,120,96,16),
  ('ORD1252','2026-05-31','Ahmedabad','Snacks','Biscuits',1,30,24,14),
  ('ORD1253','2026-05-31','Surat','Dairy','Milk 1L',4,68,62,17),
  ('ORD1254','2026-06-01','Ahmedabad','Dairy','Milk 1L',3,68,62,13),
  ('ORD1255','2026-06-02','Vadodara','Dairy','Milk 1L',5,68,62,15),
  ('ORD1256','2026-06-02','Ahmedabad','Snacks','Biscuits',4,30,24,11),
  ('ORD1257','2026-06-03','Gandhinagar','Snacks','Biscuits',2,30,24,23),
  ('ORD1258','2026-06-04','Surat','Dairy','Milk 1L',2,68,62,19),
  ('ORD1259','2026-06-04','Gandhinagar','Dairy','Curd 400g',5,45,37,14),
  ('ORD1260','2026-06-05','Surat','Fruits & Veg','Tomato',3,40,30,20),
  ('ORD1261','2026-06-05','Ahmedabad','Personal Care','Toothpaste',3,95,72,16),
  ('ORD1262','2026-06-05','Gandhinagar','Beverages','Coffee 100g',3,280,225,21),
  ('ORD1263','2026-06-08','Gandhinagar','Snacks','Chocolate',3,80,64,22),
  ('ORD1264','2026-06-10','Vadodara','Beverages','Cold Drink 2L',4,95,80,14),
  ('ORD1265','2026-06-10','Surat','Personal Care','Toothpaste',3,95,72,19),
  ('ORD1266','2026-06-10','Ahmedabad','Snacks','Chocolate',2,80,64,15),
  ('ORD1267','2026-06-11','Vadodara','Dairy','Milk 1L',3,68,62,13),
  ('ORD1268','2026-06-12','Surat','Personal Care','Shampoo',1,180,140,14),
  ('ORD1269','2026-06-12','Vadodara','Dairy','Paneer 200g',1,95,78,18),
  ('ORD1270','2026-06-13','Vadodara','Fruits & Veg','Banana',2,60,48,19),
  ('ORD1271','2026-06-13','Vadodara','Dairy','Butter 100g',5,58,50,17),
  ('ORD1272','2026-06-13','Surat','Dairy','Butter 100g',1,58,50,17),
  ('ORD1273','2026-06-14','Surat','Snacks','Chocolate',3,80,64,15),
  ('ORD1274','2026-06-14','Ahmedabad','Beverages','Coffee 100g',5,280,225,12),
  ('ORD1275','2026-06-15','Ahmedabad','Snacks','Namkeen',3,50,40,13),
  ('ORD1276','2026-06-15','Surat','Beverages','Coffee 100g',2,280,225,15),
  ('ORD1277','2026-06-16','Ahmedabad','Personal Care','Face Wash',2,220,165,16),
  ('ORD1278','2026-06-16','Ahmedabad','Beverages','Tea 250g',3,130,105,12),
  ('ORD1279','2026-06-16','Vadodara','Snacks','Namkeen',3,50,40,18),
  ('ORD1280','2026-06-17','Gandhinagar','Beverages','Tea 250g',5,130,105,15),
  ('ORD1281','2026-06-17','Ahmedabad','Household','Dishwash Liquid',2,135,105,14),
  ('ORD1282','2026-06-18','Surat','Dairy','Paneer 200g',5,95,78,16),
  ('ORD1283','2026-06-18','Gandhinagar','Fruits & Veg','Apple',4,180,140,13),
  ('ORD1284','2026-06-19','Vadodara','Fruits & Veg','Onion',1,45,34,15),
  ('ORD1285','2026-06-21','Surat','Beverages','Tea 250g',2,130,105,20),
  ('ORD1286','2026-06-21','Surat','Fruits & Veg','Tomato',4,40,30,13),
  ('ORD1287','2026-06-21','Surat','Household','Detergent 1kg',5,210,170,18),
  ('ORD1288','2026-06-22','Ahmedabad','Beverages','Tea 250g',5,130,105,15),
  ('ORD1289','2026-06-23','Gandhinagar','Personal Care','Toothpaste',2,95,72,16),
  ('ORD1290','2026-06-24','Ahmedabad','Dairy','Milk 1L',5,68,62,11),
  ('ORD1291','2026-06-24','Surat','Snacks','Biscuits',2,30,24,14),
  ('ORD1292','2026-06-24','Surat','Personal Care','Shampoo',1,180,140,12),
  ('ORD1293','2026-06-24','Ahmedabad','Household','Floor Cleaner',2,160,125,14),
  ('ORD1294','2026-06-25','Surat','Snacks','Namkeen',3,50,40,13),
  ('ORD1295','2026-06-26','Surat','Personal Care','Face Wash',3,220,165,12),
  ('ORD1296','2026-06-26','Ahmedabad','Dairy','Butter 100g',2,58,50,15),
  ('ORD1297','2026-06-28','Ahmedabad','Fruits & Veg','Onion',3,45,34,15),
  ('ORD1298','2026-06-29','Ahmedabad','Snacks','Namkeen',2,50,40,13),
  ('ORD1299','2026-06-29','Surat','Dairy','Paneer 200g',2,95,78,19),
  ('ORD1300','2026-06-30','Ahmedabad','Snacks','Biscuits',2,30,24,15);
GO

-- Helper view: Revenue = qty*price, Profit = qty*(price-cost), order_month = 'YYYY-MM'
-- (CREATE VIEW must be the first statement in its batch, hence the GO lines)
CREATE VIEW vw_orders AS
SELECT *,
       quantity * unit_price                          AS revenue,
       quantity * (unit_price - cost_price)           AS profit,
       LEFT(CONVERT(VARCHAR(10), order_date, 23), 7)  AS order_month
FROM orders;
GO

-- ================= PART 2: DASHBOARD QUERIES =================

-- TILE 1: KPI cards (Revenue, Profit, Margin, Orders, AOV, Avg Delivery)
SELECT
  SUM(revenue)                                AS total_revenue,
  SUM(profit)                                 AS total_profit,
  ROUND(SUM(profit) * 100.0 / SUM(revenue), 1) AS profit_margin_pct,
  COUNT(*)                                    AS total_orders,
  ROUND(SUM(revenue) / COUNT(*), 0)           AS avg_order_value,
  ROUND(AVG(delivery_mins * 1.0), 1)          AS avg_delivery_mins
FROM vw_orders;

-- TILE 2: Revenue and profit by category
SELECT category,
       SUM(revenue) AS revenue,
       SUM(profit)  AS profit,
       ROUND(SUM(revenue) * 100.0 / SUM(SUM(revenue)) OVER (), 1) AS revenue_share_pct
FROM vw_orders
GROUP BY category
ORDER BY revenue DESC;

-- TILE 3: City performance (revenue, orders, delivery time, rank)
SELECT city,
       SUM(revenue)                         AS revenue,
       COUNT(*)                             AS orders,
       ROUND(AVG(delivery_mins * 1.0), 1)   AS avg_delivery_mins,
       RANK() OVER (ORDER BY SUM(revenue) DESC) AS revenue_rank
FROM vw_orders
GROUP BY city
ORDER BY revenue_rank;

-- TILE 4: Monthly trend with month-over-month growth %
SELECT order_month,
       SUM(revenue) AS revenue,
       COUNT(*)     AS orders,
       ROUND((SUM(revenue) - LAG(SUM(revenue)) OVER (ORDER BY order_month))
             * 100.0 / LAG(SUM(revenue)) OVER (ORDER BY order_month), 1) AS mom_growth_pct
FROM vw_orders
GROUP BY order_month
ORDER BY order_month;

-- TILE 5: Top 5 products by revenue
SELECT TOP 5 product, category,
       SUM(quantity) AS units_sold,
       SUM(revenue)  AS revenue,
       SUM(profit)   AS profit
FROM vw_orders
GROUP BY product, category
ORDER BY revenue DESC;

-- TILE 6: Best category in each city (CTE + window function)
WITH city_cat AS (
  SELECT city, category, SUM(revenue) AS revenue,
         ROW_NUMBER() OVER (PARTITION BY city ORDER BY SUM(revenue) DESC) AS rn
  FROM vw_orders
  GROUP BY city, category
)
SELECT city, category AS top_category, revenue
FROM city_cat
WHERE rn = 1
ORDER BY revenue DESC;

-- TILE 7: Slow deliveries (orders taking more than 20 minutes) by city
SELECT city,
       COUNT(*) AS total_orders,
       SUM(CASE WHEN delivery_mins > 20 THEN 1 ELSE 0 END) AS slow_orders,
       ROUND(SUM(CASE WHEN delivery_mins > 20 THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) AS slow_order_pct
FROM vw_orders
GROUP BY city
ORDER BY slow_order_pct DESC;
