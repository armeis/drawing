#如果存在数据库DrawingNo，则删除。否则创建数据库
drop database if exists DrawingNo;
create database DrawingNo DEFAULT CHARSET utf8 COLLATE utf8_general_ci;
use DrawingNo;
#如果存在数据表，则删除，否则创建
drop table if exists blueprint;
#创建图纸图号表
create table blueprint(
	drawing_id int(5) not null AUTO_INCREMENT primary key,
	drawing_name varchar(50) not null,
	drawing_code varchar(25),
	drawing_identifier varchar(20),
	drawing_parameter varchar(50),
	drawing_edition varchar(4),
	drawing_classification varchar(20),
	orderDate date,
	document_name varchar(255),
	document_type varchar(4),
	document_path varchar(255),
	drawing_remarks varchar(255)
);
drop table if exists partnumber;
#创建料号表
create table partnumber(
	item_id int(5) not null AUTO_INCREMENT primary key,
	item_No varchar(11) not null,
	item_name varchar(20) not null,
	item_specs varchar(50) not null,
	blueprint_id int(5) not null,
	constraint fk_partnumber foreign key(blueprint_id) references blueprint(drawing_id)
);
#向图纸图号表添加数据
insert into blueprint (drawing_id,drawing_name,drawing_code,drawing_identifier,drawing_parameter,drawing_edition,drawing_classification,orderDate,document_name,document_type,document_path,drawing_remarks)
values
(0,'8000755243','','22','DH/2-32,DW/4+8','A','客供原稿','2020/01/01','8000755243','pdf','F:\图纸图号管理\客户图纸\8000755243.pdf','公差：长宽-3+0，厚-1.5+0'),
(0,'8000755246','','22','DH/2-32,DW/4-2','A','客供原稿','2020/01/01','8000755246','pdf','F:\图纸图号管理\客户图纸\8000755246.pdf','公差：长宽-3+0，厚-1.5+0'),
(0,'P375086B109P8286','LEHY-MRL-II(NL2L1)','375','','','客供原稿','2020/01/01','P375086B109P8286','pdf','F:\图纸图号管理\客户图纸\P375086B109P8286.pdf','上海三菱电梯非标产品图纸'),
(0,'NS375036D008','CMB375-36EI','375','','','客供原稿','2021/10/22','NS375036D008','png','F:\图纸图号管理\NS375036D008\NS375036D008.png','公差：长-6+0，宽包含纤维布边厚度，厚-0+1'),
(0,'NS375036D008-03HH2100JJ1500','','03','HH2100JJ1500','A01','切割','2022/06/16','NS375036D008-03HH2100JJ1500','dxf','F:\图纸图号管理\NS375036D008\NS375036D008-03HH2100JJ1500.dxf','公差：长-6+0，宽包含纤维布边厚度，厚-0+1');
#向料号表添加数据
insert into partnumber (item_id,item_No,item_name,item_specs,blueprint_id)
values
(0,'1107-001387','M5100','1988*142.5*15-开孔',5);