<%
'系统配置
dim Sysname,Versions,Official,Author,Contact,IsShow
    Sysname  = "CM实验室(cmlabs)"
    Versions = "4.0.0" '版本
    Official = "http://www.5cmlabs.com"
	Author   = "CM"
	Contact  = "admin@5cmlabs.com"
	IsShow   = true '是否显示在后台导航版权信息
	
'多语言设置
dim urlPath,urlStr
    urlPath = "../../system/models/"
    urlStr  = "中文,cn|英文,en"
	
'数据库路径设置
dim dbpath,sysdbpath,dberr
    dbpath    = "../system/data/KM_09_12_20.mdb"
    sysdbpath = "system/data/cmcms_db_12_08_05.asa"
	dberr     = "database err!"

'分页参数
dim pageadds
    pageAdds = 3 '分页浮动值
	
'× --------------------------------------
'× -----------  后台基本配置  -------------
'× --------------------------------------
dim sys_str_add,sys_str_edit,sys_tip_errtype,sys_tip_none,sys_tip_ok,sys_tip_false
	sys_str_add    = "添加"
	sys_str_edit   = "修改"
	sys_tip_errtype= "分类有误!请重新选择."
	sys_tip_none   = "未找到相应的数据!"
	sys_tip_ok     = "成功!"
	sys_tip_false  = "失败!"
%>