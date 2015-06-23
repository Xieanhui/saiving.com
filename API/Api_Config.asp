<%
'=========================================================
' File: Api_Config.asp
' Version:1.0
' Date: 2007-3-8
' Code by Terry
'=========================================================

'=========================================================
'多系统整合设置
'=========================================================
'API_Enable 是否打开系统整合（默认闭关: False ,打开：True ）
Const MsxmlVersion=".3.0"
Const API_Enable	= False
'API_SysKey 设置系统密钥 (系统整合，必须保证与其它系统设置的密钥一致。)
Const API_SysKey	= "API_TEST"
'API_Urls :整合的其它程序的接口文件路径。多个程序接口之间用半角"|"分隔。
'例如：API_Urls = "http://你的网站地址/博客安装目录/API/API_Response.asp|http://你的网站地址/论坛安装目录/dv_dpo.asp"
Const API_Urls	= "http://192.168.1.181/dvbbs/dv_dpo.asp"
%>





