<%Option Explicit%>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<%
Response.Buffer = True
Response.CacheControl = "no-cache"
Dim Conn,str_Url,Go_Url,mf_sys,MF_Site_Name,MF_Site_lock
MF_Default_Conn
MF_Session_TF
str_Url = NoSqlHack(Request.QueryString("URLs"))
if trim(str_Url)="" or  isnull(str_Url) then:Go_Url="sysinfo.asp":else:Go_Url=Replace(str_Url,"||","&"):end if
if instr(1,str_Url,""& G_ADMIN_DIR &"/index.asp",1)>0 then Go_Url = "sysinfo.asp"
set mf_sys = Conn.execute("select top 1 MF_Site_Name,MF_Site_lock from FS_MF_Config")
MF_Site_Name = mf_sys(0)
MF_Site_lock = mf_sys(1)
mf_sys.close:set mf_sys=nothing
%>
<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><% = MF_Site_Name %>--管理后台</title>
<link rel="icon" href="../favicon.ico" type="image/x-icon" />
<meta name="keywords" content="网站内容管理系统">
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<SCRIPT language="javascript">
<%
Dim Temp_Admin_Is_Super,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
if Temp_Admin_Is_Super =1 then%>
var Str_Status="当前用户：<%=Temp_Admin_Name%> ,系统管理员/超级管理员";
window.status=Str_Status;
<%else%>
var Str_Status="当前用户：<%=Temp_Admin_Name%> ,一般管理员";
window.status=Str_Status;
<%end if%>
</SCRIPT>
</HEAD>
<FRAMESET id="Frame" rows="51,*" cols="*" border="0">
  <FRAME id="TopFrame" src="TopFrame.asp?SessionID=<%= Session.SessionID %>" name="topFrame" scrolling="NO" noresize >
  <FRAMESET id="MainFrame" cols="170,*" frameborder="NO" border="0" framespacing="0"  scrolling="yes"  noresize>
		<FRAME id="MenuFrame" src="shortCutMenu.asp" name="MenuFrame" scrolling="yes" frameborder="0">
		<FRAME id="ContentFrame" src="<% = Go_Url %>" name="ContentFrame" scrolling="yes" frameborder="0" marginheight="0" marginwidth="0" >
  </FRAMESET>
</FRAMESET>
<NOFRAMES>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
  <p>你的浏览器版本过低！！！本系统要求IE5及以上版本才能使用本系统。</p>
  </body>
</NOFRAMES>
</HTML>
<%
set Conn = nothing
%>
<script language="JavaScript" type="text/javascript" src="http://PassPort.foosun.net/passport?User=<%=MF_Site_lock%>&URL=<%=Request.ServerVariables("SERVER_NAME")%>&Email=<%=request.Cookies("FoosunMFCookies")("FoosunMFEmail")%>"></script>






