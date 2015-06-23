<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
'session判断
MF_Session_TF 
if not MF_Check_Pop_TF("NS_Templet") then Err_Show
if not MF_Check_Pop_TF("NS036") then Err_Show
dim Fs_news,strShowErr
set Fs_news = new Cls_News
Fs_News.GetSysParam()
if Request.Form("Action") = "Templet_News" then
	Dim str_s_classIDarray,tmp_splitarrey,tmp_i,str_Templet,str_NewsTemplet
	str_s_classIDarray =Replace(Request.Form("s_Classid")," ","")
	str_Templet = Trim(Replace(Request.Form("Templet"),"//","/"))
	str_NewsTemplet = Trim(Replace(Request.Form("NewsTemplet"),"//","/"))
	if Trim(str_s_classIDarray)="" then
		strShowErr = "<li>请选择栏目</li><li>至少要选择一个要捆绑的栏目!</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End if
	tmp_splitarrey = split(str_s_classIDarray,",")
	for tmp_i = LBound(tmp_splitarrey) to UBound(tmp_splitarrey)
			Dim Tmp_rs
			Set Tmp_rs=server.CreateObject(G_FS_RS)
		    Tmp_rs.open "select isUrl,Templet,NewsTemplet,Classid From [FS_NS_NewsClass] where ReycleTF=0 and ClassID='"&NoSqlHack(tmp_splitarrey(tmp_i))&"' order by id desc",Conn,1,3
			Do while Not Tmp_rs.eof 
				if Tmp_rs("isUrl")=1 then
					Tmp_rs.movenext
				Else
					Conn.execute("Update FS_NS_NewsClass set Templet='"& NoSqlHack(str_Templet) &"',NewsTemplet='"& NoSqlHack(str_NewsTemplet) &"' where ClassID='"& NoSqlHack(Tmp_rs("ClassID")) &"'")
					Tmp_rs.movenext
				End if
			Loop
	Next
	Tmp_rs.close:set Tmp_rs=nothing
		Call MF_Insert_oper_Log("模板捆绑","批量进行了模板的捆绑",now,session("admin_name"),"NS")
		strShowErr = "<li>捆绑成功</li><li>需要重新生成才生效!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>标签管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu"><a href="#" class="sd"><strong>模板管理</strong></a><a href="../../help?Lable=NS_Class_Templet" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
  </tr>
  <tr> 
    <td height="18" class="hback"><div align="left"><a href="Class_ToTemplet.asp">首页</a> 
        &nbsp;|&nbsp; <a href="Class_ToTempletRead.asp">栏目模板查看</a></div></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form_m" method="post" action="">
    <tr> 
      <td width="38%" align="center" class="hback"><div align="left"> 
          <select name="s_Classid" id="select" multiple style="width:100%" size="18">
            <%
		  	Dim rs_movelist_rs,str_tmp_move
			Set rs_movelist_rs = server.CreateObject(G_FS_RS)
			rs_movelist_rs.Open "Select ID,ClassID,ClassName,ParentID,ReycleTF from FS_NS_NewsClass where ParentID='0'  and ReycleTF=0",Conn,1,3
			str_tmp_move = ""
			do while not rs_movelist_rs.eof
				str_tmp_move = str_tmp_move & "<option value="""& rs_movelist_rs ("ClassID") &""">"& rs_movelist_rs ("ClassName") &"</option>"
			   str_tmp_move = str_tmp_move & Fs_news.News_ChildNewsList(rs_movelist_rs("ClassID"),"")
			  rs_movelist_rs.movenext
		  Loop
		  	Response.Write str_tmp_move
		  rs_movelist_rs.close:set rs_movelist_rs=nothing
          %>
          </select>
          <input type="button" name="Submit" value="选定所有栏目" onClick="SelectAllClass()">
          <input type="button" name="Submit" value="取消选定栏目" onClick="UnSelectAllClass()">
        </div></td>
      <td width="6%" align="center" class="hback"> <strong>批<br>
        量<br>
        捆<br>
        绑</strong></td>
      <td width="56%" class="hback">栏目模板： 
        <input type="text" name="Templet" value="<%=Replace("/"& G_TEMPLETS_DIR &"/NewsClass/class.htm","//","/")%>" style="width:60%"> 
        <input name="Submit53" type="button" id="selNewsTemplet" value="选择模板"  onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=Replace("/"&G_VIRTUAL_ROOT_DIR&"/"& G_TEMPLETS_DIR,"//","/") %>',400,300,window,document.form_m.Templet);document.form_m.Templet.focus();"> 
        <br> <br>
        新闻模板： 
        <input name="NewsTemplet" type="text" id="NewsTemplet" style="width:60%" value="<%=Replace("/"& G_TEMPLETS_DIR &"/NewsClass/news.htm","//","/")%>"> 
        <input name="Submit532" type="button" id="Submit53" value="选择模板"  onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=Replace("/"&G_VIRTUAL_ROOT_DIR&"/"& G_TEMPLETS_DIR,"//","/") %>',400,300,window,document.form_m.NewsTemplet);document.form_m.NewsTemplet.focus();"></td>
    </tr>
    <tr>
      <td colspan="3" class="hback"><strong>注意：</strong>按住&quot;CTRL&quot;键或者&quot;shift&quot;键可以对栏目进行批量选择。如果您更换了某个栏目的模板，请分别刷新一下分类。否则前台不会变化</td>
    </tr>
    <tr> 
      <td colspan="3" class="hback"><div align="center"> 
          <input name="Action" type="hidden" id="Action" value="Templet_News">
          <input type="submit" name="Submit6" value="确定开始捆绑">
          <input type="reset" name="Submit7" value="重新设定">
        </div></td>
    </tr>
  </form>
</table>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
</body>
</html>
<%
set Fs_news = nothing
%>
<script language="JavaScript" type="text/JavaScript" src="js/Public.js"></script>
<script language="JavaScript" type="text/JavaScript">
function SelectAllClass(){
  for(var i=0;i<document.form_m.s_Classid.length;i++){
    document.form_m.s_Classid.options[i].selected=true;}
}
function UnSelectAllClass(){
  for(var i=0;i<document.form_m.s_Classid.length;i++){
    document.form_m.s_Classid.options[i].selected=false;}
}

function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = myForm.elements[i];  
    if (e.name != 'chkall')  
       e.checked = myForm.chkall.checked;  
    }  
	}
</script>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





