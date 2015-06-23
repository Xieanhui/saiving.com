<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="lib/cls_main.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%'Copyright (c) 2006 Foosun Inc.  
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
MF_User_Conn
'session判断
MF_Session_TF
'权限判断
	if not Get_SubPop_TF("","NS_Special","NS","specail") then Err_Show
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
int_RPP=15 '设置每页显示数目
int_showNumberLink_=10 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings>9</font>"   			'首页 
toP10_=" <font face=webdings>7</font>"			'上十
toP1_=" <font face=webdings>3</font>"			'上一
toN1_=" <font face=webdings>4</font>"			'下一
toN10_=" <font face=webdings>8</font>"			'下十
toL_="<font face=webdings>:</font>"				'尾页
Dim  DelIDList,strShowErr
''=====================
If request.QueryString("Action")="del" Then 
	if not Get_SubPop_TF("","NS027","NS","specail") then Err_Show
	IF request.QueryString("SpecialID")<>"" then 
		Conn.execute("delete from FS_NS_Special where SpecialID = "&CintStr(request.QueryString("SpecialID")))
	Else
		DelIDList=FormatIntArr(request.Form("Cid"))
		if DelIDList="" then 
			strShowErr = "<li>你必须至少选择一项才能删除。</li>"
			Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		Conn.execute("delete from FS_NS_Special where SpecialID in ("& DelIDList &")")
		Call MF_Insert_oper_Log("专题","批量删除了专题,专题ID："& replace(request.Form("Cid")," ","") &"",now,session("admin_name"),"NS")
		strShowErr = "<li>恭喜，删除成功。</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If
End If	
''=====================
if request.QueryString("Act") = "OtherEdit" then
	if not Get_SubPop_TF("","NS027","NS","specail") then Err_Show
 
	if request.QueryString("EditSql")<>"" then 
		Conn.execute(Decrypt(request.QueryString("EditSql")))	
	end if
	if request.QueryString("Red_Url")<>"" then 
		response.Redirect(request.QueryString("Red_Url"))
	else
		response.Redirect("Special_Manage.asp")	
	end if	
end if
dim Fs_news
set Fs_news = new Cls_News
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>专题管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
//点击标题排序
/////////////////////////////////////////////////////////
var Old_Sql = document.URL;
function OrderByName(FildName)
{
	var New_Sql='';
	var oldFildName="";
	if (Old_Sql.indexOf("&filterorderby=")==-1&&Old_Sql.indexOf("?filterorderby=")==-1)
	{
		if (Old_Sql.indexOf("=")>-1)
			New_Sql = Old_Sql+"&filterorderby=" + FildName + "csed";
		else
			New_Sql = Old_Sql+"?filterorderby=" + FildName + "csed";
	}
	else
	{	
		var tmp_arr_ = Old_Sql.split('?')[1].split('&');
		for(var ii=0;ii<tmp_arr_.length;ii++)
		{
			if (tmp_arr_[ii].indexOf("filterorderby=")>-1)
			{
				oldFildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby=") + "filterorderby=".length , tmp_arr_[ii].length);
				break;	
			}
		}
		oldFildName.indexOf("csed")>-1?New_Sql = Old_Sql.replace('='+oldFildName,'='+FildName):New_Sql = Old_Sql.replace('='+oldFildName,'='+FildName+"csed");
	}	
	//alert(New_Sql);
	location = New_Sql;
}
/////////////////////////////////////////////////////////
</script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu"><a href="#" class="sd"><strong>专题管理</strong></a></td>
  </tr>
  <tr> 
    <td height="18" class="hback"><a href="Special_Manage.asp">管理首页</a> 
		| <a href="Special_Add.asp?Action=add">新增专题</a> 		
		| <a href="../../help?Lable=Special_Add" target="_blank" style="cursor:help;">帮助</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <form name="MainForm" method="post" action="Special_Manage.asp?Action=del">
    <tr class="xingmu"> 
      <td width="6%" align="center" class="xingmu" ><a href="javascript:OrderByName('SpecialID')" class="sd"><b>〖ID〗</b></a> 
        <span id="Show_Oder_SpecialID"></span></td>
      <td colspan="2" align="center" class="xingmu" ><a href="javascript:OrderByName('SpecialCName')" class="sd"><b>栏目中文[栏目英文]</b></a> 
        <span id="Show_Oder_SpecialCName"></span></td>
      <td width="18%" align="center" class="xingmu"><a href="javascript:OrderByName('AddTime')" class="sd"><b>添加时间</b></a> 
        <span id="Show_Oder_AddTime"></span></td>
      <td width="12%" align="center" class="xingmu"><a href="javascript:OrderByName('isLock')" class="sd"><b>属性</b></a> 
        <span id="Show_Oder_isLock"></span></td>
      <td width="20%" class="xingmu"> <div align="center">操作</div></td>
    </tr>
    <%
	Dim obj_special_rs,obj_special_rs_1,isUrlStr,This_Fun_Sql,orderby
	orderby = request.QueryString("filterorderby")
	This_Fun_Sql = "Select SpecialID,SpecialCName,SpecialEName,SpecialSize,SpecialContent,SavePath,Templet,ExtName,isLock,Addtime,GroupID,sPoint From FS_NS_Special "
	if orderby<>"" then This_Fun_Sql = This_Fun_Sql &"  Order By "& replace(orderby,"csed"," Desc")
	''response.Write(This_Fun_Sql)
	Set obj_special_rs = server.CreateObject(G_FS_RS)
	obj_special_rs.Open This_Fun_Sql,Conn,1,1
	IF obj_special_rs.eof THEN
	 	response.Write("<tr class=""hback""><td colspan=15>暂无数据.</td></tr>") 
	else	
	obj_special_rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo>obj_special_rs.PageCount Then cPageNo=obj_special_rs.PageCount 
	If cPageNo<=0 Then cPageNo=1
	obj_special_rs.AbsolutePage=cPageNo
	  FOR int_Start=1 TO int_RPP 
	%>
    <tr class="hback"> 
      <td height="22" class="hback"> <div align="center">
          <% = obj_special_rs("SpecialID")%>
        </div></td>
      <td width="26%" height="22" class="hback"> <% 
			Response.Write  "<a href=Special_add.asp?SpecialID="&obj_special_rs("SpecialID")&"&Action=edit>"&obj_special_rs("SpecialCName")&"</a>&nbsp;<font style=""font-size:9px;"">["& obj_special_rs("SpecialEName") &"]</font>"
		%> </td>
      <td width="18%" class="hback"><div align="center"><a href="News_Manage.asp?SpecialEName=<% = obj_special_rs("SpecialEName") %>&SpecialCName=<%=server.URLEncode(obj_special_rs("SpecialCName"))%>">查看此专题下新闻</a></div></td>
      <td class="hback" align="center"><% = obj_special_rs("Addtime")%></td>
      <td class="hback" align="center"> <%
		if obj_special_rs("isLock")=1 then 
			''锁定,需要解锁
			response.Write  "<input type=button value=""锁 定"" onclick=""javascript:location='Special_Manage.asp?Act=OtherEdit&EditSql="&server.URLEncode( Encrypt("Update FS_NS_Special set isLock=0 where SpecialID="&obj_special_rs("SpecialID")) )&"&Red_Url='"" title=""点击解锁"" style=""color:red"">" & vbcrlf
		else
			response.Write  "<input type=button value=""正 常"" onclick=""javascript:location='Special_Manage.asp?Act=OtherEdit&EditSql="&server.URLEncode( Encrypt("Update FS_NS_Special set isLock=1 where SpecialID="&obj_special_rs("SpecialID")) )&"&Red_Url='"" title=""点击锁定"">" & vbcrlf
		end if
	  %> </td>
      <td class="hback"><div align="center"> <a href="Special_add.asp?SpecialID=<% = obj_special_rs("SpecialID")%>&Action=edit">修改</a>  
          | <a href="Special_Manage.asp?SpecialID=<% = obj_special_rs("SpecialID")%>&Action=del"  onClick="{if(confirm('确定删除您所选择的专题吗？\n\n此专题下的新闻也将被删除!!')){return true;}return false;}">删除</a>&nbsp;&nbsp; 
          <input name="Cid" type="checkbox" id="Cid" value="<% = obj_special_rs("SpecialID")%>">
        </div></td>
    </tr>
    <%
		obj_special_rs.MoveNext
 		if obj_special_rs.eof then exit for
      NEXT
	End IF  
%>
    <tr> 
      <td height="30" colspan="5" align="left" class="hback"> 
		<%=fPageCount(obj_special_rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf%>       </td>
      <td height="30" class="hback"><input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
        选择所有 
        <input name="Submit" type="submit" id="Submit"  onClick="{if(confirm('确实要进行删除吗?')){this.document.MainForm.submit();return true;}return false;}" value=" 删除 "></td>
    </tr>
  </form>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr> 
    <td height="18" class="hback"><div align="left"> 
        <p><span class="tx"><strong>说明</strong></span>:点击标题排序<br>
        </p>
        </div></td>
  </tr>
</table>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
</body>
</html>
<%
obj_special_rs.close
set obj_special_rs =nothing
set Fs_news = nothing
%>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = MainForm.elements[i];  
    if (e.name != 'chkall')  
       e.checked = MainForm.chkall.checked;  
    }  
  }
//打开后根据规则显示箭头
var Req_FildName;
if (Old_Sql.indexOf("filterorderby=")>-1)
{
	var tmp_arr_ = Old_Sql.split('?')[1].split('&');
	for(var ii=0;ii<tmp_arr_.length;ii++)
	{
		if (tmp_arr_[ii].indexOf("filterorderby=")>-1)
		{
			if(Old_Sql.indexOf("csed")>-1)
				{Req_FildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby=") + "filterorderby=".length , tmp_arr_[ii].indexOf("csed"));break;}
			else
				{Req_FildName = tmp_arr_[ii].substring(tmp_arr_[ii].indexOf("filterorderby=") + "filterorderby=".length , tmp_arr_[ii].length);break;}	
		}
	}	
	if (document.getElementById('Show_Oder_'+Req_FildName)!=null)  
	{
		if(Old_Sql.indexOf(Req_FildName + "csed")>-1)
		{
			eval('Show_Oder_'+Req_FildName).innerText = '↓';
		}
		else
		{
			eval('Show_Oder_'+Req_FildName).innerText = '↑';
		}
	}	
}
///////////////////////////////////////////////////////// 
</script>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





