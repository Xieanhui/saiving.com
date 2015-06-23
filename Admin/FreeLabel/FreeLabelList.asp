<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%
'页面设置及权限判断
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,strShowErr
MF_Default_Conn
MF_Session_TF 
If Not MF_Check_Pop_TF("MF_sPublic") Then Err_Show
'分页参数设置
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage
int_RPP = 20 '设置每页显示数目
int_showNumberLink_ = 8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_ = "<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_ = " <font face=webdings title=""上十页"">7</font>"			'上十
toP1_ = " <font face=webdings title=""上一页"">3</font>"			'上一
toN1_ = " <font face=webdings title=""下一页"">4</font>"			'下一
toN10_ = " <font face=webdings title=""下十页"">8</font>"			'下十
toL_ = "<font face=webdings title=""最后一页"">:</font>"

'===========================================
'单个删除记录
Dim ActionStr,DelID
ActionStr = Request.QueryString("Act")
If ActionStr = "Del" Then
	DelID = Request.QueryString("LableID")
	IF DelID = "" Then
		Response.Write "<script>alert('没有选择需要删除的记录');</script>"
		Response.End
	Else
		Conn.ExeCute("Delete From FS_MF_FreeLabel Where LabelID = '" & NoSqlHack(DelID) & "'")
	End If
End If

'批量删除
Dim AllDelIDStr,IDStr,i
If Request.Form("Action") = "del" Then
	AllDelIDStr = NoSqlHack(Request.Form("FreeID"))
	If AllDelIDStr = "" Then
		Response.Write "<script>alert('没有选择需要删除的记录');</script>"
		Response.End
	Else
		AllDelIDStr = Replace(Replace(AllDelIDStr," ,",","),", ",",")
		If Right(AllDelIDStr,1) = "," Then
			AllDelIDStr = Left(AllDelIDStr,Len(AllDelIDStr) - 1)
		End If
	End If
	If Instr(AllDelIDStr,",") > 0 Then
		IDStr = "'" & Replace(AllDelIDStr,",","','") & "'"
	Else
		IDStr = "'" & AllDelIDStr & "'"
	End If
	AllDelIDStr = "'" & Replace(AllDelIDStr,",","','") & "'"
	if AllDelIDStr <> "" then Conn.ExeCute("Delete From FS_MF_FreeLabel Where LabelID In(" & AllDelIDStr & ")")	
	Response.Write "<script>alert('删除成功');parent.location.reload();</script>"	
End If
%>
<html>
<head>
<title>自由标签管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<body>
<table width="98%" height="66" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
  <tr class="hback" > 
    <td width="100%" height="20"  align="Left" class="xingmu">标签库</td>
  </tr>
  <tr class="hback" > 
     <td height="27" align="center" class="hback"><div align="left"><a href="../Label/All_Label_Stock.asp">所有标签</a>┆<a href="FreeLabelList.asp"><font color="#FF0000">自由标签</font></a>┆<a href="../Label/All_Label_Stock.asp?isDel=1">备份库</a>┆<a href="../Label/label_creat.asp">创建标签</a>┆<a href="../Label/Label_Class.asp" target="_self">标签分类</a>&nbsp;┆<a href="../Label/All_label_style.asp">样式管理</a>&nbsp;<a href="../../help?Label=MF_Label_Stock" target="_blank" style="cursor:help;"><img src="../Images/_help.gif" width="50" height="17" border="0"></a></div></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr>
    <td width="10%" class="xingmu"><div align="center">序号</div></td>
    <td width="30%" class="xingmu"><div align="center">标签名称</div></td>
    <td width="15%" class="xingmu"><div align="center">子系统</div></td>
    <td width="35%" class="xingmu"><div align="center">操作</div></td>
	<td width="10%" class="xingmu"><div align="center">选择</div></td>
  </tr>
  <tr class="hback_1">
	<td colspan="5" height="2"></td>
  </tr>
 <form name="FreeListForm" id="FreeListForm" method="post" action="" style="margin:0px;" target="TempFream"> 
<%
'取得数据列表
Dim FreeRs,FreeSql,FreeListNum,HaveVTF
Set FreeRs = Server.CreateObject(G_FS_RS)
FreeSql = "Select LabelID,LabelName,LabelSQl,NSFields,NCFields,LabelContent,selectNum,DesCon,SysType From FS_MF_FreeLabel Where ID > 0 Order By ID Desc"
FreeRs.Open FreeSql,Conn,1,1
IF FreeRs.Eof Then
	HaveVTF = False
%>  
  <tr class="hback_1">
	<td colspan="5" height="20">暂时无数据.</td>
  </tr>
<%
Else
	HaveVTF = True
	FreeRs.PageSize = int_RPP
	cPageNo = NoSqlHack(Request.QueryString("Page"))
	If cPageNo = "" Then cPageNo = 1
	If Not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo > FreeRs.PageCount Then cPageNo = FreeRs.PageCount 
	FreeRs.AbsolutePage = cPageNo
	For FreeListNum = 1 To FreeRs.pagesize
	If FreeRs.Eof Then Exit For
%>  
  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
    <td class="hback"><div align="center"><% = FreeListNum %></div></td>
    <td class="hback"><div align="left"><span style="cursor:hand;" onClick="opencat(lable_<% = FreeRs(0) %>)"><% = FreeRs(1) %></span></div></td>
    <td class="hback"><div align="center">
	<%
		If FreeRs(8) = "NS" Then
			Response.Write "新闻"
		ElseiF FreeRs(8) = "DS" Then
			Response.Write "下载"
		ElseIf FreeRs(8) = "MS" Then
			Response.Write "商城"
		End If			
	%>
	</div></td>
	<td class="hback"><div align="center"><a href="AddFreeOne.asp?Act=Edit&LableID=<% = FreeRs(0) %>">修改</a>┆<a href="FreeLabelList.asp?Act=Del&LableID=<% = FreeRs(0) %>" onClick="{if(confirm('确定删除此记录吗？')){return true;}return false;}">删除</a>┆<span onClick="Dis_Style('../<%=G_ADMIN_DIR%>/FreeLabel/FlableDIsStyle.asp?ConStr=<% = Server.URLEncode(FreeRs(5)) %>',500,350,'obj');" style="cursor:hand;">查看样式</span></div></td>
    <td class="hback"><div align="center"><input name="FreeID" type="checkbox" id="FreeID" value="<% = FreeRs(0) %>"></div></td>
  </tr>
  <tr id="lable_<% = FreeRs(0) %>" style="display:none;">
    <td colspan="5" class="hback" style="WORD-BREAK: break-all; TABLE-LAYOUT: fixed;">
	<div style="width:100%; line-height:20px; text-align:left;">
		<span style="width:100%; line-height:20px; text-align:left;">标签描述：</span><br />
		<span style="width:100%; line-height:20px; text-align:left;"><% = Server.HTMLEncode(FreeRs(7)) %></span><br />
		<span style="width:100%; line-height:20px; text-align:left;">查询语句：</span><br />
		<span style="width:100%; line-height:20px; text-align:left;"><% = Server.HTMLEncode(Replace(FreeRs(2),"*",",")) %></span>
	</div>
	</td>
  </tr>
<%
	FreeRs.MoveNext
	Next
End If	
%>  
   <tr>
     <td height="26" colspan="5" class="hback" valign="middle"><div style="height:25px;"><span style="width:50%; height:25px; line-height:25px; text-align:left; float:left;"><input type="button" name="AddNew" value="新建" onClick="JavaScript:location.href='AddFreeOne.asp?Act=Add'"></span>
	 <span style="width:48%; height:25px; line-height:25px; text-align:right; float:left;"> <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
	 选中所有
	 <input type="button" name="Submit" value="删除"  onClick="document.FreeListForm.Action.value='del';if(confirm('确定清除您所选择的记录吗？')){this.document.FreeListForm.submit();}"></span>
	 </div></td>
   </tr>
   <input name="Action" type="hidden" id="Action" value="">
  </form>
</table>
<% If HaveVTF = True Then %>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">  
   <tr class="hback">
    <td height="30" align="right" valign="middle">
<% response.Write fPageCount(FreeRs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo) %>
    </td>
   </tr>
</table>
<% End If %>
<iframe name="TempFream" id="TempFream" src="" width="0" height="0" frameborder="0"></iframe>
</body>
<% 
FreeRs.Close : Set FreeRs = Nothing
Conn.Close : Set Conn=nothing
%>
</html>
<script language="JavaScript" type="text/JavaScript">
function opencat(cat)
{
  if(cat.style.display=="none")
  {
     cat.style.display="";
  } 
  else
  {
     cat.style.display="none"; 
  }
}
function CheckAll(form)  
{  
  	for (var i=0;i<form.elements.length;i++)  
	{  
		var e = FreeListForm.elements[i];  
		if (e.name != 'chkall')
		{
			e.checked = FreeListForm.chkall.checked;  
		}  
	}  
}

function Dis_Style(URL,widthe,heighte,obj)
{
  var obj=window.OpenWindowAndSetValue("../../Fs_Inc/convert.htm?"+URL,widthe,heighte,'window',obj)
  if (obj==undefined)return false;
}
</script>