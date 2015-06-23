<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Cls_Cache.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<%
	Dim Conn,strShowErr,rs,Levels,str_Levels,str_Group_Name,str_PopList,rs_up
	MF_Default_Conn
	MF_Session_TF
	if not MF_Check_Pop_TF("MF_Pop") then Err_Show
	if Request.Form("Action")="Save" then
		if trim(Request.Form("Levels"))="" then
			strShowErr = "<li>错误的参数</li>"
			Response.Redirect("Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		set rs_up= Server.CreateObject(G_FS_RS)
		rs_up.open "select Levels,Group_Name,PopList From FS_MF_AdminGroup where Levels="& NoSqlHack(Request.Form("Levels"))&"",Conn,1,3
		if not rs_up.eof then
			rs_up("PopList")= Replace(NoSqlHack(Request.Form("PopList_MF"))&","&NoSqlHack(Request.Form("PopList_NS"))&","&NoSqlHack(Request.Form("PopList_MS"))&","&NoSqlHack(Request.Form("PopList_DS"))&","&NoSqlHack(Request.Form("PopList_ME"))&","&NoSqlHack(Request.Form("PopList_AP"))&","&NoSqlHack(Request.Form("PopList_SD"))&","&NoSqlHack(Request.Form("PopList_CS"))&","&NoSqlHack(Request.Form("PopList_SS"))&","&NoSqlHack(Request.Form("PopList_HS"))&","&NoSqlHack(Request.Form("PopList_VS"))&","&NoSqlHack(Request.Form("PopList_AS"))&","&NoSqlHack(Request.Form("PopList_WS"))&","&NoSqlHack(Request.Form("PopList_FL"))&","&NoSqlHack(Request.Form("PopList"))," ","")
			rs_up.update
			rs_up.close:set rs_up = nothing
		else
			strShowErr = "<li>错误的参数</li>"
			rs_up.close:set rs_up = nothing
			Response.Redirect("Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
		end if
		strShowErr = "<li>更新成功</li>"
		Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	end if
	select case CintStr(Request.QueryString("Leves"))
		case 1
			Levels = " and Levels=1"
		case 2
			Levels = " and Levels=2"
		case 3
			Levels = " and Levels=3"
		case 4
			Levels = " and Levels=4"
		case 5
			Levels = " and Levels=5"
		case else
			Levels = " and Levels=1"
	end select
	set rs= Server.CreateObject(G_FS_RS)
	rs.open "select Levels,Group_Name,PopList From FS_MF_AdminGroup where 1=1 "& Levels &"",Conn,1,3
	if rs.eof then
			strShowErr = "<li>找不到初始数据，请与程序供应商Foosun Inc.联系</li>"
			Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
			Response.end
	else
		str_Levels =rs(0)
		str_Group_Name =rs(1)
		str_PopList =rs(2)
	end if
Dim Rs_NewsClass
Set Rs_NewsClass= Server.CreateObject(G_FS_RS)
Rs_NewsClass.open "SELECT ClassID,ClassName From FS_NS_NewsClass",Conn,1,3
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
</HEAD>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/Prototype.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/CheckJs.js" type="text/JavaScript"></script>
<script language="JavaScript" type="text/JavaScript">
function SwitchPopType(Html_express)
{
	var Str_express=Html_express.substring(0,2);
	if ($("PopList_"+Str_express).checked==true)
	{
		$(Str_express+'_ID').style.display='';
	}
	else
	{
		$(Str_express+'_ID').style.display='none';
	}
	return true;
}
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = PopForm.elements[i];
    if (e.name != 'chkall')
       {
	   e.checked = PopForm.chkall.checked;
	   document.getElementById('MF_ID').style.display='';
	   }
	}
	}
function ChooseadminType(Type)
{
	switch (Type)
	{
	case "0":
		document.getElementById('all_id').style.display='';
	break;
	default:
		document.getElementById('all_id').style.display='none';
	break;
	}
}
var thisPop="";
function UpdateNsClass(Obj)
{
	var PopId=Obj.id;
	PopId=PopId.substring(6,PopId.length);
	if (thisPop!=PopId)
	{
		if ($(PopId).checked)
		{
			var Popvalue=$(PopId).value;
			var checkarr="";
			var str_Classlist=Popvalue.substring(8,Popvalue.length);
			SelectClear($("ClassList"));
			if (str_Classlist!="")
			{
				checkarr=str_Classlist.split("|");
				;
				for (var j=0;j<checkarr.length;j++)
				{
					if (checkarr[j]!="")
					{
						Selectone($("ClassList"),checkarr[j])
					}
				}
			}
			else
			{
				SelectClear($("ClassList"));
			}
			$("ClassListarea").style.display="";
			if (thisPop!="")
			{
				$("Update"+thisPop).innerHTML="设置栏目";
				$("Update"+thisPop).style.color="#000000";
			}
			thisPop=PopId;
			Obj.innerHTML="设置完毕";
			Obj.style.color="#FF0000";
		}
		else
		{
			if (thisPop!="")
			{
				$("Update"+thisPop).innerHTML="设置栏目";
				$("Update"+thisPop).style.color="#000000";
			}
			thisPop="";
			$("ClassListarea").style.display="none";
			alert("必须勾选此功能后才能设置栏目");
		}
	}
	else
	{
		Obj.innerHTML="设置栏目";
		Obj.style.color="#000000";
		thisPop="";
		$("ClassListarea").style.display="none";
	}
}
function Selectone(Obj,value)
{
	for (var i=0 ;i<Obj.length ;i++)
	{
		if (Obj[i].value==value)
			Obj[i].selected=true;
	}
}
function SelectClear(Obj)
{
	for (var i=0 ;i<Obj.length ;i++)
	{
			Obj[i].selected=false;
	}
}
function ClassListC(Obj)
{
	if (thisPop!="" && GetSelectValue(Obj)!="")
	{
		$(thisPop).value=thisPop+"$$$"+GetSelectValue(Obj);
	}
}
function GetSelectValue(Obj)
{
	var ReturnValue="";
	for (var i=0;i<Obj.length;i++)
	{
		if (Obj[i].selected)
		{
			if (ReturnValue=="")
			{
				ReturnValue=Obj[i].value;
			}else
			{
				ReturnValue+="|"+Obj[i].value;
			}
		}
	}
	return ReturnValue;
}
function Selectall(Obj)
{
	for (var i=0 ;i<Obj.length ;i++)
	{
		Obj[i].selected=true;
	}
}
</script>
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	<tr class="xingmu">
		<td colspan="2" class="xingmu">设置固定管理组:<span class="tx"><b><%=str_Group_Name%></b></span>的权限</td>
	</tr>
	<tr class="hback">
		<td colspan="2"><a href="SysAdmin_List.asp">返回管理员列表</a><a href="SysAdmin_List.asp?my=1"></a> </td>
	</tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	<form name="PopForm" method="post" action="">
		<tr>
			<td class="hback"><div align="left">
					<input name="Levels" type="hidden" id="Levels" value="<%=str_Levels%>">
					<input name="Action" type="hidden" id="Action" value="Save">
					管理员类型: <a href="SysAdmin_SetPop_Group.asp?Leves=1">超级管理员</a>┆<a href="SysAdmin_SetPop_Group.asp?Leves=2">一般管理员</a>┆<a href="SysAdmin_SetPop_Group.asp?Leves=3">总编辑</a>┆<a href="SysAdmin_SetPop_Group.asp?Leves=4">责任编辑</a>┆<a href="SysAdmin_SetPop_Group.asp?Leves=5">记者</a>┆
					<input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form);">
					选中/取消所有</div></td>
		</tr>
		<tr>
			<td class="hback"><table width="100%" border="0" cellspacing="1" cellpadding="5" style="display:;" id="all_id">
					<%if Request.Cookies("FoosunSUBCookie")("FoosunSUBMF")=1 then%>
					<tr>
						<td colspan="2" class="hback_1"><div align="left">
								<input name="PopList_MF" type="checkbox" id="PopList" onClick="SwitchPopType('MF_Pop');"   value="MF_Pop" <%if InStr(1,str_PopList,"MF_Pop",1)<>0 then response.Write("checked")%>>
								MF主系统 </div></td>
					</tr>
					<tr>
						<td colspan="2" class="hback"><table width="100%" border="0" cellpadding="4" cellspacing="0" style="display: <%if InStr(1,str_PopList,"MF_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>"  id="MF_ID">
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td width="14%" class="hback"><div align="left">
											<input name="PopList" type="checkbox"  value="MF_Templet" <%if InStr(1,str_PopList,"MF_Templet",1)<>0 then:response.Write("checked")%>>
											模板管理 </div></td>
									<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" value="MF001" <%if InStr(1,str_PopList,"MF001",1)>0 then response.Write "checked"%>>
													修改文件
													<input name="PopList" type="checkbox" id="PopList" value="MF002"  <%if InStr(1,str_PopList,"MF002",1)>0 then response.Write "checked"%>>
													改名文件
													<input name="PopList" type="checkbox" id="PopList" value="MF003" <%if InStr(1,str_PopList,"MF003",1)>0 then response.Write "checked"%>>
													删除文件
													<input name="PopList" type="checkbox" id="PopList" value="MF004" <%if InStr(1,str_PopList,"MF004",1)>0 then response.Write "checked"%>>
													创建目录
													<input name="PopList" type="checkbox" id="PopList" value="MF005" <%if InStr(1,str_PopList,"MF005",1)>0 then response.Write "checked"%>>
													导入文件
													<input name="PopList" type="checkbox" id="PopList" value="MF006" <%if InStr(1,str_PopList,"MF006",1)>0 then response.Write "checked"%>>
													在线编辑(修改模板) </td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><div align="left">
											<input name="PopList" type="checkbox" id="PopList" value="MF_Style"  <%if InStr(1,str_PopList,"MF_Style",1)>0 then response.Write "checked"%>>
											标签样式</div></td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="MF007" <%if InStr(1,str_PopList,"MF007",1)>0 then response.Write "checked"%>>
													增加样式
													<input name="PopList" type="checkbox" id="PopList" value="MF008" <%if InStr(1,str_PopList,"MF008",1)>0 then response.Write "checked"%>>
													修改样式 </td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><div align="left">
											<input name="PopList" type="checkbox" id="PopList" value="MF_Public" <%if InStr(1,str_PopList,"MF_Public",1)>0 then response.Write "checked"%>>
											发布管理</div></td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="MF009" <%if InStr(1,str_PopList,"MF009",1)>0 then response.Write "checked"%>>
													开启 </td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><div align="left">
											<input name="PopList" type="checkbox" id="PopList" value="MF_sPublic" <%if InStr(1,str_PopList,"MF_sPublic",1)>0 then response.Write "checked"%>>
											标签管理</div></td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="MF025" <%if InStr(1,str_PopList,"MF025",1)>0 then response.Write "checked"%>>
													创建标签 </td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><div align="left">
											<input name="PopList" type="checkbox" id="PopList" value="MF_SysSet" <%if InStr(1,str_PopList,"MF_SysSet",1)>0 then response.Write "checked"%>>
											参数设置</div></td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="MF010" <%if InStr(1,str_PopList,"MF010",1)>0 then response.Write "checked"%>>
													开启 </td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><div align="left">
											<input name="PopList" type="checkbox" id="PopList" value="MF_Const" <%if InStr(1,str_PopList,"MF_Const",1)>0 then response.Write "checked"%>>
											配置文件</div></td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="MF011" <%if InStr(1,str_PopList,"MF011",1)>0 then response.Write "checked"%>>
													全局变量设置
													<input name="PopList" type="checkbox" id="PopList" value="MF012" <%if InStr(1,str_PopList,"MF012",1)>0 then response.Write "checked"%>>
													自动刷新配置文件</td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><div align="left">
											<input name="PopList" type="checkbox" id="PopList" value="MF_SubSite" <%if InStr(1,str_PopList,"MF_SubSite",1)>0 then response.Write "checked"%>>
											子系统维护</div></td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="MF013" <%if InStr(1,str_PopList,"MF013",1)>0 then response.Write "checked"%>>
													开启 </td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MF_DataFix" <%if InStr(1,str_PopList,"MF_DataFix",1)>0 then response.Write "checked"%>>
										数据库维护</td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="MF015" <%if InStr(1,str_PopList,"MF015",1)>0 then response.Write "checked"%>>
													数据库压缩
													<input name="PopList" type="checkbox" id="PopList" value="MF016" <%if InStr(1,str_PopList,"MF016",1)>0 then response.Write "checked"%>>
													数据库备份
													<input name="PopList" type="checkbox" id="PopList" value="MF017" <%if InStr(1,str_PopList,"MF017",1)>0 then response.Write "checked"%>>
													SQL语句查询操作</td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MF_Log" <%if InStr(1,str_PopList,"MF_Log",1)>0 then response.Write "checked"%>>
										日志管理</td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="MF018" <%if InStr(1,str_PopList,"MF018",1)>0 then response.Write "checked"%>>
													操作日志
													<input name="PopList" type="checkbox" id="PopList" value="MF019" <%if InStr(1,str_PopList,"MF019",1)>0 then response.Write "checked"%>>
													安全日志
													<input name="PopList" type="checkbox" id="PopList" value="MF020" <%if InStr(1,str_PopList,"MF020",1)>0 then response.Write "checked"%>>
													删除日志</td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MF_Define" <%if InStr(1,str_PopList,"MF_Define",1)>0 then response.Write "checked"%>>
										自定义字段</td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="MF021" <%if InStr(1,str_PopList,"MF021",1)>0 then response.Write "checked"%>>
													增加
													<input name="PopList" type="checkbox" id="PopList" value="MF022" <%if InStr(1,str_PopList,"MF022",1)>0 then response.Write "checked"%>>
													删除</td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MF_other" <%if InStr(1,str_PopList,"MF_other",1)>0 then response.Write "checked"%>>
										其他</td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="MF024" <%if InStr(1,str_PopList,"MF024",1)>0 then response.Write "checked"%>>
													修改密码
													<input name="PopList" type="checkbox" id="PopList" value="MF025" <%if InStr(1,str_PopList,"MF025",1)>0 then response.Write "checked"%>>
													上传文件
													<input name="PopList" type="checkbox" id="PopList" value="MF026" <%if InStr(1,str_PopList,"MF026",1)>0 then response.Write "checked"%>>
													重命名文件/文件夹
													<input name="PopList" type="checkbox" id="PopList" value="MF027" <%if InStr(1,str_PopList,"MF027",1)>0 then response.Write "checked"%>>
													删除文件/目录
													<input name="PopList" type="checkbox" id="PopList" value="MF028" <%if InStr(1,str_PopList,"MF028",1)>0 then response.Write "checked"%>>
													新建目录</td>
											</tr>
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="MF029" <%if InStr(1,str_PopList,"MF029",1)>0 then response.Write "checked"%>>
													管理员管理&nbsp;</td>
											</tr>
										</table></td>
								</tr>
							</table></td>
					</tr>
					<%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBNS")=1 then
		  %>
					<tr  class="hback_1">
						<td colspan="2" class="hback_1"><div align="left"> <strong>
								<input name="PopList_NS" type="checkbox" id="PopList" value="NS_Pop"  onClick="SwitchPopType('NS_Pop');" <%if InStr(1,str_PopList,"NS_Pop",1)>0 then response.Write "checked"%>>
								</strong>NS新闻系统</div></td>
					</tr>
					<tr  class="hback">
						<td colspan="2" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"NS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="NS_ID">
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td width="14%" class="hback"><div align="left">
											<input name="PopList" type="checkbox" id="PopList" value="NS_News" <%if InStr(1,str_PopList,"NS_News",1)>0 then response.Write "checked"%>>
											新闻管理</div></td>
									<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2" <%if InStr(1,str_PopList,"MF010",1)>0 then response.Write "checked"%>>
											<tr>
												<td><table width="98%" border="0" cellspacing="0" cellpadding="2" <%if InStr(1,str_PopList,"MF010",1)>0 then response.Write "checked"%>>
														<tr>
															<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
																	<tr>
																		<td width="30%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
																				<tr>
																					<td><input name="PopList" type="checkbox" id="NS013" value="<%= GetNewsPopStr(str_PopList,"NS013") %>" <%if InStr(1,str_PopList,"NS013",1)>0 then response.Write "checked"%> />
																						<label for="NS013">查看</label></td>
																					<td><button id="UpdateNS013" onClick="UpdateNsClass(this);">设置栏目</button></td>
																				</tr>
																				<tr>
																					<td width="50%"><input name="PopList" type="checkbox" id="NS001" value="<%= GetNewsPopStr(str_PopList,"NS001") %>" <%if InStr(1,str_PopList,"NS001",1)>0 then response.Write "checked"%> />
																						<label for="NS001">添加</label></td>
																					<td><button id="UpdateNS001" onClick="UpdateNsClass(this);">设置栏目</button></td>
																				</tr>
																				<tr>
																					<td><input name="PopList" type="checkbox" id="NS002" value="<%= GetNewsPopStr(str_PopList,"NS002") %>" <%if InStr(1,str_PopList,"NS002",1)>0 then response.Write "checked"%> />
																						<label for="NS002">修改</label></td>
																					<td><button id="UpdateNS002" onClick="UpdateNsClass(this);">设置栏目</button></td>
																				</tr>
																				<tr>
																					<td><input name="PopList" type="checkbox" id="NS003" value="<%= GetNewsPopStr(str_PopList,"NS003") %>" <%if InStr(1,str_PopList,"NS003",1)>0 then response.Write "checked"%> />
																						<label for="NS003">删除</label></td>
																					<td><button id="UpdateNS003" onClick="UpdateNsClass(this);">设置栏目</button></td>
																				</tr>
																				<tr>
																					<td><input name="PopList" type="checkbox" id="NS004" value="<%= GetNewsPopStr(str_PopList,"NS004") %>" <%if InStr(1,str_PopList,"NS004",1)>0 then response.Write "checked"%> />
																						<label for="NS004">审核</label></td>
																					<td><button id="UpdateNS004" onClick="UpdateNsClass(this);">设置栏目</button></td>
																				</tr>
																				<tr>
																					<td><input name="PopList" type="checkbox" id="NS005" value="<%= GetNewsPopStr(str_PopList,"NS005") %>" <%if InStr(1,str_PopList,"NS005",1)>0 then response.Write "checked"%> />
																						<label for="NS005">锁定</label></td>
																					<td><button id="UpdateNS005" onClick="UpdateNsClass(this);">设置栏目</button></td>
																				</tr>
																				<tr>
																					<td><input name="PopList" type="checkbox" id="NS007" value="<%= GetNewsPopStr(str_PopList,"NS007") %>" <%if InStr(1,str_PopList,"NS007",1)>0 then response.Write "checked"%> />
																						<label for="NS007">复制</label></td>
																					<td><button id="UpdateNS007" onClick="UpdateNsClass(this);">设置栏目</button></td>
																				</tr>
																				<tr>
																					<td><input name="PopList" type="checkbox" id="NS009" value="<%= GetNewsPopStr(str_PopList,"NS009") %>" <%if InStr(1,str_PopList,"NS009",1)>0 then response.Write "checked"%> />
																						<label for="NS009">移动</label></td>
																					<td><button id="UpdateNS009" onClick="UpdateNsClass(this);">设置栏目</button></td>
																				</tr>
																				<tr>
																					<td><input name="PopList" type="checkbox" id="NS006" value="<%= GetNewsPopStr(str_PopList,"NS006") %>" <%if InStr(1,str_PopList,"NS006",1)>0 then Response.Write "checked"%> />
																						<label for="NS006">设置权重</label></td>
																					<td><button id="UpdateNS006" onClick="UpdateNsClass(this);">设置栏目</button></td>
																				</tr>
																				<tr>
																					<td><input name="PopList" type="checkbox" id="NS010" value="<%= GetNewsPopStr(str_PopList,"NS010") %>" <%if InStr(1,str_PopList,"NS010",1)>0 then response.Write "checked"%> />
																						<label for="NS010">批量替换</label></td>
																					<td><button id="UpdateNS010" onClick="UpdateNsClass(this);">设置栏目</button></td>
																				</tr>
																				<tr>
																					<td><input name="PopList" type="checkbox" id="NS011" value="<%= GetNewsPopStr(str_PopList,"NS011") %>" <%if InStr(1,str_PopList,"NS011",1)>0 then response.Write "checked"%> />
																						<label for="NS011">生成</label></td>
																					<td><button id="UpdateNS011" onClick="UpdateNsClass(this);">设置栏目</button></td>
																				</tr>
																				<tr>
																					<td><input name="PopList" type="checkbox" id="NS012" value="<%= GetNewsPopStr(str_PopList,"NS012") %>" <%if InStr(1,str_PopList,"NS012",1)>0 then response.Write "checked"%> />
																						<label for="NS012">加入JS</label></td>
																					<td><button id="UpdateNS012" onClick="UpdateNsClass(this);">设置栏目</button></td>
																				</tr>
																			</table></td>
																		<td><div id="ClassListarea" style="display:none;">
																				<table width="100%" border="0" cellspacing="0" cellpadding="0">
																					<tr>
																						<td width="20%" valign="middle"><select name="ClassList" onchange="ClassListC(this)" size="18" multiple id="ClassList" onChange="">
																							<% While Not Rs_NewsClass.Eof %>
																							<option value="<%= Rs_NewsClass("ClassID") %>"><%= Rs_NewsClass("ClassName") %></option>
																							<%  	Rs_NewsClass.Movenext %>
																							<% Wend %>
																							</select></td>
																						<td valign="middle"><button onClick="Selectall(this.form.ClassList);ClassListC(this.form.ClassList)">选择所有栏目</button></td>
																					</tr>
																				</table>
																			</div>&nbsp;</td>
																	</tr>
																</table></td>
														</tr>
													</table></td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><div align="left">
											<input name="PopList" type="checkbox" id="PopList" value="NS_Class" <%if InStr(1,str_PopList,"NS_Class",1)>0 then response.Write "checked"%>>
											栏目管理</div></td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="NS016" <%if InStr(1,str_PopList,"NS016",1)>0 then response.Write "checked"%>>
													添加
													<input name="PopList" type="checkbox" id="PopList" value="NS017" <%if InStr(1,str_PopList,"NS017",1)>0 then response.Write "checked"%>>
													修改
													<input name="PopList" type="checkbox" id="PopList" value="NS018" <%if InStr(1,str_PopList,"NS018",1)>0 then response.Write "checked"%>>
													复位
													<input name="PopList" type="checkbox" id="PopList" value="NS019" <%if InStr(1,str_PopList,"NS019",1)>0 then response.Write "checked"%>>
													合并
													<input name="PopList" type="checkbox" id="PopList" value="NS020" <%if InStr(1,str_PopList,"NS020",1)>0 then response.Write "checked"%>>
													转移
													<input name="PopList" type="checkbox" id="PopList" value="NS021" <%if InStr(1,str_PopList,"NS021",1)>0 then response.Write "checked"%>>
													删除
													<input name="PopList" type="checkbox" id="PopList" value="NS022" <%if InStr(1,str_PopList,"NS022",1)>0 then response.Write "checked"%>>
													xml(RSS)
													<input name="PopList" type="checkbox" id="PopList" value="NS023" <%if InStr(1,str_PopList,"NS023",1)>0 then response.Write "checked"%>>
													清空
													<input name="PopList" type="checkbox" id="PopList" value="NS024" <%if InStr(1,str_PopList,"NS024",1)>0 then response.Write "checked"%>>
													排序</td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_Special" <%if InStr(1,str_PopList,"NS_Special",1)>0 then response.Write "checked"%>>
										专题管理</td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="NS026" <%if InStr(1,str_PopList,"NS026",1)>0 then response.Write "checked"%>>
													添加
													<input name="PopList" type="checkbox" id="PopList" value="NS027" <%if InStr(1,str_PopList,"NS027",1)>0 then response.Write "checked"%>>
													修改
													<input name="PopList" type="checkbox" id="PopList" value="NS028" <%if InStr(1,str_PopList,"NS028",1)>0 then response.Write "checked"%>>
													锁定 </td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_Constr" <%if InStr(1,str_PopList,"NS_Constr",1)>0 then response.Write "checked"%>>
										投稿管理</td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="NS030" <%if InStr(1,str_PopList,"NS030",1)>0 then response.Write "checked"%>>
													审核
													<input name="PopList" type="checkbox" id="PopList" value="NS031" <%if InStr(1,str_PopList,"NS031",1)>0 then response.Write "checked"%>>
													锁定
													<input name="PopList" type="checkbox" id="PopList" value="NS032" <%if InStr(1,str_PopList,"NS032",1)>0 then response.Write "checked"%>>
													删除
													<input name="PopList" type="checkbox" id="PopList" value="NS033" <%if InStr(1,str_PopList,"NS033",1)>0 then response.Write "checked"%>>
													退稿
													<input name="PopList" type="checkbox" id="PopList" value="NS034" <%if InStr(1,str_PopList,"NS034",1)>0 then response.Write "checked"%>>
													统计</td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_Templet" <%if InStr(1,str_PopList,"NS_Templet",1)>0 then response.Write "checked"%>>
										捆绑模板</td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="NS036" <%if InStr(1,str_PopList,"NS036",1)>0 then response.Write "checked"%>>
													开启</td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_Freejs" <%if InStr(1,str_PopList,"NS_Freejs",1)>0 then response.Write "checked"%>>
										自由JS管理</td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="NS037" <%if InStr(1,str_PopList,"NS037",1)>0 then response.Write "checked"%>>
													增加
													<input name="PopList" type="checkbox" id="PopList" value="NS038" <%if InStr(1,str_PopList,"NS038",1)>0 then response.Write "checked"%>>
													修改
													<input name="PopList" type="checkbox" id="PopList" value="NS039" <%if InStr(1,str_PopList,"NS039",1)>0 then response.Write "checked"%>>
													删除 </td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_Sysjs" <%if InStr(1,str_PopList,"NS_Sysjs",1)>0 then response.Write "checked"%>>
										系统JS管理</td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="NS040" <%if InStr(1,str_PopList,"NS040",1)>0 then response.Write "checked"%>>
													增加
													<input name="PopList" type="checkbox" id="PopList" value="NS041" <%if InStr(1,str_PopList,"NS041",1)>0 then response.Write "checked"%>>
													修改
													<input name="PopList" type="checkbox" id="PopList" value="NS042" <%if InStr(1,str_PopList,"NS042",1)>0 then response.Write "checked"%>>
													删除 </td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_Recyle" <%if InStr(1,str_PopList,"NS_Recyle",1)>0 then response.Write "checked"%>>
										回收站管理</td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="NS043" <%if InStr(1,str_PopList,"NS043",1)>0 then response.Write "checked"%>>
													恢复
													<input name="PopList" type="checkbox" id="PopList" value="NS044" <%if InStr(1,str_PopList,"NS044",1)>0 then response.Write "checked"%>>
													删除 </td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_UnRl" <%if InStr(1,str_PopList,"NS_UnRl",1)>0 then response.Write "checked"%>>
										不规则新闻</td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="NS045" <%if InStr(1,str_PopList,"NS045",1)>0 then response.Write "checked"%>>
													增加
													<input name="PopList" type="checkbox" id="PopList" value="NS046" <%if InStr(1,str_PopList,"NS046",1)>0 then response.Write "checked"%>>
													修改
													<input name="PopList" type="checkbox" id="PopList" value="NS047" <%if InStr(1,str_PopList,"NS047",1)>0 then response.Write "checked"%>>
													删除</td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_Genal" <%if InStr(1,str_PopList,"NS_Genal",1)>0 then response.Write "checked"%>>
										常规管理</td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="NS048" <%if InStr(1,str_PopList,"NS048",1)>0 then response.Write "checked"%>>
													开启</td>
											</tr>
										</table></td>
								</tr>
								<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
									<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_Param" <%if InStr(1,str_PopList,"NS_Param",1)>0 then response.Write "checked"%>>
										系统参数设置</td>
									<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
											<tr>
												<td><input name="PopList" type="checkbox" id="PopList" value="NS049" <%if InStr(1,str_PopList,"NS049",1)>0 then response.Write "checked"%>>
													开启</td>
											</tr>
										</table></td>
								</tr>
							</table></td>
					</tr>
					<%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBMS")=1 then
		  %>
					<tr  class="hback_1">
						<td colspan="2" class="hback_1"><div align="left"> <strong>
								<input name="PopList_MS" type="checkbox" id="PopList" value="MS_Pop"  onClick="SwitchPopType('MS_Pop');"  <%if InStr(1,str_PopList,"MS_Pop",1)>0 then response.Write "checked"%>>
								</strong>MS商城系统</div></td>
					</tr>
					<tr class="hback">
						<td colspan="2" class="hback"><div align="left">
								<table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"MS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="MS_ID">
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="MS_Products" <%if InStr(1,str_PopList,"MS_Products",1)>0 then response.Write "checked"%>>
												商品管理</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="MS001" <%if InStr(1,str_PopList,"MS001",1)>0 then response.Write "checked"%>>
														添加
														<input name="PopList" type="checkbox" id="PopList" value="MS002" <%if InStr(1,str_PopList,"MS002",1)>0 then response.Write "checked"%>>
														修改
														<input name="PopList" type="checkbox" id="PopList" value="MS003" <%if InStr(1,str_PopList,"MS003",1)>0 then response.Write "checked"%>>
														删除
														<input name="PopList" type="checkbox" id="PopList" value="MS004" <%if InStr(1,str_PopList,"MS004",1)>0 then response.Write "checked"%>>
														审核
														<input name="PopList" type="checkbox" id="PopList" value="MS005" <%if InStr(1,str_PopList,"MS005",1)>0 then response.Write "checked"%>>
														锁定
														<input name="PopList" type="checkbox" id="PopList" value="MS006" <%if InStr(1,str_PopList,"MS006",1)>0 then response.Write "checked"%>>
														置顶
														<input name="PopList" type="checkbox" id="PopList" value="MS007" <%if InStr(1,str_PopList,"MS007",1)>0 then response.Write "checked"%>>
														热点推荐</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="MS_Class" <%if InStr(1,str_PopList,"MS_Class",1)>0 then response.Write "checked"%>>
												类别栏目</div></td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="MS010" <%if InStr(1,str_PopList,"MS010",1)>0 then response.Write "checked"%>>
														添加
														<input name="PopList" type="checkbox" id="PopList" value="MS011" <%if InStr(1,str_PopList,"MS011",1)>0 then response.Write "checked"%>>
														修改
														<input name="PopList" type="checkbox" id="PopList" value="MS012" <%if InStr(1,str_PopList,"MS012",1)>0 then response.Write "checked"%>>
														删除
														<input name="PopList" type="checkbox" id="PopList" value="MS013" <%if InStr(1,str_PopList,"MS013",1)>0 then response.Write "checked"%>>
														排序
														<input name="PopList" type="checkbox" id="PopList" value="MS014" <%if InStr(1,str_PopList,"MS014",1)>0 then response.Write "checked"%>>
														复位
														<input name="PopList" type="checkbox" id="PopList" value="MS015" <%if InStr(1,str_PopList,"MS015",1)>0 then response.Write "checked"%>>
														合并
														<input name="PopList" type="checkbox" id="PopList" value="MS016" <%if InStr(1,str_PopList,"MS016",1)>0 then response.Write "checked"%>>
														清空 </td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MS_Special" <%if InStr(1,str_PopList,"MS_Special",1)>0 then response.Write "checked"%>>
											专区管理</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="MS017" <%if InStr(1,str_PopList,"MS017",1)>0 then response.Write "checked"%>>
														添加
														<input name="PopList" type="checkbox" id="PopList" value="MS018" <%if InStr(1,str_PopList,"MS018",1)>0 then response.Write "checked"%>>
														修改
														<input name="PopList" type="checkbox" id="PopList" value="MS019" <%if InStr(1,str_PopList,"MS019",1)>0 then response.Write "checked"%>>
														删除
														<input name="PopList" type="checkbox" id="PopList" value="MS020" <%if InStr(1,str_PopList,"MS020",1)>0 then response.Write "checked"%>>
														锁定</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MS_order" <%if InStr(1,str_PopList,"MS_order",1)>0 then response.Write "checked"%>>
											定单管理</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="MS021" <%if InStr(1,str_PopList,"MS021",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MS_WrOut" <%if InStr(1,str_PopList,"MS_WrOut",1)>0 then response.Write "checked"%>>
											退货管理</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="MS022" <%if InStr(1,str_PopList,"MS022",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MS_Company" <%if InStr(1,str_PopList,"MS_Company",1)>0 then response.Write "checked"%>>
											物流公司</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="MS023" <%if InStr(1,str_PopList,"MS023",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MS_Recycle" <%if InStr(1,str_PopList,"MS_Recycle",1)>0 then response.Write "checked"%>>
											回收站管理</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="MS024" <%if InStr(1,str_PopList,"MS024",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MS_Param" <%if InStr(1,str_PopList,"MS_Param",1)>0 then response.Write "checked"%>>
											系统参数</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="MS025" <%if InStr(1,str_PopList,"MS025",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
								</table>
							</div></td>
					</tr>
					<%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBDS")=1 then
		  %>
					<tr  class="hback_1">
						<td colspan="2" class="hback_1"><div align="left"><strong>
								<input name="PopList_DS" type="checkbox" id="PopList_DS" value="DS_Pop"  onClick="SwitchPopType('DS_Pop');" <%if InStr(1,str_PopList,"DS_Pop",1)>0 then response.Write "checked"%>>
								</strong>DS下载系统</div></td>
					</tr>
					<tr class="hback">
						<td colspan="2" class="hback"><div align="left">
								<table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"DS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="DS_ID">
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="Down_List" <%if InStr(1,str_PopList,"Down_List",1)>0 then response.Write "checked"%>>
												下载列表管理</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="DS001" <%if InStr(1,str_PopList,"DS001",1)>0 then response.Write "checked"%>>
														添加
														<input name="PopList" type="checkbox" id="PopList" value="DS002" <%if InStr(1,str_PopList,"DS002",1)>0 then response.Write "checked"%>>
														修改
														<input name="PopList" type="checkbox" id="PopList" value="DS003" <%if InStr(1,str_PopList,"DS003",1)>0 then response.Write "checked"%>>
														删除
														<input name="PopList" type="checkbox" id="PopList" value="DS005" <%if InStr(1,str_PopList,"DS005",1)>0 then response.Write "checked"%>>
														锁定 </td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="DS_Class" <%if InStr(1,str_PopList,"DS_Class",1)>0 then response.Write "checked"%>>
												栏目管理</div></td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="DS010" <%if InStr(1,str_PopList,"DS010",1)>0 then response.Write "checked"%>>
														添加
														<input name="PopList" type="checkbox" id="PopList" value="DS011" <%if InStr(1,str_PopList,"DS011",1)>0 then response.Write "checked"%>>
														修改
														<input name="PopList" type="checkbox" id="PopList" value="DS012" <%if InStr(1,str_PopList,"DS012",1)>0 then response.Write "checked"%>>
														删除
														<input name="PopList" type="checkbox" id="PopList" value="DS013" <%if InStr(1,str_PopList,"DS013",1)>0 then response.Write "checked"%>>
														排序
														<input name="PopList" type="checkbox" id="PopList" value="DS014" <%if InStr(1,str_PopList,"DS014",1)>0 then response.Write "checked"%>>
														复位
														<input name="PopList" type="checkbox" id="PopList" value="DS015"  <%if InStr(1,str_PopList,"DS015",1)>0 then response.Write "checked"%>>
														合并
														<input name="PopList" type="checkbox" id="PopList" value="DS016"  <%if InStr(1,str_PopList,"DS016",1)>0 then response.Write "checked"%>>
														清空 </td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="DS_speical" <%if InStr(1,str_PopList,"DS_Class",1)>0 then response.Write "checked"%>>
											专区管理&nbsp;</td>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="DS018" <%if InStr(1,str_PopList,"DS018",1)>0 then response.Write "checked"%>>
											管理                          &nbsp;</td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="DS_Param" <%if InStr(1,str_PopList,"DS_Param",1)>0 then response.Write "checked"%>>
											参数设置</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="DS017" <%if InStr(1,str_PopList,"DS017",1)>0 then response.Write "checked"%>>
														开启
														<input name="PopList" type="checkbox" id="PopList" value="DS_KunBang" <%if InStr(1,str_PopList,"DS_KunBang",1)>0 then response.Write "checked"%>>
														模版捆绑</td>
												</tr>
											</table></td>
									</tr>
								</table>
							</div></td>
					</tr>
					<%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBME")=1 then
		  %>
					<tr  class="hback_1">
						<td colspan="2" class="hback_1"><div align="left"> <strong>
								<input name="PopList_ME" type="checkbox" id="PopList" value="ME_Pop"  onClick="SwitchPopType('ME_Pop');"  <%if InStr(1,str_PopList,"ME_Pop",1)>0 then response.Write "checked"%>>
								</strong>ME会员系统</div></td>
					</tr>
					<tr class="hback">
						<td colspan="2" class="hback"><div align="left">
								<table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"ME_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="ME_ID">
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="ME_List" <%if InStr(1,str_PopList,"ME_List",1)>0 then response.Write "checked"%>>
												会员管理</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME001" <%if InStr(1,str_PopList,"ME001",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="ME_Intergel" <%if InStr(1,str_PopList,"ME_Intergel",1)>0 then response.Write "checked"%>>
												积分管理</div></td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME002" <%if InStr(1,str_PopList,"ME002",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Card" <%if InStr(1,str_PopList,"ME_Card",1)>0 then response.Write "checked"%>>
											点卡管理</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME017" <%if InStr(1,str_PopList,"ME017",1)>0 then response.Write "checked"%>>
														添加
														<input name="PopList" type="checkbox" id="PopList" value="ME018" <%if InStr(1,str_PopList,"ME018",1)>0 then response.Write "checked"%>>
														修改
														<input name="PopList" type="checkbox" id="PopList" value="ME019" <%if InStr(1,str_PopList,"ME019",1)>0 then response.Write "checked"%>>
														删除</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_News" <%if InStr(1,str_PopList,"ME_News",1)>0 then response.Write "checked"%>>
											公告管理</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME021" <%if InStr(1,str_PopList,"ME021",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Form" <%if InStr(1,str_PopList,"ME_Form",1)>0 then response.Write "checked"%>>
											社群管理</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME022" <%if InStr(1,str_PopList,"ME022",1)>0 then response.Write "checked"%>>
														新建
														<input name="PopList" type="checkbox" id="PopList" value="ME023" <%if InStr(1,str_PopList,"ME023",1)>0 then response.Write "checked"%>>
														修改
														<input name="PopList" type="checkbox" id="PopList" value="ME024" <%if InStr(1,str_PopList,"ME024",1)>0 then response.Write "checked"%>>
														删除
														<input name="PopList" type="checkbox" id="PopList" value="ME025"  <%if InStr(1,str_PopList,"ME025",1)>0 then response.Write "checked"%>>
														锁定</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_HY" <%if InStr(1,str_PopList,"ME_HY",1)>0 then response.Write "checked"%>>
											行业管理</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME026" <%if InStr(1,str_PopList,"ME026",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_award" <%if InStr(1,str_PopList,"ME_award",1)>0 then response.Write "checked"%>>
											抽奖管理</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME027" <%if InStr(1,str_PopList,"ME027",1)>0 then response.Write "checked"%>>
														新建
														<input name="PopList" type="checkbox" id="PopList" value="ME028" <%if InStr(1,str_PopList,"ME028",1)>0 then response.Write "checked"%>>
														修改
														<input name="PopList" type="checkbox" id="PopList" value="ME029" <%if InStr(1,str_PopList,"ME029",1)>0 then response.Write "checked"%>>
														删除
														<input name="PopList" type="checkbox" id="PopList" value="ME030"  <%if InStr(1,str_PopList,"ME030",1)>0 then response.Write "checked"%>>
														开奖</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Order" <%if InStr(1,str_PopList,"ME_Order",1)>0 then response.Write "checked"%>>
											定单管理(在线支付)</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME031" <%if InStr(1,str_PopList,"ME031",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Mproducts" <%if InStr(1,str_PopList,"ME_Mproducts",1)>0 then response.Write "checked"%>>
											添加会员商品</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME032" <%if InStr(1,str_PopList,"ME032",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Horder" <%if InStr(1,str_PopList,"ME_Horder",1)>0 then response.Write "checked"%>>
											交易明晰</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME033" <%if InStr(1,str_PopList,"ME033",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_GUser" <%if InStr(1,str_PopList,"ME_GUser",1)>0 then response.Write "checked"%>>
											会员组</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME034" <%if InStr(1,str_PopList,"ME034",1)>0 then response.Write "checked"%>>
														新建
														<input name="PopList" type="checkbox" id="PopList" value="ME035" <%if InStr(1,str_PopList,"ME035",1)>0 then response.Write "checked"%>>
														修改
														<input name="PopList" type="checkbox" id="PopList" value="ME036" <%if InStr(1,str_PopList,"ME036",1)>0 then response.Write "checked"%>>
														删除</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Jubao" <%if InStr(1,str_PopList,"ME_Jubao",1)>0 then response.Write "checked"%>>
											举报管理</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME037" <%if InStr(1,str_PopList,"ME037",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Review" <%if InStr(1,str_PopList,"ME_Review",1)>0 then response.Write "checked"%>>
											评论管理</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME038" <%if InStr(1,str_PopList,"ME038",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Log" <%if InStr(1,str_PopList,"ME_Log",1)>0 then response.Write "checked"%>>
											日志管理</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME039" <%if InStr(1,str_PopList,"ME039",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Photo" <%if InStr(1,str_PopList,"ME_Photo",1)>0 then response.Write "checked"%>>
											相册管理</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME040" <%if InStr(1,str_PopList,"ME040",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Param" <%if InStr(1,str_PopList,"ME_Param",1)>0 then response.Write "checked"%>>
											参数设置</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME041" <%if InStr(1,str_PopList,"ME041",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Pay" <%if InStr(1,str_PopList,"ME_Pay",1)>0 then response.Write "checked"%>>
											在线支付</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="ME042" <%if InStr(1,str_PopList,"ME042",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
								</table>
							</div></td>
					</tr>
					<%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBAP")=1 then
		  %>
					<tr  class="hback_1">
						<td colspan="2" class="hback_1"><div align="left"><strong>
								<input name="PopList_AP" type="checkbox" id="PopList_AP" value="AP_Pop"  onClick="SwitchPopType('AP_Pop');" <%if InStr(1,str_PopList,"AP_Pop",1)>0 then response.Write "checked"%>>
								</strong>AP招聘求职系统</div></td>
					</tr>
					<tr class="hback">
						<td colspan="2" class="hback"><div align="left">
								<table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"AP_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="AP_ID">
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="AP_Param" <%if InStr(1,str_PopList,"AP_Param",1)>0 then response.Write "checked"%>>
												系统参数设置</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="AP001" <%if InStr(1,str_PopList,"AP001",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="AP_Province" <%if InStr(1,str_PopList,"AP_Province",1)>0 then response.Write "checked"%>>
												省份设置</div></td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="AP002" <%if InStr(1,str_PopList,"AP002",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="AP_city" <%if InStr(1,str_PopList,"AP_city",1)>0 then response.Write "checked"%>>
											城市设置</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="AP003" <%if InStr(1,str_PopList,"AP003",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="AP_Search" <%if InStr(1,str_PopList,"AP_Search",1)>0 then response.Write "checked"%>>
												会员记录查询</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="AP004" <%if InStr(1,str_PopList,"AP004",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="AP_check" <%if InStr(1,str_PopList,"AP_check",1)>0 then response.Write "checked"%>>
												注册信息审核</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="AP005" <%if InStr(1,str_PopList,"AP005",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
								</table>
							</div></td>
					</tr>
					<%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBSD")=1 then
		  %>
					<tr  class="hback_1">
						<td colspan="2" class="hback_1"><div align="left"><strong>
								<input name="PopList_SD" type="checkbox" id="PopList_SD" value="SD_Pop"  onClick="SwitchPopType('SD_Pop');" <%if InStr(1,str_PopList,"SD_Pop",1)>0 then response.Write "checked"%>>
								</strong>SD供求系统</div></td>
					</tr>
					<tr class="hback">
						<td colspan="2" class="hback"><div align="left">
								<table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"SD_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="SD_ID">
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="SD_List" <%if InStr(1,str_PopList,"SD_List",1)>0 then response.Write "checked"%>>
												供求信息</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="SD001" <%if InStr(1,str_PopList,"SD001",1)>0 then response.Write "checked"%>>
														新建
														<input name="PopList" type="checkbox" id="PopList" value="SD002" <%if InStr(1,str_PopList,"SD002",1)>0 then response.Write "checked"%>>
														修改
														<input name="PopList" type="checkbox" id="PopList" value="SD003" <%if InStr(1,str_PopList,"SD003",1)>0 then response.Write "checked"%>>
														删除
														<input name="PopList" type="checkbox" id="PopList" value="SD004" <%if InStr(1,str_PopList,"SD004",1)>0 then response.Write "checked"%>>
														审核</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="SD_Class" <%if InStr(1,str_PopList,"SD_Class",1)>0 then response.Write "checked"%>>
												类别栏目</div></td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="SD005" <%if InStr(1,str_PopList,"SD005",1)>0 then response.Write "checked"%>>
														新建
														<input name="PopList" type="checkbox" id="PopList" value="SD006" <%if InStr(1,str_PopList,"SD006",1)>0 then response.Write "checked"%>>
														修改
														<input name="PopList" type="checkbox" id="PopList" value="SD007" <%if InStr(1,str_PopList,"SD007",1)>0 then response.Write "checked"%>>
														删除</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="AP_area" <%if InStr(1,str_PopList,"AP_area",1)>0 then response.Write "checked"%>>
											区域管理</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="SD008" <%if InStr(1,str_PopList,"SD008",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="AP_param" <%if InStr(1,str_PopList,"AP_param",1)>0 then response.Write "checked"%>>
												系统设置</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="SD009" <%if InStr(1,str_PopList,"SD009",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
								</table>
							</div></td>
					</tr>
					<%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBCS")=1 then
		  %>
					<tr  class="hback_1">
						<td colspan="2" class="hback_1"><div align="left"><strong>
								<input name="PopList_CS" type="checkbox" id="PopList_CS" value="CS_Pop"  onClick="SwitchPopType('CS_Pop');" <%if InStr(1,str_PopList,"CS_Pop",1)>0 then response.Write "checked"%>>
								</strong>CS采集系统</div></td>
					</tr>
					<tr class="hback">
						<td colspan="2" class="hback"><div align="left">
								<table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"CS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="CS_ID">
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="CS_site" <%if InStr(1,str_PopList,"CS_site",1)>0 then response.Write "checked"%>>
												设置站点</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="CS001" <%if InStr(1,str_PopList,"CS001",1)>0 then response.Write "checked"%>>
														修改/新建
														<input name="PopList" type="checkbox" id="PopList" value="CS_collect" <%if InStr(1,str_PopList,"CS_collect",1)>0 then response.Write "checked"%>>
														采集</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="CS_Ink" <%if InStr(1,str_PopList,"CS_Ink",1)>0 then response.Write "checked"%>>
											采入数据库</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="CS003" <%if InStr(1,str_PopList,"CS003",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
								</table>
							</div></td>
					</tr>
					<%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBSS")=1 then
		  %>
					<tr  class="hback_1">
						<td colspan="2" class="hback_1"><div align="left"><strong>
								<input name="PopList_SS" type="checkbox" id="PopList_SS" value="SS_Pop"  onClick="SwitchPopType('SS_Pop');"  <%if InStr(1,str_PopList,"SS_Pop",1)>0 then response.Write "checked"%>>
								</strong>SS站点统计</div></td>
					</tr>
					<tr class="hback">
						<td colspan="2" class="hback"><div align="left">
								<table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"SS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="SS_ID">
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="SS_site" <%if InStr(1,str_PopList,"SS_site",1)>0 then response.Write "checked"%>>
												站点统计</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="SS001" <%if InStr(1,str_PopList,"SS001",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
								</table>
							</div></td>
					</tr>
					<%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBHS")=1 then
		  %>
					<tr  class="hback_1">
						<td colspan="2" class="hback_1"><div align="left"><strong>
								<input name="PopList_HS" type="checkbox" id="PopList_HS" value="HS_Pop"  onClick="SwitchPopType('HS_Pop');" <%if InStr(1,str_PopList,"HS_Pop",1)>0 then response.Write "checked"%>>
								</strong>HS房产楼盘系统</div></td>
					</tr>
					<tr class="hback">
						<td colspan="2" class="hback"><div align="left">
								<table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"HS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="HS_ID">
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="HS_Loup" <%if InStr(1,str_PopList,"HS_Loup",1)>0 then response.Write "checked"%>>
												楼盘管理</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="HS001" <%if InStr(1,str_PopList,"HS001",1)>0 then response.Write "checked"%>>
														新建
														<input name="PopList" type="checkbox" id="PopList" value="HS002" <%if InStr(1,str_PopList,"HS002",1)>0 then response.Write "checked"%>>
														修改
														<input name="PopList" type="checkbox" id="PopList" value="HS003" <%if InStr(1,str_PopList,"HS003",1)>0 then response.Write "checked"%>>
														删除
														<input name="PopList" type="checkbox" id="PopList" value="HS004" <%if InStr(1,str_PopList,"HS004",1)>0 then response.Write "checked"%>>
														审核</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="HS_Ero" <%if InStr(1,str_PopList,"HS_Ero",1)>0 then response.Write "checked"%>>
												二手房</div></td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="HS005" <%if InStr(1,str_PopList,"HS005",1)>0 then response.Write "checked"%>>
														新建
														<input name="PopList" type="checkbox" id="PopList" value="HS006" <%if InStr(1,str_PopList,"HS006",1)>0 then response.Write "checked"%>>
														修改
														<input name="PopList" type="checkbox" id="PopList" value="HS007" <%if InStr(1,str_PopList,"HS007",1)>0 then response.Write "checked"%>>
														删除</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td class="hback"><input name="PopList" type="checkbox" id="PopList" value="HS_Zu" <%if InStr(1,str_PopList,"HS_Zu",1)>0 then response.Write "checked"%>>
											租赁信息</td>
										<td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="HS008" <%if InStr(1,str_PopList,"HS008",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="HS_param" <%if InStr(1,str_PopList,"HS_param",1)>0 then response.Write "checked"%>>
												系统设置</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="HS009" <%if InStr(1,str_PopList,"HS009",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="HS_FY" <%if InStr(1,str_PopList,"HS_FY",1)>0 then response.Write "checked"%>>
												房源审核</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="HS010" <%if InStr(1,str_PopList,"HS010",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="HS_Search" <%if InStr(1,str_PopList,"HS_Search",1)>0 then response.Write "checked"%>>
												查询统计</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="HS011" <%if InStr(1,str_PopList,"HS011",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="HS_CJ" <%if InStr(1,str_PopList,"HS_CJ",1)>0 then response.Write "checked"%>>
												厂商管理</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="HS012" <%if InStr(1,str_PopList,"HS012",1)>0 then response.Write "checked"%>>
														开启</td>
												</tr>
											</table></td>
									</tr>
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="HS_Other" <%if InStr(1,str_PopList,"HS_Other",1)>0 then response.Write "checked"%>>
												其他</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="HS013" <%if InStr(1,str_PopList,"HS013",1)>0 then response.Write "checked"%>>
														回收站管理
														<input name="PopList" type="checkbox" id="PopList" value="HS014" <%if InStr(1,str_PopList,"HS014",1)>0 then response.Write "checked"%>>
														清理过期房源
														<input name="PopList" type="checkbox" id="PopList" value="HS015" <%if InStr(1,str_PopList,"HS015",1)>0 then response.Write "checked"%>>
														捆绑模板</td>
												</tr>
											</table></td>
									</tr>
								</table>
							</div></td>
					</tr>
					<%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBVS")=1 then
		  %>
					<tr  class="hback_1">
						<td colspan="2" class="hback_1"><div align="left"><strong>
								<input name="PopList_VS" type="checkbox" id="PopList_VS" value="VS_Pop"  onClick="SwitchPopType('VS_Pop');" <%if InStr(1,str_PopList,"VS_Pop",1)>0 then response.Write "checked"%>>
								</strong>VS投票管理</div></td>
					</tr>
					<tr class="hback">
						<td colspan="2" class="hback"><div align="left">
								<table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"VS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="VS_ID">
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="VS_site" <%if InStr(1,str_PopList,"VS_site",1)>0 then response.Write "checked"%>>
												开启 </div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="VS001" <%if InStr(1,str_PopList,"VS001",1)>0 then response.Write "checked"%>>
														参数设置
														<input name="PopList" type="checkbox" id="PopList" value="VS002" <%if InStr(1,str_PopList,"VS002",1)>0 then response.Write "checked"%>>
														管理
														<input name="PopList" type="checkbox" id="PopList" value="VS003" <%if InStr(1,str_PopList,"VS003",1)>0 then response.Write "checked"%>>
														查看</td>
												</tr>
											</table></td>
									</tr>
								</table>
							</div></td>
					</tr>
					<%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBAS")=1 then
		  %>
					<tr  class="hback_1">
						<td colspan="2" class="hback_1"><div align="left"><strong>
								<input name="PopList_AS" type="checkbox" id="PopList_AS" value="AS_Pop"  onClick="SwitchPopType('AS_Pop');" <%if InStr(1,str_PopList,"AS_Pop",1)>0 then response.Write "checked"%>>
								</strong>AS广告管理系统</div></td>
					</tr>
					<tr class="hback">
						<td colspan="2" class="hback"><div align="left">
								<table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"AS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="AS_ID">
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="AS_site" <%if InStr(1,str_PopList,"AS_site",1)>0 then response.Write "checked"%>>
												开启</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="AS001" <%if InStr(1,str_PopList,"AS001",1)>0 then response.Write "checked"%>>
														管理
														<input name="PopList" type="checkbox" id="PopList" value="AS002" <%if InStr(1,str_PopList,"AS002",1)>0 then response.Write "checked"%>>
														统计
														<input name="PopList" type="checkbox" id="PopList" value="AS003" <%if InStr(1,str_PopList,"AS003",1)>0 then response.Write "checked"%>>
														分类</td>
												</tr>
											</table></td>
									</tr>
								</table>
							</div></td>
					</tr>
					<%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBWS")=1 then
		  %>
					<tr  class="hback_1">
						<td colspan="2" class="hback_1"><div align="left"><strong>
								<input name="PopList_WS" type="checkbox" id="PopList_WS" value="WS_Pop"  onClick="SwitchPopType('WS_Pop');" <%if InStr(1,str_PopList,"WS_Pop",1)>0 then response.Write "checked"%>>
								</strong>WS留言系统</div></td>
					</tr>
					<tr class="hback">
						<td colspan="2" class="hback"><div align="left">
								<table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"WS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="WS_ID">
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="WS_site" <%if InStr(1,str_PopList,"WS_site",1)>0 then response.Write "checked"%>>
												开启</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="WS001" <%if InStr(1,str_PopList,"WS001",1)>0 then response.Write "checked"%>>
														查看
														<input name="PopList" type="checkbox" id="PopList" value="WS002" <%if InStr(1,str_PopList,"WS002",1)>0 then response.Write "checked"%>>
														管理</td>
												</tr>
											</table></td>
									</tr>
								</table>
							</div></td>
					</tr>
					<%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBFL")=1 then
		  %>
					<tr  class="hback_1">
						<td colspan="2" class="hback_1"><div align="left"><strong>
								<input name="PopList_FL" type="checkbox" id="PopList_FL" value="FL_Pop"  onClick="SwitchPopType('FL_Pop');" <%if InStr(1,str_PopList,"FL_Pop",1)>0 then response.Write "checked"%>>
								</strong>FL友情联接系统</div></td>
					</tr>
					<tr class="hback">
						<td colspan="2" class="hback"><div align="left">
								<table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"FL_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="FL_ID">
									<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
										<td width="14%" class="hback"><div align="left">
												<input name="PopList" type="checkbox" id="PopList" value="FL_site" <%if InStr(1,str_PopList,"FL_site",1)>0 then response.Write "checked"%>>
												开启</div></td>
										<td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
												<tr>
													<td><input name="PopList" type="checkbox" id="PopList" value="FL001" <%if InStr(1,str_PopList,"FL001",1)>0 then response.Write "checked"%>>
														查看
														<input name="PopList" type="checkbox" id="PopList" value="FL002" <%if InStr(1,str_PopList,"FL002",1)>0 then response.Write "checked"%>>
														管理</td>
												</tr>
											</table></td>
									</tr>
								</table>
							</div></td>
					</tr>
					<%end if%>
				</table></td>
		</tr>
		<tr >
			<td class="hback"><div align="left"></div>
				<input type="submit" name="Submit" value="确定设置权限">
				<input type="reset" name="Submit2" value="重置"></td>
		</tr>
	</form>
</table>
</body>
</html>
<%
Set Conn = Nothing
%>