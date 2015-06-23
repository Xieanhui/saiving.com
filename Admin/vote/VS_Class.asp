<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/VS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%'Copyright (c) 2006 Foosun Inc. 
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,VS_Rs,VS_Sql
Dim AutoDelete,Months
MF_Default_Conn 
MF_Session_TF
if not MF_Check_Pop_TF("VS003") then Err_Show

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
''得到相关表的值。
Function Get_OtherTable_Value(This_Fun_Sql)
	Dim This_Fun_Rs,f_Tmpstr_
	f_Tmpstr_ = ""
	if instr(This_Fun_Sql," FS_ME_")>0 then 
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	else
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	end if			
	if not This_Fun_Rs.eof then 
		if instr(This_Fun_Sql," in ") > 0 then 
			do while not This_Fun_Rs.eof
				f_Tmpstr_ = f_Tmpstr_ & This_Fun_Rs(0) &","
				This_Fun_Rs.movenext
			loop		
		else	
			f_Tmpstr_ = This_Fun_Rs(0)
		end if	
	else
		f_Tmpstr_ = "0"
	end if
	if Err.Number>0 then 
		Err.Clear
		response.Redirect("../error.asp?ErrCodes=<li>Get_OtherTable_Value未能得到相关数据。错误描述："&Err.Description&"</li>") : response.End()
	end if
	set This_Fun_Rs=nothing 
	if right(f_Tmpstr_,1)="," then f_Tmpstr_ = left(f_Tmpstr_,len(f_Tmpstr_) - 1)
	Get_OtherTable_Value = f_Tmpstr_
End Function
Function Get_FildValue_List(This_Fun_Sql,EquValue,Get_Type)
'''This_Fun_Sql 传入sql语句,EquValue与数据库相同的值如果是<option>则加上selected,Get_Type=1为<option>
Dim Get_Html,This_Fun_Rs,Text
On Error Resume Next
set This_Fun_Rs = Conn.execute(This_Fun_Sql)
If Err.Number <> 0 then Err.clear : response.Redirect("../error.asp?ErrCodes=<li>抱歉,传入的Sql语句有问题.或表和字段不存在.</li>")
do while not This_Fun_Rs.eof 
	select case Get_Type
	  case 1
		''<option>		
		if instr(This_Fun_Sql,",") >0 then 
			Text = This_Fun_Rs(1)
		else
			Text = This_Fun_Rs(0)
		end if	
		if trim(EquValue) = trim(This_Fun_Rs(0)) then 
			Get_Html = Get_Html & "<option value="""&This_Fun_Rs(0)&"""  style=""color:#0000FF"" selected>"&Text&"</option>"&vbNewLine
		else
			Get_Html = Get_Html & "<option value="""&This_Fun_Rs(0)&""">"&Text&"</option>"&vbNewLine
		end if		
	  case else
		exit do : Get_FildValue_List = "Get_Type值传入错误" : exit Function 
	end select
	This_Fun_Rs.movenext
loop
This_Fun_Rs.close
Get_FildValue_List = Get_Html
End Function 

Function Get_While_Info(Add_Sql,orderby)
	Dim Get_Html,This_Fun_Sql,ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp
	Dim CID,ClassName,Description
	CID  = NoSqlHack(request("CID"))
	ClassName   = NoSqlHack(request("ClassName"))
	Description = NoSqlHack(request("Description"))
	Str_Tmp = "CID,ClassName,Description"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_VS_Class"
	if CID<>"" then New_Search_Str = and_where(New_Search_Str) & Search_TextArr(CID,"CID","")
	if ClassName<>"" then New_Search_Str = and_where(New_Search_Str) & Search_TextArr(ClassName,"ClassName","")
	if Description<>"" then New_Search_Str = and_where(New_Search_Str) & Search_TextArr(Description,"Description","")
	if New_Search_Str<>"" then This_Fun_Sql = and_where(This_Fun_Sql) & replace(New_Search_Str," where ","")
	Str_Tmp = ""
	if Add_Sql<>"" then This_Fun_Sql = and_where(This_Fun_Sql) &" "& Decrypt(Add_Sql)
	if orderby<>"" then This_Fun_Sql = This_Fun_Sql &"  Order By "& replace(orderby,"csed"," Desc")
	'response.Write(This_Fun_Sql)
	On Error Resume Next
	Set VS_Rs = CreateObject(G_FS_RS)
	VS_Rs.Open This_Fun_Sql,Conn,1,1	
	if Err<>0 then 
		Err.Clear
		response.Redirect("../error.asp?ErrCodes=<li>查询出错："&Err.Description&"</li><li>请检查字段类型是否匹配.</li>")
		response.End()
	end if
	IF VS_Rs.eof THEN
	 	response.Write("<tr class=""hback""><td colspan=15>暂无数据.</td></tr>") 
	else	
	VS_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("Page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>VS_Rs.PageCount Then cPageNo=VS_Rs.PageCount 
	VS_Rs.AbsolutePage=cPageNo
	
	  FOR int_Start=1 TO int_RPP 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&VS_Rs("CID")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center""><a href=""VS_Class.asp?Act=Edit&CID="&VS_Rs("CID")&""">"&VS_Rs("ClassName")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&left(VS_Rs("Description"),100)&"...</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" name=""CID"" id=""CID"" value="""&VS_Rs("CID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		VS_Rs.MoveNext
 		if VS_Rs.eof or VS_Rs.bof then exit for
      NEXT
	END IF
	Get_Html = Get_Html & "<tr class=""hback""><td colspan=20 align=""center"" class=""ischeck"">"& vbcrlf &"<table width=""100%"" border=0><tr><td height=30>" & vbcrlf
	Get_Html = Get_Html & fPageCount(VS_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf
	Get_Html = Get_Html & "</td><td align=right><input type=""submit"" name=""submit"" value="" 删除 "" onclick=""javascript:return confirm('确定要删除所选项目吗?');""></td>"
	Get_Html = Get_Html &"</tr></table>"&vbNewLine&"</td></tr>"
	Get_Html = Get_Html &"</td></tr>"
	VS_Rs.close
	Get_While_Info = Get_Html
End Function

Sub Del()
	if not MF_Check_Pop_TF("VS002") then Err_Show
	Dim Str_Tmp
	if request.QueryString("CID")<>"" then 
		Conn.execute("Delete from FS_VS_Class where CID = "&Cintstr(request.QueryString("CID")))
		Conn.execute("Delete from FS_VS_Items where TID in ( "&Get_OtherTable_Value("select TID from FS_VS_Theme where CID in ( "&FormatIntArr(request.QueryString("CID"))&")")&" )")
		Conn.execute("Delete from FS_VS_Items_Result where TID in ( "&Get_OtherTable_Value("select TID from FS_VS_Theme where CID in ( "&FormatIntArr(request.QueryString("CID"))&")")&" )")
		Conn.execute("Delete from FS_VS_Steps where TID in ( "&Get_OtherTable_Value("select TID from FS_VS_Theme where CID in ( "&FormatIntArr(request.QueryString("CID"))&")")&" )")
		Conn.execute("Delete from FS_VS_Theme where CID = "&Cintstr(request.QueryString("CID")))
	else
		Str_Tmp = FormatIntArr(request.form("CID"))
		if Str_Tmp="" then response.Redirect("../error.asp?ErrCodes=<li>你必须至少选择一个进行删除。</li>"):response.End()
		Str_Tmp = replace(Str_Tmp," ","")
		Conn.execute("Delete from FS_VS_Class where CID in ("&Str_Tmp&")")
		Conn.execute("Delete from FS_VS_Items where TID in ( "&Get_OtherTable_Value("select TID from FS_VS_Theme where CID in ( "&FormatIntArr(Str_Tmp)&")")&" )")
		Conn.execute("Delete from FS_VS_Items_Result where TID in ( "&Get_OtherTable_Value("select TID from FS_VS_Theme where CID in ( "&FormatIntArr(Str_Tmp)&")")&" )")
		Conn.execute("Delete from FS_VS_Steps where TID in ( "&Get_OtherTable_Value("select TID from FS_VS_Theme where CID in ( "&FormatIntArr(Str_Tmp)&")")&" )")
		Conn.execute("Delete from FS_VS_Theme where CID in ("&Str_Tmp&")")
	end if
	response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Vote/VS_Class.asp?Act=View" )&"&ErrCodes=<li>恭喜，删除成功。</li>")
End Sub
''================================================================
Sub Save()
	if not MF_Check_Pop_TF("VS002") then Err_Show
	Dim Str_Tmp,Arr_Tmp,CID
	Str_Tmp = "ClassName,Description"
	Arr_Tmp = split(Str_Tmp,",")	
	CID = NoSqlHack(request.Form("CID"))
	if not isnumeric(CID) or not CID<>"" then CID = 0
	VS_Sql = "select "&Str_Tmp&"  from FS_VS_Class  where CID = "&CintStr(CID)
	Set VS_Rs = CreateObject(G_FS_RS)
	VS_Rs.Open VS_Sql,Conn,3,3
	if not VS_Rs.eof then 
	''修改
		for each Str_Tmp in Arr_Tmp
			VS_Rs(Str_Tmp) = NoSqlHack(request.Form(Str_Tmp))
		next
		VS_Rs.update
		VS_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Vote/VS_Class.asp?Act=Edit&CID="&CID )&"&ErrCodes=<li>恭喜，修改成功。</li>")
	else
	''新增
		if Conn.execute("Select Count(*) from FS_VS_Class where ClassName='"&NoSqlHack(request.Form("ClassName"))&"'")(0)>0 then 
			response.Redirect("../error.asp?ErrCodes=<li>你提交的数据已经存在，属于重复提交，请更换关键字。</li>"):response.End()
		end if
		VS_Rs.addnew
		for each Str_Tmp in Arr_Tmp
			'response.Write(Str_Tmp&":"&NoSqlHack(request.Form(Str_Tmp))&"<br>")
			VS_Rs(Str_Tmp) = NoSqlHack(request.Form(Str_Tmp))
		next
		VS_Rs.update
		VS_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Vote/VS_Class.asp?Act=Add&ClassName="&request.Form("ClassName") ) &"&ErrCodes=<li>恭喜，新增成功。</li>")
	end if
End Sub
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js"></script>
<script language="JavaScript" src="../../FS_Inc/CheckJs.js"></script>
<script language="JavaScript">
<!--
function chkinput()
{
	return isEmpty('Description','Description_Alt') && isEmpty('ClassName','ClassName_Alt');
}
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
-->
</script>
<head>
<body>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
     <tr  class="hback"> 
            
    <td colspan="10" align="left" class="xingmu" >投票分类管理</td>
	</tr>
  <tr  class="hback"> 
    <td colspan="10" height="25">
	 <a href="VS_Class.asp">管理首页</a> | <a href="VS_Class.asp?Act=Add">新增</a> | <a href="VS_Class.asp?Act=Search">查询</a>	
	</td>
  </tr>
</table>
<%
'******************************************************************
select case request.QueryString("Act")
	case "Add","Edit","Search"
		Add_Edit_Search
	case "View","SearchGo",""
		View
	case "Save"
		Save
	case "Del"
		Del
	case "OtherSet"
		OtherSet(request.QueryString("Sql"))
	case else
	response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>错误的参数传递。</li>") : response.End()
end select
'******************************************************************
Sub View()
if not MF_Check_Pop_TF("VS_site") then Err_Show
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
	<form name="form1" id="form1" method="post" action="?Act=Del">
   <tr  class="hback"> 
      <td width="10%" align="center" class="xingmu"><a href="javascript:OrderByName('CID')" class="sd"><b>ID</b></a> 
        <span id="Show_Oder_CID"></span></td>
      <td width="30%" align="center" class="xingmu" ><a href="javascript:OrderByName('ClassName')" class="sd"><b>类别名称</b></a> 
        <span id="Show_Oder_ClassName"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('Description')" class="sd"><b>描述</b></a> 
        <span id="Show_Oder_Description"></span></td>
      <td width="2%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_While_Info( request.QueryString("Add_Sql"),request.QueryString("filterorderby") ) )
	%>
   </form>	
</table>
<%End Sub
Sub Add_Edit_Search()
if not MF_Check_Pop_TF("VS_site") then Err_Show
Dim Bol_IsEdit,CID,ClassName
Bol_IsEdit = false
if request.QueryString("Act")="Edit" then
	CID = request.QueryString("CID")
	if CID="" then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>必要的CID没有提供。</li>") : response.End()
	VS_Sql = "select CID,ClassName,Description from FS_VS_Class where CID = "& CintStr(CID)
	Set VS_Rs	= CreateObject(G_FS_RS)
	VS_Rs.Open VS_Sql,Conn,1,1
	if not VS_Rs.eof then 
		Bol_IsEdit = True
		ClassName = VS_Rs("ClassName")
	end if
else
	
end if
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" id="form1" method="post" <%if request.QueryString("Act")<>"Search" then response.Write("action=""?Act=Save"" onsubmit=""return chkinput();""") else response.Write("action=""?Act=SearchGo"" onsubmit=""SearchAdd();""") end if%>>
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" >投票分类信息<%if Bol_IsEdit then	 response.Write("<input type=""hidden"" name=""CID"" id=""CID"" value="""&VS_Rs("CID")&""">") end if%></td>
	</tr>
<%if request.QueryString("Act")="Search" then %>

    <tr class="hback"> 
      <td width="100" align="right">自动编号</td>
      <td>
	  	<input type="text" name="CID" id="CID" size="11" maxlength="11">
      </td>
    </tr>
<%end if%>
    <tr  class="hback"> 
      <td align="right">类别名称</td>
      <td>
		<input type="text" name="ClassName" id="ClassName" size="30" maxlength="20" onFocus="Do.these('ClassName',function(){return isEmpty('ClassName','ClassName_Alt')})" onKeyUp="Do.these('ClassName',function(){return isEmpty('ClassName','ClassName_Alt')})" value="<%if Bol_IsEdit then response.Write(VS_Rs("ClassName")) end if%>">
		<span id="ClassName_Alt"></span>		
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">描述</td>
      <td>
		<input type="text" name="Description" id="Description" size="30" maxlength="40" value="<%if Bol_IsEdit then response.Write(VS_Rs("Description")) end if%>">
        <span id="Description_Alt"></span>
	  </td>
    </tr>
   <tr class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit"value=" 确定提交 "/> 
              &nbsp; <input type="reset" value=" 重置 " />
  			  &nbsp; <input type="button" name="btn_todel" value=" 删除 " onClick="if(confirm('确定删除该项目吗？')) location='<%="VS_Class.asp?Act=Del&CID="&CID%>'">
            </td>
          </tr>
        </table>
      </td>
    </tr>	
  </form>
</table>
<%
End Sub
set VS_Rs = Nothing
Conn.close
%>
<script language="JavaScript">
<!--//判断后将排序完善.字段名后面显示指示
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
-->
</script>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





