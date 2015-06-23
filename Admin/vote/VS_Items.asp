<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/VS_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/md5.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<%'Copyright (c) 2006 Foosun Inc. 
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Conn,VS_Rs,VS_Sql,VS_Rs1 ,sRootDir,str_CurrPath
Dim AutoDelete,Months
Dim Temp_Admin_Is_Super,Temp_Admin_FilesTF,Temp_Admin_Name
Temp_Admin_Name = Session("Admin_Name")
Temp_Admin_Is_Super = Session("Admin_Is_Super")
Temp_Admin_FilesTF = Session("Admin_FilesTF")
MF_Default_Conn 
MF_Session_TF
if not MF_Check_Pop_TF("VS003") then Err_Show

if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""
If Temp_Admin_Is_Super = 1 then
	str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
Else
	If Temp_Admin_FilesTF = 0 Then
		str_CurrPath = Replace(sRootDir &"/"&G_UP_FILES_DIR&"/adminfiles/"&UCase(md5(Temp_Admin_Name,16)),"//","/")
	Else
		str_CurrPath = sRootDir &"/"&G_UP_FILES_DIR
	End If	
End if


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
	Dim This_Fun_Rs
	if instr(This_Fun_Sql," FS_ME_")>0 then 
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	else
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	end if			
	if not This_Fun_Rs.eof then 
		Get_OtherTable_Value = This_Fun_Rs(0)
	else
		Get_OtherTable_Value = ""
	end if
	if Err.Number>0 then 
		Err.Clear
		response.Redirect("../error.asp?ErrCodes=<li>Get_OtherTable_Value未能得到相关数据。错误描述："&Err.Type&"</li>") : response.End()
	end if
	set This_Fun_Rs=nothing 
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
	Dim Get_Html,This_Fun_Sql,ii,db_ii,Str_Tmp,Arr_Tmp,New_Search_Str,Req_Str,regxp
	Str_Tmp = "IID,TID,ItemName,ItemValue,ItemMode,PicSrc,DisColor,VoteCount,ItemDetail"
	This_Fun_Sql = "select "&Str_Tmp&" from FS_VS_Items"
	if request.QueryString("Act")="SearchGo" then 
		Arr_Tmp = split(Str_Tmp,",")
		for each Str_Tmp in Arr_Tmp
			Req_Str = NoSqlHack(Trim(request(Str_Tmp)))
			if Req_Str<>"" then 				
				select case Str_Tmp
					case "IID","TID","ItemMode","VoteCount"
					''数字,日期
						regxp = "|<|>|=|<=|>=|<>|"
						if instr(regxp,"|"&left(Req_Str,1)&"|")>0 or instr(regxp,"|"&left(Req_Str,2)&"|")>0 then 
							New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" "& Req_Str
						elseif instr(Req_Str,"*")>0 then 
							if left(Req_Str,1)="*" then Req_Str = "%"&mid(Req_Str,2)
							if right(Req_Str,1)="*" then Req_Str = mid(Req_Str,1,len(Req_Str) - 1) & "%"							
							New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" like '"& Req_Str &"'"							
						else	
							New_Search_Str = and_where( New_Search_Str ) & Str_Tmp &" = "& Req_Str
						end if		
					case else
					''字符
						New_Search_Str = and_where(New_Search_Str) & Search_TextArr(Req_Str,Str_Tmp,"")
				end select 		
			end if
		next
		if New_Search_Str<>"" then This_Fun_Sql = and_where(This_Fun_Sql) & replace(New_Search_Str," where ","")
	end if
	Str_Tmp = ""
	'response.Write(This_Fun_Sql)
	if Add_Sql<>"" then This_Fun_Sql = and_where(This_Fun_Sql) &" "& Decrypt(Add_Sql)
	if orderby<>"" then This_Fun_Sql = This_Fun_Sql &"  Order By "& replace(orderby,"csed"," Desc")
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
		Get_Html = Get_Html & "<td align=""center""><a href=""VS_Items.asp?Act=Edit&IID="&VS_Rs("IID")&""" title=""点击修改或查看详细"">"&VS_Rs("ItemName")&"</a></td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" style=""cursor:hand"" onclick=""javascript:if(TD_U_"&VS_Rs("IID")&".style.display=='') TD_U_"&VS_Rs("IID")&".style.display='none'; else {TD_U_"&VS_Rs("IID")&".style.display='';ReImgSize('TD_Img_"&VS_Rs("IID")&"');}"" title='点击查看详细情况'>"&Get_OtherTable_Value("select Theme from FS_VS_Theme where TID ="&VS_Rs("TID"))&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Replacestr(VS_Rs("ItemMode"),"1:文字描述模式,2:<span class=tx>自主填写模式</span>,3:<b>图片模式</b>")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&VS_Rs("PicSrc")&"</td>" & vbcrlf
		Get_Html = Get_Html & Replacestr(VS_Rs("DisColor"),":<td align=""center"">无</td>,else:<td align=""center"" bgcolor="""&VS_Rs("DisColor")&""">"&VS_Rs("DisColor")&"</td>") & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&VS_Rs("VoteCount")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"">"&Replacestr(VS_Rs("ItemDetail"),":无,else:"&left(VS_Rs("ItemDetail"),80)&"...")&"</td>" & vbcrlf
		Get_Html = Get_Html & "<td align=""center"" class=""ischeck""><input type=""checkbox"" name=""IID"" id=""IID"" value="""&VS_Rs("IID")&""" /></td>" & vbcrlf
		Get_Html = Get_Html & "</tr>" & vbcrlf
		''++++++++++++++++++++++++++++++++++++++点开时显示详细信息。
		set VS_Rs1 = Conn.execute("select TID,CID,Theme,Type,DisMode,StartDate,EndDate,ItemMOde from FS_VS_Theme where TID ="&VS_Rs("TID"))
		Get_Html = Get_Html & "<tr class=""hback"" id=""TD_U_"& VS_Rs("IID") &""" style=""display:'none'""><td colspan=20>" & vbcrlf
		Get_Html = Get_Html & "<table width=""100%"" height=""30"" border=""0"" cellspacing=""1"" cellpadding=""2"" class=""table"">" & vbcrlf 
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td colspan=3>投票分类:"&Get_OtherTable_Value("select ClassName from FS_VS_Class where CID ="&VS_Rs1("CID"))& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td colspan=3>分类描述:"&Get_OtherTable_Value("select Description from FS_VS_Class where CID ="&VS_Rs1("CID"))& "</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>调查主题:"&VS_Rs1("Theme")&"</td><td>项目类型:"&Replacestr(VS_Rs1("Type"),"1:单选,2:多选,3:多步")&"</td><td>显示方式:"&Replacestr(VS_Rs1("DisMode"),"1:直方图,2:饼图,3:折线图")&"</td></tr>" & vbcrlf
		Get_Html = Get_Html & "<tr class=""hback"">" & vbcrlf &"<td>排列方式:"&Replacestr(VS_Rs1("ItemMOde"),":无,0:横向排列,1:1选项/行(纵向),2:2选项/行,3:3选项/行,4:4选项/行,5:5选项/行,6:6选项/行,7:7选项/行,8:8选项/行,9:9选项/行,10:10选项/行,11:11选项/行,12:12选项/行")&"</td><td>开始时间:"&VS_Rs1("StartDate")&"</td><td>结束时间:"&VS_Rs1("EndDate")&"</td></tr>" & vbcrlf
		Get_Html = Get_Html & "</table>" & vbcrlf
		Get_Html = Get_Html &"</td></tr>" & vbcrlf
		VS_Rs1.close
		''+++++++++++++++++++++++++++++++++++++++		
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
	if request.QueryString("IID")<>"" then 
		Conn.execute("Delete from FS_VS_Items where IID = "&CintStr(request.QueryString("IID")))
		Conn.execute("Delete from FS_VS_Items_Result where IID = "&CintStr(request.QueryString("IID")))
	else
		Str_Tmp = FormatIntArr(request.form("IID"))
		if Str_Tmp="" then response.Redirect("../error.asp?ErrCodes=<li>你必须至少选择一个进行删除。</li>"):response.End()
		
		Conn.execute("Delete from FS_VS_Items where IID in ("&Str_Tmp&")")
		Conn.execute("Delete from FS_VS_Items_Result where IID in ("&Str_Tmp&")")
	end if
	response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Vote/VS_Items.asp?Act=View" )&"&ErrCodes=<li>恭喜，删除成功。</li>")
End Sub
''================================================================
Sub Save()
	Dim Str_Tmp,Arr_Tmp,IID
	Str_Tmp = "TID,ItemName,ItemValue,ItemMode,PicSrc,DisColor,VoteCount,ItemDetail"
	Arr_Tmp = split(Str_Tmp,",")
	IID = NoSqlHack(request.Form("IID"))	
	if not isnumeric(IID) or not IID<>"" then IID = 0
	VS_Sql = "select "&Str_Tmp&"  from FS_VS_Items  where IID = "&CintStr(IID)
	Set VS_Rs = CreateObject(G_FS_RS)
	VS_Rs.Open VS_Sql,Conn,3,3
	if not VS_Rs.eof then 
	''修改
		for each Str_Tmp in Arr_Tmp
			VS_Rs(Str_Tmp) = NoSqlHack(request.Form(Str_Tmp))
		next
		VS_Rs.update
		VS_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Vote/VS_Items.asp?Act=Edit&IID="&IID )&"&ErrCodes=<li>恭喜，修改成功。</li>")
	else
	''新增
		if Conn.execute("Select Count(*) from FS_VS_Items where ItemName='"&NoSqlHack(request.Form("ItemName"))&"' and TID = "&NoSqlHack(request.Form("TID")))(0)>0 then 
			response.Redirect("../error.asp?ErrCodes=<li>你提交的数据已经存在，属于重复提交，请更换关键字。</li>"):response.End()
		end if
		VS_Rs.addnew
		for each Str_Tmp in Arr_Tmp
			'response.Write(Str_Tmp&":"&NoSqlHack(request.Form(Str_Tmp))&"<br>")
			VS_Rs(Str_Tmp) = NoSqlHack(request.Form(Str_Tmp))
		next
		'response.End()
		VS_Rs.update
		VS_Rs.close
		response.Redirect("../Success.asp?ErrorUrl="&server.URLEncode( "Vote/VS_Items.asp?Act=Add&VoteCount="&request.form("VoteCount")&"&TID="&request.form("TID")&"&ItemValue="&request.form("ItemValue") ) &"&ErrCodes=<li>恭喜，新增成功。</li>")
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
<script language="JavaScript" src="../../FS_Inc/coolWindowsCalendar.js"></script>
<script language="JavaScript">
<!--
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
<iframe width="260" height="165" id="colorPalette" src="../CommPages/selcolor.htm" style="visibility:hidden; position: absolute;border:1px gray solid" frameborder="0" scrolling="no" ></iframe>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
     <tr  class="hback"> 
            
    <td colspan="10" align="left" class="xingmu" >投票选项管理</td>
	</tr>
  <tr  class="hback"> 
    <td colspan="10" height="25">
	 <a href="VS_Items.asp">管理首页</a> | <a href="VS_Items.asp?Act=Add">新增</a> | <a href="VS_Items.asp?Act=Search" title="数字和日期型的字段,支持<=<>=><>等等运算符号如:查过期天数>2 ; 其它类型支持 A B ,A* *B ,*A* *B* ,AB等模式.">查询</a>	
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
Sub View
if not MF_Check_Pop_TF("VS003") then Err_Show
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
	<form name="form1" id="form1" method="post" action="?Act=Del">
   <tr  class="hback"> 
      <td align="center" class="xingmu"><a href="javascript:OrderByName('ItemName')" class="sd"><b>选项描述</b></a> 
        <span id="Show_Oder_ItemName"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('TID')" class="sd"><b>所属调查</b></a> 
        <span id="Show_Oder_TID"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('ItemMode')" class="sd"><b>选项模式</b></a> 
        <span id="Show_Oder_ItemMode"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('PicSrc')" class="sd"><b>图片位置</b></a> 
        <span id="Show_Oder_PicSrc"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('DisColor')" class="sd"><b>显示颜色</b></a> 
        <span id="Show_Oder_DisColor"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('VoteCount')" class="sd"><b>票数</b></a> 
        <span id="Show_Oder_VoteCount"></span></td>
      <td align="center" class="xingmu" ><a href="javascript:OrderByName('ItemDetail')" class="sd"><b>选项说明</b></a> 
        <span id="Show_Oder_ItemDetail"></span></td>
      <td width="2%" align="center" class="xingmu"><input name="ischeck" type="checkbox" value="checkbox" onClick="selectAll(this.form)" /></td>
    </tr>
    <%
		response.Write( Get_While_Info( request.QueryString("Add_Sql"),request.QueryString("filterorderby") ) )
	%>
   </form>	
</table>
<%End Sub
Sub Add_Edit_Search()
Dim Bol_IsEdit,IID,TID,ItemValue,ItemMode,DisColor,VoteCount
Bol_IsEdit = false
if request.QueryString("Act")="Edit" then
if not MF_Check_Pop_TF("VS002") then Err_Show
	IID = request.QueryString("IID")
	if IID="" then response.Redirect("../error.asp?ErrorUrl=&ErrCodes=<li>必要的IID没有提供。</li>") : response.End()
	VS_Sql = "select IID,TID,ItemName,ItemValue,ItemMode,PicSrc,DisColor,VoteCount,ItemDetail from FS_VS_Items where IID = "& CintStr(IID)
	Set VS_Rs	= CreateObject(G_FS_RS)
	VS_Rs.Open VS_Sql,Conn,1,1
	if not VS_Rs.eof then 
		Bol_IsEdit = True
		TID = VS_Rs("TID")
		ItemValue = VS_Rs("ItemValue")
		ItemMode = VS_Rs("ItemMode")
		DisColor = VS_Rs("DisColor")
		VoteCount = VS_Rs("VoteCount")
	end if
elseif request.QueryString("Act") = "Add" then 
	if not MF_Check_Pop_TF("VS002") then Err_Show
	TID = NoSqlHack(request.QueryString("TID"))
	ItemValue = NoSqlHack(request.QueryString("ItemValue"))
	if ItemValue = "" then	ItemValue = "1-9"
	ItemMode = 1
	DisColor = ""
	VoteCount = NoSqlHack(request.QueryString("VoteCount"))
	if VoteCount = "" then 
		randomize		
		VoteCount = CStr(Int((99* Rnd) + 1))
	end if	
end if
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" id="form1" method="post" <%if request.QueryString("Act")<>"Search" then response.Write("action=""?Act=Save"" onsubmit=""return chkinput();""") else response.Write("action=""?Act=SearchGo""") end if%>>
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" >投票选项信息<%if Bol_IsEdit then	 response.Write("<input type=""hidden"" name=""IID"" id=""IID"" value="""&VS_Rs("IID")&""">") end if%></td>
	</tr>
<%if request.QueryString("Act")="Search" then %>

    <tr class="hback"> 
      <td width="100" align="right">自动编号</td>
      <td>
	  	<input type="text" name="IID" id="IID" size="11" maxlength="11">
      </td>
    </tr>
<%end if%>
    <tr  class="hback"> 
      <td align="right">所属投票</td>
      <td>
		<select name="TID" id="TID" onChange="Do.these('TID',function(){return isEmpty('TID','TID_Alt')})">
		<option value="">请选择</option>
		<%=Get_FildValue_List("select TID,'分类:'+ClassName+'--主题:'+Theme from FS_VS_Theme A,FS_VS_Class B where A.CID=B.CID",NoSqlHack(TID),1)%>
		</select>
		<span id="TID_Alt"></span>
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">选项描述</td>
      <td>
		<input type="text" name="ItemName" id="ItemName" size="50" maxlength="100" onFocus="Do.these('ItemName',function(){return isEmpty('ItemName','ItemName_Alt')})" onKeyUp="Do.these('ItemName',function(){return isEmpty('ItemName','ItemName_Alt')})" value="<%if Bol_IsEdit then response.Write(VS_Rs("ItemName")) end if%>">
		<span id="ItemName_Alt"></span>		
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">项目符号</td>
      <td>
		<select name="ItemValue" id="ItemValue">
		<%=PrintOption(ItemValue,":请选择,A-Z:A-Z,a-z:a-z,1-9:1-9,·:·,else:"&ItemValue)%>
		</select>
		<span  class="tx">A-Z,a-z,1-9或其它不递增的符号&nbsp;</span>
		<span id="ItemValue_Alt"></span>		
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">选项模式</td>
      <td> <select name="ItemMode" id="ItemMode" onChange="Do.these('ItemMode',function(){return isEmpty('ItemMode','ItemMode_Alt')}); this.options[this.selectedIndex].value=='3'?PicSrc.disabled=false:PicSrc.disabled=true;">
          <%=PrintOption(ItemMode,":请选择,1:文字描述模式,2:自主填写模式,3:图片模式")%> 
        </select>
		<span  class="tx">选择自主填写模式,文字后可以多个录入框,建议选择&nbsp;</span>
        <span id="ItemMode_Alt"></span></td>
    </tr>
    <tr  class="hback"> 
      <td align="right">图片位置</td>
      <td>
		<input type="text" name="PicSrc" id="PicSrc" readonly="" size="50" maxlength="200" value="<%if Bol_IsEdit then response.Write(VS_Rs("PicSrc")) end if%>">
		<input type="button" name="bnt_ChoosePic_rowBettween"  value="选择图片" onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath=<%=str_CurrPath %>',500,300,window,document.form1.PicSrc);">
		<span  class="tx">图片位置(针对图片模式而言)&nbsp;</span>
		<span id="PicSrc_Alt"></span>		
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">显示颜色</td>
      <td>
		<input type="text" name="DisColor" id="DisColor" size="15" maxlength="7" <%if DisColor<>"" then response.Write("style=""background-color:"&DisColor&"""") end if%> value="<%=DisColor%>">
        <img src="../Images/rectNoColor.gif" width="18" height="17" border=0 align="absmiddle" id="TitleFontColor_Show" style="cursor:pointer;background-color:;" title="选取颜色!" onClick="GetColor(this,'DisColor');"> 
        <span  class="tx">统计时显示颜色如#FF0000&nbsp;</span> <span id="DisColor_Alt"></span>	
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">当前票数</td>
      <td>
		<input type="text" name="VoteCount" id="VoteCount" size="15" maxlength="5" onFocus="Do.these('VoteCount',function(){return isEmpty('VoteCount','VoteCount_Alt')&&isNumber('VoteCount','VoteCount_Alt','必须数字',false)})" onKeyUp="Do.these('VoteCount',function(){return isEmpty('VoteCount','VoteCount_Alt')&&isNumber('VoteCount','VoteCount_Alt','必须数字',false)})" value="<%=VoteCount%>">
		<span id="VoteCount_Alt"></span>		
	  </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">选项详细说明</td>
      <td>
		<textarea name="ItemDetail" cols="50" rows="15" id="ItemDetail"><%if Bol_IsEdit then response.Write(VS_Rs("ItemDetail")) end if%></textarea>
		<span id="ItemDetail_Alt"></span>		
	  </td>
    </tr>
   <tr class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit"value=" 确定提交 " onClick="ItemDetail.value=ItemDetail.value.substring(0,300);" /> 
              &nbsp; <input type="reset" value=" 重置 " />
  			  &nbsp; <input type="button" name="btn_todel" value=" 删除 " onClick="if(confirm('确定删除该项目吗？')) location='<%="VS_Items.asp?Act=Del&IID="&IID%>'">
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
function chkinput()
{
	return isEmpty('TID','TID_Alt') && isEmpty('ItemName','ItemName_Alt') && isEmpty('ItemMode','ItemMode_Alt') && isEmpty('VoteCount','VoteCount_Alt');
}
function getOffsetTop(elm) {
	var mOffsetTop = elm.offsetTop;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent){
		mOffsetTop += mOffsetParent.offsetTop;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetTop;
}
function getOffsetLeft(elm) {
	var mOffsetLeft = elm.offsetLeft;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent) {
		mOffsetLeft += mOffsetParent.offsetLeft;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetLeft;
}
function GetColor(img_val,input_val)
{
	var PaletteLeft,PaletteTop
	var obj = document.getElementById("colorPalette");
	ColorImg = img_val;
	ColorValue = document.getElementById(input_val);	
	if (obj){
		PaletteLeft = getOffsetLeft(ColorImg)
		PaletteTop = (getOffsetTop(ColorImg) + ColorImg.offsetHeight)
		if (PaletteLeft+150 > parseInt(document.body.clientWidth)) PaletteLeft = parseInt(event.clientX)-260;
		obj.style.left = PaletteLeft + "px";
		obj.style.top = PaletteTop + "px";
		if (obj.style.visibility=="hidden")
		{
			obj.style.visibility="visible";
		}else {
			obj.style.visibility="hidden";
		}
	}
}
function setColor(color)
{
	if(ColorImg.id=="FontColorShow"&&color=="#") color='#000000';
	if(ColorImg.id=="FontBgColorShow"&&color=="#") color='#FFFFFF';
	if (ColorValue){ColorValue.value = color.substr(1);}
	if (ColorImg && color.length>1){
		ColorImg.src='../Images/Rect.gif';
		ColorImg.style.backgroundColor = color;
	}else if(color=='#'){ ColorImg.src='../Images/rectNoColor.gif';}
	document.getElementById("colorPalette").style.visibility="hidden";
}
-->
</script>
<!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





