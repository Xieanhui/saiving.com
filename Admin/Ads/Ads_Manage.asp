<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="Cls_ads.asp"-->
<%
response.buffer=true	
Response.CacheControl = "no-cache"
Dim Conn,User_Conn
MF_Default_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("AS_site") then Err_Show
if not MF_Check_Pop_TF("AS001") then Err_Show
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_
Dim Page,cPageNo

int_RPP=5 '设置每页显示数目
int_showNumberLink_=8 '数字导航显示数目
showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#999999" '非热链接颜色
toF_="<font face=webdings title=""首页"">9</font>"  			'首页 
toP10_=" <font face=webdings title=""上十页"">7</font>"			'上十
toP1_=" <font face=webdings title=""上一页"">3</font>"			'上一
toN1_=" <font face=webdings title=""下一页"">4</font>"			'下一
toN10_=" <font face=webdings title=""下十页"">8</font>"			'下十
toL_="<font face=webdings title=""最后一页"">:</font>"			'尾页

Dim Ad_OP,ID,strShowErr,CheckAllID,i,TempID,str_ClassName,str_ClassMode,o_ClassInfo_Rs,str_ClassInfo,str_AdRemarks,Temp_C_ID
Dim o_Ad_Rs,str_Ad_Sql,str_TempAType,str_Temp_Lock,str_date,str_ClassType,str_Classtype_Sql,f_AdclassFlag,MyFile_Obj,G_Ads_FILES_DIR
G_Ads_FILES_DIR="/ads"
Dim Str_SysDir
Str_SysDir=""
if G_VIRTUAL_ROOT_DIR<>"" then
	Str_SysDir="/"&G_VIRTUAL_ROOT_DIR&"/Ads"
Else
	Str_SysDir="/Ads"
end if
str_ClassType=NoSqlHack(Request.queryString("ClassID"))
Ad_OP=Request.QueryString("Ad_OP")
ID=Request.QueryString("ID")
CheckAllID=Request.Form("Checkallbox")
Select Case Ad_OP
	Case "AdLock"
			Conn.execute("update FS_AD_Info set AdLock=1 where AdID="&CintStr(ID)&"")
			UpdateAdsJsContent CintStr(ID)
	Case "AdUnLock"
		Set f_AdclassFlag=Conn.execute("Select Lock From FS_AD_Class Where AdClassID=(Select AdClassID from FS_AD_Info where AdID="&CintStr(ID)&")")
		If Not f_AdclassFlag.Eof Then
			If f_AdclassFlag("Lock")=1 Then
				strShowErr = "<li>所属栏目已被锁定,要想解锁此广告,请先解锁此广告所在栏目</li>"
				Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
				Response.end
			Else
				Conn.execute("update FS_AD_Info set AdLock=0 where AdID="&CintStr(ID)&"")
				UpdateAdsJsContent CintStr(ID)
			End If
		Else
			Conn.execute("update FS_AD_Info set AdLock=0 where AdID="&CintStr(ID)&"")
			UpdateAdsJsContent CintStr(ID)
		End If
		Set f_AdclassFlag=Nothing
	Case "AdDelOne"
		Conn.execute("delete  from FS_AD_Info where AdID="&CintStr(ID)&"")
		Conn.execute("delete  from FS_AD_TxtInfo where AdID="&CintStr(ID)&"")
		Set MyFile_Obj=Server.CreateObject(G_FS_FSO)
		on error resume next
		If MyFile_Obj.FileExists(Server.MapPath(G_Ads_FILES_DIR)&"\"& CintStr(ID) &".js") then
			MyFile_Obj.DeleteFile(Server.MapPath(G_Ads_FILES_DIR)&"\"& CintStr(ID) &".js")
		End If
		Set MyFile_Obj=Nothing
		strShowErr = "<li>删除成功!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_Manage.asp?Page="&Request.QueryString("Page")&"")
		Response.end
	Case "P_Lock"
		CheckAllIDFLag("锁定")
		Conn.execute("update FS_AD_Info set AdLock=1 where AdID in (" & FormatIntArr(CheckAllID) & ")")
		Dim Ads_p_Lock_i,Ads_p_Lock_Arr
		If DelHeadAndEndDot(CheckAllID) <> "" Then
			Ads_p_Lock_Arr = Split(DelHeadAndEndDot(CheckAllID),",")
			For Ads_p_Lock_i = LBound(Ads_p_Lock_Arr) To UBound(Ads_p_Lock_Arr)
				If Ads_p_Lock_Arr(Ads_p_Lock_i) <> "" And IsNumeric(Ads_p_Lock_Arr(Ads_p_Lock_i)) Then
					UpdateAdsJsContent Clng(Ads_p_Lock_Arr(Ads_p_Lock_i))
				End if	
			Next
		End If
		strShowErr = "<li>批量锁定成功!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_Manage.asp?Page="&NoSqlHack(Request.QueryString("Page"))&"")
		Response.end
	Case "P_UnLock"
		Dim YesLockNum,temp_num
		YesLockNum=0
		CheckAllIDFLag("解锁")
		Temp_C_ID=Replace(CheckAllID," ","")
		CheckAllID=split(Temp_C_ID,",")
		For i=0 to Ubound(CheckAllID)
			Set f_AdclassFlag=Conn.execute("Select Lock From FS_AD_Class Where AdClassID In (Select AdClassID from FS_AD_Info where AdID = "&CintStr(CheckAllID(i))&")")
			If Not f_AdclassFlag.Eof Then
				If f_AdclassFlag("Lock")=0 Then
					YesLockNum=YesLockNum+1
					Conn.execute("update FS_AD_Info set AdLock=0 where AdID = "&CintStr(CheckAllID(i))&"")
					UpdateAdsJsContent CintStr(CheckAllID(i))
				End If
			Else
				YesLockNum=YesLockNum+1
				Conn.execute("update FS_AD_Info set AdLock=0 where AdID = "&CintStr(CheckAllID(i))&"")
				UpdateAdsJsContent CintStr(CheckAllID(i))
			End If
			Set f_AdclassFlag=Nothing
		Next
		If YesLockNum=0 Then
			temp_num=0
		Else
			temp_num=YesLockNum-1
		End if
		strShowErr = "<li>批量解锁成功!</li>"
		strShowErr = strShowErr&"<li>共选中<font color=red>"&Ubound(CheckAllID)&"</font>个广告</li>"
		strShowErr = strShowErr&"<li>共解锁<font color=red>"&temp_num&"</font>个广告</li>"
		If Ubound(CheckAllID)-YesLockNum+1<>0 Then
			strShowErr = strShowErr&"<li><font color=red>"&Ubound(CheckAllID)-YesLockNum+1&"</font>个广告解锁失败,原因：此广告栏目已锁定</li>"
		End If
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_Manage.asp?Page="&NoSqlHack(Request.QueryString("Page"))&"")
		Response.end
	Case "P_Del"
		on error resume next
		CheckAllIDFLag("删除")		
		Set MyFile_Obj=Server.CreateObject(G_FS_FSO)
		Temp_C_ID=CheckAllID
		CheckAllID=split(CheckAllID,",")
		For i=0 to Ubound(CheckAllID)
			TempID=TempID&CheckAllID(i)&","		
			If MyFile_Obj.FileExists(Server.MapPath(G_Ads_FILES_DIR)&"\"& CheckAllID(i) &".js") then
				MyFile_Obj.DeleteFile(Server.MapPath(G_Ads_FILES_DIR)&"\"& CheckAllID(i) &".js")
			End If
		Next
		'Response.Write("delete  from FS_AD_Info where AdID in ("&DelHeadAndEndDot(Temp_C_ID)&")")
		'Response.End()
		Conn.execute("delete  from FS_AD_Info where AdID in (" & FormatIntArr(Temp_C_ID) & ")")
		Conn.execute("delete  from FS_AD_TxtInfo where AdID in (" & FormatIntArr(Temp_C_ID) & ")")
		Set MyFile_Obj=Nothing
		strShowErr = "<li>批量删除成功!</li>"
		Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../Ads_Manage.asp?Page="&NoSqlHack(Request.QueryString("Page"))&"")
		Response.end
End Select
Sub CheckAllIDFLag(Showstr)
	If CheckAllID="" or IsNull(CheckAllID) Then
		strShowErr = "<li>请选择要"&Showstr&"的文件!</li>"
		Response.Redirect("lib/Error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
		Response.end
	End If
End Sub
if str_ClassType="" Then str_ClassType=0
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>广告管理___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script src="Public.js" language="JavaScript"></script>
</head>
<body>
<%
	Dim str_AdClass_Sql,o_AdClass_Rs,str_AdClass_str,str_Selected
	str_AdClass_Sql="Select AdClassID,AdClassName,Lock from FS_AD_Class"
	Set o_AdClass_Rs=Conn.execute(str_AdClass_Sql)

	If Not o_AdClass_Rs.Eof Then
		While Not o_AdClass_Rs.Eof
			str_Selected=""
			If CintStr(str_ClassType)=CintStr(o_AdClass_Rs("AdClassID")) Then
				str_Selected=" selected"
			End If
			str_AdClass_str=str_AdClass_str&"<option value="&o_AdClass_Rs("AdClassID")&str_Selected&">"&o_AdClass_Rs("AdClassName")&"</option>"			
		o_AdClass_Rs.MoveNext
		Wend			
	End If
	Set o_AdClass_Rs=Nothing	
%>
<form name="AdInfo" action="" method="post">
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback"> 
    <td class="xingmu">广告管理</td>
  </tr>
  <tr> 
      <td width="100%" height="18" class="hback"><div algin="Center"><a href="Ads_Add.asp?OpType=Add">添加广告</a> | <a href="javascript:P_Lock();">批量锁定</a> | <a href="javascript:P_UnLock();">批量解锁</a> | <a href="javascript:P_Del();">批量删除</a> | 
        <select name="ClassID" size="1" style="width:150px" onChange="javascript:location='Ads_Manage.asp?ClassID='+this.value;">
		  <option value="-1">查看所有栏目</option>
		  <%=str_AdClass_str%>
        </select>
        </div></td>
  </tr>
</table>
<%

If Cint(str_ClassType) >0 Then
	str_Classtype_Sql="Where AdClassID="&CintStr(str_ClassType)&""
Else
	str_Classtype_Sql=""
End If
str_Ad_Sql="Select AdID,AdName,AdType,AdAddDate,AdMaxClickNum,AdMaxShowNum,AdLock,AdRemarks,AdClassID from FS_AD_Info "&str_Classtype_Sql&" order by AdID Desc"
Set o_Ad_Rs= CreateObject(G_FS_RS)
o_Ad_Rs.Open str_Ad_Sql,Conn,1,1
If Not o_Ad_Rs.Eof Then
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="hback">
    <td width="20%" align="center" class="xingmu" height="43">广告名称</td>
    <td width="12%" align="center" class="xingmu" height="43">广告类型</td>
    <td width="8%" align="center" class="xingmu" height="43">添加时间</td>
    <td width="10%" align="center" class="xingmu" title="最大点击次数" height="43">最大点击数</td>
    <td width="10%" align="center" class="xingmu" title="最大显示次数" height="43">最大显示数</td>
    <td width="6%" align="center" class="xingmu" height="43">状态</td>
    <td width="34%" align="center" class="xingmu" height="43">操作</td>
  </tr>
<%
	o_Ad_Rs.PageSize=int_RPP
	cPageNo=NoSqlHack(Request.QueryString("page"))
	If cPageNo="" Then cPageNo = 1
	If not isnumeric(cPageNo) Then cPageNo = 1
	cPageNo = Clng(cPageNo)
	If cPageNo<=0 Then cPageNo=1
	If cPageNo>o_Ad_Rs.PageCount Then cPageNo=o_Ad_Rs.PageCount 
	o_Ad_Rs.AbsolutePage=cPageNo

	For int_Start=1 TO int_RPP  
		Select Case Clng(o_Ad_Rs("AdType"))
			Case 0
				str_TempAType="普通显示广告"
			Case 1
				str_TempAType="弹出新窗口"
			Case 2
				str_TempAType="打开新窗口"
			Case 3
				str_TempAType="渐隐消失"
			Case 4
				str_TempAType="网页对话框"
			Case 5
				str_TempAType="透明对话框"
			Case 6
				str_TempAType="满屏浮动"
			Case 7
				str_TempAType="左下底端"
			Case 8
				str_TempAType="右下底端"
			Case 9
				str_TempAType="对联广告"
			Case 10
				str_TempAType="循环广告"
			Case 11
				str_TempAType="文字广告"
		End Select
		Select Case Clng(o_Ad_Rs("AdLock"))
			Case 0
				str_Temp_Lock="正常"
			Case 1
				str_Temp_Lock="<font color=""red"">锁定</font>"
		End Select
%>
  <tr class="hback">
    <td width="20%" align="center" class="hback" onClick="javascript:ShowClassInfo('<%=o_Ad_Rs("AdID")%>')" title="点击显示栏目信息"><%=o_Ad_Rs("AdName")%></td>
    <td width="12%" align="center" class="hback" onClick="javascript:ShowClassInfo('<%=o_Ad_Rs("AdID")%>')" title="点击显示栏目信息"><%=str_TempAType%></td>
	<%
		str_date = Year(o_Ad_Rs("AdAddDate"))&"-"&Month(o_Ad_Rs("AdAddDate"))&"-"&Day(o_Ad_Rs("AdAddDate"))
	%>
    <td width="8%" align="center" class="hback" onClick="javascript:ShowClassInfo('<%=o_Ad_Rs("AdID")%>')" title="点击显示栏目信息"><%=str_date%></td>
    <td width="10%" align="center" class="hback" onClick="javascript:ShowClassInfo('<%=o_Ad_Rs("AdID")%>')" title="点击显示栏目信息"><%=o_Ad_Rs("AdMaxClickNum")%></td>
    <td width="10%" align="center" class="hback" onClick="javascript:ShowClassInfo('<%=o_Ad_Rs("AdID")%>')" title="点击显示栏目信息"><%=o_Ad_Rs("AdMaxShowNum")%></td>
    <td width="6%" align="center" class="hback" onClick="javascript:ShowClassInfo('<%=o_Ad_Rs("AdID")%>')" title="点击显示栏目信息"><%=str_Temp_Lock%></td>
    <td width="34%" align="center" class="hback"><a href="javascript:AdLock('<%=Clng(o_Ad_Rs("AdID"))%>');">锁定</a> | <a href="javascript:AdUnLock('<%=Clng(o_Ad_Rs("AdID"))%>');">解锁</a> | <a href="javascript:AdDelOne('<%=Clng(o_Ad_Rs("AdID"))%>');">删除</a> | <a href="javascript:AdUpdate('<%=Clng(o_Ad_Rs("AdID"))%>','<%=o_Ad_Rs("AdClassID")%>','<%=Request.QueryString("Page")%>');">修改</a> | <a href="javascript:Use('<%=Clng(o_Ad_Rs("AdID"))%>');">调用代码</a> |
      <input type="checkbox" name="Checkallbox" value="<%=o_Ad_Rs("AdID")%>"></td>
  </tr>
  	<%
  		str_ClassInfo=""
  		str_ClassMode=""
  		str_ClassName=""
  		str_AdRemarks=""
  		If o_Ad_Rs("AdRemarks")="" Or IsNull(o_Ad_Rs("AdRemarks")) Then
  			str_AdRemarks=""
  		Else
  			str_AdRemarks="&nbsp;&nbsp;广告备注："&Left(o_Ad_Rs("AdRemarks"),50)
  		End If
  		If Clng(o_Ad_Rs("AdClassID"))=-1 Then
  			str_ClassInfo="此广告无所属栏目"&str_AdRemarks
  		Else
  			Set o_ClassInfo_Rs=Conn.execute("Select AdClassID,AdClassName,Lock From FS_AD_Class Where AdClassID="&CintStr(o_Ad_Rs("AdClassID"))&"")
  			If Not o_ClassInfo_Rs.Eof Then
  				str_ClassName=o_ClassInfo_Rs("AdClassName")
  				If o_ClassInfo_Rs("Lock")=0 Then
  					str_ClassMode="正常"
  				Else
  					str_ClassMode="<font color=""red"">被锁定</font>"
  				End If
  				str_ClassInfo="所属栏目："&str_ClassName&"&nbsp;&nbsp;栏目状态："&str_ClassMode&""&str_AdRemarks

	  		Else
  				str_ClassInfo="此广告无所属栏目"&str_AdRemarks
  			End If
  			Set o_ClassInfo_Rs=Nothing
  		End If
  	%>
	<tr id="<%=o_Ad_Rs("AdID")%>" style="display:none" title="点击隐藏栏目信息">
    <td width="100%" align="left" class="hback" onClick="javascript:HideClassInfo('<%=o_Ad_Rs("AdID")%>')" colspan="7" ><%=str_ClassInfo%></td>
    </tr>
<%
		o_Ad_Rs.MoveNext
		If o_Ad_Rs.Eof or o_Ad_Rs.Bof Then Exit For
	Next
	Response.Write "<tr><td class=""hback"" colspan=""7"" align=""left"">"&fPageCount(o_Ad_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;全选<input type=""checkbox"" name=""Checkallbox"" onclick=""javascript:CheckAll('Checkallbox');"" value=""0""></td></tr>"
%>
  </table>
<%
	Else
		Response.write"<table width=""98%"" border=0 align=center cellpadding=2 cellspacing=1 class=table><tr class=""hback""><td>当前没有广告!</td></tr></table>"
End If
%>	
</form>
</body>
</html>
<script language="javascript">
function AdLock(ID)
{
	ID=parseInt(ID);
	location='?Ad_OP=AdLock&ID='+ID+'&Page=<%=NoSqlHack(Request.QueryString("Page"))%>';
}
function AdUnLock(ID)
{
	ID=parseInt(ID);
	location='?Ad_OP=AdUnLock&ID='+ID+'&Page=<%=NoSqlHack(Request.QueryString("Page"))%>';
}
function AdDelOne(ID)
{
	if(confirm('此操作将删除选中的内容？\n你确定删除吗？'))
	{
		ID=parseInt(ID);
		location='?Ad_OP=AdDelOne&ID='+ID+'&Page=<%=NoSqlHack(Request.QueryString("Page"))%>';
	}
}
function CheckAll(CheckType)
{
	var checkBoxArray=document.all(CheckType)
	if(checkBoxArray[checkBoxArray.length-1].checked)
	{
		for(var i=0;i<checkBoxArray.length-1;i++)
		{
			checkBoxArray[i].checked=true;
			
		}
	}else
	{
		for(var i=0;i<checkBoxArray.length-1;i++)
		{
			checkBoxArray[i].checked=false;
		}
	}
}
function P_Lock()
{
	if(confirm('此操作将锁定选中的广告？\n你确定锁定吗？'))
	{
		document.AdInfo.action="?Ad_OP=P_Lock"+'&Page=<%=NoSqlHack(Request.QueryString("Page"))%>';
		document.AdInfo.submit();
	}
}
function P_UnLock()
{
	if(confirm('此操作将解除锁定选中的广告？\n你确定解除锁定吗？'))
	{
		document.AdInfo.action="?Ad_OP=P_UnLock"+'&Page=<%=NoSqlHack(Request.QueryString("Page"))%>';
		document.AdInfo.submit();
	}
}
function P_Del()
{
	if(confirm('此操作将删除选中的广告？\n你确定删除吗？'))
	{
		document.AdInfo.action="?Ad_OP=P_Del"+'&Page=<%=NoSqlHack(Request.QueryString("Page"))%>';
		document.AdInfo.submit();
	}
}
function AdUpdate(ID,ClassID,Page)
{
	ID=parseInt(ID);
	location='Ads_Add.asp?OpType=Update&ID='+ID+'&AdClassID='+ClassID+'&OpPage='+Page;
}
function Use(ID)
{
	OpenWindow('Ad_UseShow.asp?PageTitle=获取调用代码&ID='+ID+"&rnd="+Math.random(),300,130,window);
}
function ShowClassInfo(TrID)
{
	document.all(TrID).style.display="";
}
function HideClassInfo(TrID)
{
	document.all(TrID).style.display="none";
}
</script>
<%
Set Conn=nothing
%><!-- Powered by: FoosunCMS5.0系列,Company:Foosun Inc. -->





