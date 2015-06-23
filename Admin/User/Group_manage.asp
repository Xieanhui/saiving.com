<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
'on error resume next
Dim Conn,User_Conn,ManageGroupRs,GType,GroupIndex,sRootDir
MF_Default_Conn
MF_User_Conn
MF_Session_TF 
if not MF_Check_Pop_TF("ME_GUser") then Err_Show 
if G_VIRTUAL_ROOT_DIR<>"" then sRootDir="/"+G_VIRTUAL_ROOT_DIR else sRootDir=""

Dim GroupName,UpfileNum,UpfileSize,GroupDate,GroupPoint,GroupMoney,GroupType,CorpTemplet,LimitInfoNum,GroupDebateNum,JuniorDomain,KeywordsNumber,Ishtml,BcardNumber,Templetwatermark,Str_ID,CheckRs,CheckSql
if Request.QueryString("Act")="update" then
	if not MF_Check_Pop_TF("ME034") then Err_Show
	Str_ID = NoSqlHack(Request.QueryString("Str_ID"))  
	CheckSql="Select GroupName From FS_ME_Group Where GroupName='"&NoSqlHack(Request.Form("GroupName"))&"' and not GroupID="&Str_ID
	Set CheckRs=Server.CreateObject(G_FS_RS)
	CheckRs.Open CheckSql,User_Conn,1,1
	If CheckRs.RecordCount>0 Then 
		CheckRs.Close
		Set CheckRs=Nothing
		Response.Redirect("../error.asp?ErrCodes=<li>组名不能重复</li>")
		Response.end
	End If
	CheckRs.Close
	Set CheckRs=Nothing
	'Response.Write Str_ID : response.End
	GType=NoSqlHack(Request.Form("GType"))
	GroupIndex=NoSqlHack(Request.Form("GroupIndex"))
	if GType="all" then
		If Str_ID <> "" And IsNumeric(Str_ID) Then
			User_Conn.execute("Update FS_ME_Group set GroupName='"&NoSqlHack(Request.Form("GroupName"))&"',UpfileNum="&NoSqlHack(Request.Form("UpfileNum"))&",UpfileSize="&NoSqlHack(Request.Form("UpfileSize"))&",GroupDate="&NoSqlHack(Request.Form("GroupDate"))&",GroupPoint="&NoSqlHack(Request.Form("GroupPoint"))&",GroupMoney="&NoSqlHack(Request.Form("GroupMoney"))&",GroupType="&NoSqlHack(Request.Form("GroupType"))&",CorpTemplet='CorpTemplet',LimitInfoNum="&NoSqlHack(Request.Form("LimitInfoNum"))&",GroupDebateNum='"&NoSqlHack(Request.Form("GroupDebateNum_1"))&","&NoSqlHack(Request.Form("GroupDebateNum_2"))&"',ProductDiscount="&NoSqlHack(request.Form("ProductDiscount"))&",JuniorDomain=1,KeywordsNumber=0,isHtml="&NoSqlHack(Request.Form("isHtml"))&",BcardNumber="&Replacestr(NoSqlHack(Request.Form("BcardNumber")),":0,else:"&NoSqlHack(Request.Form("BcardNumber")))&",Templetwatermark="&Request.Form("Templetwatermark")&" where GroupID="&Str_ID)
		Else
			User_Conn.execute("Update FS_ME_Group set GroupName='"&NoSqlHack(Request.Form("GroupName"))&"',UpfileNum="&NoSqlHack(Request.Form("UpfileNum"))&",UpfileSize="&NoSqlHack(Request.Form("UpfileSize"))&",GroupDate="&NoSqlHack(Request.Form("GroupDate"))&",GroupPoint="&NoSqlHack(Request.Form("GroupPoint"))&",GroupMoney="&NoSqlHack(Request.Form("GroupMoney"))&",GroupType="&NoSqlHack(Request.Form("GroupType"))&",CorpTemplet='CorpTemplet',LimitInfoNum="&NoSqlHack(Request.Form("LimitInfoNum"))&",GroupDebateNum='"&NoSqlHack(Request.Form("GroupDebateNum_1"))&","&NoSqlHack(Request.Form("GroupDebateNum_2"))&"',ProductDiscount="&NoSqlHack(request.Form("ProductDiscount"))&",JuniorDomain=1,KeywordsNumber=0,isHtml="&NoSqlHack(Request.Form("isHtml"))&",BcardNumber="&NoSqlHack(Request.Form("BcardNumber"))&",Templetwatermark="&NoSqlHack(Request.Form("Templetwatermark")))
		End if
	elseif GroupIndex="user" then
		User_Conn.execute("Update FS_ME_Group set GroupName='"&NoSqlHack(Request.Form("GroupName"))&"',UpfileNum="&NoSqlHack(Request.Form("UpfileNum"))&",UpfileSize="&NoSqlHack(Request.Form("UpfileSize"))&",GroupDate="&NoSqlHack(Request.Form("GroupDate"))&",GroupPoint="&NoSqlHack(Request.Form("GroupPoint"))&",GroupMoney="&NoSqlHack(Request.Form("GroupMoney"))&",GroupType="&NoSqlHack(Request.Form("GroupType"))&",CorpTemplet='CorpTemplet',LimitInfoNum="&NoSqlHack(Request.Form("LimitInfoNum"))&",GroupDebateNum='"&NoSqlHack(Request.Form("GroupDebateNum_1"))&","&NoSqlHack(Request.Form("GroupDebateNum_2"))&"',ProductDiscount="&NoSqlHack(request.Form("ProductDiscount"))&",JuniorDomain=1,KeywordsNumber=0,isHtml="&NoSqlHack(Request.Form("isHtml"))&",BcardNumber="&NoSqlHack(Request.Form("BcardNumber"))&",Templetwatermark="&NoSqlHack(Request.Form("Templetwatermark"))&" where GroupType=1")
	elseif GroupIndex="corp" then
		User_Conn.execute("Update FS_ME_Group set GroupName='"&NoSqlHack(Request.Form("GroupName"))&"',UpfileNum="&NoSqlHack(Request.Form("UpfileNum"))&",UpfileSize="&NoSqlHack(Request.Form("UpfileSize"))&",GroupDate="&NoSqlHack(Request.Form("GroupDate"))&",GroupPoint="&NoSqlHack(Request.Form("GroupPoint"))&",GroupMoney="&NoSqlHack(Request.Form("GroupMoney"))&",GroupType="&NoSqlHack(Request.Form("GroupType"))&",CorpTemplet='CorpTemplet',LimitInfoNum="&NoSqlHack(Request.Form("LimitInfoNum"))&",GroupDebateNum='"&NoSqlHack(Request.Form("GroupDebateNum_1"))&","&NoSqlHack(Request.Form("GroupDebateNum_2"))&"',ProductDiscount="&NoSqlHack(request.Form("ProductDiscount"))&",JuniorDomain=1,KeywordsNumber=0,isHtml="&NoSqlHack(Request.Form("isHtml"))&",BcardNumber="&NoSqlHack(Request.Form("BcardNumber"))&",Templetwatermark="&NoSqlHack(Request.Form("Templetwatermark"))&" where GroupType=0")	
	else
	User_Conn.execute("Update FS_ME_Group set GroupName='"&NoSqlHack(Request.Form("GroupName"))&"',UpfileNum="&NoSqlHack(Request.Form("UpfileNum"))&",UpfileSize="&NoSqlHack(Request.Form("UpfileSize"))&",GroupDate="&NoSqlHack(Request.Form("GroupDate"))&",GroupPoint="&NoSqlHack(Request.Form("GroupPoint"))&",GroupMoney="&NoSqlHack(Request.Form("GroupMoney"))&",GroupType="&NoSqlHack(Request.Form("GroupType"))&",CorpTemplet='CorpTemplet',LimitInfoNum="&NoSqlHack(Request.Form("LimitInfoNum"))&",GroupDebateNum='"&trim(NoSqlHack.Form("GroupDebateNum_1"))&","&NoSqlHack(Request.Form("GroupDebateNum_2"))&"',ProductDiscount="&NoSqlHack(request.Form("ProductDiscount"))&",JuniorDomain=1,KeywordsNumber=0,isHtml="&NoSqlHack(Request.Form("isHtml"))&",BcardNumber="&Replacestr(NoSqlHack(Request.Form("BcardNumber")),":0,else:"&NoSqlHack(Request.Form("BcardNumber")))&",Templetwatermark="&NoSqlHack(Request.Form("Templetwatermark"))&" where GroupID="&NoSqlHack(GroupIndex))	
	End if
	if err.number>0 then
		Response.Redirect("../error.asp?ErrCodes="&err.description&"&ErrorUrl=./user/Group_Manage.asp")
		Response.End()
	else
		Response.Redirect("../success.asp?ErrCodes=<li>操作成功</li>&ErrorUrl=./user/Group_Manage.asp")
		Response.End()
	end if
elseif Request.QueryString("Act")="delete" then
	if not MF_Check_Pop_TF("ME036") then Err_Show 
	if Request.QueryString("tf")="all" then
		User_Conn.execute("Delete from FS_ME_Group")
	elseif Request.QueryString("tf")="user" then
		User_Conn.execute("Delete from FS_ME_Group where GroupType=1")
	elseif Request.QueryString("tf")="corp" then
		User_Conn.execute("Delete from FS_ME_Group where GroupType=0")
	else
		User_Conn.execute("Delete from FS_ME_Group where GroupID="&NoSqlHack(Request.QueryString("tf")))
	end if
	if err.number>0 then
		Response.Redirect("../error.asp?ErrCodes="&err.description)
		Response.End()
	else
		Response.Redirect("../success.asp?ErrCodes=<li>删除成功</li>&ErrorUrl=./user/Group_Manage.asp")
		Response.End()
	end if
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="lib/UserJS.js" type="text/JavaScript"></script>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes > 
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table"> 
 <tr class="hback"> 
    <td align="right" class="xingmu" colspan="2"><div align="left"><a href="Group_Manage.asp">会员组管理</a>  <a href="Group_Add.asp">会员组添加</a>&nbsp;&nbsp;<a href="javascript:history.back();">后退</a></div></td></tr></table>
	<!-----------------------------2/7 by chen------------------------------------------------------------------------->
	<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr>
    <td width="7%" class="xingmu"><div align="center">会员组名称</div></td>
    <td width="39%" class="xingmu"><div align="center">会员组类别</div></td>
    <td width="27%" class="xingmu"><div align="center">会员组操作</div></td>
  </tr>
  <%
  dim rs_class
  Set rs_class=server.CreateObject(G_FS_RS)
	rs_class.open "select GroupID,GroupName,GroupType From FS_ME_Group order by GroupID desc",User_Conn,1,3
  do while not rs_class.eof 
  %>
  <tr class="hback">
    <td align="center"><a href="Group_Manage.asp?Action=GroupUpdate&GroupID=<% =rs_class("GroupID")%>">
      <% = rs_class("GroupName")%>
      </a></td>
    <td align="center"><% if rs_class("GroupType")=0 then Response.Write("企业会员组") Else Response.Write("个人会员组") %></td>
	<td class="hback"><div align="center"><a href="Group_Manage.asp?Action=GroupUpdate&GroupID=<% =rs_class("GroupID")%>">会员组修改</a>｜<a href="Group_Manage.asp?Action=GroupDelete&GroupID=<% =rs_class("GroupID")%>" onClick="{if(confirm('确定要删除该会员组吗？')){return true;}return false;}">会员组删除</a></div></td>
  </tr>
  <%
  rs_class.movenext
  loop
  rs_class.close:set rs_class = nothing
  %>
</table>
<% 
	Dim str_Action
	str_Action = Request("Action")
	Select Case str_Action
	Case "GroupUpdate"
		Call GroupUpdate()
	'Case "GroupDelete"
		'Call GroupDelete()
	End Select
%>
	<!-------------------------2/7 by chen会员组管理新增---------------------------------------------------------------->
	<% sub GroupUpdate()
  dim rs_class1,GroupID,GroupName,UpfileNum,UpfileSize,GroupPoint,GroupDate,GroupMoney,GroupType,LimitInfoNum,GroupDebateNum,GroupDebateNum1,GroupDebateNum2,ProductDiscount
  GroupID=Clng(Request.QueryString("GroupID"))
  Set rs_class1=server.CreateObject(G_FS_RS)
	rs_class1.open "select GroupID,GroupNumber,GroupPoint,GroupDate,GroupRule,GroupMoney,UpfileNum,UpfileSize,GroupPopList,LimitInfoNum,CorpTemplet,GroupDebateNum,JuniorDomain,KeywordsNumber,isHtml,BcardNumber,Templetwatermark,ProductDiscount,GroupName,GroupType,GroupMoney From FS_ME_Group where GroupID = "&GroupID&" order by GroupID desc",User_Conn,1,3
  if not rs_class1.eof then
   GroupName=rs_class1("GroupName")
   UpfileNum=rs_class1("UpfileNum")
   UpfileSize=rs_class1("UpfileSize")
   ProductDiscount=rs_class1("ProductDiscount")
   GroupPoint=rs_class1("GroupPoint")
   GroupDate=rs_class1("GroupDate")
   GroupMoney=rs_class1("GroupMoney")
   GroupType=rs_class1("GroupType")
   LimitInfoNum=rs_class1("LimitInfoNum")
   GroupDebateNum=rs_class1("GroupDebateNum")
   GroupDebateNum1=split(GroupDebateNum,",")(0)
   GroupDebateNum2=split(GroupDebateNum,",")(1)
 end if
  rs_class1.close:set rs_class1 = nothing 
  %> 
	<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table"> 
<form action="?Act=update&Str_ID=<% = GroupID %>" method="post" name="ManageGroup" id="ManageGroup">  
  <tr class="hback" style="display:none">
    <td width="183" align="right">会员组选择：</td>
    <td>会员组类型：      
      <select name="GType" id="GType" onChange="getFormInfo(this)">
        <option value="all">所有会员组</option>
        <option value="1">个人会员组</option>
        <option value="0">企业会员组</option>
      </select> 
      &nbsp;
      <span id="GroupIndexContent"></span></td>
  </tr> 
        <tr class="hback"> 
          <td width="183" align="right">组名称：</td> 
          <td width="791"> <input name="GroupName" type="text" id="GroupName" size="50" value="<% =GroupName%>"/>
          <font color="#FF0000">*</font> <span class="style1" id="GroupName_Alert"></span></td> 
        </tr> 
      
<tr class="hback">
    <td align="right">文件个数限制： </td>
    <td><input name="UpfileNum" type="text" id="UpfileNum"  value="<% =UpfileNum%>" size="50" onChange="if(length(this.value)<=9){alert('文件大小限制最多只能输入9位整数');this.value='0';}"	>
    <span id="UpfileNum_Alert"></span></td>
  </tr>
<tr class="hback">
    <td align="right">文件大小限制：</td>
    <td><input name="UpfileSize" type="text" id="UpfileSize" value="<% =UpfileSize%>" size="50">
    k<span id="UpfileSize_Alert"> 会员空间占用</span></td>
	  <tr class="hback">
  <td align="right">会员折扣：</td>
  <td><input name="ProductDiscount" type="text" id="ProductDiscount" value="<% =ProductDiscount%>" size="50" onChange="if(length(this.value)<=9){alert('最多只能输入4个字符!');this.value='0';}">比如：打八折，该处则填写0.8 <span id="ProductDiscount_Alert"></span></td>
  </tr>
<tr class="hback"> 
                <td align="right">组有效期限：</td> 
                <td><input name="GroupDate" type="text" id="GroupDate"  value="<% =GroupDate%>" size="50" onChange="if(length(this.value)<=4){alert('组有效期限最多只能输入4位整数');this.value='0';}"/> 
                天 <span id="GroupDate_Alert"></span></td> 
    </tr> 
      <tr class="hback"> 
          
      <td align="right">积分：</td> 
          <td><input name="GroupPoint" type="text" id="GroupPoint"  value="<% =GroupPoint%>" size="50" onChange="if(this.value>32500){alert('积分不能大于32500');this.value='0';}"/>
          <span id="GroupPoint_Alert"></span></td> 
    </tr>
        <tr class="hback">
          
      <td align="right">金币：</td>
          <td><input name="GroupMoney" type="text" id="GroupMoney"  value="<% =GroupMoney%>" size="50" onChange="if(length(this.value)>4){alert('金币数不能大于9999');this.value='0';}"/>
          <span id="GroupMoney_Alert"></span></td>
        </tr> 
      <tr class="hback"> 
          <td align="right">组类型：</td> 
        <td><label>
            <input name="GroupType" type="radio" value="1" <%if GroupType=1 then Response.Write("checked")%>> 
            个人会员组</label>
            <label>
            <input type="radio" name="GroupType" value="0" <%if GroupType=0 then Response.Write("checked")%> >
企业会员组</label>&nbsp;<span id="GroupType_Alert">(企业会员组暂时应用于以后扩展使用。您目前建立并不能使用)</span></td> 
    </tr> 
      <tr class="hback"> 
          <td align="right">信息发布数量上限：</td> 
          <td><input name="LimitInfoNum" type="text" id="LimitInfoNum" value="<% =LimitInfoNum%>" size="50"/>
          <span id="LimitInfoNum_Alert"></span></td> 
    </tr>
        <tr class="hback" style="display:none">
          
      <td align="right">会员模板地址：</td>
          <td><input name="CorpTemplet" type="text" size="50">
		  <input name="Submit5" type="button" id="selNewsTemplet" value="选择模板"  onClick="OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectTemplet.asp?CurrPath=<%=sRootDir %>/<% = G_TEMPLETS_DIR %>',400,300,window,document.all.CorpTemplet);document.all.CorpTemplet.focus();">
		  <span id="CorpTemplet_Alert"></span></td>
        </tr>
        <tr class="hback">                 
          <td align="right">社群参数：</td>
          <td>社群最大数量：
            <input name="GroupDebateNum_1" type="text" id="GroupDebateNum_1" value="<% =GroupDebateNum1%>" size="15"> 
          &nbsp;社群最大人数
          <input name="GroupDebateNum_2" type="text" id="GroupDebateNum_2" value="<% =GroupDebateNum2%>" size="15" >
          <span id="GroupDebateNum1_Alert"></span> &nbsp;<span id="GroupDebateNum2_Alert"></span></td>
        </tr>
        <tr class="hback" style="display:none">
          <td align="right">开通二级域名：</td>
          <td><p>
            <label>
            <input type="radio" name="JuniorDomain" value="1" <%if JuniorDomain=1 then Response.Write("checked") end if%>>
  是</label>
            <label>
            <input name="JuniorDomain" type="radio" value="0" checked <%if JuniorDomain=0 then Response.Write("checked") end if%>>
  否</label>
            <br>
          </p></td>
        </tr>
        <tr class="hback" style="display:none">
          <td align="right">信息关键字个数：</td>
          <td><input name="KeywordsNumber" type="text" id="KeywordsNumber" value="0" size="50" onChange="if(length(this.value)>3){alert('信息关键字个数不能大于999');this.value='0';}"/>
          <span id="KeywordsNumber_Alert"></span></td>
        </tr>
        <tr class="hback" style="display:none">
          <td align="right">生成静态文件：</td>
          <td><label>
            <input type="radio" name="Ishtml" value="1"/>
是</label>
            <label>
            <input name="Ishtml" type="radio" value="0" checked >
否</label></td>
        </tr>
        <tr class="hback" style="display:none" id="TR_BcardNumber">
          <td align="right">名片收藏个数限制：</td>
          <td><input name="BcardNumber" type="text" id="BcardNumber" value="0" size="50" onChange="if(length(this.value)>4){alert('名片收藏个数不能大于9999');this.value='0';}"/>
          <span id="BcardNumber_Alert"></span></td>
        </tr>
        <tr class="hback" style="display:none">
          <td align="right">开通水印：</td>
          <td><label>
            <input type="radio" name="Templetwatermark" value="1" >
是</label>
            <label>
            <input name="Templetwatermark" type="radio" value="0" checked>
否</label></td>
        </tr> 
      <tr class="hback"> 
          <td align="right">&nbsp;</td> 
        <td><input type="Button" name="ManageGroupButton" value=" 保存 " onClick="MySubmit()"/> 
            <input type="reset" name="Submit2" value=" 重置 " />
            <input type="Button" name="DeleteGroup_Button" value=" 删 除 " onClick="AlertBeforeDelete()" style="display:none"></td> 
    </tr> 
  </form> 
  </tr> 
</table> 
<% end sub %>
<!---------------------------删除会员组-----2/7 by chen---------------------------->
<% 
	if request.QueryString("Action")="GroupDelete" then
		User_Conn.execute("Delete from FS_ME_Group where GroupID="&Clng(Request.QueryString("GroupID"))&"")
		Response.Redirect("../success.asp?ErrCodes=<li>删除成功</li>&ErrorUrl=user/Group_Manage.asp")
		Response.End()
	end if		
%>
<!------------------------------------------------------------------------------------------------->
</body>
<%
	Conn.close
	Set Conn=nothing
	User_Conn.close
	Set User_Conn=nothing
%>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->

function MySubmit()
{
	var flag1=isNumber('UpfileNum','UpfileNum_Alert','文件个数应为正整数',true)
	var flag2=isNumber('UpfileSize','UpfileSize_Alert','文件大小应为正整数',true)
	var flag3=isNumber('GroupDate','GroupDate_Alert','文件大小应为正整数',true)
	var flag4=isNumber('GroupMoney','GroupMoney_Alert','金币数量应为正整数',true)
	var flag5=isNumber('LimitInfoNum','LimitInfoNum_Alert','信息数量应为正整数',true)
	var flag6=isNumber('GroupDebateNum_1','GroupDebateNum1_Alert','社群数量应为正整数',true)
	var flag7=isNumber('GroupDebateNum_2','GroupDebateNum2_Alert','社群人数应为正整数',true)
	//var flag8=isNumber('KeywordsNumber','KeywordsNumber_Alert','关键字个数应为正整数',true)
	var flag9=(document.getElementById("TR_BcardNumber").style.display=='none')?true:isNumber('BcardNumber','BcardNumber_Alert','关键字个数应为正整数',true)
	var flag10=isEmpty('GroupName','GroupName_Alert','组名不能为空')
	//var flag11=isEmpty('CorpTemplet','CorpTemplet_Alert','模版地址不能为空')
	//var flag12=isNumber('ProductDiscount','ProductDiscount_Alert',false)
	if(document.ManageGroup.GroupType[0].checked|document.ManageGroup.GroupType[1].checked)
	{
		document.getElementById("GroupType_Alert").innerHTML=""
		if(flag1&&flag2&&flag3&&flag4&&flag5&&flag6&&flag7&&flag9&&flag10)
		{
			if(document.getElementById("GType").value=="all")
			{
				if(confirm("确定修改所有用户组？"))
				{
					document.ManageGroup.submit();
				}
			}else if(document.getElementById("GroupIndex").value=="user")
			{
				if(confirm("确定修改所有个人会员组？"))
				{
					document.ManageGroup.submit();
				}
			}
			else if(document.getElementById("GroupIndex").value=="corp")
			{
				if(confirm("确定修改所有企业会员组？"))
				{
					document.ManageGroup.submit();
				}
			}
			else
			document.ManageGroup.submit();
		}
	}else
	{
		document.getElementById("GroupType_Alert").innerHTML="<font color='F43631'>组类性质必须选择</font>";
	}
}
//Ajax
var request=true;
var result;
var ParamArray;
try
{
	request=new XMLHttpRequest();
}catch(trymicrosoft)
{
try
{
	request=new ActiveXObject("Msxml2.XMLHTTP")
}catch(othermicrosoft)
{
try
{
	request=new ActiveXObject("Microsoft.XMLHTTP")
}catch(filed)
{
	request=false;
}
}
}
if(!request) alert("Error initializing XMLHttpRequest!");
function getFormInfo(Obj)
{
	var typeID=Obj.value;
	if(isNaN(typeID))
	{
		document.getElementById("GroupIndexContent").innerHTML="";
		return ;
	}
	var url="getUserGroup.asp?page=UserGroup&id="+typeID+"&r="+Math.random();//构造url
	request.open("GET",url,true);//建立连接
	request.onreadystatechange = getFormInfoResult;
	request.send(null);//传送数据，因为数据通过url传递了，所以这里传递的是null
}
function getFormInfoResult()//当服务器响应的时候就使用这个方法
{
	if(request.readyState ==4)//根据HTTP 就绪状态判断响应是否完成
	{
		if(request.status == 200)//判断请求是否成功
		{
			result=request.responseText;//获得响应的结果，也就是新的<select>
			document.getElementById("GroupIndexContent").innerHTML="|&nbsp;&nbsp;会员组："+result;//将这个结果现实在客户端
		}
	}
}
function getGroupParam(Obj)
{
	var GroupID=Obj.value;
	if(!isNaN(GroupID))
	{
		var url="getUserGroupParam.asp?id="+GroupID+"&r="+Math.random();//构造url
		request.open("GET",url,true);//建立连接
		request.onreadystatechange = getGroupParamResult;
		request.send(null);//传送数据，因为数据通过url传递了，所以这里传递的是null
	}

}
//ajax end
function getGroupParamResult()//当服务器响应的时候就使用这个方法
{
	if(request.readyState ==4)//根据HTTP 就绪状态判断响应是否完成
	{
		if(request.status == 200)//判断请求是否成功
		{
			result=request.responseText;//获得响应的结果，也就是新的<select>
			//获取原有设置
			ParamArray=result.split("|");
			document.getElementById("GroupName").value=ParamArray[0];
			document.getElementById("UpfileNum").value=ParamArray[1];
			document.getElementById("UpfileSize").value=ParamArray[2];
			document.getElementById("GroupDate").value=ParamArray[3];
			document.getElementById("GroupPoint").value=ParamArray[4];
			document.getElementById("GroupMoney").value=ParamArray[5];
			if(ParamArray[6]==1)
			{
				document.ManageGroup.GroupType[0].checked=true;
			}
			else
			{
				document.ManageGroup.GroupType[1].checked=true;
			}
			document.getElementById("LimitInfoNum").value=ParamArray[7];
			document.getElementById("CorpTemplet").value=ParamArray[8];
			if(ParamArray[9]!=null && ParamArray[9]!="")
			{
				var TempArray=ParamArray[9].split(",");
				document.getElementById("GroupDebateNum_1").value=TempArray[0]
				document.getElementById("GroupDebateNum_2").value=TempArray[1]
			}
			if(ParamArray[10]==1)
			{
				document.ManageGroup.JuniorDomain[0].checked=true;
			}
			else
			{
				document.ManageGroup.JuniorDomain[1].checked=true;
			}
			document.getElementById("KeywordsNumber").value=ParamArray[11];
			if(ParamArray[12]==1)
			{
				document.ManageGroup.Ishtml[0].checked=true;
			}
			else
			{
				document.ManageGroup.Ishtml[1].checked=true;
			}
			document.getElementById("BcardNumber").value=ParamArray[13];			
			if(ParamArray[14]==1)
			{
				document.ManageGroup.Templetwatermark[0].checked=true;
			}
			else
			{
				document.ManageGroup.Templetwatermark[1].checked=true;
			}
			document.getElementById("ProductDiscount").value=ParamArray[15];
			
		}
	}
}
//end
function AlertBeforeDelete()
{
	if(document.getElementById("GType").value=="all")
	{
		if(confirm("确定要删除所有用户组！"))
			location='Group_manage.asp?act=delete&tf=all'
	}else if(document.getElementById("GType").value==1&&document.getElementById("GroupIndex").value=="user")
	{
		if(confirm("确定要删除所有个人用户组！"))
			location='Group_manage.asp?act=delete&tf=user'
	}else  if(document.getElementById("GType").value==0&&document.getElementById("GroupIndex").value=="corp")
	{
		if(confirm("确定要删除所有企业用户组！"))
			location="Group_manage.asp?act=delete&tf=corp"
	}else
	{
		if(confirm("确定要删除该用户组！"))
			location='Group_manage.asp?act=delete&tf='+document.getElementById("GroupIndex").value
	}
}
</script>
</html>






