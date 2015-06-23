<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
on error resume next
Dim Conn,User_Conn,ManageGroupRs,GType,GroupIndex
Dim GroupName,UpfileNum,UpfileSize,GroupDate,GroupPoint,GroupMoney,GroupType,CorpTemplet,LimitInfoNum,GroupDebateNum,JuniorDomain,KeywordsNumber,Ishtml,BcardNumber,Templetwatermark
'************************************Update
if Request("Act")="update" then
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF
	GType=Request.Form("GType")
	GroupIndex=Request.Form("GroupIndex")
	if GType="all" then
		User_Conn.execute("Update FS_ME_Group set GroupName='"&NoSqlHack(Request.Form("GroupName"))&"',UpfileNum="&NoSqlHack(Request.Form("UpfileNum"))&",UpfileSize="&NoSqlHack(Request.Form("UpfileSize"))&",GroupDate="&NoSqlHack(Request.Form("GroupDate"))&",GroupPoint="&NoSqlHack(Request.Form("GroupPoint"))&",GroupMoney="&NoSqlHack(Request.Form("GroupMoney"))&",GroupType="&NoSqlHack(Request.Form("GroupType"))&",CorpTemplet='"&NoSqlHack(Request.Form("CorpTemplet"))&"',LimitInfoNum="&NoSqlHack(Request.Form("LimitInfoNum"))&",GroupDebateNum='"&NoSqlHack(Request.Form("GroupDebateNum_1"))&","&NoSqlHack(Request.Form("GroupDebateNum_2"))&"',JuniorDomain="&NoSqlHack(Request.Form("JuniorDomain"))&",KeywordsNumber="&NoSqlHack(Request.Form("KeywordsNumber"))&",isHtml="&NoSqlHack(Request.Form("isHtml"))&",BcardNumber="&NoSqlHack(Request.Form("BcardNumber"))&",Templetwatermark="&NoSqlHack(Request.Form("Templetwatermark")))
	elseif GroupIndex="user" then
		User_Conn.execute("Update FS_ME_Group set GroupName='"&NoSqlHack(Request.Form("GroupName"))&"',UpfileNum="&NoSqlHack(Request.Form("UpfileNum"))&",UpfileSize="&NoSqlHack(Request.Form("UpfileSize"))&",GroupDate="&NoSqlHack(Request.Form("GroupDate"))&",GroupPoint="&NoSqlHack(Request.Form("GroupPoint"))&",GroupMoney="&NoSqlHack(Request.Form("GroupMoney"))&",GroupType="&NoSqlHack(Request.Form("GroupType"))&",CorpTemplet='"&NoSqlHack(Request.Form("CorpTemplet"))&"',LimitInfoNum="&NoSqlHack(Request.Form("LimitInfoNum"))&",GroupDebateNum='"&NoSqlHack(Request.Form("GroupDebateNum_1"))&","&NoSqlHack(Request.Form("GroupDebateNum_2"))&"',JuniorDomain="&NoSqlHack(Request.Form("JuniorDomain"))&",KeywordsNumber="&NoSqlHack(Request.Form("KeywordsNumber"))&",isHtml="&NoSqlHack(Request.Form("isHtml"))&",BcardNumber="&NoSqlHack(Request.Form("BcardNumber"))&",Templetwatermark="&NoSqlHack(Request.Form("Templetwatermark"))&" where GroupType=1")
	elseif GroupIndex="corp" then
		User_Conn.execute("Update FS_ME_Group set GroupName='"&NoSqlHack(Request.Form("GroupName"))&"',UpfileNum="&NoSqlHack(Request.Form("UpfileNum"))&",UpfileSize="&NoSqlHack(Request.Form("UpfileSize"))&",GroupDate="&NoSqlHack(Request.Form("GroupDate"))&",GroupPoint="&NoSqlHack(Request.Form("GroupPoint"))&",GroupMoney="&NoSqlHack(Request.Form("GroupMoney"))&",GroupType="&NoSqlHack(Request.Form("GroupType"))&",CorpTemplet='"&NoSqlHack(Request.Form("CorpTemplet"))&"',LimitInfoNum="&NoSqlHack(Request.Form("LimitInfoNum"))&",GroupDebateNum='"&NoSqlHack(Request.Form("GroupDebateNum_1"))&","&NoSqlHack(Request.Form("GroupDebateNum_2"))&"',JuniorDomain="&NoSqlHack(Request.Form("JuniorDomain"))&",KeywordsNumber="&NoSqlHack(Request.Form("KeywordsNumber"))&",isHtml="&NoSqlHack(Request.Form("isHtml"))&",BcardNumber="&NoSqlHack(Request.Form("BcardNumber"))&",Templetwatermark="&NoSqlHack(Request.Form("Templetwatermark"))&" where GroupType=0")	
	else
	User_Conn.execute("Update FS_ME_Group set GroupName='"&NoSqlHack(Request.Form("GroupName"))&"',UpfileNum="&NoSqlHack(Request.Form("UpfileNum"))&",UpfileSize="&NoSqlHack(Request.Form("UpfileSize"))&",GroupDate="&NoSqlHack(Request.Form("GroupDate"))&",GroupPoint="&NoSqlHack(Request.Form("GroupPoint"))&",GroupMoney="&NoSqlHack(Request.Form("GroupMoney"))&",GroupType="&NoSqlHack(Request.Form("GroupType"))&",CorpTemplet='"&NoSqlHack(Request.Form("CorpTemplet"))&"',LimitInfoNum="&NoSqlHack(Request.Form("LimitInfoNum"))&",GroupDebateNum='"&NoSqlHack(Request.Form("GroupDebateNum_1"))&","&NoSqlHack(Request.Form("GroupDebateNum_2"))&"',JuniorDomain="&NoSqlHack(Request.Form("JuniorDomain"))&",KeywordsNumber="&NoSqlHack(Request.Form("KeywordsNumber"))&",isHtml="&NoSqlHack(Request.Form("isHtml"))&",BcardNumber="&NoSqlHack(Request.Form("BcardNumber"))&",Templetwatermark="&NoSqlHack(Request.Form("Templetwatermark"))&" where GroupID="&NoSqlHack(GroupIndex))	
End if
	if err.number=0 then 
		Response.Redirect("../success.asp")
	else
		Response.Redirect("../error.asp?ErrCodes=<li>"&err.description&"</li>")
	end if
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
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
	var flag8=isNumber('KeywordsNumber','KeywordsNumber_Alert','关键字个数应为正整数',true)
	var flag9=isNumber('BcardNumber','BcardNumber_Alert','关键字个数应为正整数',true)
	var flag10=isEmpty('GroupName','GroupName_Alert','组名不能为空')
	var flag11=isEmpty('CorpTemplet','CorpTemplet_Alert','模版地址不能为空')
	if(document.ManageGroup.GroupType[0].checked|document.ManageGroup.GroupType[1].checked)
	{
		document.getElementById("GroupType_Alert").innerHTML=""
		if(flag1&&flag2&&flag3&&flag4&&flag5&&flag6&&flag7&&flag8&&flag9&&flag10&flag11)
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
		document.getElementById("GroupType_Alert").innerHTML="<font color='F43631'>组类性处必须选择</font>";
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
		}
	}
}

//end
</script>
</HEAD>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="UserJS.js" type="text/JavaScript"></script>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes  oncontextmenu="return false;"> 
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table"> 
<form action="?Act=update" method="post" name="ManageGroup" id="ManageGroup">  
  <tr class="hback"> 
    <td align="right" class="xingmu" colspan="2"><div align="left">会员组管理</div></td></tr>
  <tr class="hback">
    <td align="right">会员组选择：</td>
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
          <td align="right">组名称：</td> 
          <td width="537"> <input name="GroupName" type="text" id="GroupName" size="50" />
          <font color="#FF0000">*</font> <span class="style1" id="GroupName_Alert"></span></td> 
        </tr> 
      
<tr class="hback">
    <td align="right">文件个数限制： </td>
    <td><input name="UpfileNum" type="text" id="UpfileNum"  value="0" size="50">
    <span id="UpfileNum_Alert"></span></td>
  </tr>
<tr class="hback">
    <td align="right">文件大小限制：</td>
    <td><input name="UpfileSize" type="text" id="UpfileSize" value="0" size="50">
    k<span id="UpfileSize_Alert"></span></td>
  </tr>
<tr class="hback"> 
                <td align="right">组有效期限：</td> 
                <td><input name="GroupDate" type="text" id="GroupDate"  value="0" size="50"/> 
                天 <span id="GroupDate_Alert"></span></td> 
    </tr> 
      <tr class="hback"> 
          <td align="right">该组所需积分：</td> 
          <td><input name="GroupPoint" type="text" id="GroupPoint"  value="0" size="50"/>
          <span id="GroupPoint_Alert"></span></td> 
    </tr>
        <tr class="hback">
          <td align="right">该组所需金币：</td>
          <td><input name="GroupMoney" type="text" id="GroupMoney"  value="0" size="50"/>
          <span id="GroupMoney_Alert"></span></td>
        </tr> 
      <tr class="hback"> 
          <td align="right">组类型：</td> 
          <td><label>
            <input name="GroupType" type="radio" value="1"> 
            个人会员组</label>
            <label>
            <input type="radio" name="GroupType" value="0" >
企业会员组</label>&nbsp;<span id="GroupType_Alert"></span></td> 
    </tr> 
      <tr class="hback"> 
          <td align="right">信息发布数量上限：</td> 
          <td><input name="LimitInfoNum" type="text" id="LimitInfoNum" value="10" size="50"/>
          <span id="LimitInfoNum_Alert"></span></td> 
    </tr>
        <tr class="hback">
          <td align="right">企业会员模板地址：</td>
          <td><input name="CorpTemplet" type="text" size="50"><span id="CorpTemplet_Alert"></span></td>
        </tr>
        <tr class="hback">
          <td align="right">社群参数：</td>
          <td>社群最大数量：
            <input name="GroupDebateNum_1" type="text" id="GroupDebateNum_1" value="0" size="15"> 
          &nbsp;社群最大人数
          <input name="GroupDebateNum_2" type="text" id="GroupDebateNum_2" value="0" size="15" >
          <span id="GroupDebateNum1_Alert"></span> &nbsp;<span id="GroupDebateNum2_Alert"></span></td>
        </tr>
        <tr class="hback">
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
        <tr class="hback">
          <td align="right">信息关键字个数：</td>
          <td><input name="KeywordsNumber" type="text" id="KeywordsNumber" value="0" size="50"/>
          <span id="KeywordsNumber_Alert"></span></td>
        </tr>
        <tr class="hback">
          <td align="right">生成静态文件：</td>
          <td><label>
            <input type="radio" name="Ishtml" value="1"/>
是</label>
            <label>
            <input name="Ishtml" type="radio" value="0" checked >
否</label></td>
        </tr>
        <tr class="hback">
          <td align="right">名片收藏个数限制：</td>
          <td><input name="BcardNumber" type="text" id="BcardNumber" value="0" size="50"/>
          <span id="BcardNumber_Alert"></span></td>
        </tr>
        <tr class="hback">
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
            <input type="reset" name="Submit2" value=" 重置 " /></td> 
    </tr> 
  </form> 
  </tr> 
</table> 
</body>
<%
if Request("Act")="update" then
	Conn.close
	Set Conn=nothing
	User_Conn.close
	Set User_Conn=nothing
end if
%>
</html>






