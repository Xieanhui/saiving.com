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
	var flag1=isNumber('UpfileNum','UpfileNum_Alert','�ļ�����ӦΪ������',true)
	var flag2=isNumber('UpfileSize','UpfileSize_Alert','�ļ���СӦΪ������',true)
	var flag3=isNumber('GroupDate','GroupDate_Alert','�ļ���СӦΪ������',true)
	var flag4=isNumber('GroupMoney','GroupMoney_Alert','�������ӦΪ������',true)
	var flag5=isNumber('LimitInfoNum','LimitInfoNum_Alert','��Ϣ����ӦΪ������',true)
	var flag6=isNumber('GroupDebateNum_1','GroupDebateNum1_Alert','��Ⱥ����ӦΪ������',true)
	var flag7=isNumber('GroupDebateNum_2','GroupDebateNum2_Alert','��Ⱥ����ӦΪ������',true)
	var flag8=isNumber('KeywordsNumber','KeywordsNumber_Alert','�ؼ��ָ���ӦΪ������',true)
	var flag9=isNumber('BcardNumber','BcardNumber_Alert','�ؼ��ָ���ӦΪ������',true)
	var flag10=isEmpty('GroupName','GroupName_Alert','��������Ϊ��')
	var flag11=isEmpty('CorpTemplet','CorpTemplet_Alert','ģ���ַ����Ϊ��')
	if(document.ManageGroup.GroupType[0].checked|document.ManageGroup.GroupType[1].checked)
	{
		document.getElementById("GroupType_Alert").innerHTML=""
		if(flag1&&flag2&&flag3&&flag4&&flag5&&flag6&&flag7&&flag8&&flag9&&flag10&flag11)
		{
			if(document.getElementById("GType").value=="all")
			{
				if(confirm("ȷ���޸������û��飿"))
				{
					document.ManageGroup.submit();
				}
			}else if(document.getElementById("GroupIndex").value=="user")
			{
				if(confirm("ȷ���޸����и��˻�Ա�飿"))
				{
					document.ManageGroup.submit();
				}
			}
			else if(document.getElementById("GroupIndex").value=="corp")
			{
				if(confirm("ȷ���޸�������ҵ��Ա�飿"))
				{
					document.ManageGroup.submit();
				}
			}
			else
			document.ManageGroup.submit();
		}
	}else
	{
		document.getElementById("GroupType_Alert").innerHTML="<font color='F43631'>�����Դ�����ѡ��</font>";
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
	var url="getUserGroup.asp?page=UserGroup&id="+typeID+"&r="+Math.random();//����url
	request.open("GET",url,true);//��������
	request.onreadystatechange = getFormInfoResult;
	request.send(null);//�������ݣ���Ϊ����ͨ��url�����ˣ��������ﴫ�ݵ���null
}
function getFormInfoResult()//����������Ӧ��ʱ���ʹ���������
{
	if(request.readyState ==4)//����HTTP ����״̬�ж���Ӧ�Ƿ����
	{
		if(request.status == 200)//�ж������Ƿ�ɹ�
		{
			result=request.responseText;//�����Ӧ�Ľ����Ҳ�����µ�<select>
			document.getElementById("GroupIndexContent").innerHTML="|&nbsp;&nbsp;��Ա�飺"+result;//����������ʵ�ڿͻ���
		}
	}
}
function getGroupParam(Obj)
{
	var GroupID=Obj.value;
	if(!isNaN(GroupID))
	{
		var url="getUserGroupParam.asp?id="+GroupID+"&r="+Math.random();//����url
		request.open("GET",url,true);//��������
		request.onreadystatechange = getGroupParamResult;
		request.send(null);//�������ݣ���Ϊ����ͨ��url�����ˣ��������ﴫ�ݵ���null
	}

}
//ajax end
function getGroupParamResult()//����������Ӧ��ʱ���ʹ���������
{
	if(request.readyState ==4)//����HTTP ����״̬�ж���Ӧ�Ƿ����
	{
		if(request.status == 200)//�ж������Ƿ�ɹ�
		{
			result=request.responseText;//�����Ӧ�Ľ����Ҳ�����µ�<select>
			//��ȡԭ������
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
    <td align="right" class="xingmu" colspan="2"><div align="left">��Ա�����</div></td></tr>
  <tr class="hback">
    <td align="right">��Ա��ѡ��</td>
    <td>��Ա�����ͣ�      
      <select name="GType" id="GType" onChange="getFormInfo(this)">
        <option value="all">���л�Ա��</option>
        <option value="1">���˻�Ա��</option>
        <option value="0">��ҵ��Ա��</option>
      </select> 
      &nbsp;
      <span id="GroupIndexContent"></span></td>
  </tr> 
        <tr class="hback"> 
          <td align="right">�����ƣ�</td> 
          <td width="537"> <input name="GroupName" type="text" id="GroupName" size="50" />
          <font color="#FF0000">*</font> <span class="style1" id="GroupName_Alert"></span></td> 
        </tr> 
      
<tr class="hback">
    <td align="right">�ļ��������ƣ� </td>
    <td><input name="UpfileNum" type="text" id="UpfileNum"  value="0" size="50">
    <span id="UpfileNum_Alert"></span></td>
  </tr>
<tr class="hback">
    <td align="right">�ļ���С���ƣ�</td>
    <td><input name="UpfileSize" type="text" id="UpfileSize" value="0" size="50">
    k<span id="UpfileSize_Alert"></span></td>
  </tr>
<tr class="hback"> 
                <td align="right">����Ч���ޣ�</td> 
                <td><input name="GroupDate" type="text" id="GroupDate"  value="0" size="50"/> 
                �� <span id="GroupDate_Alert"></span></td> 
    </tr> 
      <tr class="hback"> 
          <td align="right">����������֣�</td> 
          <td><input name="GroupPoint" type="text" id="GroupPoint"  value="0" size="50"/>
          <span id="GroupPoint_Alert"></span></td> 
    </tr>
        <tr class="hback">
          <td align="right">���������ң�</td>
          <td><input name="GroupMoney" type="text" id="GroupMoney"  value="0" size="50"/>
          <span id="GroupMoney_Alert"></span></td>
        </tr> 
      <tr class="hback"> 
          <td align="right">�����ͣ�</td> 
          <td><label>
            <input name="GroupType" type="radio" value="1"> 
            ���˻�Ա��</label>
            <label>
            <input type="radio" name="GroupType" value="0" >
��ҵ��Ա��</label>&nbsp;<span id="GroupType_Alert"></span></td> 
    </tr> 
      <tr class="hback"> 
          <td align="right">��Ϣ�����������ޣ�</td> 
          <td><input name="LimitInfoNum" type="text" id="LimitInfoNum" value="10" size="50"/>
          <span id="LimitInfoNum_Alert"></span></td> 
    </tr>
        <tr class="hback">
          <td align="right">��ҵ��Աģ���ַ��</td>
          <td><input name="CorpTemplet" type="text" size="50"><span id="CorpTemplet_Alert"></span></td>
        </tr>
        <tr class="hback">
          <td align="right">��Ⱥ������</td>
          <td>��Ⱥ���������
            <input name="GroupDebateNum_1" type="text" id="GroupDebateNum_1" value="0" size="15"> 
          &nbsp;��Ⱥ�������
          <input name="GroupDebateNum_2" type="text" id="GroupDebateNum_2" value="0" size="15" >
          <span id="GroupDebateNum1_Alert"></span> &nbsp;<span id="GroupDebateNum2_Alert"></span></td>
        </tr>
        <tr class="hback">
          <td align="right">��ͨ����������</td>
          <td><p>
            <label>
            <input type="radio" name="JuniorDomain" value="1" <%if JuniorDomain=1 then Response.Write("checked") end if%>>
  ��</label>
            <label>
            <input name="JuniorDomain" type="radio" value="0" checked <%if JuniorDomain=0 then Response.Write("checked") end if%>>
  ��</label>
            <br>
          </p></td>
        </tr>
        <tr class="hback">
          <td align="right">��Ϣ�ؼ��ָ�����</td>
          <td><input name="KeywordsNumber" type="text" id="KeywordsNumber" value="0" size="50"/>
          <span id="KeywordsNumber_Alert"></span></td>
        </tr>
        <tr class="hback">
          <td align="right">���ɾ�̬�ļ���</td>
          <td><label>
            <input type="radio" name="Ishtml" value="1"/>
��</label>
            <label>
            <input name="Ishtml" type="radio" value="0" checked >
��</label></td>
        </tr>
        <tr class="hback">
          <td align="right">��Ƭ�ղظ������ƣ�</td>
          <td><input name="BcardNumber" type="text" id="BcardNumber" value="0" size="50"/>
          <span id="BcardNumber_Alert"></span></td>
        </tr>
        <tr class="hback">
          <td align="right">��ͨˮӡ��</td>
          <td><label>
            <input type="radio" name="Templetwatermark" value="1" >
��</label>
            <label>
            <input name="Templetwatermark" type="radio" value="0" checked>
��</label></td>
        </tr> 
      <tr class="hback"> 
          <td align="right">&nbsp;</td> 
          <td><input type="Button" name="ManageGroupButton" value=" ���� " onClick="MySubmit()"/> 
            <input type="reset" name="Submit2" value=" ���� " /></td> 
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






