<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
'on error resume next
Dim Conn,User_Conn,awardRs,prizeIDs,EndDate,awardUser,AwardID,ArrayIndex,AwardsUserArray,UserInfoRs,AwardsUserRs
Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,i
int_RPP=20 '����ÿҳ��ʾ��Ŀ
int_showNumberLink_=8 '���ֵ�����ʾ��Ŀ
showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
str_nonLinkColor_="#999999" '����������ɫ
toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
toL_="<font face=webdings title=""���һҳ"">:</font>"				'βҳ

MF_Default_Conn
MF_User_Conn
MF_Session_TF
Set awardRs=Server.CreateObject(G_FS_RS)
awardRs.open "select AwardID,AwardName,AwardPic,StartDate,EndDate,PrizeIDS,Opened from FS_ME_award",User_Conn,1,1
'���ֳ齱
function activeAward()
	Dim active_TF_Rs,sql_cmd,activeTF
	activeTF=false
	sql_cmd="select AwardID from FS_ME_award where opened=0"
	Set active_TF_Rs=User_Conn.execute(sql_cmd)
	if not active_TF_Rs.eof then
		activeTF=true
	End if
	activeAward=activeTF
	active_TF_Rs.close
	set active_TF_Rs=Nothing
End function
'���ֶһ�
Function activeAwardPoint
	Dim active_TF_Rs,sql_cmd,activeTF
	activeTF=false
	if  G_IS_SQL_DB=0 then
		sql_cmd="select AID from FS_ME_AnswerForPoint where DateDiff(d,Convert(nvarchar(10),AEndDate,120),#"&DateValue(Now)&"#)>0"
	Else
		sql_cmd="select AID from FS_ME_AnswerForPoint where DateDiff('d',Convert(nvarchar(10),AEndDate,120),'#"&DateValue(Now)&"#')>0"
	End if
	Set active_TF_Rs=User_Conn.execute(sql_cmd)
	if not active_TF_Rs then
		activeTF=true
	End if
	activeAwardPoint=activeTF
	active_TF_Rs.close
	set active_TF_Rs=nothing
End function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
</HEAD>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="lib/UserJS.js" type="text/JavaScript"></script>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<BODY>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr> 
    <td class="xingmu">�齱����</td>
  </tr>
  <tr>
    <td class="hback">���ֳ齱&nbsp;|&nbsp;<a href="ChangePrize.asp">���ֶһ�</a> | <a href="AnswerForPoint.asp">���־���</a> 
      | <a href="#" onClick="history.back()">����</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form action="awardAction.asp?act=delete" method="post" name="awardForm" id="awardForm">
    <tr class="xingmu"> 
      <td width="16%" align="center">����</td>
      <td width="30%" align="center">�н���Ա</td>
      <td width="11%" align="center">״̬</td>
      <td width="17%" align="center">��ֹʱ��</td>
      <td width="14%" align="center"><input type="checkbox" name="Delete_CheckAll" value="all" onClick="CheckAll(this,'DeleteAwards')"></td>
    </tr>
    <%
			if not awardRs.eof then
				awardRs.PageSize=int_RPP
				cPageNo=NoSqlHack(Request.QueryString("Page"))
				If cPageNo="" Then cPageNo = 1
				If not isnumeric(cPageNo) Then cPageNo = 1
				cPageNo = Clng(cPageNo)
				If cPageNo>awardRs.PageCount Then cPageNo=awardRs.PageCount 
				If cPageNo<=0 Then cPageNo=1
				awardRs.AbsolutePage=cPageNo
			end if
			for i=0 to int_RPP
				if awardRs.eof then exit for
				Response.Write("<tr class='hback'>"&chr(10)&chr(13)) 
				Response.Write("<td width='30%' align='center'><a href='award_AddEdit.asp?act=edit&awardid="&awardRs("awardID")&"'>"&awardRs("awardName")&"</a></td>"&chr(10)&chr(13))
				Response.Write("<td width='20%'>")
				Response.Write("<select name='Grade_"&awardRs("AwardID")&"' onchange=""getAwardUser('"&awardRs("AwardID")&"',this.value)"">")
				prizeIDs=split(awardRs("prizeIDs"),",")
				for ArrayIndex=0 to Ubound(prizeIDs)
					Response.Write("<option value='"&prizeIDs(ArrayIndex)&"'>"&(ArrayIndex+1)&"�Ƚ�</option>"&chr(10)&chr(13))
				next
				Response.Write("</select>")
				Response.Write(" | <span id='PrizeUsers_"&awardRs("awardID")&"'>"&chr(10)&chr(13))
				Response.Write("<select name='AwardUsers_"&AwardID&"'>"&chr(10)&chr(13))
				Set AwardsUserRs=User_Conn.execute("Select UserNumber,winner From FS_ME_User_Prize where PrizeID="&CintStr(prizeIDs(0))&" And awardID="&awardRs("awardID")&" and winner=1")
				if not AwardsUserRs.eof then
					while not AwardsUserRs.eof  
						Set UserInfoRs=User_Conn.execute("Select UserName from FS_ME_Users where UserNumber='"&AwardsUserRs("UserNumber")&"'")
						Response.Write("<option value='"&AwardsUserRs("UserNumber")&"'>"&UserInfoRs("UserName")&"</option>"&Chr(10)&Chr(13))
						AwardsUserRs.movenext
					Wend
				ELse
					Response.Write("<option value='-1'>�����н�</option>"&Chr(10)&Chr(13))
				End if
				AwardsUserRs.close
				Set AwardsUserRs=nothing
				Set UserInfoRs=nothing
				Response.Write("</span>")
				Response.Write("</td>")
				EndDate=awardRs("EndDate")
				if awardRs("Opened")=1 then
					Response.Write("<td align='center'>�ѹ���</td>")
				Elseif DateValue(EndDate)=DateValue(Now()) Then
					Response.Write("<td align='center'><button onClick=""openAward("&awardRs("AwardID")&")"">��  ��</button></td>")
				ElseIf DateValue(EndDate)<DateValue(Now()) And awardRs("Opened")=0 then
					Response.Write("<td align='center'><button onClick=""openAward("&awardRs("AwardID")&")"">�ѹ��ڣ��뿪��</button>")
				Else
					Response.Write("<td align='center'><font color=""red"">δ����</font></td>")
				end If
				Response.Write("<td align='center'>"&EndDate&"</td>")
				Response.Write("<td align='center'><input type='checkbox' name='DeleteAwards' value='"&awardRs("AwardID")&"'></td>")
				Response.Write("</tr>")
				awardRs.movenext
			next
		%>
  </form>
  <tr> 
    <td align="right" colspan="6" class="hback">
	<%
		Dim displayTF
		if activeAward then
			displayTF="disabled"
		End if
	%>
	<input name="AddAward" type="button" value="�� ��" onClick="location='award_AddEdit.asp?act=add'" <%=displayTF%>> 
      <input type="Button" name="deleteAward" onClick="AlertBeforeSubmite('DeleteAwards')" value="ɾ ��"> 
      <%
	response.Write fPageCount(awardRs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
	%>
    </td>
  </tr>
</table> 
</body>
<%
if Request.QueryString("Act")="addGroup" then
	AddGroupRs.close
	set AddGroupRs=nothing
	Conn.close
	Set Conn=nothing
	User_Conn.close
	Set User_Conn=nothing
end if
%>
<script language="JavaScript" type="text/JavaScript">
var request=true;
var result;
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
function getAwardUser(Obj1,Obj2)
{
	var url="getAwardUser.asp?AwardId="+Obj1+"&PrizeID="+Obj2+"&r="+Math.random();//����url
	request.open("GET",url,true);//��������
	request.onreadystatechange = getResult;
	request.send(null);//�������ݣ���Ϊ����ͨ��url�����ˣ��������ﴫ�ݵ���nulla
}
function getResult(Obj)//����������Ӧ��ʱ���ʹ���������
{
	if(request.readyState ==4)//����HTTP ����״̬�ж���Ӧ�Ƿ����
	{
		if(request.status == 200)//�ж������Ƿ�ɹ�
		{
			result=request.responseText;//�����Ӧ�Ľ����Ҳ�����µ�<select>
			var contaner=result.substring(0,result.indexOf("*"));
			var selectContent=result.substring(result.indexOf("*")+1,result.length);
			document.getElementById(contaner).innerHTML=selectContent;

		}
	}
}
function CheckAll(Obj,TargetName)
{
	var CheckBoxArray;
	CheckBoxArray=document.getElementsByName(TargetName);
	for(var i=0;i<CheckBoxArray.length;i++)
	{
		if(Obj.checked)
		{
			CheckBoxArray[i].checked=true;
		}
		else
		{
			CheckBoxArray[i].checked=false;
		}
	}
}
function AlertBeforeSubmite(TargetName)
{
	var checkGroup=document.getElementsByName(TargetName);
	var flag=false;
	for(var i=0;i<checkGroup.length;i++)
	{
		if(checkGroup[i].checked)
		{
			flag=true;
		}
	}
	if(flag)
	{
		if(confirm("ȷ��Ҫɾ���ü�¼?�ò�������ɾ���û��н���¼��"))
		{
			document.awardForm.submit();
		}
	}
	else
	{
		alert("��ѡ��Ҫɾ���ļ�¼")
	}
}
function openAward(awardID)
{
	location="awardAction.asp?act=open&awardID="+awardID+"&rnd="+Math.random();
}
</script>
</html>






