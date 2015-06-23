<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<%
'on error resume next
Dim Conn,User_Conn,OperateDebateRs,DebateID,title,content,ParentID,parentIDRs,Classid,AppointUserNumber,AppointUserGroup,isLock,AccessFile,TopicRs,TopicRs1,TopicRs2,ClassRs,UserGroupRs,CorpGroupRs,ArrayCount,selectedTF,currentTopicRs,i,ChildDebateArray
Admin_Login_State
MF_Default_Conn
MF_User_Conn
MF_Session_TF
AppointUserGroup=null
function getChildID(ID,self)
	Dim F_Rs,F_ChildID
	Set F_Rs=User_Conn.Execute("select DebateID,ParentID from FS_ME_GroupDebate where parentID="&CintStr(ID))
	while not F_Rs.eof
		F_ChildID=F_Rs("DebateID")&","&getChildID(F_Rs("DebateID"),false)
		F_Rs.movenext
	wend
	F_Rs.close
	Set F_Rs=nothing
	if self then
		getChildID=ID&","&F_ChildID
	else
		getChildID=F_ChildID
	end if
end function
function getParentNum(parentID)
	Dim F_Parent_Count,F_Rs
	F_Parent_Count=1
	if not isnumeric(parentID) then exit function
	Set F_Rs=User_Conn.execute("Select DebateID,ParentID from FS_ME_GroupDebate where DebateID="&CintStr(parentID))
	if not F_Rs.eof then
		F_Parent_Count=F_Parent_Count+getParentNum(F_Rs("ParentID"))
	end if
	F_Rs.close
	Set F_Rs=nothing
	getParentNum=F_Parent_Count
end function
'************************************Update
if Request.QueryString("Act")="edit" then'修改界面
	DebateID=Request.QueryString("DebateID")
	Set OperateDebateRs=server.CreateObject(G_FS_RS)
	OperateDebateRs.open "select DebateID,title,content,addtime,ParentID,classid,AppointUserNumber,AppointUserGroup,AddTime,isLock,AccessFile from FS_ME_GroupDebate where DebateID="&NoSqlHack(DebateID),User_Conn,1,3
	if not OperateDebateRs.eof then
		title=OperateDebateRs("title")
		content=OperateDebateRs("content")
		ParentID=OperateDebateRs("ParentID")
		classid=OperateDebateRs("classid")
		AppointUserNumber=OperateDebateRs("AppointUserNumber")
		if OperateDebateRs("AppointUserGroup")<>"" then
			AppointUserGroup=split(OperateDebateRs("AppointUserGroup"),",")
		end if
		isLock=OperateDebateRs("isLock")
		AccessFile=OperateDebateRs("AccessFile")
	end if
elseif Request.QueryString("Act")="EditDebate" then'修改动作
	if Request.Form("currentValue")<>"" then
		parentID=NoSqlHack(Request.Form("currentValue"))
	elseif Request.Form("topic2")<>"" then
		parentID=NoSqlHack(Request.Form("topic2"))
	elseif Request.Form("topic1")<>"" then
		parentID=NoSqlHack(Request.Form("topic1"))
	elseif Request.Form("topic")<>"" then
		parentID=NoSqlHack(Request.Form("topic"))
	end if 
	if request.Form("userGroup")="" and Request.Form("corpGroup")<>"" then
		AppointUserGroup=NoSqlHack(Request.Form("corpGroup"))
	elseif request.Form("userGroup") <> "" and Request.Form("corpGroup")="" then
		AppointUserGroup=NoSqlHack(Request.Form("userGroup"))
	elseif request.Form("userGroup") <> "" and Request.Form("corpGroup")<>"" then
		AppointUserGroup=NoSqlHack(trim(request.Form("userGroup")))&","&trimNoSqlHack((Request.Form("corpGroup")))
	end if
	
	'****************验证修改数据的合理性
	ChildDebateArray=split(DelHeadAndEndDot(getChildID(NoSqlHack(Request("DebateID")),true)),",")
	for i=0 to Ubound(ChildDebateArray)
		if parentID=ChildDebateArray(i) then
		Response.Redirect("../error.asp?ErrCodes=<li>1.自己不能是自己的父类</li><li>2.自己的子类不能作为父类</li>")
		Response.End()
		end if
	next
	 if getParentNum(parentID)+Ubound(ChildDebateArray)>4 then
		Response.Redirect("../error.asp?ErrCodes=<li>总层级数不能大于4</li>")
		Response.End()
	 end if
	'****************************************
	
	User_Conn.Execute("Update FS_ME_GroupDebate set title='"&NoSqlHack(Request.Form("title"))&"',content='"&NoSqlHack(Request.Form("Content"))&"',parentid="&NoSqlHack(parentid)&",classid="&NoSqlHack(Request.Form("DebateClass"))&",AppointUserNumber='"&NoSqlHack(Request.Form("AppointUserNumber"))&"',AppointUserGroup='"&NoSqlHack(AppointUserGroup)&"',addtime='"&Now()&"',isLock="&NoSqlHack(Request.Form("isLock"))&",AccessFile='"&NoSqlHack(Request.Form("AccessFile"))&"' where DebateID="&NoSqlHack(Request("DebateID")))
	if err.number>0 then
		Response.Redirect("../error.asp?ErrCodes="&err.description&"&ErrorUrl=./user/GroupDebate_manage.asp")
		Response.End()
	else
		Response.Redirect("../success.asp?ErrCodes=<li>操作成功</li>&ErrorUrl=./user/GroupDebate_manage.asp")
		Response.End()
	end if
elseif Request("Act")="AddDebate" then'添加动作
	Set OperateDebateRs=server.CreateObject(G_FS_RS)
	if Request("currentValue")<>"" then
		parentID=NoSqlHack(Request("currentValue"))
	elseif Request("topic2")<>"" then
		parentID=NoSqlHack(Request("topic2"))
	elseif Request("topic1")<>"" then
		parentID=NoSqlHack(Request("topic1"))
	elseif Request("topic")<>"" then
		parentID=NoSqlHack(Request("topic"))
	end if 
	if request.Form("userGroup")="" and Request.Form("corpGroup")<>"" then
		AppointUserGroup=NoSqlHack(Request.Form("corpGroup"))
	elseif request.Form("userGroup") <> "" and Request.Form("corpGroup")="" then
		AppointUserGroup=NoSqlHack(Request.Form("userGroup"))
	elseif request.Form("userGroup") <> "" and Request.Form("corpGroup")<>"" then
		AppointUserGroup=NoSqlHack(request.Form("userGroup"))&","&NoSqlHack(Request.Form("corpGroup"))
	end if
	OperateDebateRs.open "select title,content,ParentID,ClassID,addtime,AppointUserNumber,AppointUserGroup,isLock,AccessFile from FS_ME_GroupDebate",User_Conn,1,3
	OperateDebateRs.addNew
	OperateDebateRs("title")=NoSqlHack(Request.Form("title"))
	OperateDebateRs("content")=NoSqlHack(Request.Form("Content"))
	OperateDebateRs("ParentID")=parentID
	OperateDebateRs("ClassID")=NoSqlHack(Request.Form("DebateClass"))
	OperateDebateRs("addtime")=Now()
	if Request.Form("AppointUserNumber")<>"" then
		OperateDebateRs("AppointUserNumber")=NoSqlHack(Request.Form("AppointUserNumber"))
	end if
	if AppointUserGroup<>"" then
	OperateDebateRs("AppointUserGroup")=AppointUserGroup
	end if
	OperateDebateRs("islock")=NoSqlHack(Request.Form("islock"))
	OperateDebateRs("AccessFile")=NoSqlHack(Request.Form("AccessFile"))
	OperateDebateRs.update
	OperateDebateRs.close
	if err.number>0 then
		Response.Redirect("../error.asp?ErrCodes="&err.description&"&ErrorUrl=./user/GroupDebate_manage.asp")
		Response.End()
	else
		Response.Redirect("../success.asp?ErrCodes=<li>操作成功</li>&ErrorUrl=./user/GroupDebate_manage.asp")
		Response.End()
	end if
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>FoosunCMS</title>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>

<script language="JavaScript" src="lib/UserJS.js" type="text/JavaScript"></script>

<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css"
    rel="stylesheet" type="text/css">
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" scroll="yes"
    oncontextmenu="return false;">
    <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
        <%
	if Request.QueryString("Act")="edit" then
	Response.Write("<form action='?Act=EditDebate&DebateID="&NoSqlHack(DebateID)&"' method='post' name='operateDebate' id='operateDebate'>")
	elseif Request.QueryString("Act")="add" or Request.QueryString("Act")="AddChild" then
	Response.Write("<form action='?Act=AddDebate' method='post' name='operateDebate' id='operateDebate'>")
	end if
        %>
        <tr class="hback">
            <td align="right" class="xingmu">
                <div align="left">
                    社群管理</div>
            </td>
            <td align="right" class="xingmu">
                <div align="right">
                    <a href="GroupDebate_manage.asp">返回</a></div>
            </td>
        </tr>
        <tr class="hback">
            <td align="right">
                社群主题：</td>
            <td width="631">
                <input name="Title" type="text" id="Title" value="<%=title%>" size="50" />
                <font color="#FF0000">*</font><span id="Title_Alert"></span>
            </td>
        </tr>
        <tr class="hback">
            <td align="right">
                讨论内容：
            </td>
            <td>
                <textarea name="Content" cols="80" rows="10" id="Content"><%=content%></textarea>
                <font color="#FF0000">*</font><span id="Content_Alert"></span></td>
        </tr>
        <tr class="hback">
            <td align="right">
                所属主题：</td>
            <td>
                <%
	if Request.QueryString("act")<>"AddChild" then
		Set TopicRs=User_Conn.Execute("Select DebateID,Title,ParentID from FS_ME_GroupDebate where ParentID=0 order by addtime desc")
		Response.Write("<select name='Topic' id='topic' onchange=""getDebate(this,'Topic1')"">"&Chr(10)&Chr(13))
		Response.Write("<option value=''>请选择一级主题</option>")
		Response.Write("<option value='0'>一级主题</option>")
		while not TopicRs.eof
			Response.Write("<option value='"&TopicRs("DebateID")&"'>"&TopicRs("title")&"</option>")
			TopicRs.movenext
		wend
		Response.Write("</select>")
	end if
                %>
                <span id="Topic1_Container"></span>&nbsp; <span id="Topic2_Container"></span><span
                    id="current">
                    <%
	Response.Write("<select name='currentValue' id='currentValue' size='1' multiple>")
	if Request.QueryString("act")="AddChild" then
		Set currentTopicRs=User_Conn.Execute("Select DebateID,title from FS_ME_Groupdebate where  DebateID="&CintStr(Request.QueryString("DebateID")))
		if not currentTopicRs.eof then
			ParentID=currentTopicRs("DebateID")
			if ParentID=0 then
				Response.Write("<option value='0' selected>一级主题</option>")
			else
				Response.Write("<option value='"&ParentID&"' selected>"&currentTopicRs("title")&"</option>")
			end if
		end if
	else
		if ParentID=0 then
			Response.Write("<option value='0' selected>一级主题</option>")
		else
			Response.Write("<option value='"&ParentID&"' selected>"&User_Conn.execute("select title from FS_ME_Groupdebate where DebateID="&CintStr(ParentID))(0)&"</option>")
		end if
	end if
	Response.Write("</select>")
                    %>
                </span>
            </td>
        </tr>
        <tr class="hback">
            <td align="right">
                所属分类：</td>
            <td>
                <%	
	  	i=0
		Set ClassRs=User_Conn.Execute("Select ClassID,Title from FS_ME_GroupDebateClass")
		Response.Write("<select name='DebateClass'>"&Chr(10)&Chr(13))
		while not ClassRs.eof
			if classid=ClassRs("ClassID") then
				Response.Write("<option value='"&ClassRs("ClassID")&"' selected>"&ClassRs("title")&"</option>")
			else
				Response.Write("<option value='"&ClassRs("ClassID")&"'>"&ClassRs("title")&"</option>")
			end if
			ClassRs.movenext
		wend
                %>
            </td>
        </tr>
        <tr class="hback">
            <td align="right">
                指定用户查看：</td>
            <td>
                <textarea name="AppointUserNumber" cols="80" id="AppointUserNumber" onKeyUp="ReplaceDot('AppointUserNumber')"><%=AppointUserNumber%></textarea>
            </td>
        </tr>
        <tr class="hback">
            <td align="right">
                指定用户组查看：</td>
            <td>
                <select name="userGroup" size="10" multiple>
                    <option value="user_all" style="background-color: #014952; color: #FFFFFF;">-个人用户组-</option>
                    <%
		  	Set UserGroupRs=User_Conn.Execute("Select GroupID,GroupName from FS_ME_Group where GroupType=1")
			while not UserGroupRs.eof
				selectedTF=false
				if  not isNull(AppointUserGroup) then
					for ArrayCount=0 to Ubound(AppointUserGroup)
						if not isnumeric(AppointUserGroup(ArrayCount)) then exit for
						if Cint(trim(AppointUserGroup(ArrayCount)))=UserGroupRs("GroupID") then
							selectedTF=true
							exit for
						end if
					next
				end if
				if selectedTF then
					Response.Write("<option value='"&UserGroupRs("GroupID")&"' selected>"&UserGroupRs("GroupName")&"</option>")
				else
					Response.Write("<option value='"&UserGroupRs("GroupID")&"'>"&UserGroupRs("GroupName")&"</option>")
				end if
				UserGroupRs.movenext
			wend
                    %>
                </select>
                <select name="corpGroup" size="10" multiple>
                    <option value="corp_all" style="background-color: #014952; color: #FFFFFF;">-企业用户组-</option>
                    <%
			Set UserGroupRs=User_Conn.Execute("Select GroupID,GroupName from FS_ME_Group where GroupType=0")
			while not UserGroupRs.eof
				selectedTF=false
				if  not isNull(AppointUserGroup) then
					for ArrayCount=0 to Ubound(AppointUserGroup)
						if not isnumeric(AppointUserGroup(ArrayCount)) then exit for
						if Cint(trim(AppointUserGroup(ArrayCount)))=UserGroupRs("GroupID") then
							selectedTF=true
							exit for
						end if
					next
				end if
				if selectedTF then
					Response.Write("<option value='"&UserGroupRs("GroupID")&"' selected>"&UserGroupRs("GroupName")&"</option>")
				else
					Response.Write("<option value='"&UserGroupRs("GroupID")&"'>"&UserGroupRs("GroupName")&"</option>")
				end if
				UserGroupRs.movenext
			wend
                    %>
                </select>
            </td>
        </tr>
        <tr class="hback">
            <td align="right">
                锁定：</td>
            <td>
                <%
			if Request.QueryString("Act")="edit" then
				if isLock=1 then
					Response.Write("<input type='radio' name='isLock' value='1' checked>是"&Chr(10)&Chr(13))					
					Response.Write("<input type='radio' name='isLock' value='0' >否"&Chr(10)&Chr(13))
				elseif isLock=0 then
					Response.Write("<input type='radio' name='isLock' value='1' >是"&Chr(10)&Chr(13))					
					Response.Write("<input type='radio' name='isLock' value='0' checked>否"&Chr(10)&Chr(13))
				end if
			else
				Response.Write("<input type='radio' name='isLock' value='1' >是"&Chr(10)&Chr(13))					
				Response.Write("<input type='radio' name='isLock' value='0' checked>否"&Chr(10)&Chr(13))
			end if
                %>
            </td>
        </tr>
        <tr class="hback">
            <td align="right">
                附件地址：</td>
            <td>
                <input name="AccessFile" type="text" id="AccessFile" size="50" value="<%=AccessFile%>"><span
                    id="AccessFile_Alert"></span></td>
        </tr>
        <tr class="hback">
            <td align="right">&nbsp;
                </td>
            <td>
                <input type="Button" name="OperateDebateButton" value=" 保存 " onClick="OperateDebateSubmit()" />
                <input type="reset" name="Submit2" value=" 重置 " /></td>
        </tr>
        </form>
    </table>
</body>
<%
if Request.QueryString("Act")="edit" then
	TopicRs.close
	Set TopicRs=nothing
	ClassRs.close
	Set ClassRs=nothing
	OperateDebateRs.close
	Set OperateDebateRs=nothing
	UserGroupRs.close
	Set UserGroupRs=nothing
	Conn.close
	Set Conn=nothing
	User_Conn.close
	Set User_Conn=nothing
elseif  Request.QueryString("Act")="EditDebate" then
	parentIDRs.close
	Set parentIDRs=nothing
elseif Request.QueryString("Act")="AddChile" then
	currentTopicRs.close
	Set currentTopicRs=nothing
end if
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
function getDebate(Obj,Str)
{
	var debateID=Obj.value;
	if(isNaN(debateID))
	{
		document.getElementById("Topic1_Container").innerHTML="";
		document.getElementById("Topic2_Container").innerHTML="";
		return ;
	}else if(debateID==0)
	{
		document.getElementById("Topic1_Container").innerHTML="";
		document.getElementById("Topic2_Container").innerHTML="";
		document.getElementById("current").innerHTML="";
		return ;
	}
	document.getElementById("current").innerHTML="";
	var url="getGroupDebate.asp?name="+Str+"&debateid="+debateID+"&r="+Math.random();//构造url
	request.open("GET",url,true);//建立连接
	request.onreadystatechange = getDebateResult;
	request.send(null);//传送数据，因为数据通过url传递了，所以这里传递的是null
}
function getDebateResult()//当服务器响应的时候就使用这个方法
{
	if(request.readyState ==4)//根据HTTP 就绪状态判断响应是否完成
	{
		if(request.status == 200)//判断请求是否成功
		{
			result=request.responseText;//获得响应的结果，也就是新的<select>
			if(result.indexOf("name='Topic1'")>0)
			{
				document.getElementById("Topic1_Container").innerHTML="|&nbsp;&nbsp;"+result;//将这个结果现实在客户端
				document.getElementById("Topic2_Container").innerHTML=""
			}
			if(result.indexOf("name='Topic2'")>0)
			{
				document.getElementById("Topic2_Container").innerHTML="|&nbsp;&nbsp;"+result;//将这个结果现实在客户端
			}
		}
	}
}
function OperateDebateSubmit()
{	
	var flag1=isEmpty('Title','Title_Alert');
	var flag2=isEmpty('Content','Content_Alert');
	var flag3=isEmpty('AccessFile','AccessFile_Alert');
	if(flag1&&flag2&&flag3)
	{
		document.operateDebate.submit();
	}
}
</script>

</html>
