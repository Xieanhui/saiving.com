<%@ codepage=936%>
<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_Page.asp" -->
<!--#include file="lib/cls_main.asp" -->

<%
	response.buffer=true	
	Response.CacheControl = "no-cache"
	Dim Conn,User_Conn,Fs_news,i
	Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo,strpage
	int_RPP=20 '����ÿҳ��ʾ��Ŀ
	int_showNumberLink_=10 '���ֵ�����ʾ��Ŀ
	showMorePageGo_Type_ = 1 '�������˵���������ֵ��ת������ε���ʱֻ��ѡ1
	str_nonLinkColor_="#999999" '����������ɫ
	toF_="<font face=webdings title=""��ҳ"">9</font>"  			'��ҳ 
	toP10_=" <font face=webdings title=""��ʮҳ"">7</font>"			'��ʮ
	toP1_=" <font face=webdings title=""��һҳ"">3</font>"			'��һ
	toN1_=" <font face=webdings title=""��һҳ"">4</font>"			'��һ
	toN10_=" <font face=webdings title=""��ʮҳ"">8</font>"			'��ʮ
	toL_="<font face=webdings title=""���һҳ"">:</font>"
	MF_Default_Conn
	MF_User_Conn
	MF_Session_TF 
	strpage=request("page")
	if len(strpage)=0 Or not isnumeric(strpage) or trim(strpage)=""Then:strpage="1":end if
	if clng(strpage)<1 then strpage=1
	set Fs_news = new Cls_News
	
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ŀ����___Powered by foosun Inc.</title>
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/Prototype.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	<tr class="hback">
		<td class="xingmu"><a href="#" class="sd"><strong>��Ŀ����</strong></a><a href="../../help?Lable=NS_Class_Action" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></td>
	</tr>
	<tr>
		<td height="18" class="hback">
			<div align="left"><a href="Class_Manage.asp">������ҳ</a>��<a href="Class_add.asp?ClassID=&Action=add">��Ӹ���Ŀ</a>��<a href="Class_Action.asp?Action=one">һ����Ŀ����</a>��<a href="Class_Action.asp?Action=n">N����Ŀ����</a>��<a href="Class_Action.asp?Action=reset"   onClick="{if(confirm('ȷ�ϸ�λ������Ŀ��\n\n���ѡ��ȷ�������е���Ŀ������Ϊһ������!!')){return true;}return false;}">��λ������Ŀ</a>��<a href="Class_Action.asp?Action=unite">��Ŀ�ϲ�</a>��<a href="Class_Action.asp?Action=allmove">��Ŀת��</a> �� <a href="Class_Action.asp?Action=clearClass"  onClick="{if(confirm('ȷ�����������Ŀ���������\n\n���ѡ��ȷ��,���е���Ŀ�����Ž����ŵ�����վ��!!')){return true;}return false;}">ɾ��������Ŀ</a>��<a href="../../help?Lable=NS_Class_Action_1" target="_blank" style="cursor:help;'" class="sd"><img src="../Images/_help.gif" border="0"></a></div>
		</td>
	</tr>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class="table">
	<form name="ClassForm" method="post" action="Class_Action.asp">
		<tr class="xingmu">
			<td width="5%" height="25" class="xingmu">
				<div align="center">ID</div>
			</td>
			<td width="30%" class="xingmu">
				<div align="center">��Ŀ����[��ĿӢ��]</div>
			</td>
			<td width="7%" class="xingmu">
				<div align="center">Ȩ��</div>
			</td>
			<td width="22%" class="xingmu">
				<div align="center">����</div>
			</td>
			<td width="36%" class="xingmu">
				<div align="center">����</div>
			</td>
		</tr>
		<%
	Dim AndSQL
	AndSQL = GetAndSQLOfSearchClass("NS013")
	
	Dim obj_news_rs,obj_news_rs_1,isUrlStr
	Set obj_news_rs = server.CreateObject(G_FS_RS)
	obj_news_rs.Open "Select Orderid,id,ClassID,ClassName,ClassEName,IsUrl,isConstr,isShow,[Domain],FileExtName,isPop from FS_NS_NewsClass where Parentid  = '0'   and ReycleTF=0 " & AndSQL & " Order by Orderid desc,ID desc",Conn,1,1
	if obj_news_rs.eof then
	   Response.Write"<TR  class=""hback""><TD colspan=""5""  class=""hback"" height=""40"">û�м�¼��</TD></TR>"
	else
		obj_news_rs.PageSize=int_RPP
		cPageNo=NoSqlHack(Request.QueryString("Page"))
		If cPageNo="" Then cPageNo = 1
		If not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo>obj_news_rs.PageCount Then cPageNo=obj_news_rs.PageCount 
		If cPageNo<=0 Then cPageNo=1
		obj_news_rs.AbsolutePage=cPageNo
		for i=1 to int_RPP
			if obj_news_rs.eof Then exit For 
	%>
		<tr  onMouseOver=overColor(this) onMouseOut=outColor(this)>
			<td height="22" class="hback">
				<div align="center">
					<% = obj_news_rs("id")%>
				</div>
			</td>
			<td height="22" class="hback" style="white-space:normal; word-break:break-all;overflow:hidden;">
				<% 
		if obj_news_rs("IsUrl") = 1 then
			isUrlStr = ""
		Else
			isUrlStr = "["&obj_news_rs("ClassEName")&"]"
		End if
		Set obj_news_rs_1 = server.CreateObject(G_FS_RS)
		obj_news_rs_1.Open "Select Count(ID) from FS_NS_NewsClass where ParentID='"& obj_news_rs("ClassID") &"'",Conn,1,1
		If Get_SubPop_TF(obj_news_rs("Classid"),"NS016","NS","news") then
	        if obj_news_rs_1(0)>0 then
		        Response.Write  "<a href=""javascript:void(0);"" onclick=""getChildClassID('"&obj_news_rs("ClassID")&"','')"" title=""չ������""><img border=""0"" src=""images/jia.gif""></img></a><a href=Class_add.asp?ClassID="&obj_news_rs("ClassID")&"&Action=edit>"&obj_news_rs("ClassName")&"</a>&nbsp;<font style=""font-size:11.5px;"">"& isUrlStr &"</font>"
	        Else
		        Response.Write  "<img src=""images/-.gif""></img><a href=Class_add.asp?ClassID="&obj_news_rs("ClassID")&"&Action=edit>"&obj_news_rs("ClassName")&"</a>&nbsp;<font style=""font-size:11.5px;"">"& isUrlStr &"</font>"
	        End if
	    Else
	        if obj_news_rs_1(0)>0 then
		        Response.Write  "<a href=""javascript:void(0);"" onclick=""getChildClassID('"&obj_news_rs("ClassID")&"','')"" title=""չ������""><img border=""0"" src=""images/jia.gif""></img></a>"&obj_news_rs("ClassName")&"&nbsp;<font style=""font-size:11.5px;"">"& isUrlStr &"</font>"
	        Else
		        Response.Write  "<img src=""images/-.gif""></img>"&obj_news_rs("ClassName")&"&nbsp;<font style=""font-size:11.5px;"">"& isUrlStr &"</font>"
	        End if
	    End If
		obj_news_rs_1.close:set obj_news_rs_1 =nothing
		%>
			</td>
			<td class="hback" align="center">
				<% = obj_news_rs("OrderID")%>
			</td>
			<td class="hback">
				<div align="center">
					<%
	  if obj_news_rs("IsUrl") = 1 then
		  Response.Write("<font color=red>�ⲿ</font>&nbsp;��&nbsp;")
	  ElseIf obj_news_rs("IsUrl") = 2 then
		  Response.Write("<font color=red>��ҳ</font>&nbsp;��&nbsp;")
	  Else
		  Response.Write("ϵͳ&nbsp;��&nbsp;")
	  End if
	  if obj_news_rs("isConstr") = 1 then
		  Response.Write("<font color=red>��</font>&nbsp;��&nbsp;")
	  Else
		  Response.Write("<strike>��</strike>&nbsp;��&nbsp;")
	  End if
	  if obj_news_rs("isShow") = 1 then
		  Response.Write("<font color=red>��ʾ</font>&nbsp;��&nbsp;")
	  Else
		  Response.Write("����&nbsp;��&nbsp;")
	  End if
	  if len(obj_news_rs("Domain")) >5 then
		  Response.Write("<font color=red>��</font>&nbsp;��&nbsp;")
	  Else
		  Response.Write("��&nbsp;��&nbsp;")
	  End if
	  %>
				</div>
			</td>
			<td class="hback">
		        <div align="center">
				<a href="NewClass_review.asp?id=<% = obj_news_rs("ClassID")%>" target="_blank">Ԥ��</a>
			<% If Get_SubPop_TF(obj_news_rs("Classid"),"NS016","NS","class") then%>
				��<a href="Class_add.asp?ClassID=<% = obj_news_rs("ClassID")%>&Action=add">�������Ŀ</a>
			<%else %>
			 ��<% = GetDisableSpanCode("�������Ŀ") %>
			<%end if %>
			<% If Get_SubPop_TF(obj_news_rs("Classid"),"NS017","NS","class") then%>
				��<a href="Class_add.asp?ClassID=<% = obj_news_rs("ClassID")%>&Action=edit">�޸�</a>
			<%else %>
			 ��<% = GetDisableSpanCode("�޸�") %>
			 <%end if %>
			<% If Get_SubPop_TF(obj_news_rs("Classid"),"NS023","NS","class") then%>
				��<a href="Class_Action.asp?ClassID=<% = obj_news_rs("ClassID")%>&Action=clear" onClick="{if(confirm('ȷ����մ���Ŀ����Ϣ��?')){return true;}return false;}">���</a>
			<%else %>
			 ��<% = GetDisableSpanCode("���") %>
			 <%end if %>
			<% If Get_SubPop_TF(obj_news_rs("Classid"),"NS021","NS","class") then%>
				��<a href="Class_Action.asp?ClassID=<% = obj_news_rs("ClassID")%>&Action=del"  onClick="{if(confirm('ȷ��ɾ������ѡ�����Ŀ��?\n\n����Ŀ�µ�����Ҳ����ɾ��!!')){return true;}return false;}">ɾ��</a>
			<%else %>
			��<% = GetDisableSpanCode("ɾ��") %>
			<%end if %>
			<% If Get_SubPop_TF(obj_news_rs("Classid"),"NS022","NS","class") then%>
				��<a href="Class_makerss.asp?signxml=one&cid=<%= obj_news_rs("Classid")%>" title="���ɴ���ĿRss">Rss</a>
			<%else %>
			��<% = GetDisableSpanCode("Rss") %>
			<%end if %>
					<input name="Cid" type="checkbox" id="Cid" value="<% = obj_news_rs("ClassID")%>">
				</div>
			</td>
		</tr>
        <tr  class="hback"><td colspan="5"><div  id="c_<%=obj_news_rs("ClassID")%>"></div></td></tr>
		<%
		'Response.Write(Fs_news.GetChildNewsList(obj_news_rs("ClassID"),""))
		obj_news_rs.MoveNext
	next
%>
		<tr>
			<td height="30" colspan="5" class="hback">
				<div align="right">
					<input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">
					ѡ������
					<input name="Action" type="hidden" id="Action">
					<input type="button" name="Submit" value="ɾ��"  onClick="document.ClassForm.Action.value='DelClass';{if(confirm('ȷ���������ѡ��ļ�¼��')){this.document.ClassForm.submit();return true;}return false;}">
					<input type="button" name="Submit" value="����RSS"  onClick="document.ClassForm.Action.value='makeXML';{if(confirm('ȷʵҪ����RSS��')){this.document.ClassForm.submit();return true;}return false;}">
					<input type="button" name="Submit" value="����������ĿRSS"  onClick="location='Class_makerss.asp?signxml=All';">
				</div>
			</td>
		</tr>
		<tr>
			<td height="30" colspan="5" class="hback">
				<%
		response.Write "<p>"&  fPageCount(obj_news_rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)
	%>
			</td>
		</tr>
		<%end if%>
	</form>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	<tr>
		<td height="18" class="hback">
			<div align="left">
				<p><span class="tx"><strong>˵��</strong></span>:<br>
					ϵͳ -------ϵͳĿ¼<br>
					�ⲿ-------�ⲿ��Ŀ <br>
					��--------��Ŀ����Ͷ��<br>
					��ʾ----��������ʾ<br>
					����----���������� <br>
					��--------�����˶���������Ŀ¼</p>
			</div>
		</td>
	</tr>
</table>
</body>
</html>
<%
'obj_news_rs.close
'set obj_news_rs =nothing
set Fs_news = nothing
%>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = ClassForm.elements[i];  
    if (e.name != 'chkall')  
       e.checked = ClassForm.chkall.checked;  
    }  
  }
  function getChildClassID(cid,oooo)
  {
    	var  options={  
				   method:'get',  
				   parameters:"cc="+oooo+"&classid="+cid,
				   onComplete:function(transport)
					{  
						var returnvalue=transport.responseText;
						
						//alert(returnvalue)
						if (returnvalue.indexOf("??")>-1)
						{
                                return false;
                         }
						else
						{
						    document.getElementById("c_"+cid).innerHTML="<table style='width:100%;' cellpadding=\"0\" cellspacing=\"1\"  class=\"table1\">"+returnvalue+"</table>";
						}
					}  
			   }; 
	    new  Ajax.Request('class_ajax.asp?no-cache='+Math.random(),options);    
  }
</script>



	


