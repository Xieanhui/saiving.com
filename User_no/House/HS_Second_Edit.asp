<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="lib/cls_second.asp"-->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
Dim straction
straction = Request("action")
if straction="Unmessage" then
	User_Conn.execute("update FS_ME_Users set ismessage= 0 where UserNumber='"& Fs_User.UserNumber &"'")
	strShowErr = "<li>���ı�վ����ȡ��</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../main.asp")
	Response.end
Elseif straction = "ismessage" then
	User_Conn.execute("update FS_ME_Users set ismessage= 1 where UserNumber='"& Fs_User.UserNumber &"'")
	strShowErr = "<li>���ı�վ���ϳɹ�</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../main.asp")
	Response.end
Elseif straction = "Open" then
	User_Conn.execute("update FS_ME_Users set isOpen= 1 where UserNumber='"& Fs_User.UserNumber &"'")
	strShowErr = "<li>���⿪�����Ͽ���</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../main.asp")
	Response.end
Elseif straction = "Close" then
	User_Conn.execute("update FS_ME_Users set isOpen= 0 where UserNumber='"& Fs_User.UserNumber &"'")
	strShowErr = "<li>���⿪������ȡ��</li>"
	Response.Redirect("lib/Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=../main.asp")
	Response.end
End if
%>
<%
'on error resume next
Dim second_rs,action,str_CurrPath,SecondObj,sid,i,tmpArray
set SecondObj=New cls_Second
action=request.QueryString("action")
sid=CintStr(request.QueryString("id"))
if action="edit" then
	SecondObj.getSecondInfo(sid)
End if
str_CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<script language="javascript" src="../../FS_Inc/CheckJs.js"></script>
<script language="javascript" src="../../FS_Inc/prototype.js"></script>
<script language="javascript" src="../../FS_Inc/PublicJS.js"></script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
	<tr>
		<td>
			<!--#include file="../top.asp" -->
		</td>
	</tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
	<tr class="back">
		<td   colspan="2" class="xingmu" height="26">
			<!--#include file="../Top_navi.asp" -->
		</td>
	</tr>
	<tr class="back">
		<td width="18%" valign="top" class="hback">
			<div align="left">
				<!--#include file="../menu.asp" -->
			</div>
		</td>
		<td width="82%" valign="top" class="hback">
			<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table">
				<tr>
					<td class="xingmu">�������ַ���Ϣ</td>
				</tr>
				<tr>
					<td class="xingmu"><a href="HS_Second.asp">���ַ���Ϣ</a>&nbsp;|&nbsp;<a href='#' onClick="history.back()">����</a></td>
				</tr>
			</table>
			<form action="HS_Second_Action.asp?action=<%=action%>" method="post" id="TenancyForm">
				<input type="hidden" value="<%=sid%>" name="id"/>
				<table width="98%" border="0" align="center" cellpadding="4" cellspacing="1" class="table" id="inputpanel">
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td width="17%" align="right" class="hback">��;��</td>
						<td width="83%" align="left" class="hback">
							<select name="sel_usefor" id="sel_usefor">
								<option value="1" <%if SecondObj.UseFor="1" then response.Write("selected")%>>ס��</option>
								<option value="2" <%if SecondObj.UseFor="2" then response.Write("selected")%>>д�ּ�</option>
							</select>
							<select name="sel_class" id="sel_class">
								<option value="1" <%if SecondObj.sclass="1" then response.Write("selected")%>>����</option>
								<option value="2" <%if SecondObj.sclass="2" then response.Write("selected")%>>�н�</option>
							</select>
							&nbsp;��Ȩ���ʣ�
							<select name="sel_BelongType" id="sel_BelongType">
								<option value="1" <%if SecondObj.BelongType="1" then response.Write("selected")%>>��Ʒ��</option>
								<option value="2" <%if SecondObj.BelongType="2" then response.Write("selected")%>>���ʷ�</option>
							</select>
							</span></td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">���ݱ�ţ�</td>
						<td align="left" class="hback">
							<input name="txt_Label" maxlength="50" type="text" id="txt_Label" style="width:60%" value="<%=secondObj.label%>"  onFocus="this.className='RightInput'"/>
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">��Դ��ַ��</td>
						<td align="left" class="hback">
							<input name="txt_address" maxlength="100" type="text" id="txt_address" style="width:60%" value="<%=secondObj.address%>"  onFocus="this.className='RightInput'"/>
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">�ضΣ�</td>
						<td align="left" class="hback">
							<input name="txt_Position" maxlength="100" type="text" id="txt_Position" style="width:60%" value="<%=secondObj.Position%>" onFocus="this.className='RightInput'"/>
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">���أ�</td>
						<td align="left" class="hback">
							<input name="txt_cityarea" maxlength="20" type="text" id="txt_cityarea" style="width:60%" value="<%=secondObj.cityarea%>" onFocus="this.className='RightInput'"/>
						</td>
					</tr>
					<%
'���ʹ���
Dim HouseStyleArray(3)
HouseStyleArray(0)="1"
HouseStyleArray(1)="1"
HouseStyleArray(2)="1"
if SecondObj.HouseStyle<>"" then
	tmpArray=split(SecondObj.HouseStyle,",")
	for i=0 to Ubound(tmpArray)
		if i>2 then exit for
		HouseStyleArray(i)=tmpArray(i)
	next
End if
%>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">���ͣ�</td>
						<td align="left" class="hback">
							<input name="txt_HouseStyle_1" type="text" id="txt_HouseStyle_1" size="15"  maxlength="6" onFocus="Do.these('txt_HouseStyle_1',function(){return isNumber('txt_HouseStyle_1','span_housestyle','����д��ȷ�ĸ�ʽ',true);})"
  onkeyup="Do.these('txt_HouseStyle_1',function(){return isNumber('txt_HouseStyle_1','span_housestyle','����д��ȷ�ĸ�ʽ',true);})" value="<%=HouseStyleArray(0)%>"/>
							��
							<input name="txt_HouseStyle_2" type="text" id="txt_HouseStyle_2" size="15" maxlength="6" onFocus="Do.these('txt_HouseStyle_2',function(){return isNumber('txt_HouseStyle_2','span_housestyle','����д��ȷ�ĸ�ʽ',true);})" onKeyUp="Do.these('txt_HouseStyle_2',function(){return isNumber('txt_HouseStyle_2','span_housestyle','����д��ȷ�ĸ�ʽ',true);})" value="<%=HouseStyleArray(1)%>"/>
							��
							<input name="txt_HouseStyle_3" type="text" id="txt_HouseStyle_3" size="15" maxlength="6" onFocus="Do.these('txt_HouseStyle_3',function(){return isNumber('txt_HouseStyle_3','span_housestyle','����д��ȷ�ĸ�ʽ',true);})" onKeyUp="Do.these('txt_HouseStyle_3',function(){return isNumber('txt_HouseStyle_3','span_housestyle','����д��ȷ�ĸ�ʽ',true);})" value="<%=HouseStyleArray(2)%>"/>
							�� <span id="span_housestyle"></span></td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">���������</td>
						<td align="left" class="hback">
							<input name="txt_BuildDate" maxlength="4" type="text" id="txt_BuildDate" style="width:60%"  onFocus="Do.these('txt_BuildDate',function(){ return isNumber('txt_BuildDate','span_BuildDate','����д��ȷ�ĸ�ʽ',true);})" onKeyUp="Do.these('txt_BuildDate',function(){ return isNumber('txt_BuildDate','span_BuildDate','����д��ȷ�ĸ�ʽ',true);})" value="<%if secondObj.BuildDate<>"" then Response.Write(secondObj.BuildDate) else Response.Write("0")%>"/>
							<span id="span_BuildDate"></span> 
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">�����</td>
						<td align="left" class="hback">
							<input name="txt_Area" type="text" id="txt_Area" style="width:60%" maxlength="10" onFocus="Do.these('txt_Area',function(){ return isNumber('txt_Area','span_area','����д��ȷ�ĸ�ʽ',false);})" onKeyUp="Do.these('txt_Area',function(){ return isNumber('txt_Area','span_area','����д��ȷ�ĸ�ʽ',false);})" value="<%if secondObj.Area<>"" then Response.Write(secondObj.Area) else Response.Write("0")%>"/>
							<span id="span_area"></span></td>
					</tr>
					<%
'¥�㴦��
Dim FloorArray(2)
FloorArray(0)="0"
FloorArray(1)="0"
if SecondObj.Floor<>"" then
	tmpArray=split(secondObj.Floor,",")
	for i=0 to Ubound(tmpArray)
		if i>1 then exit for
		FloorArray(i)=tmpArray(i)
	next
End if

%>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">¥�㣺</td>
						<td align="left" class="hback">
							<select name="sel_FloorType" id="sel_FloorType">
								<option value="1" >�� ��</option>
								<option value="2">�� ��</option>
								<option value="3">С�߲�</option>
							</select>
							&nbsp;��
							<input name="txt_Floor_1" type="text" id="txt_Floor_1" maxlength="6" size="15" onFocus="Do.these('txt_Floor_1',function(){ return isNumber('txt_Floor_1','span_floor','����д��ȷ�ĸ�ʽ',true);})" onKeyUp="Do.these('txt_Floor_1',function(){ return isNumber('txt_Floor_1','span_floor','����д��ȷ�ĸ�ʽ',true);})" value="<%=FloorArray(0)%>"/>
							�㣬λ��
							<input name="txt_Floor_2" type="text" id="txt_Floor_2" maxlength="6" size="15" onFocus="Do.these('txt_Floor_2',function(){ return isNumber('txt_Floor_2','span_floor','����д��ȷ�ĸ�ʽ',true);})" onKeyUp="Do.these('txt_Floor_2',function(){ return isNumber('txt_Floor_2','span_floor','����д��ȷ�ĸ�ʽ',true);})" value="<%=FloorArray(1)%>"/>
							��<span id="span_floor"></span></td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">�ṹ��</td>
						<td align="left" class="hback">
							<select name="sel_Structure" id="sel_Structure">
								<option value="1" <%if secondObj.Structure="l" then Response.Write("selected")%>>�ֽ������</option>
								<option value="2" <%if secondObj.Structure="2" then Response.Write("selected")%>>ש��ṹ</option>
								<option value="3" <%if secondObj.Structure="3" then Response.Write("selected")%>>שľ�ṹ</option>
								<option value="4" <%if secondObj.Structure="4" then Response.Write("selected")%>>���׽ṹ</option>
							</select>
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">������ʩ��</td>
						<td align="left" class="hback">
							<input name="chk_equip" type="checkbox" id="chk_equip" value="l" <%if instr(secondObj.equip,"l")>0 then Response.Write("checked")%> />
							ˮ
							<input name="chk_equip" type="checkbox" id="chk_equip" value="m" <%if instr(secondObj.equip,"m")>0 then Response.Write("checked")%> />
							��
							<input name="chk_equip" type="checkbox" id="chk_equip" value="n" <%if instr(secondObj.equip,"n")>0 then Response.Write("checked")%> />
							��
							<input name="chk_equip_x" type="checkbox" id="chk_equip_x" value="x" <%if instr(secondObj.equip,"x")>0 then Response.Write("checked")%> />
							�绰
							<input name="chk_equip" type="checkbox" id="chk_equip" value="y" <%if instr(secondObj.equip,"y")>0 then Response.Write("checked")%> />
							����
							<input name="chk_equip" type="checkbox" id="chk_equip" value="z" <%if instr(secondObj.equip,"z")>0 then Response.Write("checked")%>  />
							���</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">װ�������</td>
						<td align="left" class="hback">
							<select name="sel_Decoration" id="sel_Decoration">
								<option value="1" <%if secondObj.Decoration="1" then Response.Write("selected")%>>��װ��</option>
								<option value="2" <%if secondObj.Decoration="2" then Response.Write("selected")%>>�е�װ��</option>
								<option value="3" <%if secondObj.Decoration="3" then Response.Write("selected")%>>�ߵ�װ��</option>
								<option value="4" <%if secondObj.Decoration="4" then Response.Write("selected")%>>û��װ��</option>
							</select>
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">���ۣ�</td>
						<td align="left" class="hback">
							<input name="txt_Price" type="text" id="txt_Price" style="width:60%" maxlength="24" onFocus="Do.these('txt_Price',function(){ return isNumber('txt_Price','span_price','����д��ȷ�ĸ�ʽ',false);})" onKeyUp="Do.these('txt_Price',function(){ return isNumber('txt_Price','span_price','����д��ȷ�ĸ�ʽ',false);})" value="<%if secondObj.price<>"" then response.Write(secondObj.price) else Response.Write("0")%>"/>
							��<span id="span_price"></span></td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">��ϵ�ˣ�</td>
						<td align="left" class="hback">
							<input name="txt_LinkMan" maxlength="100" type="text" id="txt_LinkMan" style="width:60%" onFocus="this.className='RightInput'" value="<%=secondObj.linkMan%>"/>
						</td>
					</tr>
					<tr onMouseOver="overColor(this)" onMouseOut="outColor(this)">
						<td align="right" class="hback">��ϵ��ʽ��</td>
						<td align="left" class="hback">
							<input name="txt_Contact" maxlength="100" type="text" id="txt_Contact" style="width:60%" onFocus="this.className='RightInput'" value="<%=secondObj.Contact%>"/>
						</td>
					</tr>
					<tr onMouseOver="overColor(this)" onMouseOut="outColor(this)">
						<td align="right" class="hback">��ע��</td>
						<td align="left" class="hback">
							<textarea name="txt_Remark" maxlength="250" cols="80" rows="15" id="txt_Remark" onFocus="this.className='RightInput'"><%=secondObj.remark%></textarea>
						</td>
					</tr>
					<tr id="tr_pic_num" onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">ͼƬ������</td>
						<td align="left" class="hback">
							<input name="txt_PicNum" type="text" maxlength="4" id="txt_PicNum" style="width:60%" value="<%=secondObj.picNumber%>" onFocus="if(this.value==''){$('pictable').style.display='none';return;};Do.these('txt_PicNum',function(){ return isNumber('txt_PicNum','span_picNum','����д��ȷ�ĸ�ʽ',true);});"    onKeyUp="Do.these('txt_PicNum',function(){ return isNumber('txt_PicNum','span_picNum','����д��ȷ�ĸ�ʽ',true);});if(!isNaN(this.value)){addRow(this.value);$('pictable').style.display=''};" onBlur="if(this.value==''){$('txt_PicNum').value='0'};Do.these('txt_PicNum',function(){ return isNumber('txt_PicNum','span_picNum','����д��ȷ�ĸ�ʽ',true);})"/>
							<span id="span_picNum"></span></td>
					</tr>
					<tr class="hback" onMouseOver=overColor(this) onMouseOut=outColor(this) >
						<td colspan="2">
							<table width="100%" border="0" cellpadding="0" cellspacing="1" id="pictable" style="display:none">
								<%
	if isnumeric(secondObj.picNumber) then
		if Cint(secondObj.picNumber)>0 and sid<>"" then
			Dim pic_rs,picArray
			Set pic_rs=Conn.execute("select Pic from FS_HS_Picture where HS_type=3 and id="&CintStr(sid))
			for i=0 to secondObj.picNumber-1
				if pic_rs.eof then exit for
				Response.Write("<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>"&vbcrlf)
				Response.Write("<td width='15%' align='right'>ͼƬ"&(i+1)&"</td>"&vbcrlf)
				response.Write("<td width='85%' align='left'>"&vbcrlf)
				if not pic_rs.eof then
					Response.Write("<input name='txt_PicNum_"&(i+1)&"' type='text' id='txt_PicNum' style='width:75%' value='"&pic_rs("pic")&"'/><input type='button' name='PPPChoose"&(i+1)&"' value='ѡ��ͼƬ' onClick=""OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath="&str_CurrPath&"',500,300,window,this.parentNode.firstChild);"">"&vbcrlf)
				else
					Response.Write("<input name='txt_PicNum_"&(i+1)&"' type='text' id='txt_PicNum' style='width:75%'/><input type='button' name='PPPChoose"&(i+1)&"' value='ѡ��ͼƬ' onClick=""OpenWindowAndSetValue('../CommPages/SelectManageDir/SelectPic.asp?CurrPath="&str_CurrPath&"',500,300,window,this.parentNode.firstChild);"">"&vbcrlf)
				End if
				Response.Write("</td>"&vbcrlf)
				Response.Write("</tr>"&vbcrlf)
				pic_rs.movenext
			next
		End if
	End if
%>
							</table>
						</td>
					</tr>
					<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
						<td align="right" class="hback">&nbsp;</td>
						<td align="left" class="hback">
							<input type="Button" name="mysubmit" onClick="CheckThisForm(this.form)" value="�ύ" />
							<input type="Button" name="myrest" value="����" onClick="javascript:if(confirm('�Ƿ��������б��'))$('TenancyForm').reset()" />
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
	<tr class="back">
		<td height="20"  colspan="2" class="xingmu">
			<div align="left">
				<!--#include file="../Copyright.asp" -->
			</div>
		</td>
	</tr>
</table>
</body>
</html>
<%
Set SecondObj = Nothing
%>
<script language="javascript">
<!--
function CheckThisForm(formobj){
	var inputcheck;
	inputcheck=isNumber('txt_HouseStyle_1','span_housestyle','����д��ȷ�ĸ�ʽ',true);
	inputcheck=inputcheck&&isNumber('txt_HouseStyle_2','span_housestyle','����д��ȷ�ĸ�ʽ',true);
	inputcheck=inputcheck&&isNumber('txt_HouseStyle_3','span_housestyle','����д��ȷ�ĸ�ʽ',true);
	inputcheck=inputcheck&&isNumber('txt_BuildDate','span_BuildDate','����д��ȷ�ĸ�ʽ',true);
	inputcheck=inputcheck&&isNumber('txt_Area','span_area','����д��ȷ�ĸ�ʽ',false);
	inputcheck=inputcheck&&isNumber('txt_Floor_1','span_floor','����д��ȷ�ĸ�ʽ',true);
	inputcheck=inputcheck&&isNumber('txt_Floor_2','span_floor','����д��ȷ�ĸ�ʽ',true);
	inputcheck=inputcheck&&isNumber('txt_Price','span_price','����д��ȷ�ĸ�ʽ',false);
	inputcheck=inputcheck&&isNumber('txt_PicNum','span_picNum','����д��ȷ�ĸ�ʽ',true);
	if (inputcheck){
		formobj.submit();
	}else{
		alert("����д��ȷ�ĸ�ʽ");
	}
}
if(!isNaN($F('txt_PicNum')))
{
	$('pictable').style.display="";
}else
{
	$('pictable').style.display="none";
}
function addRow(number)
{
	var oBody=$("pictable");
	oBody.firstChild.innerText="";
	for(var i=0;i<number;i++)
	{
		var myTR =oBody.insertRow();
		for(var j=0;j<2;j++) 
		{
			var myTD=myTR.insertCell();
			myTD.className="hback"
			if(j==0)
			{
				myTD.style.width="15%"
				myTD.align="right"
				myTD.innerHTML="ͼƬ"+(i+1).toString()+"��";
			}
			else
			{
				myTD.style.width="80%"
				myTD.align="left"
				myTD.innerHTML="<input name='txt_PicNum_"+(i+1)+"' type='text' id='txt_PicNum' style='width:75%'/><input type='button' name='PPPChoose"+(i+1)+"' value='ѡ��ͼƬ' onClick=\"OpenWindowAndSetValue('../CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,this.parentNode.firstChild);\">";
			}
		}
	}
}
-->
</script>
</html>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->






