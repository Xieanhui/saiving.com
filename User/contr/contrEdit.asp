<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<!--#include file="lib/cls_contr.asp"-->
<%
Dim str_CurrPath,action,id,contrObj
Dim GetSysRs,EditTF
Set contrObj=New cls_Contr
action=NoSqlHack(request.QueryString("action"))
id=CintStr(request.QueryString("id"))
if id<>"" then contrObj.getContrInfo(id)

Set GetSysRs = Conn.ExeCute("Select top 1 IsEditFileTF From FS_NS_SysParam Where SysID > 0 Order By SysID")
If GetSysRs.Eof Then
	EditTF = 1
Else
	EditTF = Cint(GetSysRs(0))
End If
GetSysRs.Close : set GetSysRs = Nothing		
If EditTF = 0 Then
	If contrObj.AuditTF = 1 Then
		Response.Redirect("../lib/error.asp?ErrCodes=<li><font color=red>�Ѿ���˵�Ͷ�岻�����ٱ༭</font></li>&ErrorUrl=")
	End If
End if		
if action="" then
	action="add"
End if

str_CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>��ӭ�û�<%=Fs_User.UserName%>����<%=GetUserSystemTitle%></title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head>
<script language="javascript" src="../../FS_Inc/prototype.js"></script>
<script language="javascript" src="../../FS_Inc/PublicJS.js"></script>
<script language="javascript" src="../../FS_Inc/CheckJs.js"></script>
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
			<table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
				<tr>
					<td class="hback_1"><a href="contrManage.asp">�������</a> | <font color="#FF0000">�༭���</font></td>
				</tr>
				<tr>
					<td>
						<form name="contrForm" action="contrAction.asp?action=<%=action%>&id=<%=id%>" method="post">
							<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
								<tr>
									<td align="right" class="hback" width="15%">���⣺</td>
									<td class="hback">
										<input type="text" name="txt_ContTitle" id="txt_ContTitle" value="<%=contrObj.ContTitle%>" style="width:60%"
				onfocus="Do.these('txt_ContTitle',function(){ return isEmpty('txt_ContTitle','span_title')})"
				onKeyUp="Do.these('txt_ContTitle',function(){ return isEmpty('txt_ContTitle','span_title')})"
			/>
										<span id="span_title"></span></td>
								</tr>
								<tr>
									<td align="right" class="hback">�����⣺</td>
									<td align="left" class="hback">
										<input type="text" name="txt_subTitle" style="width:60%" value="<%=contrObj.subTitle%>" onFocus="this.className='RightInput'">
									</td>
								</tr>
								<tr class="hback">
									<td align="right">���ģ�</td>
									<td aling="left">
										<!--�༭����ʼ-->
										<iframe id='NewsContent' src='../Editer/UserEditer.asp?id=txt_content&CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>' frameborder=0 scrolling=no width='100%' height='480'></iframe>
										<input type="hidden" name="txt_content" value="<% = HandleEditorContent(contrObj.ContContent) %>">
										<!--�༭������-->
										<span id="span_content"></span></td>
								</tr>
								<tr>
									<td align="right" class="hback">��ע��</td>
									<td align="left" class="hback">
										<textarea name="txt_OtherContent" rows="5" style="width:80%" onFocus="this.className='RightInput'"><%=contrObj.OtherContent%></textarea>
									</td>
								</tr>
								<tr>
									<%
				Dim classRs,ClassCName,MainClassName
				if contrObj.ClassID<>"" then
					Set ClassRs=User_Conn.execute("select ClassCName from FS_ME_InfoClass where classid="&contrObj.ClassID)
					if not ClassRs.eof then
						ClassCName=ClassRs("ClassCName")
					Else
						ClassCName=""
					End if
				Else
					ClassCName=""
				End if
				if contrObj.MainID<>"" then
					Set ClassRs=Conn.execute("select ClassName from FS_NS_NewsClass where id="&contrObj.MainID&" and isConstr=1")
					if not ClassRs.eof then
						MainClassName=ClassRs("ClassName")
					Else
						MainClassName=""
					End if
				Else
					MainClassName=""
				End if
				if isNull(classRs) then classRs.close:set classRs=nothing
			%>
									<td align="right" class="hback">��Ϣ���ࣺ</td>
									<td align="left" class="hback">
										<input type="text" name="txt_Class" style="width:60%" readonly  value="<%=ClassCName%>"/>
										<input type="hidden" name="txt_ClassID" style="width:60%" value="<%=contrObj.ClassID%>">
										<button onClick="SelectClass()">ѡ�����</button>
									<span id="span_class"><font color="#FF0000">*</font></span></td>
								</tr>
								<tr>
									<td align="right" class="hback">��վ���ࣺ</td>
									<td align="left" class="hback">
										<input type="text" name="txt_mainClass" style="width:60%" readonly value="<%=MainClassName%>" />
<%
Dim tmpRs,Str_MainID
if contrObj.MainID<>"" then
	Set tmpRs=Conn.execute("select Classid from FS_NS_NewsClass where id="&contrObj.MainID&" and isConstr=1") 
	if not tmpRs.eof then
		Str_MainID=tmpRs(0)
	End if
	tmpRs.close
else
	Str_MainID=""
end if
Set tmpRs=nothing
%>
										<input type="hidden" name="txt_mainClassID" style="width:60%" value="<%=Str_MainID%>">
										<button onClick="SelectMainClass()">ѡ�����</button>
										<span id="span_Mainclass"><font color="#FF0000">*</font></span></td>
								</tr>
								<tr>
									<td align="right" class="hback">�ؼ��֣�</td>
									<td align="left" class="hback">
										<textarea name="txt_KeyWords" id="txt_KeyWords" rows="5" style="width:80%" onFocus="this.className='RightInput'" onKeyUp="ReplaceDot('txt_KeyWords')"><%=contrObj.KeyWords%></textarea>
									</td>
								</tr>
								<tr>
									<td align="right" class="hback">��������վ��</td>
									<td align="left" class="hback">
										<p>
											<label>
											<input type="radio" name="rad_IsPublic" value="1" <%if contrObj.IsPublic="1" then Response.Write("checked")%>>
											��</label>
											<label>
											<input type="radio" name="rad_IsPublic" value="0" <%if contrObj.IsPublic<>"1"  then Response.Write("checked")%>>
											��</label>
										</p>
									</td>
								</tr>
								<tr>
									<td align="right" class="hback">��Ϣ����</td>
									<td align="left" class="hback">
										<select name="sel_InfoType">
											<option value="0" <%if contrObj.InfoType="0" then Response.Write("selected")%>>��ͨ</option>
											<option value="1" <%if contrObj.InfoType="1" then Response.Write("selected")%>>����</option>
											<option value="2" <%if contrObj.InfoType="2" then Response.Write("selected")%>>�Ӽ�</option>
										</select>
									</td>
								</tr>
								<tr>
									<td align="right" class="hback">���ͣ�</td>
									<td align="left" class="hback">
										<select name="sel_ContSytle" id="sel_ContSytle">
											<option value="0" <%if contrObj.ContSytle="0" then Response.Write("selected")%>>ԭ��</option>
											<option value="1" <%if contrObj.ContSytle="1" then Response.Write("selected")%>>ת��</option>
											<option value="3" <%if contrObj.ContSytle="3" then Response.Write("selected")%>>����</option>
										</select>
									</td>
								</tr>
								<tr>
									<td align="right" class="hback">������</td>
									<td align="left" class="hback">
										<p>
											<label>
											<input type="radio" name="rad_IsLock" value="1" <%if contrObj.IsLock="1" then Response.Write("checked")%>>
											��</label>
											<label>
											<input type="radio" name="rad_IsLock" value="0" <%if contrObj.IsLock<>"1" then Response.Write("checked")%>>
											��</label>
										</p>
									</td>
								</tr>
								<tr>
									<td align="right" class="hback">�Ƽ���</td>
									<td align="left" class="hback">
										<label>
										<input type="radio" name="rad_isTF" value="1" <%if contrObj.isTF="1" then Response.Write("checked")%>>
										��</label>
										<label>
										<input type="radio" name="rad_isTF" value="0" <%if contrObj.isTF<>"1" then Response.Write("checked")%>>
										��</label>
									</td>
								</tr>
								<tr>
									<td align="right" class="hback">ͼƬ��</td>
									<td align="left" class="hback">
										<input type="text" name="txt_img" id="txt_img" value="<%=contrObj.PicFile%>" style="width:60%">
										<button onClick="javascript:OpenWindowAndSetValue('<%=Add_Root_Dir(G_USER_DIR)%>/CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,$('txt_img'));">ѡ��ͼƬ</button>
									</td>
								</tr>
								<tr>
									<td align="right" class="hback">&nbsp;</td>
									<td align="left" class="hback">
										<input type="Button" name="contr_Submit" value="�ύ" onClick="mySubmit(this.form)">
										<input type="Button" name="contr_reset" value="����" onClick="javascript:if(confirm('�Ƿ��������'))$('contrForm').reset()">
									</td>
								</tr>
							</table>
						</form>
					</td>
				</tr>
			</table>
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
Set Fs_User = Nothing
Set User_Conn=nothing
Set Conn=nothing
%>
<script type="text/javascript">
//�����վ����
function SelectMainClass()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('lib/SelectClassFrame.asp?rad'+Math.random(),400,300,window);
	if (ReturnValue.indexOf('***')!=-1)
	{
		TempArray = ReturnValue.split('***');
		$('txt_mainClassID').value=TempArray[0]
		$('txt_mainClass').value=TempArray[1]
	}
}
//��÷���
function SelectClass()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('lib/SelectMyClassFrame.asp?rad'+Math.random(),400,300,window);
	if (ReturnValue.indexOf('***')!=-1)
	{
		TempArray = ReturnValue.split('***');
		$('txt_ClassID').value=TempArray[0]
		$('txt_Class').value=TempArray[1]
	}
}
//�ύ��
function mySubmit(FormObj)
{
	var flag1=isEmpty('txt_ContTitle','span_title');
	var flag3=isEmpty('txt_Class','span_class');
	var flag4=isEmpty('txt_mainClass','span_Mainclass');
	if(flag1&&flag3&&flag4)
	{
		if (frames["NewsContent"].g_currmode!='EDIT') {alert('����ģʽ���޷����棬���л������ģʽ');return false;}
		FormObj.txt_content.value=frames["NewsContent"].GetNewsContentArray();
		FormObj.submit();
	}
}

</script>
<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->
