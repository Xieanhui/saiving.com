<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_Inc/Cls_Cache.asp" -->
<!--#include file="../FS_Inc/Func_page.asp" -->
<%
Dim Conn,strShowErr,rs,ar_admin,str_PopList
MF_Default_Conn
MF_Session_TF

'��ҳʹ��Ȩ���ж�
If Not MF_Check_Pop_TF("MF_Pop") Then Err_Show
Set rs= Server.CreateObject(G_FS_RS)
rs.open "select Admin_Add_Admin,Admin_Is_Super,Admin_Name,Admin_Pop_List,Admin_Parent_Admin From FS_MF_Admin where Id="&CintStr(Request.QueryString("AdminId")),Conn,1,3
ar_admin = rs(2)
str_PopList=rs(3)
If session("Admin_Is_Super")<>1 Then
	if rs("Admin_Name")<>session("Admin_Name") Then
		if rs("Admin_Is_Super")=1 Then
			strShowErr = "<li>�����ܲ���ϵͳ����Ա</li>"
			Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
			Response.End
		End If
		If rs("Admin_Parent_Admin")<>session("Admin_Name") Then
			strShowErr = "<li>�����ܲ������˵��ӹ���Ա��</li>"
			Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=SysAdmin_list.asp")
			Response.End
		End If
	End If
End If
'-----------------
rs.close:Set rs=Nothing
If Request.Form("Action")="Save" Then
	if trim(Request.Form("AdminId"))="" Then
			strShowErr = "<li>����Ĳ�����</li>"
			Response.Redirect("error.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	Else
		Dim pop_rs
		Set rs= Server.CreateObject(G_FS_RS)
		rs.open "select Admin_Pop_List,Admin_PopGroup From FS_MF_Admin where id="&CintStr(Request.Form("AdminId")),conn,1,3
		Select Case int(Request.Form("Group_n"))
			Case 1
				Set pop_rs = Conn.execute("select PopList From FS_MF_AdminGroup where Levels=1")
				If Not pop_rs.eof Then
					rs("Admin_Pop_List")=pop_rs("PopList")
				End If
				pop_rs.close:Set pop_rs = Nothing
			Case 2
				Set pop_rs = Conn.execute("select PopList From FS_MF_AdminGroup where Levels=2")
				If Not pop_rs.eof Then
					rs("Admin_Pop_List")=pop_rs("PopList")
				End If
				pop_rs.close:Set pop_rs = Nothing
			Case 3
				Set pop_rs = Conn.execute("select PopList From FS_MF_AdminGroup where Levels=3")
				If not pop_rs.eof Then
					rs("Admin_Pop_List")=pop_rs("PopList")
				End If
				pop_rs.close:Set pop_rs = Nothing
			Case 4
				Set pop_rs = Conn.execute("select PopList From FS_MF_AdminGroup where Levels=4")
				If Not pop_rs.eof Then
					rs("Admin_Pop_List")=pop_rs("PopList")
				End If
				pop_rs.close:Set pop_rs = Nothing
			Case 5
				Set pop_rs = Conn.execute("select PopList From FS_MF_AdminGroup where Levels=5")
				If Not pop_rs.eof Then
					rs("Admin_Pop_List")=pop_rs("PopList")
				End If
				pop_rs.close:Set pop_rs = Nothing
			Case 0
				rs("Admin_Pop_List")= Replace((NoSqlHack(Request.Form("PopList_MF")))&","&NoSqlHack(Request.Form("PopList_NS"))&","&NoSqlHack(Request.Form("PopList_MS"))&","&NoSqlHack(Request.Form("PopList_DS"))&","&NoSqlHack(Request.Form("PopList_ME"))&","&NoSqlHack(Request.Form("PopList_AP"))&","&NoSqlHack(Request.Form("PopList_SD"))&","&NoSqlHack(Request.Form("PopList_CS"))&","&NoSqlHack(Request.Form("PopList_SS"))&","&NoSqlHack(Request.Form("PopList_HS"))&","&NoSqlHack(Request.Form("PopList_VS"))&","&NoSqlHack(Request.Form("PopList_AS"))&","&NoSqlHack(Request.Form("PopList_WS"))&","&NoSqlHack(Request.Form("PopList_FL"))&","&NoSqlHack(Request.Form("PopList"))," ","")
		End Select
		rs("Admin_PopGroup")=int(Request.Form("Group_n"))
		rs.update
		rs.close:Set rs=Nothing
		strShowErr = "<li>����Ȩ�޳ɹ�</li>"
		Response.Redirect("Success.asp?ErrCodes="&Server.URLEncode(strShowErr)&"&ErrorUrl=")
	End If
	Response.End
End If
Dim str_ClassList
'�����Ŀ���档�����´δ�ʱ����ٶȡ�Fsj.08.11.16

Function get_ChildClassList(TypeID,CompatStr)
		Dim ChildTypeListRs,ChildTypeListStr,TempStr
		Set ChildTypeListRs = Conn.execute("Select ParentID,ClassID,ClassName from FS_NS_NewsClass where ParentID='" & TypeID & "' and ReycleTF=0 order by OrderID desc,id desc")
		TempStr = CompatStr & "��"
		do while Not ChildTypeListRs.Eof
			get_ChildClassList = get_ChildClassList &"<input type=""checkbox"" onClick=""ClassListC(this.name)"" name=""ClassList"" value="""&ChildTypeListRs("ClassId")&""" />" & TempStr
			get_ChildClassList = get_ChildClassList & ChildTypeListRs("ClassName")&"<br>"
			get_ChildClassList = get_ChildClassList & get_ChildClassList(ChildTypeListRs("ClassID"),TempStr)
			ChildTypeListRs.MoveNext
		loop
		ChildTypeListRs.Close:Set ChildTypeListRs = Nothing
End Function

'if ""=application("ClassList_NS") then
    set rs=Server.CreateObject(G_FS_RS)
    rs.open "select ClassId,ClassName From FS_NS_NewsClass where ParentId = '0' and ReycleTF=0 order by OrderId desc,id desc",Conn	
    do while not rs.eof
        str_ClassList = str_ClassList & "<input type=""checkbox"" onClick=""ClassListC(this.name)"" name=""ClassList"" value="""&rs("ClassId")&""" />"&rs("ClassName")&"<br>"
        str_ClassList = str_ClassList & get_ChildClassList(rs("ClassId"),"��")
        rs.movenext
    loop
    rs.close:set rs=nothing
    application.Lock()
    application("ClassList_NS")=str_ClassList
    application.UnLock()
'else
'    str_ClassList=application("ClassList_NS")
'end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>FoosunCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
</HEAD>
<script language="JavaScript" src="../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/Prototype.js" type="text/JavaScript"></script>
<script language="JavaScript" src="../FS_Inc/CheckJs.js" type="text/JavaScript"></script>
<script language="javascript" type="text/javascript" src="../../FS_Inc/wz_tooltip.js"></script>
<script language="JavaScript" type="text/JavaScript">
function SwitchPopType(Html_express)
{
	var Str_express=Html_express.substring(0,2);
	if ($("PopList_"+Str_express).checked==true)
	{
		$(Str_express+'_ID').style.display='';
	}
	else
	{
		$(Str_express+'_ID').style.display='none';
	}
	return true;
}
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = PopForm.elements[i];
    if (e.name != 'chkall')
       {
	   e.checked = PopForm.chkall.checked;
	   document.getElementById('MF_ID').style.display='';
	   }
	}
	}
function ChooseadminType(Type)
{
	switch (Type)
	{
	case "0":
		document.getElementById('all_id').style.display='';
	break;
	default:
		document.getElementById('all_id').style.display='none';
	break;
	}
}
var thisPop="";
function UpdateNsClass(Obj)
{
	var PopId=Obj.id;
	PopId=PopId.substring(6,PopId.length);
	if (thisPop!=PopId)
	{
		if ($(PopId).checked)
		{
			var Popvalue=$(PopId).value;
			var checkarr="";
			var str_Classlist=Popvalue.substring(8,Popvalue.length);
			SelectClear(document.getElementsByName("ClassList"));
			if (str_Classlist!="")
			{
				checkarr=str_Classlist.split("|");
				for (var j=0;j<checkarr.length;j++)
				{
					if (checkarr[j]!="")
					{
						Selectone(document.getElementsByName("ClassList"),checkarr[j])
					}
				}
			}
			else
			{
				SelectClear(document.getElementsByName("ClassList"));
			}
			$("ClassListarea").style.display="";
			if (thisPop!="")
			{
				$("Update"+thisPop).innerHTML="������Ŀ";
				$("Update"+thisPop).style.color="#000000";
			}
			thisPop=PopId;
			Obj.innerHTML="�������";
			Obj.style.color="#FF0000";
		}
		else
		{
			if (thisPop!="")
			{
				$("Update"+thisPop).innerHTML="������Ŀ";
				$("Update"+thisPop).style.color="#000000";
			}
			thisPop="";
			$("ClassListarea").style.display="none";
			alert("���빴ѡ�˹��ܺ����������Ŀ");
		}
	}
	else
	{
		Obj.innerHTML="������Ŀ";
		Obj.style.color="#000000";
		thisPop="";
		$("ClassListarea").style.display="none";
	}
}
function Selectone(Obj,value)
{
	for (var i=0 ;i<Obj.length ;i++)
	{
		if (Obj[i].value==value)
			Obj[i].checked=true;
	}
}
function SelectClear(Obj)
{
	for (var i=0 ;i<Obj.length ;i++)
	{
			Obj[i].checked=false;
	}
}
function ClassListC(Obj_Name)
{
	var Values=GetSelectValue(Obj_Name);
	if (thisPop!="" && Values!="")
	{
		$(thisPop).value=thisPop+"$$$"+Values;
	}
}
function GetSelectValue(Obj_Name)
{
	var ReturnValue="";
	var Objs=document.getElementsByName(Obj_Name);
	for (var i=0;i<Objs.length;i++)
	{
		if (Objs[i].checked)
		{
			if (ReturnValue=="")
			{
				ReturnValue=Objs[i].value;
			}else
			{
				ReturnValue+="|"+Objs[i].value;
			}
		}
	}
	return ReturnValue;
}
function Selectall(Obj,flag)
{
	for (var i=0 ;i<Obj.length ;i++)
	{
		Obj[i].checked=flag;
	}
}
var IsSelectedAll=false;
function ChangeAllStatus(Obj)
{
	IsSelectedAll=!IsSelectedAll;
	Selectall(document.getElementsByName("ClassList"),IsSelectedAll);
	Obj.value=(IsSelectedAll)?"ȡ��ȫѡ��Ŀ":"ѡ��������Ŀ";
	ClassListC("ClassList");
}
</script>
<BODY>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="xingmu">
    <td colspan="2" class="xingmu"><a href="#" class="sd"><strong>���ù���Ա</strong></a><span class="tx"><b><%=ar_admin%></b></span>��Ȩ��</td>
  </tr>
  <tr class="hback">
    <td colspan="2"><a href="SysAdmin_List.asp">���ع���Ա�б�</a><a href="SysAdmin_List.asp?my=1"></a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <form name="PopForm" method="post" action="">
    <tr>
      <td width="13%" class="hback"><div align="right">
          <input name="AdminId" type="hidden" id="AdminId" value="<%=Request.QueryString("AdminId")%>">
          <input name="Action" type="hidden" id="Action" value="Save">
          ����Ա����</div></td>
      <td width="32%" class="hback"><select name="Group_n" onChange="ChooseadminType(this.options[this.selectedIndex].value);">
          <option value="1">��������Ա</option>
          <option value="2">��ͨ����Ա</option>
          <option value="3">�ܱ༭</option>
          <option value="4">���α༭</option>
          <option value="5">����</option>
          <option value="0" selected>�Զ������Ա</option>
        </select>
        <input type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form);">
        ѡ��/ȡ������</td>
      <td width="55%" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="4%"><img src="../sys_images/icon.gif" width="18" height="16"></td>
            <td width="96%"><a href="SysAdmin_SetPop_Group.asp">����̶�����ԱȨ��</a></td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td colspan="3" class="hback"><table width="100%" border="0" cellspacing="1" cellpadding="5" style="display:;" id="all_id">
          <%if Request.Cookies("FoosunSUBCookie")("FoosunSUBMF")=1 then%>
          <tr>
            <td colspan="2" class="hback_1"><div align="left">
                <input name="PopList_MF" type="checkbox" id="PopList" onClick="SwitchPopType('MF_Pop');"   value="MF_Pop" <%if InStr(1,str_PopList,"MF_Pop",1)<>0 then response.Write("checked")%>>
                MF��ϵͳ</div></td>
          </tr>
          <tr>
            <td colspan="2" class="hback"><table width="100%" border="0" cellpadding="4" cellspacing="0" style="display:<%if InStr(1,str_PopList,"MF_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>"  id="MF_ID">
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td width="14%" class="hback"><div align="left">
                      <input name="PopList" type="checkbox"  value="MF_Templet" <%if InStr(1,str_PopList,"MF_Templet",1)<>0 then:response.Write("checked")%>>
                      ģ����� </div></td>
                  <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" value="MF001" <%if InStr(1,str_PopList,"MF001",1)>0 then response.Write "checked"%>>
                          �޸��ļ�
                          <input name="PopList" type="checkbox" id="PopList" value="MF002"  <%if InStr(1,str_PopList,"MF002",1)>0 then response.Write "checked"%>>
                          �����ļ�
                          <input name="PopList" type="checkbox" id="PopList" value="MF003" <%if InStr(1,str_PopList,"MF003",1)>0 then response.Write "checked"%>>
                          ɾ���ļ�
                          <input name="PopList" type="checkbox" id="PopList" value="MF004" <%if InStr(1,str_PopList,"MF004",1)>0 then response.Write "checked"%>>
                          ����Ŀ¼
                          <input name="PopList" type="checkbox" id="PopList" value="MF005" <%if InStr(1,str_PopList,"MF005",1)>0 then response.Write "checked"%>>
                          �����ļ�
                          <input name="PopList" type="checkbox" id="PopList" value="MF006" <%if InStr(1,str_PopList,"MF006",1)>0 then response.Write "checked"%>>
                          ���߱༭(�޸�ģ��) </td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><div align="left">
                      <input name="PopList" type="checkbox" id="PopList" value="MF_Style"  <%if InStr(1,str_PopList,"MF_Style",1)>0 then response.Write "checked"%>>
                      ��ǩ��ʽ</div></td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="MF007" <%if InStr(1,str_PopList,"MF007",1)>0 then response.Write "checked"%>>
                          ������ʽ
                          <input name="PopList" type="checkbox" id="PopList" value="MF008" <%if InStr(1,str_PopList,"MF008",1)>0 then response.Write "checked"%>>
                          �޸���ʽ </td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><div align="left">
                      <input name="PopList" type="checkbox" id="PopList" value="MF_Public" <%if InStr(1,str_PopList,"MF_Public",1)>0 then response.Write "checked"%>>
                      ��������</div></td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="MF009" <%if InStr(1,str_PopList,"MF009",1)>0 then response.Write "checked"%>>
                          ���� </td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><div align="left">
                      <input name="PopList" type="checkbox" id="PopList" value="MF_sPublic" <%if InStr(1,str_PopList,"MF_sPublic",1)>0 then response.Write "checked"%>>
                      ��ǩ����</div></td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="MF025" <%if InStr(1,str_PopList,"MF025",1)>0 then response.Write "checked"%>>
                          ������ǩ </td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><div align="left">
                      <input name="PopList" type="checkbox" id="PopList" value="MF_SysSet" <%if InStr(1,str_PopList,"MF_SysSet",1)>0 then response.Write "checked"%>>
                      ��������</div></td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="MF010" <%if InStr(1,str_PopList,"MF010",1)>0 then response.Write "checked"%>>
                          ���� </td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><div align="left">
                      <input name="PopList" type="checkbox" id="PopList" value="MF_Const" <%if InStr(1,str_PopList,"MF_Const",1)>0 then response.Write "checked"%>>
                      �����ļ�</div></td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="MF011" <%if InStr(1,str_PopList,"MF011",1)>0 then response.Write "checked"%>>
                          ȫ�ֱ�������
                          <input name="PopList" type="checkbox" id="PopList" value="MF012" <%if InStr(1,str_PopList,"MF012",1)>0 then response.Write "checked"%>>
                          �Զ�ˢ�������ļ�</td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><div align="left">
                      <input name="PopList" type="checkbox" id="PopList" value="MF_SubSite" <%if InStr(1,str_PopList,"MF_SubSite",1)>0 then response.Write "checked"%>>
                      ��ϵͳά��</div></td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="MF013" <%if InStr(1,str_PopList,"MF013",1)>0 then response.Write "checked"%>>
                          ���� </td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MF_DataFix" <%if InStr(1,str_PopList,"MF_DataFix",1)>0 then response.Write "checked"%>>
                    ���ݿ�ά��</td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="MF015" <%if InStr(1,str_PopList,"MF015",1)>0 then response.Write "checked"%>>
                          ���ݿ�ѹ��
                          <input name="PopList" type="checkbox" id="PopList" value="MF016" <%if InStr(1,str_PopList,"MF016",1)>0 then response.Write "checked"%>>
                          ���ݿⱸ��
                          <input name="PopList" type="checkbox" id="PopList" value="MF017" <%if InStr(1,str_PopList,"MF017",1)>0 then response.Write "checked"%>>
                          SQL����ѯ����</td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MF_Log" <%if InStr(1,str_PopList,"MF_Log",1)>0 then response.Write "checked"%>>
                    ��־����</td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="MF018" <%if InStr(1,str_PopList,"MF018",1)>0 then response.Write "checked"%>>
                          ������־
                          <input name="PopList" type="checkbox" id="PopList" value="MF019" <%if InStr(1,str_PopList,"MF019",1)>0 then response.Write "checked"%>>
                          ��ȫ��־
                          <input name="PopList" type="checkbox" id="PopList" value="MF020" <%if InStr(1,str_PopList,"MF020",1)>0 then response.Write "checked"%>>
                          ɾ����־</td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MF_Define" <%if InStr(1,str_PopList,"MF_Define",1)>0 then response.Write "checked"%>>
                    �Զ����ֶ�</td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="MF021" <%if InStr(1,str_PopList,"MF021",1)>0 then response.Write "checked"%>>
                          ����
                          <input name="PopList" type="checkbox" id="PopList" value="MF022" <%if InStr(1,str_PopList,"MF022",1)>0 then response.Write "checked"%>>
                          ɾ��</td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MF_CustomForm" <%if InStr(1,str_PopList,"MF_CustomForm",1)>0 then response.Write "checked"%>>
                    �Զ����</td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="MF099" <%if InStr(1,str_PopList,"MF099",1)>0 then response.Write "checked"%>>
                          ����
						  <input name="PopList" type="checkbox" id="PopList" value="MF098" <%if InStr(1,str_PopList,"MF098",1)>0 then response.Write "checked"%>>
                          �޸�
                          <input name="PopList" type="checkbox" id="PopList" value="MF097" <%if InStr(1,str_PopList,"MF097",1)>0 then response.Write "checked"%>>
                          ɾ��
						  <input name="PopList" type="checkbox" id="PopList" value="MF096" <%if InStr(1,str_PopList,"MF096",1)>0 then response.Write "checked"%>>
                          Ԥ��</td>
                      </tr>
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="MF095" <%if InStr(1,str_PopList,"MF095",1)>0 then response.Write "checked"%>>
                          �������
						  <input name="PopList" type="checkbox" id="PopList" value="MF094" <%if InStr(1,str_PopList,"MF094",1)>0 then response.Write "checked"%>>
                          ����
                          <input name="PopList" type="checkbox" id="PopList" value="MF093" <%if InStr(1,str_PopList,"MF093",1)>0 then response.Write "checked"%>>
                          �޸�
						  <input name="PopList" type="checkbox" id="PopList" value="MF092" <%if InStr(1,str_PopList,"MF092",1)>0 then response.Write "checked"%>>
                          ɾ��</td>
                      </tr>
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="MF091" <%if InStr(1,str_PopList,"MF091",1)>0 then response.Write "checked"%>>
                          �����ݹ���
						  <input name="PopList" type="checkbox" id="PopList" value="MF090" <%if InStr(1,str_PopList,"MF090",1)>0 then response.Write "checked"%>>
                          ɾ��
                          <input name="PopList" type="checkbox" id="PopList" value="MF089" <%if InStr(1,str_PopList,"MF089",1)>0 then response.Write "checked"%>>
                          ����
						  <input name="PopList" type="checkbox" id="PopList" value="MF088" <%if InStr(1,str_PopList,"MF088",1)>0 then response.Write "checked"%>>
                          ����
						  <input name="PopList" type="checkbox" id="PopList" value="MF087" <%if InStr(1,str_PopList,"MF087",1)>0 then response.Write "checked"%>>
                          ���ظ�</td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MF_other" <%if InStr(1,str_PopList,"MF_other",1)>0 then response.Write "checked"%>>
                    ����</td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="MF024" <%if InStr(1,str_PopList,"MF024",1)>0 then response.Write "checked"%>>
                          �޸�����
                          <input name="PopList" type="checkbox" id="PopList" value="MF025" <%if InStr(1,str_PopList,"MF025",1)>0 then response.Write "checked"%>>
                          �ϴ��ļ�
                          <input name="PopList" type="checkbox" id="PopList" value="MF026" <%if InStr(1,str_PopList,"MF026",1)>0 then response.Write "checked"%>>
                          �������ļ�/�ļ���
                          <input name="PopList" type="checkbox" id="PopList" value="MF027" <%if InStr(1,str_PopList,"MF027",1)>0 then response.Write "checked"%>>
                          ɾ���ļ�/Ŀ¼
                          <input name="PopList" type="checkbox" id="PopList" value="MF028" <%if InStr(1,str_PopList,"MF028",1)>0 then response.Write "checked"%>>
                          �½�Ŀ¼</td>
                      </tr>
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="MF029" <%if InStr(1,str_PopList,"MF029",1)>0 then response.Write "checked"%>>
                          ����Ա����</td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr>
          <%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBNS")=1 then
		  %>
          <tr  class="hback_1">
            <td colspan="2" class="hback_1"><div align="left"> <strong>
                <input name="PopList_NS" type="checkbox" id="PopList" value="NS_Pop"  onClick="SwitchPopType('NS_Pop');" <%if InStr(1,str_PopList,"NS_Pop",1)>0 then response.Write "checked"%>>
                </strong>NS����ϵͳ</div></td>
          </tr>
          <tr  class="hback">
            <td colspan="2" class="hback"><table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"NS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="NS_ID">
                <tr>
                  <td width="14%" class="hback"><div align="left">
                      <input name="PopList" type="checkbox" id="PopList" value="NS_News" <%if InStr(1,str_PopList,"NS_News",1)>0 then response.Write "checked"%>>
                      ���Ź���</div></td>
                  <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2" <%if InStr(1,str_PopList,"MF010",1)>0 then response.Write "checked"%>>
                      <tr>
                        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td width="30%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr>
                                    <td><input name="PopList" type="checkbox" id="NS013" value="<%= GetNewsPopStr(str_PopList,"NS013") %>" <%if InStr(1,str_PopList,"NS013",1)>0 then response.Write "checked"%> />
                                      <label for="NS013">�鿴</label></td>
                                    <td><button id="UpdateNS013" onClick="UpdateNsClass(this);">������Ŀ</button></td>
                                  </tr>
                                  <tr>
                                    <td width="50%"><input name="PopList" type="checkbox" id="NS001" value="<%= GetNewsPopStr(str_PopList,"NS001") %>" <%if InStr(1,str_PopList,"NS001",1)>0 then response.Write "checked"%> />
                                      <label for="NS001">���</label></td>
                                    <td><button id="UpdateNS001" onClick="UpdateNsClass(this);">������Ŀ</button></td>
                                  </tr>
                                  <tr>
                                    <td><input name="PopList" type="checkbox" id="NS002" value="<%= GetNewsPopStr(str_PopList,"NS002") %>" <%if InStr(1,str_PopList,"NS002",1)>0 then response.Write "checked"%> />
                                      <label for="NS002">�޸�</label></td>
                                    <td><button id="UpdateNS002" onClick="UpdateNsClass(this);">������Ŀ</button></td>
                                  </tr>
                                  <tr>
                                    <td><input name="PopList" type="checkbox" id="NS003" value="<%= GetNewsPopStr(str_PopList,"NS003") %>" <%if InStr(1,str_PopList,"NS003",1)>0 then response.Write "checked"%> />
                                      <label for="NS003">ɾ��</label></td>
                                    <td><button id="UpdateNS003" onClick="UpdateNsClass(this);">������Ŀ</button></td>
                                  </tr>
                                  <tr>
                                    <td><input name="PopList" type="checkbox" id="NS004" value="<%= GetNewsPopStr(str_PopList,"NS004") %>" <%if InStr(1,str_PopList,"NS004",1)>0 then response.Write "checked"%> />
                                      <label for="NS004">���</label></td>
                                    <td><button id="UpdateNS004" onClick="UpdateNsClass(this);">������Ŀ</button></td>
                                  </tr>
                                  <tr>
                                    <td><input name="PopList" type="checkbox" id="NS005" value="<%= GetNewsPopStr(str_PopList,"NS005") %>" <%if InStr(1,str_PopList,"NS005",1)>0 then response.Write "checked"%> />
                                      <label for="NS005">����</label></td>
                                    <td><button id="UpdateNS005" onClick="UpdateNsClass(this);">������Ŀ</button></td>
                                  </tr>
                                  <tr>
                                    <td><input name="PopList" type="checkbox" id="NS007" value="<%= GetNewsPopStr(str_PopList,"NS007") %>" <%if InStr(1,str_PopList,"NS007",1)>0 then response.Write "checked"%> />
                                      <label for="NS007">����</label></td>
                                    <td><button id="UpdateNS007" onClick="UpdateNsClass(this);">������Ŀ</button></td>
                                  </tr>
                                  <tr>
                                    <td><input name="PopList" type="checkbox" id="NS009" value="<%= GetNewsPopStr(str_PopList,"NS009") %>" <%if InStr(1,str_PopList,"NS009",1)>0 then response.Write "checked"%> />
                                      <label for="NS009">�ƶ�</label></td>
                                    <td><button id="UpdateNS009" onClick="UpdateNsClass(this);">������Ŀ</button></td>
                                  </tr>
                                  <tr>
                                    <td><input name="PopList" type="checkbox" id="NS006" value="<%= GetNewsPopStr(str_PopList,"NS006") %>" <%if InStr(1,str_PopList,"NS006",1)>0 then Response.Write "checked"%> />
                                      <label for="NS006">����Ȩ��</label></td>
                                    <td><button id="UpdateNS006" onClick="UpdateNsClass(this);">������Ŀ</button></td>
                                  </tr>
                                  <tr>
                                    <td><input name="PopList" type="checkbox" id="NS010" value="<%= GetNewsPopStr(str_PopList,"NS010") %>" <%if InStr(1,str_PopList,"NS010",1)>0 then response.Write "checked"%> />
                                      <label for="NS010">�����滻</label></td>
                                    <td><button id="UpdateNS010" onClick="UpdateNsClass(this);">������Ŀ</button></td>
                                  </tr>
                                  <tr>
                                    <td><input name="PopList" type="checkbox" id="NS011" value="<%= GetNewsPopStr(str_PopList,"NS011") %>" <%if InStr(1,str_PopList,"NS011",1)>0 then response.Write "checked"%> />
                                      <label for="NS011">����</label></td>
                                    <td><button id="UpdateNS011" onClick="UpdateNsClass(this);">������Ŀ</button></td>
                                  </tr>
                                  <tr>
                                    <td><input name="PopList" type="checkbox" id="NS012" value="<%= GetNewsPopStr(str_PopList,"NS012") %>" <%if InStr(1,str_PopList,"NS012",1)>0 then response.Write "checked"%> />
                                      <label for="NS012">����JS</label></td>
                                    <td><button id="UpdateNS012" onClick="UpdateNsClass(this);">������Ŀ</button></td>
                                  </tr>
                                </table></td>
                              <td>
								<div id="ClassListarea" style="display:none;">
                                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td width="20%" valign="middle" nowrap>
										<div id="ClassListParent" style="height:260px;border:2px inset transparent;overflow:auto;">
									  <%=str_ClassList %>
									   </div>
									  </td>
                                      <td valign="middle"><button onClick="ChangeAllStatus(this);">ѡ��������Ŀ</button></td>
                                    </tr>
                                  </table></div>&nbsp;</td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><div align="left">
                      <input name="PopList" type="checkbox" id="PopList" value="NS_Class" <%if InStr(1,str_PopList,"NS_Class",1)>0 then response.Write "checked"%>>
                      ��Ŀ����</div></td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="NS016" <%if InStr(1,str_PopList,"NS016",1)>0 then response.Write "checked"%>>
                          ���
                          <input name="PopList" type="checkbox" id="PopList" value="NS017" <%if InStr(1,str_PopList,"NS017",1)>0 then response.Write "checked"%>>
                          �޸�
                          <input name="PopList" type="checkbox" id="PopList" value="NS018" <%if InStr(1,str_PopList,"NS018",1)>0 then response.Write "checked"%>>
                          ��λ
                          <input name="PopList" type="checkbox" id="PopList" value="NS019" <%if InStr(1,str_PopList,"NS019",1)>0 then response.Write "checked"%>>
                          �ϲ�
                          <input name="PopList" type="checkbox" id="PopList" value="NS020" <%if InStr(1,str_PopList,"NS020",1)>0 then response.Write "checked"%>>
                          ת��
                          <input name="PopList" type="checkbox" id="PopList" value="NS021" <%if InStr(1,str_PopList,"NS021",1)>0 then response.Write "checked"%>>
                          ɾ��
                          <input name="PopList" type="checkbox" id="PopList" value="NS022" <%if InStr(1,str_PopList,"NS022",1)>0 then response.Write "checked"%>>
                          xml(RSS)
                          <input name="PopList" type="checkbox" id="PopList" value="NS023" <%if InStr(1,str_PopList,"NS023",1)>0 then response.Write "checked"%>>
                          ���
                          <input name="PopList" type="checkbox" id="PopList" value="NS024" <%if InStr(1,str_PopList,"NS024",1)>0 then response.Write "checked"%>>
                          ����</td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_Special" <%if InStr(1,str_PopList,"NS_Special",1)>0 then response.Write "checked"%>>
                    ר�����</td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="NS026" <%if InStr(1,str_PopList,"NS026",1)>0 then response.Write "checked"%>>
                          ���
                          <input name="PopList" type="checkbox" id="PopList" value="NS027" <%if InStr(1,str_PopList,"NS027",1)>0 then response.Write "checked"%>>
                          �޸�
                          <input name="PopList" type="checkbox" id="PopList" value="NS028" <%if InStr(1,str_PopList,"NS028",1)>0 then response.Write "checked"%>>
                          ���� </td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_Constr" <%if InStr(1,str_PopList,"NS_Constr",1)>0 then response.Write "checked"%>>
                    Ͷ�����</td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="NS030" <%if InStr(1,str_PopList,"NS030",1)>0 then response.Write "checked"%>>
                          ���
                          <input name="PopList" type="checkbox" id="PopList" value="NS031" <%if InStr(1,str_PopList,"NS031",1)>0 then response.Write "checked"%>>
                          ����
                          <input name="PopList" type="checkbox" id="PopList" value="NS032" <%if InStr(1,str_PopList,"NS032",1)>0 then response.Write "checked"%>>
                          ɾ��
                          <input name="PopList" type="checkbox" id="PopList" value="NS033" <%if InStr(1,str_PopList,"NS033",1)>0 then response.Write "checked"%>>
                          �˸�
                          <input name="PopList" type="checkbox" id="PopList" value="NS034" <%if InStr(1,str_PopList,"NS034",1)>0 then response.Write "checked"%>>
                          ͳ��</td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_Templet" <%if InStr(1,str_PopList,"NS_Templet",1)>0 then response.Write "checked"%>>
                    ����ģ��</td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="NS036" <%if InStr(1,str_PopList,"NS036",1)>0 then response.Write "checked"%>>
                          ����</td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_Freejs" <%if InStr(1,str_PopList,"NS_Freejs",1)>0 then response.Write "checked"%>>
                    ����JS����</td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="NS037" <%if InStr(1,str_PopList,"NS037",1)>0 then response.Write "checked"%>>
                          ����
                          <input name="PopList" type="checkbox" id="PopList" value="NS038" <%if InStr(1,str_PopList,"NS038",1)>0 then response.Write "checked"%>>
                          �޸�
                          <input name="PopList" type="checkbox" id="PopList" value="NS039" <%if InStr(1,str_PopList,"NS039",1)>0 then response.Write "checked"%>>
                          ɾ�� </td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_Sysjs" <%if InStr(1,str_PopList,"NS_Sysjs",1)>0 then response.Write "checked"%>>
                    ϵͳJS����</td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="NS040" <%if InStr(1,str_PopList,"NS040",1)>0 then response.Write "checked"%>>
                          ����
                          <input name="PopList" type="checkbox" id="PopList" value="NS041" <%if InStr(1,str_PopList,"NS041",1)>0 then response.Write "checked"%>>
                          �޸�
                          <input name="PopList" type="checkbox" id="PopList" value="NS042" <%if InStr(1,str_PopList,"NS042",1)>0 then response.Write "checked"%>>
                          ɾ�� </td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_Recyle" <%if InStr(1,str_PopList,"NS_Recyle",1)>0 then response.Write "checked"%>>
                    ����վ����</td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="NS043" <%if InStr(1,str_PopList,"NS043",1)>0 then response.Write "checked"%>>
                          �ָ�
                          <input name="PopList" type="checkbox" id="PopList" value="NS044" <%if InStr(1,str_PopList,"NS044",1)>0 then response.Write "checked"%>>
                          ɾ�� </td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_UnRl" <%if InStr(1,str_PopList,"NS_UnRl",1)>0 then response.Write "checked"%>>
                    ����������</td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="NS045" <%if InStr(1,str_PopList,"NS045",1)>0 then response.Write "checked"%>>
                          ����
                          <input name="PopList" type="checkbox" id="PopList" value="NS046" <%if InStr(1,str_PopList,"NS046",1)>0 then response.Write "checked"%>>
                          �޸�
                          <input name="PopList" type="checkbox" id="PopList" value="NS047" <%if InStr(1,str_PopList,"NS047",1)>0 then response.Write "checked"%>>
                          ɾ��</td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_Genal" <%if InStr(1,str_PopList,"NS_Genal",1)>0 then response.Write "checked"%>>
                    �������</td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="NS048" <%if InStr(1,str_PopList,"NS048",1)>0 then response.Write "checked"%>>
                          ����</td>
                      </tr>
                    </table></td>
                </tr>
                <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                  <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="NS_Param" <%if InStr(1,str_PopList,"NS_Param",1)>0 then response.Write "checked"%>>
                    ϵͳ��������</td>
                  <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input name="PopList" type="checkbox" id="PopList" value="NS049" <%if InStr(1,str_PopList,"NS049",1)>0 then response.Write "checked"%>>
                          ����</td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr>
          <%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBMS")=1 then
		  %>
          <tr  class="hback_1">
            <td colspan="2" class="hback_1"><div align="left"> <strong>
                <input name="PopList_MS" type="checkbox" id="PopList" value="MS_Pop"  onClick="SwitchPopType('MS_Pop');"  <%if InStr(1,str_PopList,"MS_Pop",1)>0 then response.Write "checked"%>>
                </strong>MS�̳�ϵͳ</div></td>
          </tr>
          <tr class="hback">
            <td colspan="2" class="hback"><div align="left">
                <table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"MS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="MS_ID">
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="MS_Products" <%if InStr(1,str_PopList,"MS_Products",1)>0 then response.Write "checked"%>>
                        ��Ʒ����</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="MS001" <%if InStr(1,str_PopList,"MS001",1)>0 then response.Write "checked"%>>
                            ���
                            <input name="PopList" type="checkbox" id="PopList" value="MS002" <%if InStr(1,str_PopList,"MS002",1)>0 then response.Write "checked"%>>
                            �޸�
                            <input name="PopList" type="checkbox" id="PopList" value="MS003" <%if InStr(1,str_PopList,"MS003",1)>0 then response.Write "checked"%>>
                            ɾ��
                            <input name="PopList" type="checkbox" id="PopList" value="MS004" <%if InStr(1,str_PopList,"MS004",1)>0 then response.Write "checked"%>>
                            ���
                            <input name="PopList" type="checkbox" id="PopList" value="MS005" <%if InStr(1,str_PopList,"MS005",1)>0 then response.Write "checked"%>>
                            ����
                            <input name="PopList" type="checkbox" id="PopList" value="MS006" <%if InStr(1,str_PopList,"MS006",1)>0 then response.Write "checked"%>>
                            �ö�
                            <input name="PopList" type="checkbox" id="PopList" value="MS007" <%if InStr(1,str_PopList,"MS007",1)>0 then response.Write "checked"%>>
                            �ȵ��Ƽ�</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="MS_Class" <%if InStr(1,str_PopList,"MS_Class",1)>0 then response.Write "checked"%>>
                        �����Ŀ</div></td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="MS010" <%if InStr(1,str_PopList,"MS010",1)>0 then response.Write "checked"%>>
                            ���
                            <input name="PopList" type="checkbox" id="PopList" value="MS011" <%if InStr(1,str_PopList,"MS011",1)>0 then response.Write "checked"%>>
                            �޸�
                            <input name="PopList" type="checkbox" id="PopList" value="MS012" <%if InStr(1,str_PopList,"MS012",1)>0 then response.Write "checked"%>>
                            ɾ��
                            <input name="PopList" type="checkbox" id="PopList" value="MS013" <%if InStr(1,str_PopList,"MS013",1)>0 then response.Write "checked"%>>
                            ����
                            <input name="PopList" type="checkbox" id="PopList" value="MS014" <%if InStr(1,str_PopList,"MS014",1)>0 then response.Write "checked"%>>
                            ��λ
                            <input name="PopList" type="checkbox" id="PopList" value="MS015" <%if InStr(1,str_PopList,"MS015",1)>0 then response.Write "checked"%>>
                            �ϲ�
                            <input name="PopList" type="checkbox" id="PopList" value="MS016" <%if InStr(1,str_PopList,"MS016",1)>0 then response.Write "checked"%>>
                            ��� </td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MS_Special" <%if InStr(1,str_PopList,"MS_Special",1)>0 then response.Write "checked"%>>
                      ר������</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="MS017" <%if InStr(1,str_PopList,"MS017",1)>0 then response.Write "checked"%>>
                            ���
                            <input name="PopList" type="checkbox" id="PopList" value="MS018" <%if InStr(1,str_PopList,"MS018",1)>0 then response.Write "checked"%>>
                            �޸�
                            <input name="PopList" type="checkbox" id="PopList" value="MS019" <%if InStr(1,str_PopList,"MS019",1)>0 then response.Write "checked"%>>
                            ɾ��
                            <input name="PopList" type="checkbox" id="PopList" value="MS020" <%if InStr(1,str_PopList,"MS020",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MS_order" <%if InStr(1,str_PopList,"MS_order",1)>0 then response.Write "checked"%>>
                      ��������</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="MS021" <%if InStr(1,str_PopList,"MS021",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MS_bind" <%if InStr(1,str_PopList,"MS_bind",1)>0 then response.Write "checked"%>>
ģ������</td>
                    <td class="hback">&nbsp;</td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MS_WrOut" <%if InStr(1,str_PopList,"MS_WrOut",1)>0 then response.Write "checked"%>>
                      �˻�����</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="MS022" <%if InStr(1,str_PopList,"MS022",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MS_Company" <%if InStr(1,str_PopList,"MS_Company",1)>0 then response.Write "checked"%>>
                      ������˾/�̼�</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="MS023" <%if InStr(1,str_PopList,"MS023",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MS_Recycle" <%if InStr(1,str_PopList,"MS_Recycle",1)>0 then response.Write "checked"%>>
                      ����վ����</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="MS024" <%if InStr(1,str_PopList,"MS024",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="MS_Param" <%if InStr(1,str_PopList,"MS_Param",1)>0 then response.Write "checked"%>>
                      ϵͳ����</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="MS025" <%if InStr(1,str_PopList,"MS025",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
              </div></td>
          </tr>
          <%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBDS")=1 then
		  %>
          <tr  class="hback_1">
            <td colspan="2" class="hback_1"><div align="left"><strong>
                <input name="PopList_DS" type="checkbox" id="PopList_DS" value="DS_Pop"  onClick="SwitchPopType('DS_Pop');" <%if InStr(1,str_PopList,"DS_Pop",1)>0 then response.Write "checked"%>>
                </strong>DS����ϵͳ</div></td>
          </tr>
          <tr class="hback">
            <td colspan="2" class="hback"><div align="left">
                <table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"DS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="DS_ID">
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="Down_List" <%if InStr(1,str_PopList,"Down_List",1)>0 then response.Write "checked"%>>
                        ���ع���</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="DS001" <%if InStr(1,str_PopList,"DS001",1)>0 then response.Write "checked"%>>
                            ���
                            <input name="PopList" type="checkbox" id="PopList" value="DS002" <%if InStr(1,str_PopList,"DS002",1)>0 then response.Write "checked"%>>
                            �޸�
                            <input name="PopList" type="checkbox" id="PopList" value="DS003" <%if InStr(1,str_PopList,"DS003",1)>0 then response.Write "checked"%>>
                            ɾ��
                            <input name="PopList" type="checkbox" id="PopList" value="DS005" <%if InStr(1,str_PopList,"DS005",1)>0 then response.Write "checked"%>>
                            ���� </td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="DS_Class" <%if InStr(1,str_PopList,"DS_Class",1)>0 then response.Write "checked"%>>
                        ��Ŀ����</div></td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="DS010" <%if InStr(1,str_PopList,"DS010",1)>0 then response.Write "checked"%>>
                            ���
                            <input name="PopList" type="checkbox" id="PopList" value="DS011" <%if InStr(1,str_PopList,"DS011",1)>0 then response.Write "checked"%>>
                            �޸�
                            <input name="PopList" type="checkbox" id="PopList" value="DS012" <%if InStr(1,str_PopList,"DS012",1)>0 then response.Write "checked"%>>
                            ɾ��
                            <input name="PopList" type="checkbox" id="PopList" value="DS013" <%if InStr(1,str_PopList,"DS013",1)>0 then response.Write "checked"%>>
                            ����
                            <input name="PopList" type="checkbox" id="PopList" value="DS014" <%if InStr(1,str_PopList,"DS014",1)>0 then response.Write "checked"%>>
                            ��λ
                            <input name="PopList" type="checkbox" id="PopList" value="DS015"  <%if InStr(1,str_PopList,"DS015",1)>0 then response.Write "checked"%>>
                            �ϲ�
                            <input name="PopList" type="checkbox" id="PopList" value="DS016"  <%if InStr(1,str_PopList,"DS016",1)>0 then response.Write "checked"%>>
                            ��� </td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="DS_speical" <%if InStr(1,str_PopList,"DS_Class",1)>0 then response.Write "checked"%>>
                      ר������&nbsp;</td>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="DS018" <%if InStr(1,str_PopList,"DS018",1)>0 then response.Write "checked"%>>
                      ����                          &nbsp;</td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="DS_Param" <%if InStr(1,str_PopList,"DS_Param",1)>0 then response.Write "checked"%>>
                      ��������</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="DS017" <%if InStr(1,str_PopList,"DS017",1)>0 then response.Write "checked"%>>
                            ����
                            <input name="PopList" type="checkbox" id="PopList" value="DS_KunBang" <%if InStr(1,str_PopList,"DS_KunBang",1)>0 then response.Write "checked"%>>
                            ģ������</td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
              </div></td>
          </tr>
          <%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBME")=1 then
		  %>
          <tr  class="hback_1">
            <td colspan="2" class="hback_1"><div align="left"> <strong>
                <input name="PopList_ME" type="checkbox" id="PopList" value="ME_Pop"  onClick="SwitchPopType('ME_Pop');"  <%if InStr(1,str_PopList,"ME_Pop",1)>0 then response.Write "checked"%>>
                </strong>ME��Աϵͳ</div></td>
          </tr>
          <tr class="hback">
            <td colspan="2" class="hback"><div align="left">
                <table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"ME_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="ME_ID">
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="ME_List" <%if InStr(1,str_PopList,"ME_List",1)>0 then response.Write "checked"%>>
                        ��Ա����</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME001" <%if InStr(1,str_PopList,"ME001",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="ME_Intergel" <%if InStr(1,str_PopList,"ME_Intergel",1)>0 then response.Write "checked"%>>
                        ���ֹ���</div></td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME002" <%if InStr(1,str_PopList,"ME002",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Card" <%if InStr(1,str_PopList,"ME_Card",1)>0 then response.Write "checked"%>>
                      �㿨����</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME017" <%if InStr(1,str_PopList,"ME017",1)>0 then response.Write "checked"%>>
                            ���
                            <input name="PopList" type="checkbox" id="PopList" value="ME018" <%if InStr(1,str_PopList,"ME018",1)>0 then response.Write "checked"%>>
                            �޸�
                            <input name="PopList" type="checkbox" id="PopList" value="ME019" <%if InStr(1,str_PopList,"ME019",1)>0 then response.Write "checked"%>>
                            ɾ��</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_News" <%if InStr(1,str_PopList,"ME_News",1)>0 then response.Write "checked"%>>
                      �������</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME021" <%if InStr(1,str_PopList,"ME021",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Form" <%if InStr(1,str_PopList,"ME_Form",1)>0 then response.Write "checked"%>>
                      ��Ⱥ����</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME022" <%if InStr(1,str_PopList,"ME022",1)>0 then response.Write "checked"%>>
                            �½�
                            <input name="PopList" type="checkbox" id="PopList" value="ME023" <%if InStr(1,str_PopList,"ME023",1)>0 then response.Write "checked"%>>
                            �޸�
                            <input name="PopList" type="checkbox" id="PopList" value="ME024" <%if InStr(1,str_PopList,"ME024",1)>0 then response.Write "checked"%>>
                            ɾ��
                            <input name="PopList" type="checkbox" id="PopList" value="ME025"  <%if InStr(1,str_PopList,"ME025",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_HY" <%if InStr(1,str_PopList,"ME_HY",1)>0 then response.Write "checked"%>>
                      ��ҵ����</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME026" <%if InStr(1,str_PopList,"ME026",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_award" <%if InStr(1,str_PopList,"ME_award",1)>0 then response.Write "checked"%>>
                      �齱����</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME027" <%if InStr(1,str_PopList,"ME027",1)>0 then response.Write "checked"%>>
                            �½�
                            <input name="PopList" type="checkbox" id="PopList" value="ME028" <%if InStr(1,str_PopList,"ME028",1)>0 then response.Write "checked"%>>
                            �޸�
                            <input name="PopList" type="checkbox" id="PopList" value="ME029" <%if InStr(1,str_PopList,"ME029",1)>0 then response.Write "checked"%>>
                            ɾ��
                            <input name="PopList" type="checkbox" id="PopList" value="ME030"  <%if InStr(1,str_PopList,"ME030",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Order" <%if InStr(1,str_PopList,"ME_Order",1)>0 then response.Write "checked"%>>
                      ��������(����֧��)</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME031" <%if InStr(1,str_PopList,"ME031",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Mproducts" <%if InStr(1,str_PopList,"ME_Mproducts",1)>0 then response.Write "checked"%>>
                      ��ӻ�Ա��Ʒ</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME032" <%if InStr(1,str_PopList,"ME032",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Horder" <%if InStr(1,str_PopList,"ME_Horder",1)>0 then response.Write "checked"%>>
                      ��������</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME033" <%if InStr(1,str_PopList,"ME033",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_GUser" <%if InStr(1,str_PopList,"ME_GUser",1)>0 then response.Write "checked"%>>
                      ��Ա��</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME034" <%if InStr(1,str_PopList,"ME034",1)>0 then response.Write "checked"%>>
                            �½�
                            <input name="PopList" type="checkbox" id="PopList" value="ME035" <%if InStr(1,str_PopList,"ME035",1)>0 then response.Write "checked"%>>
                            �޸�
                            <input name="PopList" type="checkbox" id="PopList" value="ME036" <%if InStr(1,str_PopList,"ME036",1)>0 then response.Write "checked"%>>
                            ɾ��</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Jubao" <%if InStr(1,str_PopList,"ME_Jubao",1)>0 then response.Write "checked"%>>
                      �ٱ�����</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME037" <%if InStr(1,str_PopList,"ME037",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Review" <%if InStr(1,str_PopList,"ME_Review",1)>0 then response.Write "checked"%>>
                      ���۹���</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME038" <%if InStr(1,str_PopList,"ME038",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Log" <%if InStr(1,str_PopList,"ME_Log",1)>0 then response.Write "checked"%>>
                      ��־����</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME039" <%if InStr(1,str_PopList,"ME039",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Photo" <%if InStr(1,str_PopList,"ME_Photo",1)>0 then response.Write "checked"%>>
                      ������</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME040" <%if InStr(1,str_PopList,"ME040",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Param" <%if InStr(1,str_PopList,"ME_Param",1)>0 then response.Write "checked"%>>
                      ��������</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME041" <%if InStr(1,str_PopList,"ME041",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="ME_Pay" <%if InStr(1,str_PopList,"ME_Pay",1)>0 then response.Write "checked"%>>
                      ����֧��</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="ME042" <%if InStr(1,str_PopList,"ME042",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
              </div></td>
          </tr>
          <%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBAP")=1 then
		  %>
          <tr  class="hback_1">
            <td colspan="2" class="hback_1"><div align="left"><strong>
                <input name="PopList_AP" type="checkbox" id="PopList_AP" value="AP_Pop"  onClick="SwitchPopType('AP_Pop');" <%if InStr(1,str_PopList,"AP_Pop",1)>0 then response.Write "checked"%>>
                </strong>AP��Ƹ��ְϵͳ</div></td>
          </tr>
          <tr class="hback">
            <td colspan="2" class="hback"><div align="left">
                <table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"AP_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="AP_ID">
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="AP_Param" <%if InStr(1,str_PopList,"AP_Param",1)>0 then response.Write "checked"%>>
                        ϵͳ��������</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="AP001" <%if InStr(1,str_PopList,"AP001",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="AP_Province" <%if InStr(1,str_PopList,"AP_Province",1)>0 then response.Write "checked"%>>
                        ʡ������</div></td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="AP002" <%if InStr(1,str_PopList,"AP002",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="AP_city" <%if InStr(1,str_PopList,"AP_city",1)>0 then response.Write "checked"%>>
                      ��������</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="AP003" <%if InStr(1,str_PopList,"AP003",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="AP_Search" <%if InStr(1,str_PopList,"AP_Search",1)>0 then response.Write "checked"%>>
                        ��Ա��¼��ѯ</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="AP004" <%if InStr(1,str_PopList,"AP004",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="AP_check" <%if InStr(1,str_PopList,"AP_check",1)>0 then response.Write "checked"%>>
                        ע����Ϣ���</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="AP005" <%if InStr(1,str_PopList,"AP005",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
              </div></td>
          </tr>
          <%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBSD")=1 then
		  %>
          <tr  class="hback_1">
            <td colspan="2" class="hback_1"><div align="left"><strong>
                <input name="PopList_SD" type="checkbox" id="PopList_SD" value="SD_Pop"  onClick="SwitchPopType('SD_Pop');" <%if InStr(1,str_PopList,"SD_Pop",1)>0 then response.Write "checked"%>>
                </strong>SD����ϵͳ</div></td>
          </tr>
          <tr class="hback">
            <td colspan="2" class="hback"><div align="left">
                <table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"SD_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="SD_ID">
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="SD_List" <%if InStr(1,str_PopList,"SD_List",1)>0 then response.Write "checked"%>>
                        ������Ϣ</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="SD001" <%if InStr(1,str_PopList,"SD001",1)>0 then response.Write "checked"%>>
                            �½�
                            <input name="PopList" type="checkbox" id="PopList" value="SD002" <%if InStr(1,str_PopList,"SD002",1)>0 then response.Write "checked"%>>
                            �޸�
                            <input name="PopList" type="checkbox" id="PopList" value="SD003" <%if InStr(1,str_PopList,"SD003",1)>0 then response.Write "checked"%>>
                            ɾ��
                            <input name="PopList" type="checkbox" id="PopList" value="SD004" <%if InStr(1,str_PopList,"SD004",1)>0 then response.Write "checked"%>>
                            ���</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="SD_Class" <%if InStr(1,str_PopList,"SD_Class",1)>0 then response.Write "checked"%>>
                        �����Ŀ</div></td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="SD005" <%if InStr(1,str_PopList,"SD005",1)>0 then response.Write "checked"%>>
                            �½�
                            <input name="PopList" type="checkbox" id="PopList" value="SD006" <%if InStr(1,str_PopList,"SD006",1)>0 then response.Write "checked"%>>
                            �޸�
                            <input name="PopList" type="checkbox" id="PopList" value="SD007" <%if InStr(1,str_PopList,"SD007",1)>0 then response.Write "checked"%>>
                            ɾ��</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="AP_area" <%if InStr(1,str_PopList,"AP_area",1)>0 then response.Write "checked"%>>
                      �������</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="SD008" <%if InStr(1,str_PopList,"SD008",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="AP_param" <%if InStr(1,str_PopList,"AP_param",1)>0 then response.Write "checked"%>>
                        ϵͳ����</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="SD009" <%if InStr(1,str_PopList,"SD009",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
              </div></td>
          </tr>
          <%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBCS")=1 then
		  %>
          <tr  class="hback_1">
            <td colspan="2" class="hback_1"><div align="left"><strong>
                <input name="PopList_CS" type="checkbox" id="PopList_CS" value="CS_Pop"  onClick="SwitchPopType('CS_Pop');" <%if InStr(1,str_PopList,"CS_Pop",1)>0 then response.Write "checked"%>>
                </strong>CS�ɼ�ϵͳ</div></td>
          </tr>
          <tr class="hback">
            <td colspan="2" class="hback"><div align="left">
                <table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"CS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="CS_ID">
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="CS_site" <%if InStr(1,str_PopList,"CS_site",1)>0 then response.Write "checked"%>>
                        ����վ��</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="CS001" <%if InStr(1,str_PopList,"CS001",1)>0 then response.Write "checked"%>>
                            �޸�/�½�
                            <input name="PopList" type="checkbox" id="PopList" value="CS_collect" <%if InStr(1,str_PopList,"CS_collect",1)>0 then response.Write "checked"%>>
                            �ɼ�</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="CS_Ink" <%if InStr(1,str_PopList,"CS_Ink",1)>0 then response.Write "checked"%>>
                      �������ݿ�</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="CS003" <%if InStr(1,str_PopList,"CS003",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
              </div></td>
          </tr>
          <%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBSS")=1 then
		  %>
          <tr  class="hback_1">
            <td colspan="2" class="hback_1"><div align="left"><strong>
                <input name="PopList_SS" type="checkbox" id="PopList_SS" value="SS_Pop"  onClick="SwitchPopType('SS_Pop');"  <%if InStr(1,str_PopList,"SS_Pop",1)>0 then response.Write "checked"%>>
                </strong>SSվ��ͳ��</div></td>
          </tr>
          <tr class="hback">
            <td colspan="2" class="hback"><div align="left">
                <table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"SS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="SS_ID">
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="SS_site" <%if InStr(1,str_PopList,"SS_site",1)>0 then response.Write "checked"%>>
                        վ��ͳ��</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="SS001" <%if InStr(1,str_PopList,"SS001",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
              </div></td>
          </tr>
          <%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBHS")=1 then
		  %>
          <tr  class="hback_1">
            <td colspan="2" class="hback_1"><div align="left"><strong>
                <input name="PopList_HS" type="checkbox" id="PopList_HS" value="HS_Pop"  onClick="SwitchPopType('HS_Pop');" <%if InStr(1,str_PopList,"HS_Pop",1)>0 then response.Write "checked"%>>
                </strong>HS����¥��ϵͳ</div></td>
          </tr>
          <tr class="hback">
            <td colspan="2" class="hback"><div align="left">
                <table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"HS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="HS_ID">
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="HS_Loup" <%if InStr(1,str_PopList,"HS_Loup",1)>0 then response.Write "checked"%>>
                        ¥�̹���</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="HS001" <%if InStr(1,str_PopList,"HS001",1)>0 then response.Write "checked"%>>
                            �½�
                            <input name="PopList" type="checkbox" id="PopList" value="HS002" <%if InStr(1,str_PopList,"HS002",1)>0 then response.Write "checked"%>>
                            �޸�
                            <input name="PopList" type="checkbox" id="PopList" value="HS003" <%if InStr(1,str_PopList,"HS003",1)>0 then response.Write "checked"%>>
                            ɾ��
                            <input name="PopList" type="checkbox" id="PopList" value="HS004" <%if InStr(1,str_PopList,"HS004",1)>0 then response.Write "checked"%>>
                            ���</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="HS_Ero" <%if InStr(1,str_PopList,"HS_Ero",1)>0 then response.Write "checked"%>>
                        ���ַ�</div></td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="HS005" <%if InStr(1,str_PopList,"HS005",1)>0 then response.Write "checked"%>>
                            �½�
                            <input name="PopList" type="checkbox" id="PopList" value="HS006" <%if InStr(1,str_PopList,"HS006",1)>0 then response.Write "checked"%>>
                            �޸�
                            <input name="PopList" type="checkbox" id="PopList" value="HS007" <%if InStr(1,str_PopList,"HS007",1)>0 then response.Write "checked"%>>
                            ɾ��</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td class="hback"><input name="PopList" type="checkbox" id="PopList" value="HS_Zu" <%if InStr(1,str_PopList,"HS_Zu",1)>0 then response.Write "checked"%>>
                      ������Ϣ</td>
                    <td class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="HS008" <%if InStr(1,str_PopList,"HS008",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="HS_param" <%if InStr(1,str_PopList,"HS_param",1)>0 then response.Write "checked"%>>
                        ϵͳ����</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="HS009" <%if InStr(1,str_PopList,"HS009",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="HS_FY" <%if InStr(1,str_PopList,"HS_FY",1)>0 then response.Write "checked"%>>
                        ��Դ���</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="HS010" <%if InStr(1,str_PopList,"HS010",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="HS_Search" <%if InStr(1,str_PopList,"HS_Search",1)>0 then response.Write "checked"%>>
                        ��ѯͳ��</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="HS011" <%if InStr(1,str_PopList,"HS011",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="HS_CJ" <%if InStr(1,str_PopList,"HS_CJ",1)>0 then response.Write "checked"%>>
                        ���̹���</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="HS012" <%if InStr(1,str_PopList,"HS012",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="HS_Other" <%if InStr(1,str_PopList,"HS_Other",1)>0 then response.Write "checked"%>>
                        ����</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="HS013" <%if InStr(1,str_PopList,"HS013",1)>0 then response.Write "checked"%>>
                            ����վ����
                            <input name="PopList" type="checkbox" id="PopList" value="HS014" <%if InStr(1,str_PopList,"HS014",1)>0 then response.Write "checked"%>>
                            ������ڷ�Դ
                            <input name="PopList" type="checkbox" id="PopList" value="HS015" <%if InStr(1,str_PopList,"HS015",1)>0 then response.Write "checked"%>>
                            ����ģ��</td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
              </div></td>
          </tr>
          <%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBVS")=1 then
		  %>
          <tr  class="hback_1">
            <td colspan="2" class="hback_1"><div align="left"><strong>
                <input name="PopList_VS" type="checkbox" id="PopList_VS" value="VS_Pop"  onClick="SwitchPopType('VS_Pop');" <%if InStr(1,str_PopList,"VS_Pop",1)>0 then response.Write "checked"%>>
                </strong>VSͶƱ����</div></td>
          </tr>
          <tr class="hback">
            <td colspan="2" class="hback"><div align="left">
                <table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"VS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="VS_ID">
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="VS_site" <%if InStr(1,str_PopList,"VS_site",1)>0 then response.Write "checked"%>>
                        ���� </div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="VS001" <%if InStr(1,str_PopList,"VS001",1)>0 then response.Write "checked"%>>
                            ��������
                            <input name="PopList" type="checkbox" id="PopList" value="VS002" <%if InStr(1,str_PopList,"VS002",1)>0 then response.Write "checked"%>>
                            ����
                            <input name="PopList" type="checkbox" id="PopList" value="VS003" <%if InStr(1,str_PopList,"VS003",1)>0 then response.Write "checked"%>>
                            �鿴</td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
              </div></td>
          </tr>
          <%end if
		  %>
          <%
		  if Request.Cookies("FoosunSUBCookie")("FoosunSUBAS")=1 then%>
          <tr  class="hback_1">
            <td colspan="2" class="hback_1"><div align="left"><strong>
                <input name="PopList_AS" type="checkbox" id="PopList_AS" value="AS_Pop"  onClick="SwitchPopType('AS_Pop');" <%if InStr(1,str_PopList,"AS_Pop",1)>0 then response.Write "checked"%>>
                </strong>AS������ϵͳ</div></td>
          </tr>
          <tr class="hback">
            <td colspan="2" class="hback"><div align="left">
                <table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"AS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="AS_ID">
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="AS_site" <%if InStr(1,str_PopList,"AS_site",1)>0 then response.Write "checked"%>>
                        ����</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="AS001" <%if InStr(1,str_PopList,"AS001",1)>0 then response.Write "checked"%>>
                            ����
                            <input name="PopList" type="checkbox" id="PopList" value="AS002" <%if InStr(1,str_PopList,"AS002",1)>0 then response.Write "checked"%>>
                            ͳ��
                            <input name="PopList" type="checkbox" id="PopList" value="AS003" <%if InStr(1,str_PopList,"AS003",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
              </div></td>
          </tr>
          <%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBWS")=1 then
		  %>
          <tr  class="hback_1">
            <td colspan="2" class="hback_1"><div align="left"><strong>
                <input name="PopList_WS" type="checkbox" id="PopList_WS" value="WS_Pop"  onClick="SwitchPopType('WS_Pop');" <%if InStr(1,str_PopList,"WS_Pop",1)>0 then response.Write "checked"%>>
                </strong>WS����ϵͳ</div></td>
          </tr>
          <tr class="hback">
            <td colspan="2" class="hback"><div align="left">
                <table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"WS_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="WS_ID">
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="WS_site" <%if InStr(1,str_PopList,"WS_site",1)>0 then response.Write "checked"%>>
                        ����</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="WS001" <%if InStr(1,str_PopList,"WS001",1)>0 then response.Write "checked"%>>
                            �鿴
                            <input name="PopList" type="checkbox" id="PopList" value="WS002" <%if InStr(1,str_PopList,"WS002",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
              </div></td>
          </tr>
          <%
			end if
			if Request.Cookies("FoosunSUBCookie")("FoosunSUBFL")=1 then
		  %>
          <tr  class="hback_1">
            <td colspan="2" class="hback_1"><div align="left"><strong>
                <input name="PopList_FL" type="checkbox" id="PopList_FL" value="FL_Pop"  onClick="SwitchPopType('FL_Pop');" <%if InStr(1,str_PopList,"FL_Pop",1)>0 then response.Write "checked"%>>
                </strong>FL��������ϵͳ</div></td>
          </tr>
          <tr class="hback">
            <td colspan="2" class="hback"><div align="left">
                <table width="100%" border="0" cellspacing="0" cellpadding="4" style="display: <%if InStr(1,str_PopList,"FL_Pop",1)<>0 then:response.Write(";"):else:Response.Write("none"):end if%>" id="FL_ID">
                  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
                    <td width="14%" class="hback"><div align="left">
                        <input name="PopList" type="checkbox" id="PopList" value="FL_site" <%if InStr(1,str_PopList,"FL_site",1)>0 then response.Write "checked"%>>
                        ����</div></td>
                    <td width="86%" class="hback"><table width="98%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><input name="PopList" type="checkbox" id="PopList" value="FL001" <%if InStr(1,str_PopList,"FL001",1)>0 then response.Write "checked"%>>
                            �鿴
                            <input name="PopList" type="checkbox" id="PopList" value="FL002" <%if InStr(1,str_PopList,"FL002",1)>0 then response.Write "checked"%>>
                            ����</td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
              </div></td>
          </tr>
          <%end if%>
        </table></td>
    </tr>
    <tr >
      <td colspan="3" class="hback"><input type="submit" name="Submit" value="ȷ������Ȩ��">
        <input type="reset" name="Submit2" value="����">
      </td>
    </tr>
  </form>
</table>
</body>
</html>
<%
Set Conn = Nothing
%>






