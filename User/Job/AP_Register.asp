<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<% 
Dim Ap_Rs,Ap_Sql
Dim UserNumber,C_Name
UserNumber = Session("FS_UserNumber")
if not UserNumber<>"" then response.Redirect("../lib/error.asp?ErrCodes=<li>����δ��½,�����.</li>&ErrorUrl=../login.asp") : response.End()
C_Name = Get_OtherTable_Value("select C_Name from FS_ME_CorpUser where UserNumber='"&UserNumber&"'")
if C_Name = "" then response.Redirect("../lib/error.asp?ErrCodes=<li>�㲻�ǹ�˾��Ա��û���ʸ����.</li>&ErrorUrl=../login.asp") : response.End()
''�õ���ر��ֵ��
Function Get_OtherTable_Value(This_Fun_Sql)
	Dim This_Fun_Rs
	if instr(This_Fun_Sql," FS_ME_")>0 then 
		set This_Fun_Rs = User_Conn.execute(This_Fun_Sql)
	else
		set This_Fun_Rs = Conn.execute(This_Fun_Sql)
	end if			
	if not This_Fun_Rs.eof then 
		Get_OtherTable_Value = This_Fun_Rs(0)
	else
		Get_OtherTable_Value = ""
	end if
	if Err.Number>0 then 
		Err.Clear
		response.Redirect("../lib/error.asp?ErrCodes=<li>Get_OtherTable_Valueδ�ܵõ�������ݡ�����������"&Err.Description&"</li>") : response.End()
	end if
	set This_Fun_Rs=nothing 
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title>ע�ṫ˾---��վ���ݹ���ϵͳ</title>
<meta name="keywords" content="��Ѷcms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,��Ѷ��վ���ݹ���ϵͳ">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<meta name="Keywords" content="Foosun,FoosunCMS,Foosun Inc.,��Ѷ,��Ѷ��վ���ݹ���ϵͳ,��Ѷϵͳ,��Ѷ����ϵͳ,��Ѷ�̳�,��Ѷb2c,����ϵͳ,CMS,�����ռ�,asp,jsp,asp.net,SQL,SQL SERVER" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../FS_Inc/PublicJS_YanZheng.js"></script>
<script language="JavaScript" src="../../FS_Inc/coolWindowsCalendar.js"></script>
<script language="JavaScript">
<!--
setTimeout('showhide()',500);
function showhide()
{
	if(document.all.GroupLevel==null||document.all.GroupLevel=='undefined') return;
	if(document.all.GroupLevel.value=='1'){document.all.PuT.style.display='';document.all.BaoYue.style.display='none';document.all.BaoYue1.style.display='none'} else {document.all.PuT.style.display='none';document.all.BaoYue.style.display='';document.all.BaoYue1.style.display=''}
}
-->
</script>
<head>
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
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="../Top_navi.asp" --> </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="../menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback"><table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
          <tr class="hback"> 
            
          <td class="hback"><strong>λ�ã�</strong><a href="../../">��վ��ҳ</a> &gt;&gt; 
            <a href="../main.asp">��Ա��ҳ</a> &gt;&gt; <a href="job_applications.asp">��Ƹ��ҳ</a>��ע�ṫ˾��չ��Ϣ</td>
          </tr>
        </table>

<%
'******************************************************************
select case request.QueryString("Act")
	case "","Add","Edit","Search"
	Add_Edit_Search
	case else
	response.Redirect("job_applications.asp")
end select
'******************************************************************
Sub Add_Edit_Search()
Dim Bol_IsEdit,UserClass,GroupLevel,DocExt,Audited
Audited = 0
Bol_IsEdit = false
if UserNumber="" then response.Redirect("../lib/error.asp?ErrorUrl=&ErrCodes=<li>��û�е�½�������</li>") : response.End()
Ap_Sql = "select top 1 UserClass,GroupLevel,LeftCount,BeginDate,EndDate,CompanyName,Introduct,DocExt,DescriptApply,DescriptAudit,Audited,Phone,Email from FS_AP_UserList where UserNumber = '"&UserNumber&"'"
Set Ap_Rs	= CreateObject(G_FS_RS)
Ap_Rs.Open Ap_Sql,Conn,1,1
if not Ap_Rs.eof then 
	Bol_IsEdit = True
	UserClass = Ap_Rs("UserClass")
	GroupLevel = Ap_Rs("GroupLevel")
	DocExt = Ap_Rs("DocExt")
	Audited = Ap_Rs("Audited")
end if	
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <form name="form1" id="form1" onSubmit="return chkinput();" method="post" action="AP_Register_Action.asp?Act=Save">
    <tr  class="hback"> 
      <td colspan="3" align="left" class="xingmu" ><%if Bol_IsEdit then
	  response.Write("�޸�ע����Ϣ")
	  else
	  response.Write("��дע����Ϣ�����ͨ����Ϳ��Է�����Ƹ��Ϣ����")
	  end if%></td>
	</tr>
  <tr  class="hback"> 
    <td colspan="3" height="25">
	 <a href="AP_Register.asp">��ҳ</a>	
	</td>
  </tr>
    <tr  class="hback"> 
      <td align="right">��Ա����</td>
      <td>
		<select name="UserClass" id="UserClass"<%if Audited=1 then response.Write(" disabled") end if%>>
		<%=PrintOption(UserClass,"1:���˵�λ,2:��ͷ��˾")%>
		</select><span id="UserClass_Alt"></span>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�ȼ�����</td>
      <td>
		<select name="GroupLevel" id="GroupLevel" onChange="showhide()"<%if Audited=1 then response.Write(" disabled") end if%>>
		<%=PrintOption(GroupLevel,"1:��ͨ��Ա,2:���»�Ա,3:VIP��Ա")%>
		</select><span id="UserClass_Alt"></span>
      </td>
    </tr>
    <tr class="hback" id="PuT" style="display:none"> 
      <td align="right">ʣ�����</td>
      <td>
	  	<%if Bol_IsEdit then response.Write(Ap_Rs("LeftCount")) else response.Write("0") end if%>
      </td>
    </tr>
    <tr class="hback" id="BaoYue" style="display:none"> 
      <td align="right">��Ч����</td>
      <td>
	  <input type="text" name="BeginDate" id="BeginDate" onFocus="setday(this)" style="WIDTH: 100px; HEIGHT: 22px" maskType="shortDate" value="<%if Bol_IsEdit then response.Write(Ap_Rs("BeginDate")) else response.Write(date()) end if%>"<%if Audited=1 then response.Write(" disabled") end if%>>
	  <IMG id="img3" onClick="BeginDate.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom"><span id="BeginDate_Alt"></span>
      </td>
    </tr>
    <tr  class="hback" id="BaoYue1" style="display:none"> 
      <td align="right">��������</td>
      <td>
	  <input type="text" name="EndDate" id="EndDate"  onfocus="setday(this)" style="WIDTH: 100px; HEIGHT: 22px" maskType="shortDate" value="<%if Bol_IsEdit then response.Write(Ap_Rs("EndDate")) end if%>"<%if Audited=1 then response.Write(" disabled") end if%>>
	  <IMG id="img4" onClick="EndDate.focus()" src="../../FS_Inc/calendar.bmp" align="absBottom"><span id="EndDate_Alt"></span>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��˾����</td>
      <td>
	  <input type="text" name="CompanyName" id="CompanyName" size="60" maxlength="30" value="<%if Bol_IsEdit then response.Write(Ap_Rs("CompanyName")) else response.Write(C_Name) end if%>"<%if request.QueryString("Act")<>"Search" then%> dataType="LimitB" min="4" max="30" msg="���ȱ�����4-30֮��"<%end if%><%if Audited=1 then response.Write(" disabled") end if%>>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">��˾���</td>
      <td>
	  <textarea name="Introduct" cols="60" rows="20" id="Introduct"<%if Audited=1 then response.Write(" disabled") end if%>><%if Bol_IsEdit then response.Write(Ap_Rs("Introduct")) end if%></textarea>
      <span id="Introduct_Alt"></span>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">ӦƸ���</td>
      <td>
		<select name="DocExt" id="DocExt">
		<%=PrintOption(DocExt,"1:.doc,2:.rar,0:�ޱ��")%>
		</select><span id="DocExt_Alt"></span>
      </td>
    </tr>
	
    <tr  class="hback"> 
      <td align="right">��˾�绰</td>
      <td>
	  <input type="text" name="phone" id="phone" size="40" maxlength="50"<%if request.QueryString("Act")<>"Search" then%> dataType="Require" msg="�绰������д"<%end if%> value="<%if Bol_IsEdit then response.Write(Ap_Rs("phone")) end if%>">
      </td>
    </tr>

    <tr  class="hback"> 
      <td align="right">Email</td>
      <td>
	  <input type="text" name="Email" id="Email" size="15" maxlength="20"<%if request.QueryString("Act")<>"Search" then%> dataType="Require" msg="�ʼ�������д"<%end if%> value="<%if Bol_IsEdit then response.Write(Ap_Rs("Email")) end if%>">
      </td>
    </tr>
	

    <tr  class="hback"> 
      <td align="right">���뱸ע</td>
      <td>
	  <textarea name="DescriptApply" cols="60" rows="5" id="DescriptApply"<%if Audited=1 then response.Write(" disabled") end if%>><%if Bol_IsEdit then response.Write(Ap_Rs("DescriptApply")) end if%></textarea>
      <span id="DescriptApply_Alt"></span>
      </td>
    </tr>
     <tr  class="hback"> 
      <td align="right">��˱�ע</td>
      <td>
	  <textarea name="DescriptAudit" disabled cols="60" rows="5" id="DescriptAudit"><%if Bol_IsEdit then response.Write(Ap_Rs("DescriptAudit")) end if%></textarea>
      <span id="DescriptAudit_Alt"></span>
      </td>
    </tr>
    <tr  class="hback"> 
      <td align="right">�Ƿ�ͨ�����?</td>
      <td>
			<%=Replacestr(Audited,"1:��ͨ��,0:<span class=tx>δͨ��</span>,else:δ�ύ")%>
      </td>
    </tr>
   <tr  class="hback"> 
      <td colspan="4">
	  <table border="0" width="100%" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="center"> <input type="submit"value=" ȷ���ύ "/> 
              &nbsp; <input type="reset" value=" ���� " />
  			  &nbsp; <input type="button" name="btn_todel" value=" ����ע�� " onClick="if(confirm('ȷ��ɾ������Ŀ��')) location='<%=server.URLEncode("AP_Register_Action.asp?Act=Del")%>'">
            </td>
          </tr>
        </table>
      </td>
    </tr>	
  </form>
</table>
<%
End Sub
%>
       </td>
    </tr>
    <tr class="back"> 
      <td height="20" colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="../Copyright.asp" -->
        </div></td>
    </tr>
 
</table>
<%
Set Fs_User = Nothing
set AP_Rs = Nothing
%>


<!-- Powered by: FoosunCMS5.0ϵ��,Company:Foosun Inc. -->





