<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_Inc/Function.asp"-->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_InterFace/NS_Function.asp" -->
<!--#include file="../../FS_Inc/Func_page.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
	Dim Conn
	MF_Default_Conn
	'session�ж�
	MF_Session_TF 
	if not MF_Check_Pop_TF("MF_sPublic") then Err_Show
%>
<html>
<head>
<title>���ű�ǩ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<base target=self>
</head>
<body class="hback">
<script language="JavaScript" src="../../FS_Inc/PublicJS.js" type="text/JavaScript"></script>
  <form  name="form1" method="post">
  <table width="98%" height="57" border="0" align=center cellpadding="3" cellspacing="1" class="table" valign=absmiddle>
    <tr class="hback" > 
      <td height="27"  align="Left" class="xingmu"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="41%" class="xingmu"><strong>����ǩ����</strong></td>
            <td width="59%"><div align="right"> 
                <input name="button4" type="button" onClick="window.returnValue='';window.close();" value="�ر�">
            </div></td>
          </tr>
        </table></td>
    </tr>
    <tr class="hback" >
      <td height="27"  align="Left" class="hback"><a href="photo_C_Label.asp?type=Photo_List" target="_self">�����ǩ</a>��<a href="photo_C_Label.asp?type=PhotoClassList" target="_self">�ռ�����</a>��<a href="photo_C_Label.asp?type=ShowPhoto" target="_self">ͼƬ���</a></td>
    </tr>
  </table>
  <%
  if Request.QueryString("type")="PhotoClassList" then
  	call PhotoClassList()
  elseif request.QueryString("type")="ShowPhoto" then
  	call ShowPhoto()
  else
  	call Photo_List()
  end if
  Sub  Photo_List()
  %>
 <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    
    <tr>
      <td width="19%" class="hback"><div align="right">����</div></td>
      <td width="81%" class="hback">
	  <select name="BookPop" id="BookPop">
        <option value="0">���Ƽ��������</option>
        <option value="1">�Ƽ�</option>
        <option value="2">���</option>
      </select>	  </td>
    </tr>
    
    <tr>
      <td class="hback"><div align="right">��������</div></td>
      <td class="hback"><input name="TitleNumber" type="text" id="TitleNumber" value="10"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��ʾ����</div></td>
      <td class="hback"><input name="leftTitle" type="text" id="leftTitle" value="30">
        ���Ϊ0����ʾ����</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">ͼƬ��С</div></td>
      <td class="hback"><input name="picsize" type="text" id="picsize" value="100,80" size="8">
        ����,��,ͼƬCSS
          <label>
          <input name="PicCSS" type="text" id="PicCSS" size="5">
        </label></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">ÿ������</div></td>
      <td class="hback"><input name="ColsNumber" type="text" id="ColsNumber" value="1" size="10">
      ��DIV+CSS�����Ч </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">����CSS</div></td>
      <td class="hback"><input name="TitleCSS" type="text" id="TitleCSS" size="10"> 
        ��DIV+CSS�����Ч </td>
    </tr>
    
    <tr> 
      <td class="hback"><div align="right">�����ʽ</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
          <option value="out_Table">��ͨ��ʽ</option>
          <option value="out_DIV">DIV+CSS��ʽ</option>
          
        </select></td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:����;display:none;" > 
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV����</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��"> 
        &quot; class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;������Ч</td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">        &quot; class=&quot; <input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:ME=Photo_List��';
		retV+='����$' + obj.BookPop.value + '��';
		retV+='����$' + obj.TitleNumber.value + '��';
		retV+='����$' + obj.ColsNumber.value + '��';
		retV+='����$' + obj.leftTitle.value + '��';
		retV+='ͼƬ��С$' + obj.picsize.value + '��';
		retV+='ͼƬCSS$' + obj.PicCSS.value + '��';
		retV+='����CSS$' + obj.TitleCSS.value + '��';
		retV+='DivID$' + obj.DivID.value + '��';
		retV+='Divclass$' + obj.Divclass.value + '��';
		retV+='ulid$' + obj.ulid.value + '��';
		retV+='ulclass$' + obj.ulclass.value + '��';
		retV+='liid$' + obj.liid.value + '��';
		retV+='liclass$' + obj.liclass.value + '��';
		retV+='�����ʽ$' + obj.out_char.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%End Sub%>
	<%Sub PhotoClassList()%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    
    <tr>
      <td width="19%" class="hback"><div align="right">��������</div></td>
      <td width="81%" class="hback"><input name="TitleNumber" type="text" id="TitleNumber" value="10"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">��ʾ����</div></td>
      <td class="hback"><input name="leftTitle" type="text" id="leftTitle" value="30">
        ���Ϊ0����ʾ����</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">ͼƬ��С</div></td>
      <td class="hback"><input name="picsize" type="text" id="picsize" value="100,80" size="8">
        ����,��,������д��ʽ(100,0)����ʾ�߶�Ϊ100����Ȳ����ƣ���֤��ͼƬ������;,(0,100)���ʾ�����Ƹ߶�
          <label></label></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">ͼƬCSS </div></td>
      <td class="hback"><input name="PicCSS" type="text" id="PicCSS" size="5"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">ÿ������</div></td>
      <td class="hback"><input name="ColsNumber" type="text" id="ColsNumber" value="1" size="10">
      ��DIV+CSS�����Ч </td>
    </tr>
    <tr>
      <td class="hback"><div align="right">����CSS</div></td>
      <td class="hback"><input name="TitleCSS" type="text" id="TitleCSS" size="10"> 
        ��DIV+CSS�����Ч </td>
    </tr>
    
    <tr> 
      <td class="hback"><div align="right">�����ʽ</div></td>
      <td class="hback">
	     <select name="out_char" id="out_char" onChange="selectHtml_express(this.options[this.selectedIndex].value);">
          <option value="out_Table">��ͨ��ʽ</option>
          <option value="out_DIV">DIV+CSS��ʽ</option>
          
        </select></td>
    </tr>
    <tr class="hback"  id="div_id" style="font-family:����;display:none;" > 
      <td rowspan="3"  align="center" class="hback"><div align="right"></div>
        <div align="right">DIV����</div></td>
      <td colspan="3" class="hback" >&lt;div id=&quot; <input name="DivID"  type="text" id="DivID" size="6" disabled style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�ID�ţ�����CSS��Ԥ�ȶ��塣����Ϊ��"> 
        &quot; class=&quot; <input name="Divclass"  type="text" id="Divclass" size="6"   style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;������Ч</td>
    </tr>
    <tr class="hback" id="ul_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;ul id=&quot; <input name="ulid"   type="text" id="ulid" size="6" style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����ul���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot; class=&quot; <input name="ulclass"  type="text" id="ulclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr class="hback"  id="li_id" style="font-family:����;display:none;"> 
      <td colspan="3" class="hback" >&lt;li id=&quot;
        <input name="liid"  type="text" id="liid" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000" title="ǰ̨����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!!">        &quot; class=&quot; <input name="liclass"  type="text" id="liclass" size="6"  style="	border-top-width: 0px;	border-right-width: 0px;	border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"  title="ǰ̨����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!!"> 
        &quot;&gt;</td>
    </tr>
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:ME=PhotoClassList��';
		retV+='����$' + obj.TitleNumber.value + '��';
		retV+='����$' + obj.ColsNumber.value + '��';
		retV+='����$' + obj.leftTitle.value + '��';
		retV+='ͼƬ��С$' + obj.picsize.value + '��';
		retV+='ͼƬCSS$' + obj.PicCSS.value + '��';
		retV+='����CSS$' + obj.TitleCSS.value + '��';
		retV+='DivID$' + obj.DivID.value + '��';
		retV+='Divclass$' + obj.Divclass.value + '��';
		retV+='ulid$' + obj.ulid.value + '��';
		retV+='ulclass$' + obj.ulclass.value + '��';
		retV+='liid$' + obj.liid.value + '��';
		retV+='liclass$' + obj.liclass.value + '��';
		retV+='�����ʽ$' + obj.out_char.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%End Sub%>
	<%Sub ShowPhoto()%>
	<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
    <tr>
      <td width="19%" class="hback"><div align="right">��ʾ����</div></td>
      <td width="81%" class="hback"><input name="leftTitle" type="text" id="leftTitle" value="30">
        ���Ϊ0����ʾ����</td>
    </tr>
    <tr>
      <td class="hback"><div align="right">ͼƬ��С</div></td>
      <td class="hback"><input name="picsize" type="text" id="picsize" value="300,200" size="8">
        ����,��,������д��ʽ(100,0)����ʾ�߶�Ϊ100����Ȳ����ƣ���֤��ͼƬ������;,(0,100)���ʾ�����Ƹ߶�
          <label></label></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">ͼƬCSS </div></td>
      <td class="hback"><input name="PicCSS" type="text" id="PicCSS" size="5"></td>
    </tr>
    <tr>
      <td class="hback"><div align="right">����CSS</div></td>
      <td class="hback"><input name="TitleCSS" type="text" id="TitleCSS" size="10"> 
        ��DIV+CSS�����Ч </td>
    </tr>
    
    <tr>
      <td class="hback"><div align="right"></div></td>
      <td class="hback"><input name="button" type="button" onClick="ok(this.form);" value="ȷ�������˱�ǩ">
        <input name="button" type="button" onClick="window.returnValue='';window.close();" value=" ȡ �� "></td>
    </tr>
  </table>
	<script language="JavaScript" type="text/JavaScript">
	function ok(obj)
	{
		var retV = '{FS:ME=ShowPhoto��';
		retV+='����$' + obj.leftTitle.value + '��';
		retV+='ͼƬ��С$' + obj.picsize.value + '��';
		retV+='ͼƬCSS$' + obj.PicCSS.value + '��';
		retV+='����CSS$' + obj.TitleCSS.value;
		retV+='}';
		window.parent.returnValue = retV;
		window.close();
	}
	</script>
	<%End Sub%>
 </form>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
function SelectClass()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('../Supply/SelectClassFrame.asp',400,300,window);
	if (ReturnValue.indexOf('***')!=-1)
	{
		TempArray = ReturnValue.split('***');
		document.all.ClassID.value=TempArray[0]
		document.all.ClassName.value=TempArray[1]
	}
}
function selectHtml_express(Html_express)
{
	switch (Html_express)
	{
	case "out_Table":
		document.getElementById('div_id').style.display='none';
		document.getElementById('li_id').style.display='none';
		document.getElementById('ul_id').style.display='none';
		document.getElementById('DivID').disabled=true;
		break;
	case "out_DIV":
		document.getElementById('div_id').style.display='';
		document.getElementById('li_id').style.display='';
		document.getElementById('ul_id').style.display='';
		document.getElementById('DivID').disabled=false;
		break;
	}
}
</script>






