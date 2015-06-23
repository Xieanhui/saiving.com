<% Option Explicit %>
<!--#include file="../../FS_Inc/Const.asp" -->
<!--#include file="../../FS_InterFace/MF_Function.asp" -->
<!--#include file="../../FS_Inc/Function.asp" -->
<!--#include file="../lib/strlib.asp" -->
<!--#include file="../lib/UserCheck.asp" -->
<%
If CheckBlogOpen=False Then 
	Response.write("<script language=""javascript"">alert('日志功能暂停使用,如需要使用请联系管理员.');history.back();</script>")
	Response.End()
End If
Dim str_CurrPath,FileName,str_FileName,rs_sys,str_FileExtName
str_CurrPath = Replace("/"&G_VIRTUAL_ROOT_DIR &"/"&G_USERFILES_DIR&"/"&Session("FS_UserNumber"),"//","/")
	set rs_sys=User_Conn.execute("select top 1 FileName,FileExtName From FS_ME_iLogSysParam")
	if rs_sys.eof then
		response.Write("系统配置出错!")
		response.End
		rs_sys.close:set rs_sys=nothing
	else
		FileName=rs_sys("FileName")
		str_FileExtName=rs_sys("FileExtName")
		rs_sys.close:set rs_sys=nothing
	end if
	if FileName=0 then
		str_FileName=GetRamCode(8)
	elseif  FileName=1 then
		str_FileName="自动编号"
	else
		str_FileName=right(year(now),2)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_"&GetRamCode(3)
	end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<title><%=GetUserSystemTitle%></title>
<meta name="keywords" content="风讯cms,cms,FoosunCMS,FoosunOA,FoosunVif,vif,风讯网站内容管理系统">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta content="MSHTML 6.00.3790.2491" name="GENERATOR" />
<link href="../images/skin/Css_<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>/<%=Request.Cookies("FoosunUserCookies")("UserLogin_Style_Num")%>.css" rel="stylesheet" type="text/css">
<head></head>
<body>
<script language="JavaScript" src="../../Editor/FS_scripts/editor.js"></script>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
  <tr>
    <td>
      <!--#include file="../top.asp" -->
    </td>
  </tr>
</table>
<table width="98%" height="135" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <tr class="back"> 
      <td   colspan="2" class="xingmu" height="26"> <!--#include file="../Top_navi.asp" -->
    </td>
    </tr>
    <tr class="back"> 
      <td width="18%" valign="top" class="hback"> <div align="left"> 
          <!--#include file="../menu.asp" -->
        </div></td>
      <td width="82%" valign="top" class="hback">
	  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <tr class="hback"> 
          <td  valign="top">你的位置：<a href="../../">网站首页</a> &gt;&gt; <a href="../main.asp">会员首页</a> 
            &gt;&gt; <a href="index.asp">日志管理</a> &gt;&gt;日志管理</td>
        </tr>
        <tr class="hback"> 
          <td  valign="top"><a href="index.asp">日志首页</a>┆<a href="PublicLog.asp">发表日志</a>┆<a href="index.asp?type=box">草稿箱</a>┆<a href="../PhotoManage.asp">相册管理</a>┆<a href="PublicParam.asp">参数设置</a>┆<a href="../Review.asp">评论管理</a></td>
        </tr>
      </table>
        
      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
        <form name="s_form" method="post" action="Public_Save.asp">
          <tr> 
            <td colspan="2" class="xingmu">发表日志</td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">日志标题</div></td>
            <td class="hback"><input name="Title" type="text" id="Title" size="57" maxlength="100" onFocus="Do.these('Title',function(){return CheckContentLen('Title','span_Title','3-100')})" onKeyUp="Do.these('Title',function(){return CheckContentLen('Title','span_Title','3-100')})"> 
              <span id="span_Title"></span></td>
          </tr>
          <tr> 
            <td width="13%" class="hback"><div align="right">主分类</div></td>
            <td width="87%" class="hback">
				<select name="MainID" id="MainID">
                <!--<option value="">选择系统分类</option>-->
				<%
				dim c_rs
				set c_rs = Server.CreateObject(G_FS_RS)
				c_rs.open "select ID,ClassName From FS_ME_iLogClass Order by id asc",User_Conn,1,3
				do while not c_rs.eof
				%>
                <option value="<%=c_rs("id")%>"><%=c_rs("ClassName")%></option>
				<%
				c_rs.movenext
				loop
				c_rs.close:set c_rs=nothing
				%>
              </select>
              专栏 <select name="ClassID" id="ClassID" style="width:38%">
                <option value="0">选择我的专栏</option>
                <%
			  dim obj_zl_rs
			  set obj_zl_rs= Server.CreateObject(G_FS_RS)
			  obj_zl_rs.open "select ClassID,ClassCName,ClassEName,ParentID,UserNumber,ClassTypes From FS_ME_InfoClass  where ParentID=0 and ClassTypes=7 and UserNumber='"& Fs_User.UserNumber&"' order by ClassID desc",User_Conn,1,3
			  do while not obj_zl_rs.eof 
			  		Response.Write"<option value="""& obj_zl_rs("ClassID")&""">"& obj_zl_rs("ClassCName")&"</option>"
				  obj_zl_rs.movenext
			  Loop
			  obj_zl_rs.close:set obj_zl_rs =nothing
			  %>
              </select></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">类型</div></td>
            <td class="hback"><select name="iLogStyle" id="iLogStyle">
                <option value="0" selected>日志</option>
                <option value="1">网摘</option>
              </select> </td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">Tags(关键字)</div></td>
            <td class="hback"><input name="keyword1" type="text" id="keyword1" size="15" maxlength="15">
              , 
              <input name="keyword2" type="text" id="keyword2" size="15" maxlength="15">
              , 
              <input name="keyword3" type="text" id="keyword3" size="15" maxlength="15"></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">来源</div></td>
            <td class="hback"><input name="iLogSource" type="text" id="iLogSource" size="57" maxlength="60"></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">内容</div></td>
            <td class="hback">
						 <!--编辑器开始-->
						<iframe id='NewsContent' src='../Editer/UserEditer.asp?id=Content' frameborder=0 scrolling=no width='100%' height='280'></iframe>
						<input type="hidden" name="Content" value="">
						<!--编辑器结束-->
						</td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">首页置顶</div></td>
            <td class="hback"><input name="isTop" type="checkbox" id="isTop" value="1"></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">心情</div></td>
            <td class="hback"><input name="EmotFace" type="radio" value="face1.gif" checked> 
              <img src="../../sys_images/emot/face1.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face3.gif"> <img src="../../sys_images/emot/face3.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face5.gif"> <img src="../../sys_images/emot/face5.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face8.gif"> <img src="../../sys_images/emot/face8.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face18.gif"> <img src="../../sys_images/emot/face18.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face11.gif"> <img src="../../sys_images/emot/face11.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face47.gif"> <img src="../../sys_images/emot/face47.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face15.gif"> <img src="../../sys_images/emot/face15.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face22.gif"> <img src="../../sys_images/emot/face22.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face29.gif"> <img src="../../sys_images/emot/face29.gif" width="19" height="19"> 
              <input type="radio" name="EmotFace" value="face34.gif"> <img src="../../sys_images/emot/face34.gif" width="19" height="19"></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">图片地址</div></td>
            <td class="hback"><table width="100%" border="0" cellspacing="1" cellpadding="5">
                <tr> 
                  <td width="29%"><div align="center"> 
                      <table width="10" border="0" cellspacing="1" cellpadding="2" class="table">
                        <tr> 
                          <td class="hback"><img src="../Images/nopic_supply.gif" width="90" height="90" id="pic_p_1"></td>
                        </tr>
                      </table>
                      <input name="pic_1" type="hidden" id="pic_1" size="40" >
                    </div></td>
                  <td width="36%"><div align="center"> 
                      <table width="10" border="0" cellspacing="1" cellpadding="2" class="table">
                        <tr> 
                          <td class="hback"><img src="../Images/nopic_supply.gif" width="90" height="90" id="pic_p_2"></td>
                        </tr>
                      </table>
                      <input name="pic_2" type="hidden" id="pic_2" size="40">
                    </div></td>
                  <td width="35%"><div align="center"> 
                      <table width="10" border="0" cellspacing="1" cellpadding="2" class="table">
                        <tr> 
                          <td class="hback"><img src="../Images/nopic_supply.gif" width="90" height="90" id="pic_p_3"></td>
                        </tr>
                      </table>
                      <input name="pic_3" type="hidden" id="pic_3" size="40">
                    </div></td>
                </tr>
                <tr> 
                  <td><div align="center"><img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.s_form.pic_1);" style="cursor:hand;"> 
                      　<img src="../Images/del_supply.gif" width="44" height="22" onClick="dels_1();" style="cursor:hand;"> 
                    </div></td>
                  <td><div align="center"><img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.s_form.pic_2);" style="cursor:hand;"> 
                      　<img src="../Images/del_supply.gif" width="44" height="22" onClick="dels_2();" style="cursor:hand;"> 
                    </div></td>
                  <td><div align="center"><img  src="../Images/upfile.gif" width="44" height="22" onClick="OpenWindowAndSetValue('../CommPages/SelectPic.asp?CurrPath=<% = str_CurrPath %>&f_UserNumber=<% = session("FS_UserNumber")%>',500,320,window,document.s_form.pic_3);" style="cursor:hand;"> 
                      　<img src="../Images/del_supply.gif" width="44" height="22" onClick="dels_3();" style="cursor:hand;"> 
                    </div></td>
                </tr>
              </table></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">图片放入相册</div></td>
            <td class="hback"><input name="PutInPhoto" type="checkbox" id="PutInPhoto" value="1"></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">日志密码</div></td>
            <td class="hback"><input name="Password" type="password" id="Password"></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right">文件名</div></td>
            <td class="hback"><input name="FileName" type="text" id="FileName" value="<%=str_FileName%>" maxlength="20"  onFocus="Do.these('FileName',function(){return CheckContentLen('FileName','span_FileName','3-18')})" onKeyUp="Do.these('FileName',function(){return CheckContentLen('FileName','span_FileName','3-18')})"><span id="span_FileName"></span>
              扩展名 
              <select name="FileExtName" id="FileExtName">
                <option value="html" <%if str_FileExtName="html" then response.Write"selected"%>>html</option>
                <option value="htm" <%if str_FileExtName="htm" then response.Write"selected"%>>htm</option>
                <option value="shtm" <%if str_FileExtName="shtm" then response.Write"selected"%>>shtm</option>
                <option value="shtml" <%if str_FileExtName="shtml" then response.Write"selected"%>>shtml</option>
                <option value="asp" <%if str_FileExtName="asp" then response.Write"selected"%>>asp</option>
              </select></td>
          </tr>
          <tr onMouseOver=overColor(this) onMouseOut=outColor(this)> 
            <td class="hback"><div align="right"></div></td>
            <td class="hback"><input name="Action" type="hidden" id="Action"> 
              <input type="button" name="button1" value="保存日志/网摘"  onClick="SaveSubmit(this.form);"> 
              <input type="button" name="Submit22" value="保存为草稿"  onClick="CheckSubmit(this.form);"> 
              <input type="reset" name="Submit3" value="重新填写"></td>
          </tr>
        </form>
      </table>
       </td>
    </tr>
    <tr class="back"> 
      <td height="20"  colspan="2" class="xingmu"> <div align="left"> 
          <!--#include file="../Copyright.asp" -->
        </div></td>
    </tr>
</table>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
function CheckSubmit(FormObj){
	FormObj.Action.value='isDraft';
	if(confirm('确定保存为草稿吗？')){
		if (frames["NewsContent"].g_currmode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return false;}
		FormObj.Content.value=frames["NewsContent"].GetNewsContentArray();
		FormObj.submit();
		return true;
	}
	return false;
}
function SaveSubmit(FormObj){
	FormObj.Action.value='Save';
	FormObj.Content.value=frames["NewsContent"].GetNewsContentArray();
	if(confirm('确定保存日志吗？')){
		if (frames["NewsContent"].g_currmode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return false;}
		FormObj.Content.value=frames["NewsContent"].GetNewsContentArray();
		FormObj.submit();
		return true;
	}
	return false;
}

new Form.Element.Observer($('pic_1'),1,pics_1);
	function pics_1()
		{
			if ($('pic_1').value=='')
			{
				$('pic_p_1').src='../Images/nopic_supply.gif'
			}
			else
			{
			$('pic_p_1').src=$('pic_1').value
			}
		} 
new Form.Element.Observer($('pic_2'),1,pics_2);
	function pics_2()
		{
			if($('pic_2').value=='')
			{
			$('pic_p_2').src='../Images/nopic_supply.gif'
			}
			else
			{
			$('pic_p_2').src=$('pic_2').value
			}
		} 
new Form.Element.Observer($('pic_3'),1,pics_3);
	function pics_3()
		{
			if($('pic_3').value=='')
			{
			$('pic_p_3').src='../Images/nopic_supply.gif'
			}
			else
			{
			$('pic_p_3').src=$('pic_3').value
			}
		}
function dels_1()
	{
		document.s_form.pic_1.value=''
	}
function dels_2()
	{
		document.s_form.pic_2.value=''
	}
function dels_3()
	{
		document.s_form.pic_3.value=''
	}
</script>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->





