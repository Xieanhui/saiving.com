<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="2"></td>
  </tr>
</table>
<table width="95%" border="0" align="center" cellpadding="2" cellspacing="1" class="leftframetable">
  <tr classid="VoteManage"> 
    <td width="2%" height="21"><img src="../Images/Folder/main_sys.gif" width="15" height="17"></td>
    <td width="98%"><table width="120" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="47"  class="titledaohang" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(Menuid_main,Menuid_sub);"  language=javascript><font style="font-size:12px">主系统</font></td>
          <td width="73"  class="titledaohang" id=item$pval[CatID]) style="CURSOR: hand"  onmouseup="opencat(Menuid_sub,Menuid_main);"  language=javascript> 
            <div align="left"><font style="font-size:12px">子系统</font></div></td>
        </tr>
      </table></td>
  </tr>
</table>
<table width="95%" border="0" align="center" cellpadding="2" cellspacing="1" class="leftframetable" Id="Menuid_main" style="display:none">
  <tr> 
    <td width="3%"><div align="right"><img src="../Images/Folder/sub_sys.gif" width="15" height="17"></div></td>
    <td width="97%" class="titledaohang"><strong>主系统导航</strong></td>
  </tr>
  <tr> 
    <td colspan="2"><table width="95"  border="0" align="left" cellpadding="2" cellspacing="0">
		<%if MF_Check_Pop_TF("MF_Templet") then %>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
          <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
          <td><a href="../Templets_List.asp" target="ContentFrame" class="lefttop">模板管理</a></td>
        </tr>
		<%
		end if
		if MF_Check_Pop_TF("MF_Style") then 
		%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
          <td><a href="../All_Label_style.asp" target="ContentFrame">样式管理</a></td>
        </tr>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;">
          <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
          <td><a href="../All_Label_Stock.asp" target="ContentFrame">标签库管理</a></td>
        </tr>
		<%
		end if
		if MF_Check_Pop_TF("MF_Public") then 
		%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
          <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
          <td><a href="../Sys_Public.asp" target="ContentFrame">发布管理</a></td>
        </tr>
		<%
		end if
		if MF_Check_Pop_TF("MF_SysSet") then 
		%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="RecyleManage"> 
          <td width="22%" valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
          <td width="78%"><a href="../SysParaSet.asp" target="ContentFrame" class="lefttop">参数设置</a></td>
        </tr>
		<%
		end if
		if MF_Check_Pop_TF("MF_Const") then 
		%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="AdsManage" style="display:;"> 
          <td valign="top"><div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
          <td><a href="../SysConstSet.asp" target="ContentFrame">配置文件</a></td>
        </tr>
		<%
		end if
		if MF_Check_Pop_TF("MF_SubSite") then 
		%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="AdsManage" style="display:;"> 
          <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
          <td><a href="../SubSysSet_List.asp" target="ContentFrame" class="lefttop">子系统维护</a></td>
        </tr>
		<%
		end if
		%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
          <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
          <td><a href="../SysAdmin_list.asp" target="ContentFrame" class="lefttop">管理员管理</a></td>
        </tr>
		<%
		if MF_Check_Pop_TF("MF_DataFix") then 
		%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
          <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
          <td><a href="../DataManage.asp" target="ContentFrame">数据库维护</a></td>
        </tr>
		<%
		end if
		if MF_Check_Pop_TF("MF_Log") then 
		%>
        <tr classid="VoteManage" style="display:;"> 
          <td valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
          <td><a href="../Sys_Oper_Log.asp" target="ContentFrame">日志管理</a></td>
        </tr>
		<%
		end if
		if MF_Check_Pop_TF("MF_Define") then 
		%>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
          <td height="20" valign="top"> <div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
          <td><a href="../DefineTable_Manage.asp" target="ContentFrame" class="lefttop">自定义字段</a></td>
        </tr>
		<%end if%>
      </table></td>
  </tr>
</table>
<table width="95%" border="0" align="center" cellpadding="2" cellspacing="1" class="leftframetable" Id="Menuid_sub" style="display:none">
  <tr> 
    <td width="3%"><div align="right"><img src="../Images/Folder/sub_sys.gif" width="15" height="17"></div></td>
    <td width="97%" class="titledaohang"><strong>子系统导航</strong></td>
  </tr>
  <tr> 
    <td colspan="2"> <table width="95" border="0" align="left" cellpadding="2" cellspacing="0">
        <%
  '得到子类列表
    Dim obj_sub_rs
	Set obj_sub_rs=server.CreateObject(G_FS_RS)
	obj_sub_rs.open "select Sub_Sys_Name,Sub_Sys_Path,Sub_Sys_Index,Sub_Sys_Installed From [FS_MF_Sub_Sys] where Sub_Sys_Installed=1 order by id asc",Conn,1,1
	do while Not obj_sub_rs.eof 
  %>
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" classid="VoteManage" style="display:;"> 
          <td width="22%" height="21" valign="top">
<div align="center"><img src="../Images/Folder/folderclosed.gif"></div></td>
          <td width="78%"><a href="../<% = obj_sub_rs("Sub_Sys_Path")%>/<% = obj_sub_rs("Sub_Sys_Index")%>" class="lefttop"> 
            <% = obj_sub_rs("Sub_Sys_Name")%>
            </a></td>
        </tr>
        <%
	   obj_sub_rs.movenext
   Loop
   obj_sub_rs.close:set obj_sub_rs = nothing
   %>
      </table></td>
  </tr>
</table>
<script language="JavaScript" type="text/JavaScript">
function opencat(cat,hidden)
{
	  if(cat.style.display=="none")
	  {
		 cat.style.display="";
		 hidden.style.display="none";
	  }
	  else
	  {
		 cat.style.display="none"; 
	  }

}
</script>






