<% Option Explicit %>
<%
Dim PreviewImagePath,FileExtName,FileIconDic,FileIcon,AvailableShowTypeStr,PicPara
PreviewImagePath = Request("FilePath")
AvailableShowTypeStr = "jpg,gif,bmp,pst,png,ico"
Set FileIconDic = CreateObject("Scripting.Dictionary")
FileIconDic.Add "txt","../../Images/FileIcon/txt.gif"
FileIconDic.Add "gif","../../Images/FileIcon/gif.gif"
FileIconDic.Add "exe","../../Images/FileIcon/exe.gif"
FileIconDic.Add "asp","../../Images/FileIcon/asp.gif"
FileIconDic.Add "html","../../Images/FileIcon/html.gif"
FileIconDic.Add "htm","../../Images/FileIcon/html.gif"
FileIconDic.Add "jpg","../../Images/FileIcon/jpg.gif"
FileIconDic.Add "jpeg","../../Images/FileIcon/jpg.gif"
FileIconDic.Add "pl","../../Images/FileIcon/perl.gif"
FileIconDic.Add "perl","../../Images/FileIcon/perl.gif"
FileIconDic.Add "zip","../../Images/FileIcon/zip.gif"
FileIconDic.Add "rar","../../Images/FileIcon/zip.gif"
FileIconDic.Add "gz","../../Images/FileIcon/zip.gif"
FileIconDic.Add "doc","../../Images/FileIcon/doc.gif"
FileIconDic.Add "xml","../../Images/FileIcon/xml.gif"
FileIconDic.Add "xsl","../../Images/FileIcon/xml.gif"
FileIconDic.Add "dtd","../../Images/FileIcon/xml.gif"
FileIconDic.Add "vbs","../../Images/FileIcon/vbs.gif"
FileIconDic.Add "js","../../Images/FileIcon/vbs.gif"
FileIconDic.Add "wsh","../../Images/FileIcon/vbs.gif"
FileIconDic.Add "sql","../../Images/FileIcon/script.gif"
FileIconDic.Add "bat","../../Images/FileIcon/script.gif"
FileIconDic.Add "tcl","../../Images/FileIcon/script.gif"
FileIconDic.Add "eml","../../Images/FileIcon/mail.gif"
FileIconDic.Add "swf","../../Images/FileIcon/flash.gif"
if PreviewImagePath = "" then
	PreviewImagePath = "../../../sys_Images/DefaultPreview.gif"
else
	FileExtName = Right(PreviewImagePath,Len(PreviewImagePath)-InStrRev(PreviewImagePath,"."))
	if InStr(AvailableShowTypeStr,FileExtName) = 0 then
		FileIcon = FileIconDic.Item(LCase(FileExtName))
		if FileIcon = "" then
			FileIcon = "../../Images/FileIcon/unknown.gif"
		end if
		PreviewImagePath = FileIcon
		PicPara = " width=""30"" height=""30"" "
	else
		PicPara = ""
	end if
		'Response.Write(PreviewImagePath & "<br>" &FileExtName)
		'Response.End
end if
Set FileIconDic = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>Õº∆¨‘§¿¿</TITLE>
</HEAD>
<BODY topmargin="0" leftmargin="0">
<TABLE width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <TR>
    <TD align="center" valign="middle" width="100%" height="100%"><div align="center"><IMG <% = PicPara %> src="<% = PreviewImagePath %>" width="185"></div></TD>
  </TR>
</TABLE>
</BODY>
</HTML>






