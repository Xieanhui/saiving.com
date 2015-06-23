<% Option Explicit %>
<!--#include file="../FS_Inc/Const.asp" -->
<!--#include file="../FS_Inc/Function.asp" -->
<!--#include file="../FS_InterFace/MF_Function.asp" -->
<!--#include file="lib/strlib.asp" -->
<!--#include file="lib/UserCheck.asp" -->
<%
dim str_type,rs,C_rs,str_Id,str_URL,mf_rs,str_URL_domain,mf_domain
str_type = NoSqlHack(Request.QueryString("type"))
str_Id = CintStr(Request.QueryString("Url"))
'获得参数MF
set mf_rs = Conn.execute("select top 1 MF_Domain From FS_MF_Config")
mf_domain = mf_rs(0)
mf_rs.close:set mf_rs = nothing
select Case str_type
	case "NS"
		dim ns_rs,NewsDir_ns,IsDomain_ns
		set ns_rs = Conn.execute("select Top 1 NewsDir,IsDomain from FS_NS_SysParam")
		if ns_rs.eof then
			response.Write "找不到配置信息"
			response.end
			ns_rs.close:set ns_rs = nothing
		else
			NewsDir_ns = ns_rs(0)
			IsDomain_ns = ns_rs(1)
			ns_rs.close:set ns_rs = nothing
		end if
		set rs = Conn.execute("select IsURL,URLAddress,SaveNewsPath,FileName,FileExtName,ClassId From FS_NS_News where islock=0 and isRecyle=0 and isdraft=0 and Id="&CintStr(str_Id))
		if rs.eof then
			response.Write("找不到记录")
			response.End
			rs.close:set rs =nothing
		else
			set c_rs = Conn.execute("select ClassEname,[Domain],SavePath from FS_NS_NewsClass where ClassiD='"&rs("ClassID")&"'")
			if c_rs.eof then
				response.Write("参数丢失")
				response.End
				c_rs.close:set c_rs = nothing
			else
				if c_rs("Domain")<>"" then
					str_URL_domain = "http://" & c_rs("Domain")
				else
					if IsDomain_ns<>"" then
						str_URL_domain = "http://" & IsDomain_ns
					else
						str_URL_domain = "http://" & mf_domain
					end if
				end if
				if rs("IsURL") = 0 then
					str_URL = str_URL_domain & c_rs("SavePath") & "/"& c_rs("ClassEname")&rs("SaveNewsPath")& "/" &rs("FileName")&"."&rs("FileExtName")
				else
					str_URL = rs("URLAddress")
				end if
				Call Url(str_URL)
				c_rs.close:set c_rs = nothing
			end if
			rs.close:set rs =nothing
		end if
	Case "MS"
		dim ms_rs,NewsDir_ms,IsDomain_ms
		set ms_rs = Conn.execute("select Top 1 SavePath,isDomain from FS_MS_SysPara")
		if ms_rs.eof then
			response.Write "找不到配置信息"
			response.end
			ms_rs.close:set ms_rs = nothing
		else
			NewsDir_ms = ms_rs(0)
			IsDomain_ms = ms_rs(1)
			ms_rs.close:set ms_rs = nothing
		end if
		set rs = Conn.execute("select SavePath,FileName,FileExtName,ClassId From FS_MS_Products where ReycleTF=0 and "&all_substring&"(StyleFlagBit,9,1)=0 and Id="&Cintstr(str_Id))
		if rs.eof then
			response.Write("找不到记录,商品已经删除或者已经被管理员锁定")
			response.End
			rs.close:set rs =nothing
		else
			set c_rs = Conn.execute("select ClassEName,[Domain],SavePath from FS_MS_ProductsClass where ClassiD='"&rs("ClassID")&"'")
			if c_rs.eof then
				response.Write("参数丢失")
				response.End
				c_rs.close:set c_rs = nothing
			else
				if c_rs("Domain")<>"" then
					str_URL_domain = "http://" & c_rs("Domain")
				else
					if IsDomain_ms<>"" then
						str_URL_domain = "http://" & IsDomain_ms
					else
						str_URL_domain = "http://" & mf_domain
					end if
				end if
				str_URL = str_URL_domain & replace(c_rs("SavePath") & "/"& c_rs("ClassEname")&rs("SavePath")& "/" &rs("FileName")&"."&rs("FileExtName"),"//","/")
				Call Url(str_URL)
				c_rs.close:set c_rs = nothing
			end if
			rs.close:set rs =nothing
		end if
	Case "DS"
		dim ds_rs,DownDir_ds,IsDomain_ds
		set ds_rs = Conn.execute("select Top 1 DownDir,IsDomain from FS_DS_SysPara")
		if ds_rs.eof then
			response.Write "找不到配置信息"
			response.end
			ds_rs.close:set ds_rs = nothing
		else
			DownDir_ds = ds_rs(0)
			IsDomain_ds = ds_rs(1)
			ds_rs.close:set ds_rs = nothing
		end if
		set rs = Conn.execute("select * From FS_DS_List where AuditTF=1 and Id="&CintStr(str_Id))
		if rs.eof then
			response.Write("找不到记录,下载已经删除或者已经被管理员锁定")
			response.End
			rs.close:set rs =nothing
		else
			set c_rs = Conn.execute("select ClassEName,[Domain],SavePath from FS_DS_Class where ClassiD='"&rs("ClassID")&"'")
			if c_rs.eof then
				response.Write("参数丢失")
				response.End
				c_rs.close:set c_rs = nothing
			else
				if c_rs("Domain")<>"" then
					str_URL_domain = "http://" & c_rs("Domain")
				else
					if IsDomain_ms<>"" then
						str_URL_domain = "http://" & IsDomain_ds
					else
						str_URL_domain = "http://" & mf_domain
					end if
				end if
				str_URL = str_URL_domain & replace(c_rs("SavePath") & "/"& c_rs("ClassEname")&rs("SavePath")& "/" &rs("FileName")&"."&rs("FileExtName"),"//","/")
				Call Url(str_URL)
				c_rs.close:set c_rs = nothing
			end if
			rs.close:set rs =nothing
		end if
	Case "LS"
		dim rs_log,rs_log_dir
		set rs_log = User_Conn.execute("select top 1 Dir From FS_ME_iLogSysParam")
		if rs_log.eof then
			rs_log_dir = "blog"
			rs_log.close:set rs_log = nothing
		else
			rs_log_dir = rs_log("Dir")
			rs_log.close:set rs_log = nothing
		end if
		str_URL = "http://"& mf_domain & "/" & rs_log_dir & "/Blog.asp?id="&str_Id
		Call Url(str_URL)
	Case "PH"
		dim rs_ph,rs_ph_dir
		set rs_ph = User_Conn.execute("select top 1 Dir From FS_ME_iLogSysParam")
		if rs_ph.eof then
			rs_ph_dir = "blog"
			rs_ph.close:set rs_ph = nothing
		else
			rs_ph_dir = rs_ph("Dir")
			rs_ph.close:set rs_ph = nothing
		end if
		str_URL = "http://"& mf_domain & "/" & rs_ph_dir & "/ShowPhoto.asp?id="&str_Id
		Call Url(str_URL)
	Case "HS"
		str_URL = "http://"& mf_domain & "/House/HouseRead.asp?id="&str_Id
		Call Url(str_URL)
	Case "SD"
		dim rs_SD,rs_sd_dir,rs_sd_domain
		set rs_SD = User_Conn.execute("select top 1 [Domain],SavePath From FS_SD_Config")
		if rs_SD.eof then
			rs_sd_dir = "Supply"
			rs_SD.close:set rs_SD = nothing
		else
			rs_sd_dir = rs_SD("SavePath")
			rs_sd_domain =  rs_SD("Domain")
			rs_SD.close:set rs_SD = nothing
		end if
		if rs_sd_domain<>"" then
			str_URL = "http://"& rs_sd_domain & "/"& rs_sd_dir &"/Supply.asp?id="&str_Id
		else
			str_URL = "http://"& mf_domain & "/"& rs_sd_dir &"/Supply.asp?id="&str_Id
		end if
		Call Url(str_URL)
end select
HouseRead.asp
Function Url(str_URL)
	response.Redirect str_URL
	response.end
End Function
%>
<!--Powsered by Foosun Inc.,Product:FoosunCMS V5.0系列-->





