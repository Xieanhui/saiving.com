<%
Class Cls_News
	Private m_Obj_news_Rs
	Private m_sysID,m_siteName,m_keyWords,m_newsDir,m_isDomain,m_fileNameRule
	Private m_ReycleTF,m_fileDirRule,m_classSaveType,m_fileExtName,m_indexPage,m_newsCheck,m_InsideLink
	Private m_refreshFile,m_isOpen,m_indexTemplet,m_isPrintPic,m_picClassid,m_linkType,m_fileChar
	Private m_isCheck,m_isReviewCheck,m_isConstrCheck,m_addNewsType,m_allInfotitle,m_CopyFileTF,m_EditFilesTF
	Private m_RSSTF,m_rssNumber,m_rssdescript,m_RSSPIC,m_rssContentNumber,IsAdPic,AdPicWH,AdPicLink,AdPicAdress
	'调用类的初始值
	Private Sub Class_Initialize() 
 	End Sub
	'释放初始值 
	Private Sub Class_Terminate()  
 	End Sub
	'得到多少位数的随机函数 
	Public Function GetRamCode(f_number)
		Randomize
		Dim f_Randchar,f_Randchararr,f_RandLen,f_Randomizecode,f_iR
		f_Randchar="0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
		f_Randchararr=split(f_Randchar,",") 
		f_RandLen=f_number '定义密码的长度或者是位数
		for f_iR=1 to f_RandLen
			f_Randomizecode=f_Randomizecode&f_Randchararr(Int((21*Rnd)))
		next 
		GetRamCode = f_Randomizecode
	End Function
	
	'得到子类新闻类别分页
	Public Function GetChildClassList(f_classid)
		
	End Function
	
	Public Function GetSysParamDir()
			Dim f_Obj_sysparm,SysParmTF
			Set f_Obj_sysparm = server.CreateObject(G_FS_RS)
			f_Obj_sysparm.Open "select top 1 NewsDir from FS_NS_SysParam",Conn,1,1
			if  not (f_Obj_sysparm.eof or f_Obj_sysparm.bof) then
				GetSysParamDir = Replace("/"& f_Obj_sysparm("NewsDir"),"//","/")
			Else
				GetSysParamDir = ""
			End if
			If GetSysParamDir = "" Then
				GetSysParamDir = "/"
			End If
			f_Obj_sysparm.CLose : Set f_Obj_sysparm = Nothing	
	End Function
	
	'得到栏目名称,返回值GetClassName
	Public Function GetClassName(f_classid)
				Dim f_obj_className_rs
				if f_classid<>"" then
					Set f_obj_className_rs = server.CreateObject(G_FS_RS)
					f_obj_className_rs.Open "select ClassID,ClassName,ParentID from FS_NS_NewsClass where ClassID='"& NoSqlHack(f_classid) &"'",Conn,1,1
					if  not (f_obj_className_rs.eof or f_obj_className_rs.bof) then
							GetClassName =f_obj_className_rs("ClassName")
					Else
							GetClassName ="根栏目"
					End if
				Else
							GetClassName ="根栏目"
				End if
				set f_obj_className_rs = nothing
	End Function
	'添加新闻的时候，获得栏目中文名称
	Public Function GetAdd_ClassName(f_classid) 
				Dim f_obj_addclassName_rs
				Set f_obj_addclassName_rs = server.CreateObject(G_FS_RS)
				f_obj_addclassName_rs.Open "select ClassID,ClassName from FS_NS_NewsClass where ClassID='"& NoSqlHack(f_classid) &"'",Conn,1,1 
				if  not (f_obj_addclassName_rs.eof or f_obj_addclassName_rs.bof) then 
						GetAdd_ClassName =f_obj_addclassName_rs("ClassName") 
				Else
						GetAdd_ClassName =""
				End if
				set f_obj_addclassName_rs = nothing
	End Function
	
	'得到自定义字段
	Public Function GetDefineClassId()
				Dim f_obj_Define_rs
				GetDefineClassId = ""
				Set f_obj_Define_rs = server.CreateObject(G_FS_RS)
				f_obj_Define_rs.Open "select DefineName,DefineID from FS_MF_DefineTableClass where ParentID=0 Order by DefineID desc",Conn,1,1
				if  not (f_obj_Define_rs.eof or f_obj_Define_rs.bof) then
					Do while Not f_obj_Define_rs.eof 
						if lng_DefineID = f_obj_Define_rs("DefineID")  then
							GetDefineClassId = GetDefineClassId & "<option value="""& f_obj_Define_rs("DefineID") &""" selected>---" & f_obj_Define_rs("DefineName") &"</option>"
						Else
							GetDefineClassId = GetDefineClassId & "<option value="""& f_obj_Define_rs("DefineID") &""" >---" & f_obj_Define_rs("DefineName") &"</option>"
						End if
						f_obj_Define_rs.movenext
					Loop
				Else
					GetDefineClassId = GetDefineClassId & "<option value="""">没有自定义分类</option>"
				End if
	End Function
	
	Public Function IsSelfRefer()
		Dim sHttp_Referer, sServer_Name
		sHttp_Referer = NoSqlHack(CStr(Request.ServerVariables("HTTP_REFERER")))
		sServer_Name = NoSqlHack(CStr(Request.ServerVariables("SERVER_NAME")))
		If Mid(sHttp_Referer, 8, Len(sServer_Name)) = sServer_Name Then
			IsSelfRefer = True
		Else
			IsSelfRefer = False
		End If
	End Function 
	'得到子类新闻栏目
	Public Function GetChildNewsList(TypeID,CompatStr)
		Dim AndSQL
		AndSQL = GetAndSQLOfSearchClass("NS013")
	
	    Dim ChildNewsRs,ChildTypeListStr,TempStr,TmpStr,f_isUrlStr
	    Set ChildNewsRs = Conn.Execute("Select id,orderid,ClassName,ClassEName,ClassID,IsUrl,isConstr,isShow,[Domain] from FS_NS_NewsClass where ParentID='" & NoSqlHack(TypeID) & "'  and ReycleTF=0 " & AndSQL & "  order by Orderid desc,id desc" )
	   
	    'TempStr =CompatStr & "<img src=""images/L.gif""></img>"
			if CompatStr="oooo" then 
				TempStr = "&nbsp;<img src=""images/L.gif""></img>"
			else
				TempStr = "<img src=""images/L.gif""></img>"
			end if
		    do while Not ChildNewsRs.Eof
             If True then '  Get_SubPop_TF(TypeID,"NS001","NS","news") or News_ChildNewsPower(TypeID) Then             
           
        	      TmpStr = ""
			      if ChildNewsRs("IsUrl") = 1 then
				    TmpStr = TmpStr & "<font color=red>外部</font>&nbsp;┆&nbsp;" 
			      Else
				    TmpStr = TmpStr & "系统&nbsp;┆&nbsp;" 
			      End if
			      if ChildNewsRs("isConstr") = 1 then
				    TmpStr = TmpStr & "<font color=red>稿</font>&nbsp;┆&nbsp;" 
			      Else
				    TmpStr = TmpStr & "<strike>稿</strike>&nbsp;┆&nbsp;" 
			      End if
			      if ChildNewsRs("isShow") = 1 then
				    TmpStr = TmpStr & "<font color=red>显示</font>&nbsp;┆&nbsp;" 
			      Else
				    TmpStr = TmpStr & "隐藏&nbsp;┆&nbsp;" 
			      End if
			      if len(ChildNewsRs("Domain")) >5 then
				    TmpStr = TmpStr & "<font color=red>域</font>&nbsp;┆&nbsp;"
			      Else
				    TmpStr = TmpStr & "本&nbsp;┆&nbsp;"
			      End if
	  		    GetChildNewsList = GetChildNewsList & "<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>"&Chr(13) & Chr(10)
			    GetChildNewsList = GetChildNewsList & "<td width=""5%"" class=""hback"" align=""center"">"& ChildNewsRs("id")&"</td>" & Chr(13) & Chr(10)
			    if ChildNewsRs("IsUrl") = 1 then
				    f_isUrlStr = ""
			    Else
				    f_isUrlStr = "["&ChildNewsRs("ClassEName")&"]"
			    End if
				dim obj_news_rs_1,tf
				Set obj_news_rs_1 = server.CreateObject(G_FS_RS)
				obj_news_rs_1.Open "Select Count(ID) from FS_NS_NewsClass where ParentID='"& ChildNewsRs("ClassID") &"'",Conn,1,1
				if obj_news_rs_1(0)>0 then
					tf=  "<a href=""javascript:void(0);"" onclick=""getChildClassID('"&ChildNewsRs("ClassID")&"','oooo')"" title=""展开子类""><img border=""0"" src=""images/jia.gif""></img></a>"
				Else
					tf= "<img src=""images/-.gif""></img>"
				End if
				
				
				If Get_SubPop_TF(ChildNewsRs("Classid"),"NS016","NS","news") Then  
					GetChildNewsList = GetChildNewsList & "<td width=""30%"" class=""hback"">&nbsp;"& TempStr & tf &"<a href=""Class_add.asp?ClassID="&ChildNewsRs("ClassID")&"&Action=edit"">" & ChildNewsRs("ClassName") & "</a>&nbsp;<font style=""font-size:11.5px;"">"& f_isUrlStr &"</font></td>" & Chr(13) & Chr(10) 
				else
					GetChildNewsList = GetChildNewsList & "<td width=""30%"" class=""hback"">&nbsp;"& TempStr & tf &"" & ChildNewsRs("ClassName") & "<font style=""font-size:11.5px;"">"& f_isUrlStr &"</font></td>" & Chr(13) & Chr(10) 
				end if
				GetChildNewsList = GetChildNewsList & "<td width=""7%"" class=""hback"" align=""center"">"&ChildNewsRs("OrderID")&"</td>" & Chr(13) & Chr(10)
			    GetChildNewsList = GetChildNewsList & "<td width=""22%"" class=""hback"" align=""center""><div align=""center"">"& TmpStr &"</div></td>" & Chr(13) & Chr(10)
			    
				GetChildNewsList = GetChildNewsList & "<td width=""36%"" class=""hback"" align=""center""><div align=""center"">"
				GetChildNewsList = GetChildNewsList & "<a href=""NewClass_review.asp?id="&ChildNewsRs("ClassID")&""" target=""_blank"">预览</a>"
				 If Get_SubPop_TF(ChildNewsRs("Classid"),"NS016","NS","class") then
					GetChildNewsList = GetChildNewsList & "┆<a href=""Class_add.asp?ClassID="&ChildNewsRs("ClassID")&"&Action=add"">添加子栏目</a>"
				else
					GetChildNewsList = GetChildNewsList & "┆" & GetDisableSpanCode("添加子栏目")
				end if
				If Get_SubPop_TF(ChildNewsRs("Classid"),"NS017","NS","class")	 then			
					GetChildNewsList = GetChildNewsList & "┆<a href=""Class_add.asp?ClassID="&ChildNewsRs("ClassID")&"&Action=edit"">修改</a>"
				else
					GetChildNewsList = GetChildNewsList & "┆" & GetDisableSpanCode("修改")
				end if
				If Get_SubPop_TF(ChildNewsRs("Classid"),"NS023","NS","class") then
					GetChildNewsList = GetChildNewsList & "┆<a href=""Class_Action.asp?ClassID="&ChildNewsRs("ClassID")&"&Action=clear"" onClick=""{if(confirm('确定清空此栏目下信息吗?')){return true;}return false;}"">清空</a>"
				else
					GetChildNewsList = GetChildNewsList & "┆" & GetDisableSpanCode("清空")
				end if
				If Get_SubPop_TF(ChildNewsRs("Classid"),"NS021","NS","class") then
					GetChildNewsList = GetChildNewsList & "┆<a href=""Class_Action.asp?ClassID="&ChildNewsRs("ClassID")&"&Action=del""   onClick=""{if(confirm('确定删除您所选择的栏目吗？\n\n此栏目下的子类也将被删除!!')){return true;}return false;}"">删除</a>"
				else
					GetChildNewsList = GetChildNewsList & "┆" & GetDisableSpanCode("删除")
				end if
				If Get_SubPop_TF(ChildNewsRs("Classid"),"NS022","NS","class") then
					GetChildNewsList = GetChildNewsList & "┆<a href=""Class_makerss.asp?signxml=one&cid="&ChildNewsRs("Classid")&""" title=""生成此栏目Rss"">Rss</a>"
				else
					GetChildNewsList = GetChildNewsList & "┆" & GetDisableSpanCode("Rss")
				end if
				
				GetChildNewsList = GetChildNewsList & vbNewLine&"<input name=""Cid"" type=""checkbox"" id=""Cid"" value="""& ChildNewsRs("ClassID")&"""></div></td>" & Chr(13) & Chr(10)
			    GetChildNewsList = GetChildNewsList & "</tr>" & Chr(13) & Chr(10)
			  'else
			   ' GetChildNewsList = GetChildNewsList & "<td width=""30%"" class=""hback"">&nbsp;"& TempStr & tf &"" & ChildNewsRs("ClassName") & "<font style=""font-size:11.5px;"">"& f_isUrlStr &"</font></td>" & Chr(13) & Chr(10) 
			  '  GetChildNewsList = GetChildNewsList & "<td width=""7%"" class=""hback"" align=""center"">"&ChildNewsRs("OrderID")&"</td>" & Chr(13) & Chr(10)
			  '  GetChildNewsList = GetChildNewsList & "<td width=""22%"" class=""hback"" align=""center""><div align=""center"">"& TmpStr &"</div></td>" & Chr(13) & Chr(10)
			 '   GetChildNewsList = GetChildNewsList & "<td width=""36%"" class=""hback"" align=""center""><div align=""center""><font color=red>没有此目录的操作权限！</font></div></td>" & Chr(13) & Chr(10)
			 '   GetChildNewsList = GetChildNewsList & "</tr>" & Chr(13) & Chr(10)
			 ' end if
			 end if
			 	GetChildNewsList = GetChildNewsList &"<tr  class=""hback""><td colspan=""5""><div id=""c_"&ChildNewsRs("ClassID")&"""></div></td></tr>"
				'
			    'GetChildNewsList = GetChildNewsList &GetChildNewsList(ChildNewsRs("ClassID"),TempStr)
			    ChildNewsRs.MoveNext
			    
		loop
		ChildNewsRs.Close
		Set ChildNewsRs = Nothing
	End Function
		
	'获得排序号子类
	Public Function GetChildNewsList_order(TypeID,CompatStr)  
		Dim Order_ChildNewsRs,ChildTypeListStr,TempStr,TmpStr,f_isUrlStr,lng_GetCount
		Set Order_ChildNewsRs = Conn.Execute("Select id,orderid,ClassName,ClassEName,ClassID from FS_NS_NewsClass where ParentID='" & NoSqlHack(TypeID) & "'  and ReycleTF=0  order by Orderid desc,id desc" )
		TempStr =CompatStr & "<img src=""images/L.gif""></img>"
		do while Not Order_ChildNewsRs.Eof
				GetChildNewsList_order = GetChildNewsList_order & "<form name=""ClassForm"" method=""post"" action=""Class_Action.asp""><tr>"&Chr(13) & Chr(10)
				GetChildNewsList_order = GetChildNewsList_order & "<td class=""hback"">&nbsp;"& TempStr &"<Img src=""images/-.gif""></img>" & Order_ChildNewsRs("ClassName") & "</td>" & Chr(13) & Chr(10) 
				GetChildNewsList_order = GetChildNewsList_order & "<td class=""hback"" align=""center"">"& Order_ChildNewsRs("ID")&"</td>" & Chr(13) & Chr(10)
				GetChildNewsList_order = GetChildNewsList_order & "<td class=""hback"" align=""center"">" & Chr(13) & Chr(10)
				GetChildNewsList_order = GetChildNewsList_order & "<input name=""OrderID"" type=""text"" id=""OrderID"" value="& Order_ChildNewsRs("OrderID") &" size=""4"" maxlength=""3"">" & Chr(13) & Chr(10)
				GetChildNewsList_order = GetChildNewsList_order & "<input name=""ClassID"" type=""hidden"" id=""ClassID"" value="& Order_ChildNewsRs("ClassID") &">" & Chr(13) & Chr(10)
				GetChildNewsList_order = GetChildNewsList_order & "<input name=""Action"" type=""hidden"" id=""ClassID"" value=""Order_n"">" & Chr(13) & Chr(10)
				GetChildNewsList_order = GetChildNewsList_order & "<input type=""submit"" name=""Submit"" value=""更新权重(排列序号)"">" & Chr(13) & Chr(10)
				GetChildNewsList_order = GetChildNewsList_order & "</td>" & Chr(13) & Chr(10)
				GetChildNewsList_order = GetChildNewsList_order & "</tr></form>" & Chr(13) & Chr(10)
			GetChildNewsList_order = GetChildNewsList_order &GetChildNewsList_order(Order_ChildNewsRs("ClassID"),TempStr)
			Order_ChildNewsRs.MoveNext
		loop
		Order_ChildNewsRs.Close
		Set Order_ChildNewsRs = Nothing
	End Function
	'得到子类select列表,多选
	Public Function News_ChildNewsList(TypeID,f_CompatStr)  
		Dim f_ChildNewsRs_1,ChildTypeListStr,f_TempStr,f_isUrlStr,lng_GetCount
		Set f_ChildNewsRs_1 = Conn.Execute("Select id,orderid,ClassName,ClassID,ParentID from FS_NS_NewsClass where ParentID='" & NoSqlHack(TypeID) & "' and ReycleTF=0 order by Orderid desc,id desc" )
		f_TempStr =f_CompatStr & "┄"
		do while Not f_ChildNewsRs_1.Eof
				News_ChildNewsList = News_ChildNewsList & "<option value="""& f_ChildNewsRs_1("ClassID") &""">"
				News_ChildNewsList = News_ChildNewsList & "├" & f_TempStr &  f_ChildNewsRs_1("ClassName") 
				News_ChildNewsList = News_ChildNewsList & "</option>" & Chr(13) & Chr(10)
				News_ChildNewsList = News_ChildNewsList &News_ChildNewsList(f_ChildNewsRs_1("ClassID"),f_TempStr)
			f_ChildNewsRs_1.MoveNext
		loop
		f_ChildNewsRs_1.Close
		Set f_ChildNewsRs_1 = Nothing
	End Function
	'得到子类select列表,单ID
	Public Function UniteChildNewsList(TypeID,f_CompatStr)  
		Dim f_ChildNewsRs,ChildTypeListStr,f_TempStr,f_isUrlStr,lng_GetCount
		Set f_ChildNewsRs = Conn.Execute("Select id,orderid,ClassName,ClassID,ParentID from FS_NS_NewsClass where ParentID='" & NoSqlHack(TypeID) & "' and ReycleTF=0 order by Orderid desc,id desc" )
		f_TempStr =f_CompatStr & "┄"
		do while Not f_ChildNewsRs.Eof
				UniteChildNewsList = UniteChildNewsList & "<option value="""& f_ChildNewsRs("ClassID") &","& f_ChildNewsRs("ParentID") &""">"
				UniteChildNewsList = UniteChildNewsList & "├" &  f_TempStr & f_ChildNewsRs("ClassName") 
				UniteChildNewsList = UniteChildNewsList & "</option>" & Chr(13) & Chr(10)
				UniteChildNewsList = UniteChildNewsList &UniteChildNewsList(f_ChildNewsRs("ClassID"),f_TempStr)
			f_ChildNewsRs.MoveNext
		loop
		f_ChildNewsRs.Close
		Set f_ChildNewsRs = Nothing
	End Function
	
	'删除子类新闻栏目
	Public Function DelChildNewsList(TypeID,f_tmp_del_rcy)  
		Dim del_ChildNewsRs
		Set del_ChildNewsRs = Conn.Execute("Select id,ParentID,ClassID from FS_NS_NewsClass where ParentID='" & NoSqlHack(TypeID) & "' order by id desc" )
		do while Not del_ChildNewsRs.Eof
			if f_tmp_del_rcy =0 then'彻底删除
				Conn.Execute("Delete From FS_NS_NewsClass Where ClassID ='"&  del_ChildNewsRs("ClassID") &"'")
				'删除新闻
				Conn.execute("Delete From FS_NS_News Where ClassID='"& del_ChildNewsRs("ClassID") &"'") 
			Else'删除到回收站
				Conn.Execute("Update FS_NS_NewsClass set ReycleTF=1 Where ClassID ='"&  del_ChildNewsRs("ClassID") &"'")
				'删除新闻 
				Conn.execute("Update FS_NS_News set isRecyle=1 Where ClassID='"& del_ChildNewsRs("ClassID") &"'") 
			End if
			'获得下级分类列表，并进行删除操作
			DelChildNewsList = DelChildNewsList &DelChildNewsList(del_ChildNewsRs("ClassID"),f_tmp_del_rcy)
			del_ChildNewsRs.MoveNext
		loop
		del_ChildNewsRs.Close
		Set del_ChildNewsRs = Nothing	
	End Function
	'检查英文名称是否合法
   Public Function chkinputchar(f_char)
		Dim f_name, i, c
		f_name = f_char
		chkinputchar = True
		If Len(f_name) <= 0 Then
			chkinputchar = False
			Exit Function
		End If
		For i = 1 To Len(f_name)
		   c = Mid(f_name, i, 1)
			If InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ@,.0123456789|-_", c) <= 0  Then
			   chkinputchar = False
			Exit Function
		   End If
	   Next
	End Function
		
	Public Function GetSysParam()
			Dim f_Obj_sysparm,SysParmTF
			SysParmTF = True
			Set f_Obj_sysparm=server.CreateObject(G_FS_RS)
			f_Obj_sysparm.Open "select top 1 sysid,SiteName,Keywords,NewsDir,IsDomain,FileNameRule,FileDirRule,ClassSaveType,FileExtName,IndexPage,NewsCheck,isOpen,IndexTemplet,LinkType,isCheck,isReviewCheck,isConstrCheck,IsCopyFileTF,IsEditFileTF,addNewsType,AllInfotitle,InsideLink,ReycleTF,RSSTF,rssNumber,rssdescript,RSSPIC,rssContentNumber from FS_NS_SysParam",Conn,1,1
			if  not (f_Obj_sysparm.eof or f_Obj_sysparm.bof) then
				m_sysID = f_Obj_sysparm("sysID")
				m_siteName = f_Obj_sysparm("siteName")
				m_keywords= f_Obj_sysparm("Keywords")
				m_newsDir= f_Obj_sysparm("NewsDir")
				m_isDomain= f_Obj_sysparm("IsDomain")
				m_fileNameRule= f_Obj_sysparm("FileNameRule")
				m_fileDirRule= f_Obj_sysparm("FileDirRule")
				m_classSaveType= f_Obj_sysparm("ClassSaveType")
				m_fileExtName= f_Obj_sysparm("FileExtName")
				m_indexPage= f_Obj_sysparm("IndexPage")
				m_newsCheck= f_Obj_sysparm("NewsCheck")
				m_isOpen= f_Obj_sysparm("isOpen")
				m_indexTemplet= f_Obj_sysparm("IndexTemplet")
				m_linkType= f_Obj_sysparm("LinkType")
				m_isCheck= f_Obj_sysparm("isCheck")
				m_isReviewCheck= f_Obj_sysparm("isReviewCheck")
				m_isConstrCheck= f_Obj_sysparm("isConstrCheck")
				m_CopyFileTF = f_Obj_sysparm("IsCopyFileTF")
				m_EditFilesTF = f_Obj_sysparm("IsEditFileTF")
				m_addNewsType= f_Obj_sysparm("addNewsType")
				m_allInfotitle= f_Obj_sysparm("AllInfotitle")
				m_InsideLink=f_Obj_sysparm("InsideLink")
				m_reycleTF=f_Obj_sysparm("ReycleTF")
				'RSS
				m_RSSTF= f_Obj_sysparm("RSSTF")
				m_rssNumber= f_Obj_sysparm("rssNumber")
				m_rssdescript= f_Obj_sysparm("rssdescript")
				m_RSSPIC= f_Obj_sysparm("RSSPIC")
				m_rssContentNumber=f_Obj_sysparm("rssContentNumber")
				SysParmTF = True
			Else
				m_sysID = ""
				m_siteName = ""
				m_keywords= ""
				m_newsDir= ""
				m_isDomain= ""
				m_fileNameRule= "$$$$$$"
				m_fileDirRule= 0
				m_classSaveType= 0
				m_fileExtName= 0
				m_indexPage= ""
				m_newsCheck= ""
				m_isOpen= 0
				m_indexTemplet= ""
				m_linkType= 0
				m_isCheck= 0
				m_isReviewCheck=0
				m_isConstrCheck= 0
				m_CopyFileTF = 0
				m_EditFilesTF = 0
				m_addNewsType= 0
				m_allInfotitle= ""
				m_InsideLink=0
				m_reycleTF=0
				'RSS
				m_RSSTF= 0
				m_rssNumber= 0
				m_rssdescript= ""
				m_RSSPIC= ""
				m_rssContentNumber=0
				SysParmTF = false
			End if
	End Function
	'赋值
	Public Property Get sysID()				'参数ID  
		sysID = m_sysID
	End Property 
	Public Property Get siteName()				'新闻系统站点标题  
		siteName = m_siteName
	End Property 
		Public Property Get keyWords()				'站点关键字  
		keyWords = m_keyWords
	End Property 
	Public Property Get newsDir()				'新闻系统前台目录 
		newsDir = m_newsDir
	End Property 
		Public Property Get isDomain()				'是否启用供求系统二级域名  
		isDomain = m_isDomain
	End Property 
	Public Property Get fileNameRule()				'新闻文件静态文件生成规则
		fileNameRule = m_fileNameRule
	End Property 
		Public Property Get fileDirRule()				'静态文件生成目录  
		fileDirRule = m_fileDirRule
	End Property 
	Public Property Get classSaveType()				'新闻栏目目录生成首页格式  
		classSaveType = m_classSaveType
	End Property 
	Public Property Get fileExtName()				'生成静态文件扩展名  
		fileExtName = m_fileExtName
	End Property 
	Public Property Get indexPage()				'首页文件名
		indexPage = m_indexPage
	End Property 
		Public Property Get newsCheck()				'发布的新闻是否需要审核 
		newsCheck = m_newsCheck
	End Property 
	Public Property Get isOpen()				'是否开通新闻发布信息 
		isOpen = m_isOpen
	End Property 
	Public Property Get indexTemplet()				'首页模板地址 
		indexTemplet = m_indexTemplet
	End Property 
	Public Property Get linkType()				'采用绝对路径还是相对路径 
		linkType = m_linkType
	End Property 
	Public Property Get isCheck()				'添加的新闻是否审核 
		isCheck = m_isCheck
	End Property 
	Public Property Get isReviewCheck()				'发布的新闻的评论是否要审核  
		isReviewCheck = m_isReviewCheck
	End Property 
	Public Property Get isConstrCheck()				'投稿是否需要审核后才能发布  
		isConstrCheck = m_isConstrCheck
	End Property 
	Public Property Get CopyFileTF()				'投稿审核后是否拷贝文件  
		CopyFileTF = m_CopyFileTF
	End Property 
	Public Property Get EditFilesTF()				'投稿审核后是否允许编辑稿件
		EditFilesTF = m_EditFilesTF
	End Property 
	Public Property Get addNewsType()				'添加新闻采用的方式  
		addNewsType = m_addNewsType
	End Property 
	Public Property Get allInfotitle()				'所有新闻系统站点及全站
		allInfotitle = m_allInfotitle
	End Property 
	public Property get reycleTF()
		reycleTF = m_reycleTF
	End Property
	public Property get InsideLink()
		InsideLink = m_InsideLink
	End Property
	
	'RSS调用参数
	public Property get RSSTF()
		RSSTF = m_RSSTF
	End Property
	public Property get rssNumber()
		rssNumber = m_rssNumber
	End Property
	public Property get rssdescript()
		rssdescript = m_rssdescript
	End Property
	public Property get RSSPIC()
		RSSPIC = m_RSSPIC
	End Property
	public Property get rssContentNumber()
		rssContentNumber = m_rssContentNumber
	End Property
	'获得今日新闻数量
	Public Function GetTodayNewsCount(f_classID) 
			Dim f_obj_cnews_rs
			Set f_obj_cnews_rs = server.CreateObject(G_FS_RS)
			If G_IS_SQL_DB=0 Then
				f_obj_cnews_rs.Open "Select ID from FS_NS_News where ClassID='"& NoSqlHack(f_classID) &"' and datevalue(addtime)=#"&date()&"#",Conn,1,1
			Else
				f_obj_cnews_rs.Open "Select ID from FS_NS_News where ClassID='"& NoSqlHack(f_classID) &"' and datediff(dd,addtime,getdate())=0",Conn,1,1
			End If
			GetTodayNewsCount = "<span class=""tx"">"&f_obj_cnews_rs.recordcount&"</span>)"
			f_obj_cnews_rs.close
			set f_obj_cnews_rs = nothing
	End Function 
	'获得用户文件名
	Public Function strFileNameRule(str,f_idTF,f_id)
		strFileNameRule = ""
		Dim f_strFileNamearr,f_str0,f_str1,f_str2,f_str3,f_str4,Getstr,f_str5,f_str6
		f_strFileNamearr = split(str,"$")
		f_str0 = f_strFileNamearr(0)
		f_str1 = f_strFileNamearr(1)
		f_str2 = f_strFileNamearr(2)
		f_str3 = f_strFileNamearr(3)
		f_str4 = f_strFileNamearr(4)
		f_str5 = f_strFileNamearr(5)
		f_str6 = f_strFileNamearr(6)
		strFileNameRule = strFileNameRule & f_strFileNamearr(0)
		If Instr(1,f_strFileNamearr(1),"Y",1)<>0 then
			if Len(Trim(Cstr(f_strFileNamearr(4))))<>0 then
				strFileNameRule = strFileNameRule & right(year(now),2)&f_strFileNamearr(4)
			Else
				strFileNameRule = strFileNameRule & right(year(now),2)
			End if
		End if
		If Instr(1,f_strFileNamearr(1),"M",1)<>0 then
				if Len(Trim(Cstr(f_strFileNamearr(4))))<>0 then
					strFileNameRule = strFileNameRule & month(now)&f_strFileNamearr(4)
				Else
					strFileNameRule = strFileNameRule& month(now)
				End if
		End if
		If Instr(1,f_strFileNamearr(1),"D",1)<>0 then
				if Len(Trim(Cstr(f_strFileNamearr(4))))<>0 then
					strFileNameRule = strFileNameRule & day(now)&f_strFileNamearr(4)
				Else
					strFileNameRule = strFileNameRule& day(now)
				End if
		End if
		If Instr(1,f_strFileNamearr(1),"H",1)<>0 then
				if Len(Trim(Cstr(f_strFileNamearr(4))))<>0 then
					strFileNameRule = strFileNameRule & hour(now)&f_strFileNamearr(4)
				Else
					strFileNameRule = strFileNameRule& hour(now)
				End if
		End if
		If Instr(1,f_strFileNamearr(1),"I",1)<>0 then
				if Len(Trim(Cstr(f_strFileNamearr(4))))<>0 then
					strFileNameRule = strFileNameRule & minute(now)&f_strFileNamearr(4)
				Else
					strFileNameRule = strFileNameRule& minute(now)
				End if
		End if
		If Instr(1,f_strFileNamearr(1),"S",1)<>0 then
				if Len(Trim(Cstr(f_strFileNamearr(4))))<>0 then
					strFileNameRule = strFileNameRule & second(now)&f_strFileNamearr(4)
				Else
					strFileNameRule = strFileNameRule& second(now)
				End if
		End if
		Randomize
		Dim f_Randchar,f_Randchararr,f_RandLen,f_iR,f_Randomizecode
		f_Randchar="0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
		f_Randchararr=split(f_Randchar,",") 
		If f_strFileNamearr(2)="2" then
			if f_strFileNamearr(3)="1" then
				f_RandLen=2 
				for f_iR=1 to f_RandLen
				f_Randomizecode=f_Randomizecode&f_Randchararr(Int((21*Rnd)))
				next 
				strFileNameRule = strFileNameRule &  f_Randomizecode
			Else
				strFileNameRule = strFileNameRule &  CStr(Int((99 * Rnd) + 1))
			End if
		Elseif f_strFileNamearr(2)="3" then
			if f_strFileNamearr(3)="1" then
				f_RandLen=3 
				for f_iR=1 to f_RandLen
				f_Randomizecode=f_Randomizecode&f_Randchararr(Int((21*Rnd)))
				next 
				strFileNameRule = strFileNameRule &  f_Randomizecode
			Else
				strFileNameRule = strFileNameRule &  CStr(Int((999* Rnd) + 1))
			End if
		Elseif f_strFileNamearr(2)="4" then
			if f_strFileNamearr(3)="1" then
				f_RandLen=4 
				for f_iR=1 to f_RandLen
				f_Randomizecode=f_Randomizecode&f_Randchararr(Int((21*Rnd)))
				next 
				strFileNameRule = strFileNameRule &  f_Randomizecode
			Else
				strFileNameRule = strFileNameRule &  CStr(Int((9999* Rnd) + 1))
			End if
		Elseif f_strFileNamearr(2)="5" then
			if f_strFileNamearr(3)="1" then
				f_RandLen=5 
				for f_iR=1 to f_RandLen
				f_Randomizecode=f_Randomizecode&f_Randchararr(Int((21*Rnd)))
				next 
				strFileNameRule = strFileNameRule &  f_Randomizecode
			Else
				strFileNameRule = strFileNameRule &  CStr(Int((99999* Rnd) + 1))
			End if
	   End if
	 if f_strFileNamearr(5) = "1" then
		 strFileNameRule = strFileNameRule&f_strFileNamearr(4)&"自动编号ID"
	 End if
	 if f_strFileNamearr(6) = "1" then
		 strFileNameRule = strFileNameRule&f_strFileNamearr(4)&"唯一NewsID"
	 End if
		 strFileNameRule = strFileNameRule
	End Function
	'得到新闻关键字下拉菜单
	Public Function GetKeywordslist(f_char,f_number)
		GetKeywordslist = ""
		dim f_obj_kw_Rs
		Set f_obj_kw_Rs = server.CreateObject(G_FS_RS)
		f_obj_kw_Rs.Open "Select GID,G_Name,G_Type,isLock from FS_NS_General where G_Type ="& CintStr(f_number) &" and isLock=0  order by GID desc",Conn,1,1
		do while Not f_obj_kw_Rs.eof 
				if f_char = f_obj_kw_Rs("G_Name") then
					GetKeywordslist = GetKeywordslist & "<option value="""& f_obj_kw_Rs("G_Name")&""" selected>"& f_obj_kw_Rs("G_Name")&"</option>"
				Else
					GetKeywordslist = GetKeywordslist & "<option value="""& f_obj_kw_Rs("G_Name")&""">"& f_obj_kw_Rs("G_Name")&"</option>"
				End if
			f_obj_kw_Rs.movenext
		Loop
		GetKeywordslist = GetKeywordslist
		f_obj_kw_Rs.close:set f_obj_kw_Rs = nothing
	End Function
	
	'得到栏目自定义ID
	Public Function GetCustClassID(f_custclassid)
		Dim obj_cust_rs
		set obj_cust_rs = Conn.execute("select DefineID from FS_NS_NewsClass where Classid='"& NoSqlHack(f_custclassid) &"'")
		if not obj_cust_rs.eof then
			GetCustClassID = obj_cust_rs("DefineID")
		Else
			GetCustClassID = ""
		End if
		obj_cust_rs.close:set obj_cust_rs =nothing
	End Function
	'得到新闻保存路径
	Public Function SaveNewsPath(f_num)
		SaveNewsPath = ""
		Select Case f_num
				Case 0
					SaveNewsPath = "/" & year(now)&"-"&month(now)&"-"&day(now)
				Case 1
					SaveNewsPath = "/" & year(now)&"/"&month(now)&"/"&day(now)
				Case 2
					SaveNewsPath = "/" & year(now)&"/"&month(now)&"-"&day(now)
				Case 3
					SaveNewsPath = "/" & year(now)&"-"&month(now)&"/"&day(now)
				Case 4
					SaveNewsPath = ""
				Case 5
					SaveNewsPath = "/" & year(now)&"/"&month(now)
				Case 6
					SaveNewsPath = "/" & year(now)&"/"&month(now)&day(now)
				Case 7
					SaveNewsPath = "/" & year(now)&month(now)&day(now)
		End Select		
	End Function
	'取得用户名
	Public Function GetUserName(f_strNumber)
		Dim RsGetUserName
		Set RsGetUserName = User_Conn.Execute("Select UserName From FS_ME_Users Where UserNumber = '"& NoSqlHack(f_strNumber) &"'")
		If  Not RsGetUserName.eof  Then 
			GetUserName = RsGetUserName("UserName")
		Else
			GetUserName = 0
		End If 
		set RsGetUserName = nothing
	End Function 
	
	'取得用户编号
	Public Function GetUserNumber(f_strName)
		Dim RsGetUserNumber
		Set RsGetUserNumber = User_Conn.Execute("Select usernumber From FS_ME_Users Where UserName = '"& f_strName &"'")
		If  Not RsGetUserNumber.eof  Then 
			GetUserNumber = RsGetUserNumber("usernumber")
		Else
			GetUserNumber = ""
		End If 
		set RsGetUserNumber = nothing
	End Function 
	'转移新闻到其他目录
	Public Function MoveNewsToClass(SourceNewsArray,ObjectClassID)
		Dim i,j,RsNewsObj,CopyNewsObj,SqlNews,FiledObj
		Dim NewsFileNames,TempNewsID,ConfigInfo
		ConfigInfo = Conn.Execute("Select FileExtName from FS_NewsClass")(0)
		for i = LBound(SourceNewsArray) to UBound(SourceNewsArray)
			Set RsNewsObj = Conn.Execute("Select * from FS_News where NewsID='" & NoSqlHack(SourceNewsArray(i)) & "'")
			SqlNews = "Select * from FS_News where 1=0"
			Set CopyNewsObj = Server.CreateObject(G_FS_RS)
			CopyNewsObj.Open SqlNews,Conn,1,3
			CopyNewsObj.AddNew
			For Each FiledObj In CopyNewsObj.Fields
				if LCase(FiledObj.name) <> "id" then
					if LCase(FiledObj.name) = "newsid" then
						TempNewsID = GetRandomID18()
						CopyNewsObj("newsid") = NoSqlHack(TempNewsID)
					elseif LCase(FiledObj.name) = "classid" then
						CopyNewsObj("classid") = NoSqlHack(ObjectClassID)
					else
						CopyNewsObj(FiledObj.name) = RsNewsObj(FiledObj.name)
					end if
				end if
			Next
			CopyNewsObj.UpDate
			'NewsFileNames=NewsFileName(ConfigArray(19),ObjectClassID,TempNewsID,CopyNewsObj("ID"))
			CopyNewsObj.Close
			'============================
			'取ID，生成文件名称，然后写回！
			Conn.Execute("Update FS_News Set FileName='"&NoSqlHack(NewsFileNames)&"' Where NewsID='"&NoSqlHack(TempNewsID)&"'")
			'============================
		next
		Set RsNewsObj = Nothing
		Set CopyNewsObj = Nothing
	End Function
	
	'删除传入的新闻ID的相关信息
	Public Function DeleteC(str_Id,f_number)
		'更新会员相关
		'更新头条
		'更新静态文件
		'更新权限表
		'更新不规则表
		'
		str_Id = NoSqlHack(replace(str_Id," ",""))
	End Function
	
	'统计投稿数量
	public Function newsStat(UserNumber,auditTF)
		Dim statRs
		if auditTF=1 then
			Set statRs=User_Conn.execute("select count(contID) from FS_ME_InfoContribution where usernumber='"&NoSqlHack(UserNumber)&"' and AuditTF=1")
		Else
			Set statRs=User_Conn.execute("select count(contID) from FS_ME_InfoContribution where usernumber='"&NoSqlHack(UserNumber)&"'")
		End if
		newsStat=statRs(0)
	End Function
	
	public Function navigation(ClassID)
		Dim classRs,naviStr,naviUnit
		Set classRs=Conn.execute("Select ClassID,ClassName,ParentID From FS_NS_NewsClass where ClassID='"&NoSqlHack(ClassID)&"'")
		if not classRs.eof then
			naviUnit="<a href=""News_Manage.asp?Classid="&classRs("ClassID")&"&SpecialEName="&NoSqlHack(request.QueryString("SpecialEName"))&""">"&classRs("ClassName")&"</a>"
			naviStr=navigation(classRs("ParentID"))&">>"&naviUnit&naviStr
		End if
		navigation=naviStr
	End function
    Public Function News_ChildNewsPower(TypeID) 
        Dim f_ChildNewsRs_1,ChildTypeListStr,f_TempStr,f_isUrlStr,lng_GetCount
		Set f_ChildNewsRs_1 = Conn.Execute("Select ClassID from FS_NS_NewsClass where ParentID='" & NoSqlHack(TypeID) & "' and ReycleTF=0 order by Orderid desc,id desc" )
		News_ChildNewsPower=False
		do while Not f_ChildNewsRs_1.Eof
		    If Get_SubPop_TF(f_ChildNewsRs_1("Classid"),"NS001","NS","news") Then
		        News_ChildNewsPower=TRUE	
		    else 
		         News_ChildNewsPower=News_ChildNewsPower(f_ChildNewsRs_1("ClassID")) or News_ChildNewsPower           			
			end if
			f_ChildNewsRs_1.MoveNext
		loop
		f_ChildNewsRs_1.Close
		Set f_ChildNewsRs_1 = Nothing
	End Function
End Class
%>





