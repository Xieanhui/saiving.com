<%'Copyright (c) 2006 Foosun Inc. 
Class Cls_House 
	Private m_sysID,m_PubRights,m_UserAudit,m_AutoDelete,m_LinkMan,m_Contact,m_LinkShow,m_Months


	'调用类的初始值
	Private Sub Class_Initialize() 
 	End Sub
	'释放初始值 
	Private Sub Class_Terminate()  
 	End Sub
	
	'得到子类类别分页  
	Public Function GetChildClassList(f_classid)
		
	End Function
	
	Public Function GetSysParamDir()
		GetSysParamDir = "/House"
	End Function
		
	Public Function GetSysParam()
			Dim f_Obj_sysparm
			GetSysParam = False
			Set f_Obj_sysparm=server.CreateObject(G_FS_RS)
			f_Obj_sysparm.Open "select top 1 SysID,PubRights,UserAudit,AutoDelete,Months,LinkMan,Contact,LinkShow from FS_HS_SysPara",Conn,1,1
			if  not (f_Obj_sysparm.eof or f_Obj_sysparm.bof) then
				m_PubRights = f_Obj_sysparm("PubRights")
				m_UserAudit = f_Obj_sysparm("UserAudit")
				m_AutoDelete = f_Obj_sysparm("AutoDelete")
				m_Months = f_Obj_sysparm("Months")
				m_LinkMan = f_Obj_sysparm("LinkMan")
				m_Contact = f_Obj_sysparm("Contact")
				m_LinkShow = f_Obj_sysparm("LinkShow")
				GetSysParam = True
			Else
				''默认的参考设置
				m_PubRights = 0
				m_UserAudit = 0
				m_AutoDelete = 1
				m_Months = 3
				m_LinkMan = "站长"
				m_Contact = "站长电话"
				m_LinkShow = 1
				GetSysParam = false
			End if
	End Function
	'赋值
	Public Property Get sysID()				'参数ID  
		sysID = m_sysID
	End Property 
	Public Property Get PubRights()				 
		PubRights = m_PubRights
	End Property 
	Public Property Get UserAudit() 
		UserAudit = m_UserAudit
	End Property 
	Public Property Get AutoDelete()				 
		AutoDelete = m_AutoDelete
	End Property 
	Public Property Get Months()				 
		Months = m_Months
	End Property 
	Public Property Get LinkMan() 
		LinkMan = m_LinkMan
	End Property 
	Public Property Get Contact()				 
		Contact = m_Contact
	End Property 
	Public Property Get LinkShow() 
		LinkShow = m_LinkShow
	End Property 

''清理过期房源信息。
	Public Function ClearGQInfo(TableName,IDList)
		select case TableName
			case "FS_HS_Quotation"
			if instr(IDList,"datediff")>0 then 
				Conn.execute("Delete from "&NoSqlHack(TableName)&" "&NoSqlHack(IDList))
			else		
				Conn.execute("Delete from "&NoSqlHack(TableName)&" where ID in ("&FormatIntArr(IDList)&")")
			end if	
			case "FS_HS_Second"
			if instr(IDList,"datediff")>0 then 
				Conn.execute("Delete from "&NoSqlHack(TableName)&" "&IDList)
			else		
				Conn.execute("Delete from "&NoSqlHack(TableName)&" where SID in ("&FormatIntArr(IDList)&")")
			end if	
			case "FS_HS_Tenancy"
			if instr(IDList,"datediff")>0 then 
				Conn.execute("Delete from "&NoSqlHack(TableName)&" "&IDList)
			else		
				Conn.execute("Delete from "&NoSqlHack(TableName)&" where TID in ("&FormatIntArr(IDList)&")")
			end if
		end select	
		if Err.Number>0 then 
			ClearGQInfo = False
		else
			ClearGQInfo = True
		end if
	End Function
End Class
%>





