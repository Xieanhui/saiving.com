// JavaScript Document
//将中文逗号转换为英文逗号 str:待替换的字符
function ReplaceDot(str)
{
	var Obj=document.getElementById(str);
	var oldValue=Obj.value;
	while(oldValue.indexOf("，")!=-1)//寻找每一个中文逗号，并替换
	{
		Obj.value=oldValue.replace("，",",");
		oldValue=Obj.value;
	}
}

//检查字符长度，Str:待检查字符；FS_Alert:错误信息显示容器；Len:限制长度
function CheckContentLen(Str,FS_Alert,Len)
{
	var Obj=document.getElementById(Str);
	if(Obj.value.length>Len)
	{
		document.getElementById(FS_Alert).innerHTML="<font color='F43631'>长度请不要超过"+Len+"</font>";
		return false;
	}
	return true;
}
//检查字符是否为数字，Str:待检查字符；FS_Alert:错误信息显示容器；isInteger:是否为整数
function isNumber(Str,FS_Alert,Msg,isInteger)
{
	var Obj=document.getElementById(Str)
	if(Obj.value=='')
	{
		document.getElementById(FS_Alert).innerHTML="<font color='F43631'>该处不能为空</font>";
		return false;
	}
	else if(isNaN(Obj.value)||Obj.value<0)
	{
		document.getElementById(FS_Alert).innerHTML="<font color='F43631'>"+Msg+"</font>";
		return false;
	}
	else if(!isNaN(Obj.value)&&Obj.value>=0)
	{
		if(isInteger)
		{
			if(Obj.value.indexOf(".")>=0)//是否为整数
			{
				document.getElementById(FS_Alert).innerHTML="<font color='F43631'>请使用整数</font>";
				return false;
			}else
			{
				document.getElementById(FS_Alert).innerHTML="";
				return true;
			}
		}
		else
		{
			document.getElementById(FS_Alert).innerHTML="";
			return true;
		}
	}
}
//检查字符是否为空，Str:待检查字符；FS_Alert:错误信息显示容器；
function isEmpty(Str,FS_Alert)
{
	var Obj=document.getElementById(Str);
	if(Obj.value=="")
	{
		document.getElementById(FS_Alert).innerHTML="<font color='F43631'>该处不能为空</font>";
		return false;
	}else
	{
		document.getElementById(FS_Alert).innerHTML="";
		return true;
	}
}
//检查字符是否为中文，Str:待检查字符；FS_Alert:错误信息显示容器；
function isChinese(Str,FS_Alert)
{ 
	var Number = "0123456789.,abcdefghijklmnopqrstuvwxyz-\/ABCDEFGHIJKLMNOPQRSTUVWXYZ`~!@#$%^&*()_";
	var Obj=document.getElementById(Str);
	for (i = 0; i < Obj.value.length;i++)
	{   
		var c = Obj.value.charAt(i);
		if (Number.indexOf(c) == -1) 
		{
			document.getElementById(FS_Alert).innerHTML="<font color='F43631'>请不要使用中文字符</font>";
			return false;
		}
	}
	document.getElementById(FS_Alert).innerHTML="";
	return true
}