// JavaScript Document
function ReplaceDot(str)
{
	var Obj=document.getElementById(str);
	var oldValue=Obj.value;
	while(oldValue.indexOf("，")!=-1)
	{
		Obj.value=oldValue.replace("，",",");
		oldValue=Obj.value;
	}
}
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
			if(Obj.value.indexOf(".")>=0)
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