// JavaScript Document
function ReplaceDot(Obj)
{
	var oldValue=Obj.value;
	while(oldValue.indexOf("，")!=-1)
	{
		Obj.value=oldValue.replace("，",",");
		oldValue=Obj.value;
	}
}
function isNumber(Obj,FS_Alert,Msg)
{
	if(Obj.value=='')
	{
		document.getElementById(FS_Alert).innerHTML="<font color='F43631'>该处不能为空、</font>";
		Obj.focus();
		return false;
	}
	else if(isNaN(Obj.value)||Obj.value<0)
	{
		document.getElementById(FS_Alert).innerHTML="<font color='F43631'>"+Msg+"</font>";
		Obj.focus();
		return false;
	}
	else if(!isNaN(Obj.value)&&Obj.value>=0)
	{
		document.getElementById(FS_Alert).innerHTML="";
		return true;
	}
}
function MySubmit(Obj,flag)
{
	if(flag)
	document.getElementById(Obj).submit();
}
function CheckContentLen(Obj,FS_Alert,Len)
{
	if(Obj.value.length>Len)
	{
		document.getElementById("FS_Alert").innerHTML="<font color='F43631'>长度请不要超过2000</font>";
		return false;
	}

}
function isChinese(Obj,FS_Alert)
{ 
	var Number = "0123456789.abcdefghijklmnopqrstuvwxyz-\/ABCDEFGHIJKLMNOPQRSTUVWXYZ`~!@#$%^&*()_";
	for (i = 0; i < Obj.value.length;i++)
	{   
		var c = Obj.value.charAt(i);
		if (Number.indexOf(c) == -1) 
		{
			document.getElementById(FS_Alert).innerHTML="<font color='F43631'>请不要使用中文字符</font>";
			Obj.focus()
			return false;
		}
	}
	document.getElementById(FS_Alert).innerHTML="";
	return true
}
