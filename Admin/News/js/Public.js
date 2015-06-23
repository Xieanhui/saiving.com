function OpenWindow(Url,Width,Height,WindowObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
	return ReturnStr;
}
//Open Modal Window
function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
	if (ReturnStr!='') SetObj.value=ReturnStr;
	return ReturnStr;
}
function CheckNumber(Obj,DescriptionStr)
{
	if (Obj.value!='' && (isNaN(Obj.value) || Obj.value<0))
	{
		alert(DescriptionStr+"应填有效数字！");
		Obj.value="";
		Obj.focus();
	}
}
function Dokesite(KeyWords)
{
	if (KeyWords!='')
	{
		if (document.NewsForm.KeywordText.value.search(KeyWords)==-1)
		{
			if (document.NewsForm.KeyWords.value=='') document.NewsForm.KeyWords.value=KeyWords;
			else document.NewsForm.KeyWords.value=document.NewsForm.KeyWords.value+','+KeyWords;
			if (document.NewsForm.KeywordText.value=='') document.NewsForm.KeywordText.value=KeyWords;
			else document.NewsForm.KeywordText.value=document.NewsForm.KeywordText.value+','+KeyWords;
		}
	}
	if (KeyWords == 'Clean')
	{
		document.NewsForm.KeyWords.value = '';
		document.NewsForm.KeywordText.value = '';
	}
	return;
}
function Dokesite_s(KeyWords)
{
	if (KeyWords!='')
	{
		if (document.form_m.KeywordText.value.search(KeyWords)==-1)
		{
			if (document.form_m.KeyWords.value=='') document.form_m.KeyWords.value=KeyWords;
			else document.form_m.KeyWords.value=document.form_m.KeyWords.value+','+KeyWords;
			if (document.form_m.KeywordText.value=='') document.form_m.KeywordText.value=KeyWords;
			else document.form_m.KeywordText.value=document.form_m.KeywordText.value+','+KeyWords;
		}
	}
	if (KeyWords == 'Clean')
	{
		document.form_m.KeyWords.value = '';
		document.form_m.KeywordText.value = '';
	}
	return;
}
function dospclear()
	{
	document.NewsForm.SpecialID.value = '';
	document.NewsForm.SpecialID_EName.value = '';
	}
function Doauthsite(Author)
{
	var TempArray,TempStr;
	TempArray=Author.split("***");
	if (TempArray[0] != '')
	{
		if (document.NewsForm.AuthorText.value.indexOf(TempArray[0])<0)
		{
			if (typeof(TempArray[1])=='undefined') TempStr=TempArray[0];
			else TempStr='<a href='+TempArray[1].replace(/[\"\']/,'')+'>'+TempArray[0]+'</a>';
			if (document.NewsForm.AuthorText.value=='') 	document.NewsForm.AuthorText.value=TempArray[0];
			else document.NewsForm.AuthorText.value = document.NewsForm.AuthorText.value + ',' + TempArray[0];
			if (document.NewsForm.Author.value=='') 	document.NewsForm.Author.value=TempArray[0];
			else document.NewsForm.Author.value = document.NewsForm.Author.value + ',' + TempArray[0];
		}
	}
	if ((TempArray[0] == '')&&(TempArray[1] == 'Clean'))
	{
		document.NewsForm.Author.value = '';
		document.NewsForm.AuthorText.value = '';
	}
	return;
}
function Dosusite(Source)
{
	var TempArray,TempStr;
	TempArray=Source.split("***");
	if (TempArray[0] != '')
	{
		if (document.NewsForm.TxtSourceText.value.indexOf(TempArray[0])<0)
		{
			if (typeof(TempArray[1])=='undefined') TempStr=TempArray[0];
			else TempStr='<a href='+TempArray[1].replace(/[\"\']/,'')+'>'+TempArray[0]+'</a>';
			if (document.NewsForm.TxtSourceText.value=='') 	document.NewsForm.TxtSourceText.value=TempArray[0];
			else document.NewsForm.TxtSourceText.value = document.NewsForm.TxtSourceText.value + ',' + TempArray[0];
			if (document.NewsForm.Source.value=='') 	document.NewsForm.Source.value=TempArray[0];
			else document.NewsForm.Source.value = document.NewsForm.Source.value + ',' + TempArray[0];
		}
	}
	if ((TempArray[0] == '')&&(TempArray[1] == 'Clean'))
	{
		document.NewsForm.Source.value = '';
		document.NewsForm.TxtSourceText.value = '';
	}
	return;
}