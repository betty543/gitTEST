<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
<table border="0" width="1200" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
		
			<form method="post" name="inUpFrm" style="margin:0">
			<input type="hidden" name="QueryYN" value="">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr>
					<td bgcolor="#EFEFEF" class="TDCont">姥歳</td>
					<td bgcolor="#FFFFFF" colspan=3>
						<input type="radio" name="sUSEYN" value="Y" class="none" onClick="fn_YES();" checked > 榎析森鉦鯉系
						<input type="radio" name="sUSEYN" value="N" class="none" onClick="fn_YES();" > 遭楳鯉系
					</td>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont">繕噺奄娃 :</td>
			        <td  bgcolor="#FFFFFF" colspan=3>
			        	<input value="" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">	
			        </td>
			        <td colspan='2' rowspan="2" bgcolor="#FFFFFF" align="center">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="goSearch(document.inUpFrm);">
			        	<%IF SS_Login_Secgroup="A" Or SS_Login_Secgroup="B" THEN%><br><br><img src="/Images/Btn/BtnExcel.gif" style="cursor:hand;" onClick="fn_Xls();"><%END IF%>
			        </td>
				</tr>
			    <tr>

			        <td width="110" bgcolor="#EFEFEF" class="TDCont">紫闇腰硲 :</td>
			        <td bgcolor="#FFFFFF">
			        	<input value="" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);"></td>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont">社紗 :</td>
					<td bgcolor="#FFFFFF" nowrap><input type="text" name="sCTIID" value="" maxlength="15" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"> <img src="/Images/Comm/IconTip.gif" style="cursor:hand;" align="absmiddle" onClick="pCateSelect('1');" >
					</td>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont">杷税切誤 :</td>
			        <td bgcolor="#FFFFFF">
			        	<input value="" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);"></td>

			        <td width="110" bgcolor="#EFEFEF" class="TDCont">眼雁呪紫淫 :</td>
			        <td bgcolor="#FFFFFF">

			        	<select name="whereCD1" size="1" class="ComboFFFCE7">
							<option value="">識澱</option>
						</select>
					</td>



			    </tr>
			</table>
			</form>
		</td>
	</tr>
</table>

<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>

<table width="1200" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td rowspan=2>No</td>
		<td rowspan=2>紫闇腰硲</td>
		<td rowspan=2>森鉦舛左</td>
		<td rowspan=2>紫闇誤</td>
		<td rowspan=2>歎採弘<br>政巷</td>
		<td rowspan=2>窒坦析切</td>
		<td colspan=3>眼雁呪紫淫</td>
		<td colspan=3>尻喰坦政巷</td>
		<td rowspan=2>淫軒</td>
	</tr>
	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td >社紗</td>
		<td >域厭</td>
		<td >失誤</td>
		<td >杷税切</td>
		<td >杷背切</td>
		<td >走番淫</td>
	</tr>
	<tr><td colspan="17" height="1" bgcolor="#FFFFFF"></td></tr>


	<tr id="cTR1" style="cursor:hand;" bgcolor="#ffffff" >

			<td align="center">1</td>
			<td align="center">0000000000</td>
			<td align="center">杷税切(神板2獣)</td>
			<td align="center">浦奄紫壱-賑楳</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">2009-01-01</td>
			<td align="center">しし浦舘賠佐企</td>
			<td align="center">しし</td>
			<td align="center">ししし</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="呪舛" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="肢薦" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
	</tr>
	<tr id="cTR1" style="cursor:hand;" bgcolor="#ffffff" >

			<td align="center">2</td>
			<td align="center">0000000000</td>
			<td align="center">杷背切(神穿8獣)</td>
			<td align="center">照穿紫壱-託勲-嘘搭紫壱,亀稽嘘搭狛酔鋼</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">2009-01-11</td>
			<td align="center">しし浦舘賠佐企</td>
			<td align="center">しし</td>
			<td align="center">ししし</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_03.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="呪舛" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="肢薦" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
	</tr>
	<tr id="cTR1" style="cursor:hand;" bgcolor="#ffffff" >

			<td align="center">3</td>
			<td align="center">0000000000</td>
			<td align="center">杷背切(神穿10獣)</td>
			<td align="center">照穿紫壱-託勲-嘘搭紫壱,亀稽嘘搭狛酔鋼</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">2009-02-20</td>
			<td align="center">しし浦舘賠佐企</td>
			<td align="center">しし</td>
			<td align="center">ししし</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="呪舛" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="肢薦" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
	</tr>


</table>

<table border="0" cellpadding="0" cellspacing="0" width="1200" align="center">
	<tr><td height="5"></td></tr>
	<tr><td height="1" bgcolor="#D6D6D6"></td></tr>
	<tr height="22" bgcolor="#EEF6FF"><td align="center">1  2  3  4  5  6</td></tr>
	<tr><td height="1" bgcolor="#D6D6D6"></td></tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="1200" align="center">
	<tr><td height="5"></td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width="1200" align="center">
	<tr><td><iframe frameborder=0 marginheight=0 marginwidth=0 width="100%" height="0" scrolling="no" name="AsInfo1fr"></iframe></td></tr>
</table>


<script>
	function pCateSelect(cn){
		Cate1 = 'A' ; //eval("inUpFrm.ACLASS"+cn).value;
		CUSTNO = '0000000000'; //parent.MemInfoFrame.frmSearch.CUSTNO.value;

		if (cn == '1')
		{//PSEQ1
			Relation = '0';//eval("inUpFrm.RELATION"+cn).value;
			RelationSeq = '0';//eval("inUpFrm.PSEQ"+cn).value;
			GoURL ="/Include/PopUp/PCategory.asp?Cate1=" +Cate1+ "&FM=" +cn+ "&CUSTNO=" +CUSTNO+"&Relation="+Relation+"&RelationSeq="+RelationSeq;
		}
		else
		{
			GoURL ="/Include/PopUp/PCategory.asp?Cate1=" +Cate1+ "&FM=" +cn+ "&CUSTNO=" +CUSTNO;
		}

		ShowPOPLayer(GoURL,'720','380');
	}
</script>
<!-- #include virtual="/Include/Bottom.asp" -->