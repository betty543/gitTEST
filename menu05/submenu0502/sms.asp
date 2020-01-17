<!-- #include virtual="/Include/Top_PopUp.asp" -->
<%
	cellphone = request("cellphone")
	SS_Login_EXTNO = SESSION("SS_Login_EXTNO")

	sYear = Year(now)
	sMonth = Month(now)
	sDay   = Day(now)
	if cint(sMonth) < 10 then
		sMonth = "0" & sMonth
	end if
	if cint(sDay) < 10 then
		sDay = "0" & sDay
	end if
	sHour = hour(Now)
	if cint(sHour) < 10 then
		sHour = "0" & sHour
	end if
	sMin = minute(Now)
	if cint(sMin) < 10 then
		sMin = "0" & sMin
	end if

	if cellphone <> "" then

		INCODE = SESSION("SS_LoginID")

		strSQL = "DELETE FROM	temp_conference where userid = '" & INCODE & "' and datagb = '2'"
		db.Execute(strSQL)


		strSQL = "INSERT INTO temp_conference ( addr_idx, userid, cellphone, gunphone, datagb)" &_
			" values (0,'"& INCODE	& "', " &_
					"'" & cellphone		& "','','2')"
		db.Execute(strSQL)
	end if	
%>
<script language="JavaScript" src="/Include/Js/SmsSend.js"></script>
<script>
<!--//
	function selectOK(){
		//alert(arg1+","+ arg2);
		parent.HddnPOPLayer();
	}

	function checkinput(){
		return true;
	}

	function fn_SendSMS(){

		smssingle.action = "sendprocess02.asp";
		smssingle.submit();

	}
//-->
</script>

<table width="600" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#ffffff">
	<tr><td height="10"></td></tr>
</table>

<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0">
<table width="600" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
	<tr><td bgcolor="#FDE6F3" class="FBlk TDCont">◈ <b>문자전송</b></td></tr>
</table>

<table width="600" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#ffffff">
	<tr><td height="10"></td></tr>
</table>
									<table width="600" cellspacing="1" cellpadding="0" bgcolor="BDCAE5" bordercolor="#CCCCCC" bordercolorlight="#CCCCCC" align="center">

									<form name="smssingle" method=post>
									<input type="hidden" name="group_name_cnt">
									<input type="hidden" name="AddChar_choice">

										 <tr>
											<td bgcolor="#EEF6FF" class="TDCont" align=center height=20 width="100">메세지</td>
											<td bgcolor="white" align=left valign=top COLSPAN=3>
											  <table width="100%" border=0 height=150>
											   <tr><Td width=34% valign=top >&nbsp;											   
											   
													<table border='0'cellpadding="0" cellspacing="0" width="137" height="125" align="center">
													<tr>
														<td width="137" height="130" background="/Images/pds_mess_bg.gif" title="<%=char_byte%>"><p align="center"> 

														<textarea name="to_message" rows="5" cols="16" style="font-size:12px; background-image:url('/Images/text_bg.gif'); border-width:0; border-style:none; cursor:hand;" onKeyUp="check_length();" ></textarea>
														</td>
													</tr>
													</table>
											    <!--<Td width=34% valign=top >&nbsp;<textarea name="to_message" rows="10" cols="16" onKeyUp="check_length();" style="font-size:12px; background-image:url('images/text_bg.gif'); border-width:0; border-style:none; cursor:hand;" ></textarea>-->
													   <table align="left" cellpadding="1" cellspacing="0" width="100%" border=0>
														 <tr>
													       <td width="100%" height="35" align="center">
															<div id="menu1" style="display:none">
															<TABLE   cellpadding="0" cellspacing="0" >
															<TR>
																<TD>&nbsp;<textarea name="to_message2" rows="3" cols="16" onKeyUp="check_length();"  style="font-size:12px;"></textarea>
																<textarea name="to_message3" rows="3" cols="16" onKeyUp="check_length();"  style="font-size:12px;"></textarea>
																<textarea name="to_message4" rows="3" cols="16" onKeyUp="check_length();"  style="font-size:12px;"></textarea>
																<textarea name="to_message5" rows="3" cols="16" onKeyUp="check_length();"  style="font-size:12px;"></textarea>
																<textarea name="to_message6" rows="3" cols="16" onKeyUp="check_length();"  style="font-size:12px;"></textarea>
																<textarea name="to_message7" rows="3" cols="16" onKeyUp="check_length();"  style="font-size:12px;"></textarea>
																<textarea name="to_message8" rows="3" cols="16" onKeyUp="check_length();"  style="font-size:12px;"></textarea></TD>
															</TR>
															</table>
															</div>
															<div id="menu2">
															</div>
															<TABLE  border=0 cellpadding="0" cellspacing="0">
															<tr>
															   <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="msglen" maxlength="5" size="5" value="0" style="text-align:right; font-family:굴림,serif; font-size:9pt; color:rgb(51,102,255); border-width:1px; border-color:rgb(153,153,153); border-style:solid;" readonly> bytes
															   </td>
															</tr>
															<!--<TR>
																<TD><font color="#FF0000">&nbsp;[80byte로 제한]</font></TD>
															</TR>-->
															</TABLE>

													       </td>
													     </tr>
												      </table>		
											     </td>
											     <td bgcolor="white" width=33% align='center' >
															<table cellpadding="1" cellspacing="1" bordercolordark="white" bordercolorlight="#CCCCCC" align="left" bgcolor="#CCCCCC" bordercolor="#CCCCCC" width="150">
																<tr> <td width="100%" height="10" colspan=20  bgcolor="BDCAE5" align="center">특수문자</td>
																</tr>
																<tr> 
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('☆');return false;"><span style="font-size:9pt;">☆</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('○');return false;"><span style="font-size:9pt;">○</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('□');return false;"><span style="font-size:9pt;">□</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('◎');return false;"><span style="font-size:9pt;">◎</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('★');return false;"><span style="font-size:9pt;">★</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('●');return false;"><span style="font-size:9pt;">●</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('■');return false;"><span style="font-size:9pt;">■</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('⊙');return false;"><span style="font-size:9pt;">⊙</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('☏');return false;"><span style="font-size:9pt;">☏</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('☎');return false;"><span style="font-size:9pt;">☎</span></a></td>
																</tr>
																<tr> 
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('◈');return false;"><span style="font-size:9pt;">◈</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('▣');return false;"><span style="font-size:9pt;">▣</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('◐');return false;"><span style="font-size:9pt;">◐</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('◑');return false;"><span style="font-size:9pt;">◑</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('☜');return false;"><span style="font-size:9pt;">☜</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('☞');return false;"><span style="font-size:9pt;">☞</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('◀');return false;"><span style="font-size:9pt;">◀</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('▶');return false;"><span style="font-size:9pt;">▶</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('▲');return false;"><span style="font-size:9pt;">▲</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('▼');return false;"><span style="font-size:9pt;">▼</span></a></td>
																</tr>
																<tr> 
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('♠');return false;"><span style="font-size:9pt;">♠</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('♣');return false;"><span style="font-size:9pt;">♣</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('♥');return false;"><span style="font-size:9pt;">♥</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('◆');return false;"><span style="font-size:9pt;">◆</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('◁');return false;"><span style="font-size:9pt;">◁</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('▷');return false;"><span style="font-size:9pt;">▷</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('△');return false;"><span style="font-size:9pt;">△</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('▽');return false;"><span style="font-size:9pt;">▽</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('♤');return false;"><span style="font-size:9pt;">♤</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('♧');return false;"><span style="font-size:9pt;">♧</span></a></td>
																</tr>
																<tr> 
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('♡');return false;"><span style="font-size:9pt;">♡</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('◇');return false;"><span style="font-size:9pt;">◇</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('※');return false;"><span style="font-size:9pt;">※</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('♨');return false;"><span style="font-size:9pt;">♨</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('♪');return false;"><span style="font-size:9pt;">♪</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('♭');return false;"><span style="font-size:9pt;">♭</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('♩');return false;"><span style="font-size:9pt;">♩</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('♬');return false;"><span style="font-size:9pt;">♬</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('㉿');return false;"><span style="font-size:9pt;">㉿</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('㈜');return false;"><span style="font-size:9pt;">㈜</span></a></td>
																</tr>
																<tr> 
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('①');return false;"><span style="font-size:9pt;">①</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('②');return false;"><span style="font-size:9pt;">②</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('③');return false;"><span style="font-size:9pt;">③</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('④');return false;"><span style="font-size:9pt;">④</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('⑤');return false;"><span style="font-size:9pt;">⑤</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('⑥');return false;"><span style="font-size:9pt;">⑥</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('⑦');return false;"><span style="font-size:9pt;">⑦</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('⑧');return false;"><span style="font-size:9pt;">⑧</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('⑨');return false;"><span style="font-size:9pt;">⑨</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('⑩');return false;"><span style="font-size:9pt;">⑩</span></a></td>
																</tr>
																<tr> 
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓐ');return false;"><span style="font-size:9pt;">ⓐ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓑ');return false;"><span style="font-size:9pt;">ⓑ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓒ');return false;"><span style="font-size:9pt;">ⓒ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓓ');return false;"><span style="font-size:9pt;">ⓓ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓔ');return false;"><span style="font-size:9pt;">ⓔ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓕ');return false;"><span style="font-size:9pt;">ⓕ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓖ');return false;"><span style="font-size:9pt;">ⓖ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓗ');return false;"><span style="font-size:9pt;">ⓗ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓘ');return false;"><span style="font-size:9pt;">ⓘ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓙ');return false;"><span style="font-size:9pt;">ⓙ</span></a></td>
																</tr>
																<tr> 
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓚ');return false;"><span style="font-size:9pt;">ⓚ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓛ');return false;"><span style="font-size:9pt;">ⓛ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓜ');return false;"><span style="font-size:9pt;">ⓜ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓝ');return false;"><span style="font-size:9pt;">ⓝ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓞ');return false;"><span style="font-size:9pt;">ⓞ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓟ');return false;"><span style="font-size:9pt;">ⓟ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓠ');return false;"><span style="font-size:9pt;">ⓠ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓡ');return false;"><span style="font-size:9pt;">ⓡ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓢ');return false;"><span style="font-size:9pt;">ⓢ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓣ');return false;"><span style="font-size:9pt;">ⓣ</span></a></td>
																</tr>
																<tr> 
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓤ');return false;"><span style="font-size:9pt;">ⓤ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓥ');return false;"><span style="font-size:9pt;">ⓥ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓦ');return false;"><span style="font-size:9pt;">ⓦ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓧ');return false;"><span style="font-size:9pt;">ⓧ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓨ');return false;"><span style="font-size:9pt;">ⓨ</span></a></td>
																  <td width="10" height="10" bgcolor="white"><a  href="javascript:;" onClick="AddChar('ⓩ');return false;"><span style="font-size:9pt;">ⓩ</span></a></td>
																  <td height="10" colspan="4" bgcolor="white"><a  href="javascript:;" onClick="AddChar('[#NM#]');return false;"><span style="font-size:9pt;"></span></a></td>
																</tr>

															  </table>		
											      </td>
												  <td><iframe src="sub_sms_list.asp" name="DBFrame" width="100%" height="100%" frameborder=0 marginheight=0 marginwidth=0 scrolling="no"></iframe>
												  </td>
												 </TR>
											   </TABLE>
										   </TD>
										 </tr>
										 <tr>
											<td bgcolor="#EEF6FF" class="TDCont" align=center height=30>보내는사람</td>
											<td bgcolor="white" align=left>
											  <table width="100%" border=0>
											   <tr>
											    <Td width=34%><input  name="from_num" maxlength="12" size="12" value="<%=SS_Login_EXTNO%>"></td>
												<td width=66% bgcolor="white" align=left><!input type="button" class="button2"" value="설정" onclick="document.location.href = '/EditForm.html';"></td>
											   </TR>
											  </TABLE>
											</TD>
											<td bgcolor="#EEF6FF" class="TDCont" align=center height=30 width='100'>받는사람</td>
											<td bgcolor="white" align=left>&nbsp;<input  name="to_num" maxlength="12" size="12">&nbsp;<!--<input type="submit" name="insert" value="번호추가" class="button2">-->
											<img src="/Images/Btn/BtnCellAdd.gif" title="번호추가" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_add();">
											
											</TD>
										 </tr>

										 <tr>
											<td bgcolor="#EEF6FF" class="TDCont" align=center height=30>예약구분</td>
											<td bgcolor="white" align=left>&nbsp;<input type="RADIO" value="0" name="sendType" checked onClick="check_gubun(1);" class="none"> 즉시전송 <input type="RADIO" value="6" name="sendType"  onClick="check_gubun(2);" class="none"> 예약 <!--<input type="checkbox" name="sendtype2"  onClick="check_gubun(3);" > 분할전송--></td>
											<td bgcolor="#EEF6FF" class="TDCont" align=center height=30>예약일시</td>
											<td bgcolor="white" align=left>
												&nbsp;<input type=TEXT name="yy" value="<%=syear%>" maxlength="4" size="4" disabled=true onkeypress="if (event.keyCode < 26 || event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;">년 <input type=TEXT name="mm" value="<%=smonth%>" maxlength="2" size="2" disabled=true onkeypress="if (event.keyCode < 26 || event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;">월 <input type=TEXT name="dd" value="<%=sday%>" size="2" maxlength="2" disabled=true onkeypress="if (event.keyCode < 26 || event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;">일	
												&nbsp;<input type=TEXT name="h" value="<%=sHour%>" size="2" maxlength="2" disabled=true onkeypress="if (event.keyCode < 26 || event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;">시 <input type=TEXT name="m" value="<%=sMin%>" size="2" maxlength="2" disabled=true onkeypress="if (event.keyCode < 26 || event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;">분	
											</td>
										 </tr>
										 <tr>
											<td bgcolor="#EEF6FF" class="TDCont" align=center height=30>전송</td>
											<td colspan=3 bgcolor="white" align=left>&nbsp;<input type="button" value="  전송  " class="button2" onclick="fn_SendSMS();"></a> <input type="button" value="  취소  " class="button2" onclick="javascript:parent.HddnPOPLayer();" onFocus="this.blur();"></td>
										 </tr>

									 </form>
								</table>



<script>
function fn_inup(form)
{
	form.JobGb.value = "U";
	form.submit();
}
function fn_SetResult(arg){
	if (arg =='1')
	{
		if (document.inUpFrm.Select1.value =='A' || document.inUpFrm.Select1.value =='C')
		{
			document.inUpFrm.Select3.disabled = true;
			document.inUpFrm.Select2.disabled = true;
			document.inUpFrm.sSelect2.value ="";
			document.inUpFrm.sSelect3.value ="";
		}
		if (document.inUpFrm.Select1.value =='B')
		{
			document.inUpFrm.Select3.disabled = true;
			document.inUpFrm.Select2.disabled = false;
			document.inUpFrm.sSelect2.value =document.inUpFrm.Select2.value;
			document.inUpFrm.sSelect3.value ="";
		}
		if (document.inUpFrm.Select1.value =='D')
		{
			document.inUpFrm.Select3.disabled = false;
			document.inUpFrm.Select2.disabled = true;
			document.inUpFrm.sSelect3.value =document.inUpFrm.Select3.value;
			document.inUpFrm.sSelect2.value ="";
		}
	}
	if (arg =='2')
	{
		document.inUpFrm.sSelect2.value =document.inUpFrm.Select2.value;
	}
	if (arg =='3')
	{
		document.inUpFrm.sSelect3.value =document.inUpFrm.Select3.value;
	}
}
</script>

<script>document.smssingle.to_message.focus();
</script>


<script>
	function fn_add(){

		if (document.all.to_num.value.length < 10)
		{
			alert('추가될 휴대폰번호를 정확히 입력하십시오!');
			return false;
		}

		DBFrame.location="sub_sms_list.asp?idx=0&IsType=INS&CellPhone="+document.all.to_num.value;
	}

</script>
<!-- #include virtual="/Include/Bottom_PopUp.asp" -->