<!-- #include virtual="/Include/Top.asp" -->
<table width="1200" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr valign="top">
		<td width="790" height="750">
			<!--<iframe src="User_List.html" name="ListFrame" width="100%" height="100%" frameborder=0 marginheight=0 marginwidth=0 scrolling="no"></iframe>-->

	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk">◈ <b>사용자 리스트</b></td></tr>
        	</table>
        	<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr height="20" bgcolor="#EEF6FF" align="center">
        			<td>NO</td>
        			<td>아이디</td>
        			<td>소속</td>
        			<td>계급</td>
        			<td>군번</td>
        			<td>성명</td>
        			<td>비밀번호</td>
        			<td>운용업무</td>
        			<td>권한</td>
        			<td>사용여부</td>
        		</tr>
        		<tr><td colspan="10" height="1" bgcolor="#FFFFFF"></td></tr>

				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" >
					<td align="center">1</td>
					<td align="center">agent01</td>
					<td align="center">소속1-소속2-소속3</td>
					<td align="center"> </td>
					<td align="center"> </td>
					<td align="center">손민경</td>
					<td align="center">1235</td>
					<td align="center">시스템관리자</td>
					<td align="center">A권한</td>
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
				</tr>
				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" >
					<td align="center">2</td>
					<td align="center">agent02</td>
					<td align="center">소속1-소속2-소속3</td>
					<td align="center"> </td>
					<td align="center"> </td>
					<td align="center">김아무개</td>
					<td align="center">1235</td>
					<td align="center">상담원</td>
					<td align="center">B권한</td>
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
				</tr>

        	</table>       	


			<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
				<tr><td height="5"></td></tr>
				<tr><td height="1" bgcolor="#D6D6D6"></td></tr>
				<tr height="22" bgcolor="#EEF6FF">
					<td align="left">1  2  3  4 </td>
				</tr>
				<tr><td height="1" bgcolor="#D6D6D6"></td></tr>
			</table>
        	
			<table border="0" cellspacing="0" width="100%" align="center">
				<tr height="30"><td align="left" class="TDL10px" height="25"></td>
					<td align="right"><img src="/Images/Btn/BtnUserAdd.gif" style="cursor:hand;" align="absmiddle" onClick="parent.DetailFrame.location.href='User_Detail.asp?guboon=INS';"></td>
				</tr>
			</table>


		</td>
		<td width="10"></td>
		<td width="400">

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>


        	<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk">◈ <b>사용자 정보</b></td></tr>
        	</table>

			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">


<form name="inUpFrm" method="post">


				<tr>
					<td nowrap width="100" bgcolor="#FFEEF9" class="TDCont">아이디</td>
					<td bgcolor="#FFFFFF"><input type="text" name="sUSERID" value="agent01" readonly maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">소속</td>
					<td bgcolor="#FFFFFF">
						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A등급' selected>소속선택</option>
						</select>					
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">계급</td>
					<td bgcolor="#FFFFFF">
						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A등급' selected>계급선택</option>
						</select>					
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">군번</td>
					<td bgcolor="#FFFFFF"><input type="text" name="sUSERNAME" value="" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">성명</td>
					<td bgcolor="#FFFFFF"><input type="text" name="sUSERNAME" value="손민경" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">비밀번호</td>
					<td bgcolor="#FFFFFF"><input type="input" name="sPASSWORD" value="smk1414" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">운용업무</td>
					<td bgcolor="#FFFFFF">
						<select name="sSECGROUP" size="1" class="ComboFFFCE7">
							<Option value ='시스템관리자' selected>운용업무선택</option>
						</select>					
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">권한</td>
					<td bgcolor="#FFFFFF">
						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A등급' selected>권한선택</option>
						</select>					
					</td>
				</tr>				
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">CTI 사용여부</td>
					<td bgcolor="#FFFFFF">
						<input type="radio" name="sCTIYN" value="Y" class="none" checked> 사용
						<input type="radio" name="sCTIYN" value="N" class="none" > 미사용
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">CTI 아이디</td>
					<td bgcolor="#FFFFFF"><input type="text" name="sCTIID" value="cti01" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">CTI 비밀번호</td>
					<td bgcolor="#FFFFFF"><input type="text" name="sCTIPASSWORD" value="0000" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">내선번호</td>
					<td bgcolor="#FFFFFF"><input type="text" name="sEXTNO" value="4567" maxlength="3" size="3" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">사용여부</td>
					<td bgcolor="#FFFFFF">
						<input type="radio" name="sUSEYN" value="Y" class="none" onClick="fn_YES();" checked > 사용
						<input type="radio" name="sUSEYN" value="N" class="none" onClick="fn_YES();" > 미사용
					</td>
				</tr>				
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">입사일자</td>
					<td bgcolor="#FFFFFF"><input name="sIPDATE" value="2009-01-01" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);"></td>
				</tr>
				<tr>
					<td bgcolor="#FFEEF9" class="TDCont">퇴사일자</td>
					<td bgcolor="#FFFFFF"><input name="sOUTDATE" value="" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);"></td>
				</tr>

</form>

			</table>
			<table border="0" cellspacing="0" width="100%" align="center">
				<tr height="30">
					<td align="right">
						<img src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_inup(document.inUpFrm);">
						<img src="/Images/Btn/BtnReset.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:document.inUpFrm.reset();">
						<img src="/Images/Btn/BtnDel.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:self.close();">
					</td>
				</tr>
			</table>			


		</td>
	</tr>
</table>
<!-- #include virtual="/Include/Bottom.asp" -->