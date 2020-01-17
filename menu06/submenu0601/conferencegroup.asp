
<!-- #include virtual="/Include/Top.asp" -->

<%


%>

<table width="940" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr valign="top" height="450">

		<td width="340">

			<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk">◈ <b>그룹</b></td><td colspan="5" align="right" height=28><img src="/Images/Btn/BtnAdd.gif" title="그룹추가" style="cursor:hand;" align="absmiddle" onClick="ShowPOPLayer('conferencegroup_detail.asp','500','210');"></td></tr>
        	</table>
			<DIV style="OVERFLOW-Y:auto; OVERFLOW-X:auto; MARGIN: 0px 0px 0px 0px; 500; HEIGHT:420;">
        	<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr height="20" bgcolor="#EEF6FF" align="center">
        			<td>NO</td>
        			<td>구분</td>
        			<td>그룹명</td>
        			<td>사용여부</td>
        			<td align='center'>관리</td>
        		</tr>
        		<tr><td colspan="5" height="1" bgcolor="#FFFFFF"></td></tr>
<%

				SQL = "	select * from TB_smsgroup where groupgb = '3' order by idx"
				i = 0
				SET Rs = db.execute(SQL)
				do until Rs.eof
						
						i = i + 1
						if rs("groupgb") = "3" then
							groupgb_NM = "다자간용"
						else
							groupgb_NM = "공용"
						end if
						groupname=rs("groupname")
						if rs("useyn") = "Y" then
							useyn_NM = "사용"
						else
							useyn_NM = "미사용"
						end if
						idx = rs("idx")
					%>

				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">
					<td align="center"><%=i%></td>
					<td align="center"><%=groupgb_NM%></td>
					<td align="center"><%=groupname%></td>
					<td align="center"><%=useyn_NM%></td>
					<td align="center" width="60">
							<img src="/Images/Comm/IconHome.gif" title="목록보기" style="cursor:hand;" align="absmiddle" onclick="document.all.ogroupidx.value=<%=idx%>; ListFrame.location.href='conferencegroup_addresslist.asp?conferencegroupid=<%=idx%>&conferencegroupname=<%=groupgb_NM%>-<%=groupname%>';">
							<img src="/Images/Btn/BtnIconModify.gif" title="그룹수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('<%=idx%>','UP');">
							<img src="/Images/Btn/BtnIconDel.gif" title="그룹삭제" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_del('<%=idx%>','D');">
					</td>
				</tr>
					<%
					Rs.movenext
				loop


%>
				<!--<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">
					<td align="center">1</td>
					<td align="center">개인</td>
					<td align="center">친구</td>
					<td align="center">사용</td>
					<td align="center" width="60">
							<img src="/Images/Comm/IconHome.gif" title="목록보기" style="cursor:hand;" align="absmiddle" onclick="ListFrame.location.href='conferencegroup_addresslist.asp?conferencegroupid=<%=conferencegroupid%>&conferencegroupname=개인-친구';">
							<img src="/Images/Btn/BtnIconModify.gif" title="그룹수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
							<img src="/Images/Btn/BtnIconDel.gif" title="그룹삭제" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','DEL');">
					</td>
				</tr>
				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">
					<td align="center">2</td>
					<td align="center">공통</td>
					<td align="center">긴급발령</td>
					<td align="center">사용</td>
					<td align="center" width="60">
							<img src="/Images/Comm/IconHome.gif" title="목록보기" style="cursor:hand;" align="absmiddle" onclick="ListFrame.location.href='conferencegroup_addresslist.asp?conferencegroupid=<%=conferencegroupid%>&conferencegroupname=공통-긴급발령';">
							<img src="/Images/Btn/BtnIconModify.gif" title="그룹수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
							<img src="/Images/Btn/BtnIconDel.gif" title="그룹삭제" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','DEL');">
					</td>
				</tr>
				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">
					<td align="center">3</td>
					<td align="center">공통</td>
					<td align="center">긴급발령</td>
					<td align="center">사용</td>

					<td align="center">
							<img src="/Images/Btn/BtnIconModify.gif" title="그룹수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
							<img src="/Images/Btn/BtnIconDel.gif" title="그룹삭제" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','DEL');">
					</td>
				</tr>-->			
				
			</table>
			</DIV>
		</td>
		<td width="10"></td>
		<td width="590">
			<iframe src="conferencegroup_addresslist.asp?conferencegroupid=<%=idx%>&conferencegroupname=<%=groupgb_NM%>-<%=groupname%>" name="ListFrame" width="100%" height="100%" frameborder=0 marginheight=0 marginwidth=0 scrolling="no"></iframe>
		</td>

	</tr>
	<tr>
	<td colspan=3>
		<table width="940" border="0" cellspacing="0" cellpadding="0">
			<tr><td height="22" colspan="2" class="FBlk"></b></td></tr>
		</table>			
		<table width="940"  border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
			<tr bgcolor="#EEF6FF" height="40" align=center><td height="22" colspan="2" class="FBlk" align="center"> 선택한 대상을&nbsp;&nbsp;
						<select name="movegroup" size="1" class="ComboFFFCE7" >
<%					
							SQL = "	select * from tb_smsgroup where useyn = 'Y' and groupgb = '3'" '인물관련
							SET Rs = DB.execute(SQL)
							do until Rs.eof
									CODE = Rs("idx")
									if Rs("groupgb")="3" then
										CODENAME = ""&Rs("groupname")
									else
										CODENAME = "공통-"&Rs("groupname")
									end if
								%>

									<%=printSelect("" &CODENAME& "","" &CODE& "","" &movegroup& "")%>
								<%
								Rs.movenext
							loop
%>
						</select> 그룹에 
						&nbsp;&nbsp;&nbsp;<img src="/Images/Btn/BtnPlus_Add.GIF" title="추가하기" style="cursor:hand;" align="absmiddle" onclick="javascript:selectvalue();javascript:fn_insert();"> &nbsp;&nbsp;또는 &nbsp;&nbsp;<img src="/Images/Btn/BtnPlus_Mov.GIF" title="이동하기" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_move();"></b></td></tr>
		</table>
	</td>
	</tr>
	<form name="insFrm" method="post" action="conferencegroup_move.asp">
<input value="" name="JOBGB" type="hidden" size="30">
<input value="" name="selectvalue" type="hidden" size="30">
<input value="<%=idx%>" name="ogroupidx" type="hidden" size="30">
<input value="" name="ngroupidx" type="hidden" size="30">
	</form>
</table>

<!-- #include virtual="/Include/PopLayer.asp" -->
<!-- #include virtual="/Include/Bottom.asp" -->

<script>
<!--
	function fn_update(arg0, arg1){
		//sms그룹 수정
		ShowPOPLayer('conferencegroup_detail.asp?idx='+arg0,'500','210');
	}
	function fn_del(arg0, arg1){
		//sms그룹 수정
		if (confirm("선택한 그룹을 삭제하시겠습니까?"))
			ShowPOPLayer('conferencegroup_detail.asp?idx='+arg0+'&JOBGB=D','500','210');
	}
	function fn_insert(){
		//sms그룹 수정

	if ( document.insFrm.selectvalue.value == "undefined" )
		document.insFrm.selectvalue.value == "";

		if ( document.insFrm.selectvalue.value.length==0 )
			alert('선택한 자료가 없습니다');

		if (confirm("선택한 자료를 추가하시겠습니까?"))
		{
			document.insFrm.JOBGB.value="INS";
			document.all.ngroupidx.value=document.all.movegroup.value;
			insFrm.submit();
		}

	}
	function fn_move(){
		//sms그룹 수정

		if ( document.insFrm.selectvalue.value == "undefined" )
			document.insFrm.selectvalue.value == "";

		if ( document.insFrm.selectvalue.value.length==0 )
			alert('선택한 자료가 없습니다');

		if (confirm("선택한 자료를 이동하시겠습니까?"))
		{
			document.insFrm.JOBGB.value="MOV";
			document.all.ngroupidx.value=document.all.movegroup.value;
			insFrm.submit();
		}
	}


  function selectvalue() {
    var i, chked=0;
	document.insFrm.selectvalue.value = ListFrame.document.ListForm.Chk.length;

    for(i=0;i<ListFrame.document.ListForm.Chk.length;i++) 
	{
		if(ListFrame.document.ListForm.Chk[i].type=='checkbox') 
		{ 
			if(ListFrame.document.ListForm.Chk[i].checked) 
			{ 				
				if ( document.insFrm.selectvalue.value == "")
				{
					document.insFrm.selectvalue.value = ListFrame.document.ListForm.Chk[i].value;	
				}
				else 
				{
					document.insFrm.selectvalue.value = document.insFrm.selectvalue.value + ':' + ListFrame.document.ListForm.Chk[i].value;	
				} 
			}

		}

   }
	if ( document.insFrm.selectvalue.value == "undefined" && ListFrame.document.ListForm.Chk.checked )
		document.insFrm.selectvalue.value = "1:" + ListFrame.document.ListForm.Chk.value;
	

		
}

//-->
</script>