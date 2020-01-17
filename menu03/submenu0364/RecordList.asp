<!-- #include virtual="/Include/Top.asp" -->

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<script>
	function fn_Search() {
		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	function fn_Search1() {
		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	function fn_GetId(arg,arg1) {
		document.inUpFrm.agt_id.value =arg;
		document.inUpFrm.FileName.value =arg1;
	}

</script>
<table border="0" width="1200" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<form name="inUpFrm" method="post" action="RecordList.asp" onsubmit="return fn_Search(this);" style="margin:0">
			<input type="hidden" name="QueryYN" value="<%=QueryYN%>">	
			<table width="1200" border="0" cellspacing="1" cellpadding="0" style="border:#E1DED6 solid 1px">
			    <tr>
			        <td class="TDCont">조회기간 :
			        	<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
			        </td><td class="TDCont">사용자 :
						<%
							'======= 상담원 가져오기 ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y'"
							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="whereCD1" size="1" class="ComboFFFCE7">
							<option value="">사용자 선택---</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = RsCode("USERNAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD1& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						<%
							'======= 상담원 가져오기 ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='N'		and	outdate >= '"&DateAdd("d",1,DateAdd("m",-1,Date())) &"'"
							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)

								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = "[퇴사]"&RsCode("USERNAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD1& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>
						</td><td class="TDCont">전화번호 : <input value="<%=PID%>" name="PID" type="text" size="14" onfocus="setFocusColor(this);" onblur="setOutColor(this);">
						</td><td class="TDCont"><img src="/Images/Btn/BtnSearch.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Search();">
			        </td>


			    </tr>
			</table>
			</form>
		</td>
	</tr>
</table>

<table border="0" width="100%" cellpadding="0" cellspacing="0" align="center"><tr height="5"><td></td></tr></table>

<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td align="center">
			<DIV style="OVERFLOW-Y:auto; OVERFLOW-X:auto; MARGIN: 0px 0px 0px 0px; WIDTH:1200; HEIGHT:700;">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
				<tr><td colspan="100" height="1" bgcolor="#FFFFFF"></td></tr>

				<tr height="20" bgcolor="#FFFFFF">
					<td colspan='2' align='center' bgcolor="#EEF6FF">녹취일시</td>
					<td align='center' bgcolor="#EEF6FF">통화시간</td>
					<td align='center' bgcolor="#EEF6FF">파일명 </td>
					<td align='center' bgcolor="#EEF6FF"></td>
					<td colspan='2' align='center' bgcolor="#EEF6FF">사용자</td>		
					<td align='center' bgcolor="#EEF6FF">IN/OUT</td>	
					<td align='center' bgcolor="#EEF6FF">전화번호</td>						
				</tr>
				<%'####### 실제자료가 들어간다. %>
				<tr height="20" bgcolor="#FFFFFF">
					<td colspan='2' align='center'>2009-05-01 18:00:00</td>
					<td align='center'>00:20:09</td>
					<td align='center'>c:\20090501_123456.wav </td><td align='center'><img src="/Images/Comm/IconAlert.gif" style="cursor:hand;" onClick="fn_dial('1');" align="absmiddle" title="전화걸기"></td>
					<td colspan='2' align='center'>손민경</td>		
					<td align='center'>인바운드</td>	
					<td align='center'>010-999-0000</td>						
				</tr>
				<tr height="20" bgcolor="#FFFFFF">
					<td colspan='2' align='center'>2009-05-02 18:00:00</td>
					<td align='center'>00:03:09</td>
					<td align='center'>c:\20090501_123456.wav </td><td align='center'><img src="/Images/Comm/IconAlert.gif" style="cursor:hand;" onClick="fn_dial('1');" align="absmiddle" title="전화걸기"></td>
					<td colspan='2' align='center'>손민경</td>		
					<td align='center'>인바운드</td>	
					<td align='center'>010-999-0000</td>						
				</tr>

			</table>
			</DIV>
		</td>
	</tr>
</table>

<script>
<!--

	function fn_Player(){
		//파일명
		var x,y;
		x = ( screen.width - 300 )/2;
		y = ( screen.height - 200 )/2;
		window.open("/include/Popup/WavePlayer.asp?FileName="+inUpFrm.FileName.value+"&RecDate=","Player", "toolbar=no,top=100,left=300,width=300,height=200,resize=no,status=yes, scrollbars=no");
	}

//-->
</script>


<!-- #include virtual="/Include/Bottom.asp" -->