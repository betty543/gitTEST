<!-- #include virtual="/Include/Top.asp" -->
<script language="javascript">
<!--
	function ClickBG(f,c){
		for(var i=1; i<=c; i++){
			document.getElementById('cTR' +i).style.backgroundColor = (i==parseInt(f)) ? "#FFEEF9" : "#FFFFFF";
		}
	}
-->
</script>
<table width="1000" height="85%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td height="10"></td></tr>
    <tr>
    	<td width="420" valign="top">
			<!-- 종합코드관리 구분 테이블 START -->
			<DIV style="OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 100%; HEIGHT: 100%;">
			<table border="0" cellspacing="1" cellspacing="1" bgcolor="#CCCCCC">
				<tr align="center">
					<td bgcolor="#FCFAED" class="TDCont" nowrap width="50%">설명</td>
					<td bgcolor="#FCFAED" class="TDCont" nowrap width="20%">구분</td>
					<td bgcolor="#FCFAED" class="TDCont" nowrap width="30%">구분명</td>
				</tr>
				<%
					sql = "SELECT COUNT(*) as cnt FROM (SELECT CODEGROUP, GROUPNAME FROM TB_CODE GROUP BY CODEGROUP, GROUPNAME) a"
					'Response.Write sql
					set Rs = db.execute(sql)
			
					if NOT Rs.eof then
						trowcount = Rs("cnt")
					end if
			
					Rs.Close
					Set Rs = Nothing
			
					sql = "SELECT CODEGROUP, GROUPNAME FROM TB_CODE where codegroup > 'A11' GROUP BY CODEGROUP, GROUPNAME  ORDER BY CODEGROUP"
					set Rs = db.execute(sql)
					
					lv_Write = 0
					lv_Rowspan = 0

					if NOT Rs.EOF then
						i = 1
						Do until Rs.EOF

							sCodegroup = Rs("codegroup"):	sGroupname = Rs("groupname")

							If lv_Rowspan = lv_Write Then
							


									sUseName = "생명의전화 관련 코드"
									lv_Rowspan = 100
									lv_Write = 0

							End if

				%>
			    <tr bgcolor="#ffffff" id="cTR<%=i%>" style="cursor:hand;" onclick="goFrame('code_list.asp?sCodegroup=<%=sCodegroup%>&sGroupname=<%=sGroupname%>', 'ifr'); ClickBG('<%=i%>','<%=trowcount%>');">
					<%if lv_Write = 0 then%>
						<td class="TDCont" rowspan="<%=lv_Rowspan%>"><%=sUseName%></td>

					<%End if%>
					<% lv_Write = lv_Write + 1 %>
			        <td align="center"><font color="#FF0000">[<%=sCodegroup%>]</font></td>
			        <td class="TDCont"><a href="javascript:goFrame('code_list.asp?sCodegroup=<%=sCodegroup%>&sGroupname=<%=sGroupname%>', 'ifr');"><%=sGroupname%></a></td>
					
			    </tr>
				<%
							i = i + 1
							Rs.Movenext
						Loop
				
						Rs.Movefirst
						
						firstCodegroup = Rs("codegroup")
						firstGroupname = Rs("groupname")
					else
				%>
				<tr bgcolor="#FFFFFF" align="center" height="30"><td colspan="2">등록된 코드가 없습니다.</td></tr>
				<%
					End if
					
			
					Rs.Close
					Set Rs = Nothing
				%>
			</table>
			</DIV>
			<!-- 종합코드관리 구분 테이블 END -->
    	</td>
		<td width="780" valign="top" class="TDL10px">
			<table cellpadding="0" cellspacing="0" height="100%">
				<tr>
					<td width="780" >
					<!-- 아이프레임 시작 부분-->
					<iframe src="code_list.asp?sCodegroup=<%=firstCodegroup%>&sGroupname=<%=firstGroupname%>" frameborder=0 marginheight=0 marginwidth=0 width="100%" height="100%" scrolling="auto" name="masterFrame" name="ifr" id ="ifr"></iframe>
					<!-- 아이프레임 시작 end-->
					</td>
				</tr>
			</table>
		</td>
    </tr>
</table>


<!-- #include virtual="/Include/Bottom.asp" -->