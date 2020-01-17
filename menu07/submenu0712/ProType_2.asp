<!-- #include virtual="/include/top_frame.asp" -->
<script language="javascript">
<!--
	function ClickBG(f,c){
		for(var i=1; i<=c; i++){
			document.getElementById('cTR' +i).style.backgroundColor = (i==parseInt(f)) ? "#FFEEF9" : "#FFFFFF";
		}
	}
-->
</script>
<%
	sAclass = Request("Aclass")
%>
<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" onLoad="ifHeight('Pro2fr');">
<div name="ifr" id="ifr">
<table border="0" cellspacing="0" width="100%">
	<tr height="25"><td class="FBlkB" align="center">2차분류</td></tr>
	<tr>
		<td height="300" valign="top" style="border:1px solid #999999;" bgcolor="#FFFFFF">
			<!--// 2차분류 리스트// -->
			<DIV style="OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 100%; HEIGHT: 100%;">
			<table width="99%" border="0" cellspacing="0" cellpadding="0" align="center">
			<%
				if sAclass = "" then
			%>
				<tr align="center" height="25">
					<td>1차분류가 선택되지 않았습니다.</td>
				</tr>
			<%
				else
					sql = "SELECT COUNT(*) as cnt FROM TB_ARMYINFO WHERE ACLASS = '" &sAclass& "' AND BCLASS IS NOT NULL AND CCLASS IS NULL "
					'Response.Write sql
					set Rs = db.execute(sql)
			
					if NOT Rs.eof then
						trowcount = Rs("cnt")
					end if
			
					Rs.Close
					Set Rs = Nothing
					
					sql = "SELECT SEQ, BCLASS, CLASSNAME FROM TB_ARMYINFO WHERE ACLASS = '" &sAclass& "' AND BCLASS IS NOT NULL AND CCLASS IS NULL ORDER BY BCLASS"
					set Rs = db.execute(sql)
					
					if Rs.EOF OR Rs.BOF then
			%>
				<tr align="center" height="25">
					<td>입력된 2차분류가 없습니다.</td>
				</tr>
			<%
					else
						i = 1
						Do until Rs.EOF
														
							Seq = Rs("SEQ")
							sBclass = Rs("BCLASS")
							sClassname = Rs("CLASSNAME")
			%>
				<tr id="cTR<%=i%>" onclick="ClickBG('<%=i%>','<%=trowcount%>');" bgcolor="ffffff">
					<td class="TDCont">
						<a href="##" onclick="goPFrame('ProType_3.asp?Aclass=<%=sAclass%>&Bclass=<%=sBclass%>&Classname=<%=sClassname%>','Pro3fr');  parent.TempleteADD('about:blank','OFF');">
							<font color="#FF0000">[<%=sBclass%>]</font> <%=sClassname%>
						</a>
					</td>
					<td nowrap width="35" align="center">
						<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onClick="goPFrame('ProType_3.asp','Pro3fr');  parent.TempleteADD('ProType_Detail.asp?Seq=<%=Seq%>&db_flag=UP&class_gb=B&Aclass=<%=sAclass%>&Bclass=<%=sBclass%>','ON');">
						<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle" onClick="javascrip:if(confirm('해당 데이타를 삭제 하시겠습니까?')) {location.href='ProType_InsUpDel.asp?Seq=<%=Seq%>&db_flag=DEL&class_gb=B&Aclass=<%=sAclass%>&Bclass=<%=sBclass%>'}">
					</td>
				</tr>
				<tr><td colspan="2" class="TRLine" height="1"></td></tr>
			<%
							i = i + 1
							Rs.Movenext
						Loop
					
					End if
					Rs.Close
					Set Rs = Nothing
				end if
			%>
			</table>
			</DIV>
			<!--// 2차분류 리스트// -->
		</td>
	</tr>
	<% if sAclass <> "" then %>
	<tr><td height="25" align="center"><input type="button" name="BtnPrint" value="2차 추가" style="width:100%; height:100%;" class="Btn4" onClick="goPFrame('ProType_3.asp','Pro3fr');  parent.TempleteADD('ProType_Detail.asp?db_flag=INS&class_gb=B&Aclass=<%=sAclass%>','ON')"></td></tr>
	<% end if %>
</table>
</div>


<!-- #include virtual="/Include/Bottom.asp" -->
