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
	sBclass = Request("Bclass")
	sCclass = Request("Cclass")
%>
<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" onLoad="ifHeight('Pro4fr');">
<div name="ifr" id="ifr">
<table border="0" cellspacing="0" width="100%">
	<tr height="25"><td class="FBlkB" align="center">4���з�</td></tr>
	<tr>
		<td height="400" valign="top" style="border:1px solid #999999;" bgcolor="#FFFFFF">
			<!--// 4���з� ����Ʈ// -->
			<DIV style="OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 100%; HEIGHT: 100%;">
			<table width="99%" border="0" cellspacing="0" cellpadding="0" align="center">
			<%
				if sCclass = "" then
			%>
				<tr align="center" height="25">
					<td>3���з��� ���õ��� �ʾҽ��ϴ�.</td>
				</tr>
			<%
				else
					sql = "SELECT COUNT(*) as cnt FROM TB_GOODBUNU WHERE ACLASS = '" &sAclass& "' AND BCLASS = '" &sBclass& "' AND CCLASS = '" &sCclass& "'"
					sql = sql&" AND DCLASS IS NOT NULL AND ECLASS IS NULL ORDER BY DCLASS"
					'Response.Write sql
					set Rs = db.execute(sql)
			
					if NOT Rs.eof then
						trowcount = Rs("cnt")
					end if
			
					Rs.Close
					Set Rs = Nothing
					
					sql = "SELECT SEQ, DCLASS, CLASSNAME FROM TB_GOODBUNU WHERE ACLASS = '" &sAclass& "' AND BCLASS = '" &sBclass& "' AND CCLASS = '" &sCclass& "'"
					sql = sql&" AND DCLASS IS NOT NULL AND ECLASS IS NULL ORDER BY DCLASS"
					set Rs = db.execute(sql)
					
					if Rs.EOF OR Rs.BOF then
			%>
				<tr align="center" height="25">
					<td>�Էµ� 4���з��� �����ϴ�.</td>
				</tr>
			<%
					else
						i = 1
						Do until Rs.EOF
														
							Seq = Rs("SEQ")
							sDclass = Rs("DCLASS")
							sClassname = Rs("CLASSNAME")
			%>
				<tr id="cTR<%=i%>" onclick="ClickBG('<%=i%>','<%=trowcount%>');" bgcolor="ffffff">
					<td class="TDCont"><a href="##" onclick="goPFrame('ProType_5.asp?Aclass=<%=sAclass%>&Bclass=<%=sBclass%>&Cclass=<%=sCclass%>&Dclass=<%=sDclass%>&Classname=<%=sClassname%>','Pro5fr'); parent.TempleteADD('about:blank','OFF');"><font color="#FF0000">[<%=sDclass%>]</font> <%=sClassname%></a></td>
					<td nowrap width="35" align="center">
						<center>
						<img src="/Images/Btn/BtnIconModify.gif" title="����" style="cursor:hand;" align="absmiddle" onClick="goPFrame('ProType_5.asp','Pro5fr'); parent.TempleteADD('ProType_Detail.asp?Seq=<%=Seq%>&db_flag=UP&class_gb=D&Aclass=<%=sAclass%>&Bclass=<%=sBclass%>&Cclass=<%=sCclass%>&Dclass=<%=sDclass%>','ON');">
						<img src="/Images/Btn/BtnIconDel.gif" title="����" style="cursor:hand;" align="absmiddle" onClick="javascrip:if(confirm('�ش� ����Ÿ�� ���� �Ͻðڽ��ϱ�?')) {location.href='ProType_InsUpDel.asp?Seq=<%=Seq%>&db_flag=DEL&class_gb=D&Aclass=<%=sAclass%>&Bclass=<%=sBclass%>&Cclass=<%=sCclass%>&Dclass=<%=sDclass%>'}">
						</center>
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
			<!--// 4���з� ����Ʈ// -->
		</td>
	</tr>
	<% if sCclass <> "" then %>
	<tr><td height="25" align="center"><input type="button" name="BtnPrint" value="4�� �߰�" style="width:100%; height:100%;" class="Btn4" onClick="goPFrame('ProType_5.asp','Pro5fr'); parent.TempleteADD('ProType_Detail.asp?db_flag=INS&class_gb=D&Aclass=<%=sAclass%>&Bclass=<%=sBclass%>&Cclass=<%=sCclass%>','ON')"></td></tr>
	<% end if %>
</table>
</div>

<!-- #include virtual="/Include/Bottom.asp" -->