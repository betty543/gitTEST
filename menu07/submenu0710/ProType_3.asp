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
%>
<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" onLoad="ifHeight('Pro3fr');">
<div name="ifr" id="ifr">
<table border="0" cellspacing="0" width="100%">
	<tr height="25"><td class="FBlkB" align="center">3���з�</td></tr>
	<tr>
		<td height="300" valign="top" style="border:1px solid #999999;" bgcolor="#FFFFFF">
			<!--// 3���з� ����Ʈ// -->
			<DIV style="OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 100%; HEIGHT: 100%;">
			<table width="99%" border="0" cellspacing="0" cellpadding="0" align="center">
			<%
				if sBclass = "" then
			%>
				<tr align="center" height="25">
					<td>2���з��� ���õ��� �ʾҽ��ϴ�.</td>
				</tr>
			<%
				else
					sql = "SELECT COUNT(*) as cnt FROM TB_ARMYINFO WHERE ACLASS = '" &sAclass& "' AND BCLASS = '" &sBclass& "' AND CCLASS IS NOT NULL AND DCLASS IS NULL"
					'Response.Write sql
					set Rs = db.execute(sql)
			
					if NOT Rs.eof then
						trowcount = Rs("cnt")
					end if
			
					Rs.Close
					Set Rs = Nothing
					
					sql = "SELECT SEQ, CCLASS, CLASSNAME FROM TB_ARMYINFO WHERE ACLASS = '" &sAclass& "' AND BCLASS = '" &sBclass& "' AND CCLASS IS NOT NULL AND DCLASS IS NULL ORDER BY CCLASS"
					set Rs = db.execute(sql)
					
					if Rs.EOF OR Rs.BOF then
			%>
				<tr align="center" height="25">
					<td>�Էµ� 3���з��� �����ϴ�.</td>
				</tr>
			<%
					else
						i = 1
						Do until Rs.EOF
														
							Seq = Rs("SEQ")
							sCclass = Rs("CCLASS")
							sClassname = Rs("CLASSNAME")
			%>
				<tr id="cTR<%=i%>" onclick="ClickBG('<%=i%>','<%=trowcount%>');" bgcolor="ffffff">
					<td class="TDCont">
						<a href="##" onclick="goPFrame('ProType_4.asp?Aclass=<%=sAclass%>&Bclass=<%=sBclass%>&Cclass=<%=sCclass%>&Classname=<%=sClassname%>','Pro4fr');  parent.TempleteADD('about:blank','OFF');">
							<font color="#FF0000">[<%=sCclass%>]</font> <%=sClassname%>
						</a>

					</td>
					<td nowrap width="35" align="center">
						<img src="/Images/Btn/BtnIconModify.gif" title="����" style="cursor:hand;" align="absmiddle" onClick=" parent.TempleteADD('ProType_Detail.asp?Seq=<%=Seq%>&db_flag=UP&class_gb=C&Aclass=<%=sAclass%>&Bclass=<%=sBclass%>&Cclass=<%=sCclass%>','ON');">
						<img src="/Images/Btn/BtnIconDel.gif" title="����" style="cursor:hand;" align="absmiddle" onClick="javascrip:if(confirm('�ش� ����Ÿ�� ���� �Ͻðڽ��ϱ�?')) {location.href='ProType_InsUpDel.asp?Seq=<%=Seq%>&db_flag=DEL&class_gb=C&Aclass=<%=sAclass%>&Bclass=<%=sBclass%>&Cclass=<%=sCclass%>'}">
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
			<!--// 3���з� ����Ʈ// -->
		</td>
	</tr>
	<% if sBclass <> "" then %>
	<tr><td height="25" align="center"><input type="button" name="BtnPrint" value="3�� �߰�" style="width:100%; height:100%;" class="Btn4" onClick="parent.TempleteADD('ProType_Detail.asp?db_flag=INS&class_gb=C&Aclass=<%=sAclass%>&Bclass=<%=sBclass%>','ON')"></td></tr>
	<% end if %>
</table>
</div>

<!-- #include virtual="/Include/Bottom.asp" -->