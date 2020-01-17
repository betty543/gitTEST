<!-- #include virtual="/Include/Top_Frame.asp" -->

<script>
<!--

// iframe 사이즈 적용

function fn_putetc2()
{
	try{
		eval("parent.document.all.whereCD2_B").value = document.all.whereCD2_B.value;

	}
	catch(e){}
}
-->
</script>

<!-- 프레임1 시작 -->


<%
CHANNEL_A = Request("CHANNEL_A")
CHANNEL_B = Request("CHANNEL_B")

%>
<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" onLoad="ifHeight2('ChannelFrameB');">

<div name="ifr" id="ifr">
<table cellspacing="0" cellpadding="0"  >
	<tr>
		<td>

<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CCLASS CODE, CLASSNAME CODENAME FROM TB_ARMYINFO"
							SqlCode = SqlCode& " WHERE ACLASS = 'Q' AND BCLASS = '" & CHANNEL_A & "' AND CCLASS IS NOT NULL AND DCLASS IS NULL"
							SqlCode = SqlCode& " ORDER BY BCLASS"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="whereCD2_B" size="1" class="ComboFFFCE7" onclick="fn_putetc2();">
							<Option value ='' selected>상담유형2차==</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &CHANNEL_B& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>						

		</td>
	</tr>
</table>
</div>

<!-- #include virtual="/Include/Bottom.asp" -->
