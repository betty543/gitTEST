<!-- #include virtual="/Include/Top_Frame.asp" -->

<script>
<!--

// iframe 사이즈 적용

function fn_putetc2()
{
	try{
		eval("parent.document.all.whereCD6_C").value = document.all.whereCD6_C.value;

	}
	catch(e){}
}
-->
</script>

<!-- 프레임1 시작 -->


<%
LEVEL_A = Request("LEVEL_A")
LEVEL_B = Request("LEVEL_B")
LEVEL_C = Request("LEVEL_C")
%>
<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" onLoad="ifHeight2('LevelFrameC');">

<div name="ifr" id="ifr">
<table cellspacing="0" cellpadding="0"  >
	<tr>
		<td>

<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT DCLASS CODE, CLASSNAME CODENAME FROM TB_ARMYINFO"
							SqlCode = SqlCode& " WHERE ACLASS = 'P' AND BCLASS = '" &LEVEL_A&"' AND CCLASS = '" &LEVEL_B&"' AND DCLASS IS NOT NULL"
							SqlCode = SqlCode& " ORDER BY DCLASS"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="whereCD6_C" size="1" class="ComboFFFCE7" onclick="fn_putetc2();">
							<Option value ='' selected>계급3차==</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &LEVEL_C& "")%>
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
