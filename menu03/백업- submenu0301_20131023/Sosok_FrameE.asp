<!-- #include virtual="/Include/Top_Frame.asp" -->

<script>
<!--

// iframe 사이즈 적용

function fn_putetc2()
{
	try{
		eval("parent.document.all.whereCD5_E").value = document.all.whereCD5_E.value;
	}
	catch(e){}
}
-->
</script>

<!-- 프레임1 시작 -->


<%
SOSOK_A = Request("SOSOK_A")
SOSOK_B = Request("SOSOK_B")
SOSOK_C = Request("SOSOK_C")
SOSOK_D = Request("SOSOK_D")
SOSOK_E = Request("SOSOK_E")
%>
<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" onLoad="ifHeight2('SosokFrameE');">

<div name="ifr" id="ifr">
<table cellspacing="0" cellpadding="0"  >
	<tr>
		<td>

<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CCLASS CODE, CLASSNAME CODENAME FROM TB_ARMYINFO"
							SqlCode = SqlCode& " WHERE ACLASS = '" &SOSOK_A&"' AND BCLASS = '" &SOSOK_B&"' AND CCLASS = '" &SOSOK_C&"' AND DCLASS = '" &SOSOK_D&"' AND ECLASS IS NULL"
							SqlCode = SqlCode& " ORDER BY DCLASS"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="whereCD5_E" size="1" class="ComboFFFCE7" onclick="fn_putetc2();">
							<Option value ='' selected>소속5차=====</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &SOSOK_E& "")%>
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
