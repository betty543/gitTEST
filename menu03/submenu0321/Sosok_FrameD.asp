<!-- #include virtual="/Include/Top_Frame.asp" -->

<script>
<!--

// iframe 사이즈 적용

function fn_putetc2()
{
	try{
		eval("parent.document.all.whereCD5_D").value = document.all.whereCD5_D.value;
		eval("parent.document.all.whereCD5_E").value = "";
		parent.SosokFrameE.location = "/menu03/submenu0301/Sosok_FrameE.asp?SOSOK_A="+eval("parent.document.all.whereCD5_A").value+"&SOSOK_B="+eval("parent.document.all.whereCD5_B").value+"&&SOSOK_C="+eval("parent.document.all.whereCD5_C").value+"&SOSOK_D="+eval("parent.document.all.whereCD5_D").value+"&SOSOK_E=";
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
%>
<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" onLoad="ifHeight2('SosokFrameD');">

<div name="ifr" id="ifr">
<table cellspacing="0" cellpadding="0"  >
	<tr>
		<td>

<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT CCLASS CODE, CLASSNAME CODENAME FROM TB_ARMYINFO"
							SqlCode = SqlCode& " WHERE ACLASS = '" &SOSOK_A&"' AND BCLASS = '" &SOSOK_B&"' AND CCLASS = '" &SOSOK_c&"' AND DCLASS IS NOT NULL AND ECLASS IS NULL"
							SqlCode = SqlCode& " ORDER BY DCLASS"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="whereCD5_D" size="1" class="ComboFFFCE7" onclick="fn_putetc2();">
							<Option value ='' selected>소속4차=====</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &SOSOK_D& "")%>
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
