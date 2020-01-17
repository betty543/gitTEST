<!-- #include virtual="/Include/Top_Frame.asp" -->

<script>
<!--

// iframe 사이즈 적용
function ifHeight2(f){
	obj = eval("parent.document.all."+ f);
	obj.style.height = document.all.level2.offsetHeight;
	obj.style.width = document.all.level2.offsetWidth;
}


function fn_putetc2()
{
	try{
		//eval("parent.document.all.whereCD7").value = document.all.level2.value;
		eval("parent.document.all.SOSOKETCGB").value = document.all.level2.value;
	}
	catch(e){}
}
-->
</script>

<!-- 프레임1 시작 -->
<form name="frmCode" style="margin:0">

<%
SOSOKGB = Request("SOSOKGB")
SOSOKETCGB = Request("SOSOKETCGB")
%>
<body bgcolor="#FFFFFF" topmargin=10 leftmargin=0 onload="ifHeight2('frame_sosok');">
<table width="100" height="100%" border="0" cellspacing="0" cellpadding="0" align=left  bgcolor="#000000">
   <tr>
    <td bgcolor="FFFFFF" align="left" height="29" valign="center"><select name="level2" onchange="javascript:fn_putetc2();">		
		<option value="" <% if SOSOKETCGB = "" then%>selected<%end if%>>기타소속</option>

		<%
			if SOSOKGB = "H" then
				'======= 처리구분 코드 가져오기 ==================================================
				SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
				SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='C41'"
				SqlCode = SqlCode& " ORDER BY CODE"
				set RsCode = db.execute(SqlCode)

				do until RsCode.eof

					CODE = RsCode("CODE")
					CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &SOSOKETCGB& "")%>
							<%
					RsCode.MoveNext
				LOOP

				RsCode.Close
				set RsCode = NOTHING
			end if
		%></select>
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/Include/Bottom.asp" -->
