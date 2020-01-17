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
		if ( eval("parent.document.all.whereCD7") != null )
			eval("parent.document.all.whereCD7").value = document.all.level2.value;

		eval("parent.document.all.LEVEL2").value = document.all.level2.value;
	}
	catch(e){}
}
-->
</script>

<!-- 프레임1 시작 -->
<form name="frmCode" style="margin:0">

<%
level = Request("level")
level2 = Request("level2")

%>
<body bgcolor="#FFFFFF" topmargin=10 leftmargin=0 onload="ifHeight2('frame_level');">
<table width="100" height="100%" border="0" cellspacing="0" cellpadding="0" align=left  bgcolor="#000000">
   <tr>
    <td bgcolor="FFFFFF" align="left" height="29" valign="center"><select name="level2" onchange="javascript:fn_putetc2();">		
		<option value="" <% if level2 = "" then%>selected<%end if%>>계급선택</option>

		<%if level = "A" or  level = "B" then 

				'======= 처리구분 코드 가져오기 ==================================================
				SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
				if level = "A" then
					SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='C06'"
				elseif level = "B" then
					SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='C07'"
				end if
				SqlCode = SqlCode& " ORDER BY CODE"
				set RsCode = db.execute(SqlCode)

				do until RsCode.eof

					CODE = RsCode("CODE")
					CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &level2& "")%>
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
