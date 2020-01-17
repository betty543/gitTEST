<!-- #include virtual="/Include/Top_Frame.asp" -->

<script>
<!--



function fn_putetc3(arg)
{
	//try{
		//eval("parent.document.all.whereCD7").value = document.all.level3.value;
		//eval("parent.document.all.SOSOKETCGB2").value = document.all.level3.value;

		//alert(arg);
		if (document.all.level3_1(arg-1).value == '')
			eval("parent.document.all.CounselorYN").value = '';
		else
			eval("parent.document.all.CounselorYN").value = document.all.level3_1(arg-1).text;
		//alert( document.all.level3.value);
		//document.all.level3_1.options[inUpFrm.whereCD4.selectedIndex].value
	//}
	//catch(e){}
}
-->
</script>

<!-- 프레임1 시작 -->
<form name="frmCode" style="margin:0">

<%
SOSOKGB = Request("SOSOKGB")
SOSOKETCGB = Request("SOSOKETCGB")
SOSOKETCGB2 = Request("SOSOKETCGB2")
%>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0>
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" align=left  bgcolor="#000000">
   <tr>
    <td bgcolor="FFFFFF" align="left" height="29" valign="center">

		<%
			if SOSOKETCGB = "" then

				sReplyHtml = "소속(중)을 선택하시면 소분류가 표시됩니다."
				response.write sReplyHtml

			else
				'======= 처리구분 코드 가져오기 ==================================================
				SqlCode = "select * from tb_armyinfo where aclass = '"&SOSOKGB&"' and bclass = '"&SOSOKETCGB&"' and cclass is not null order by cclass"
				set RsCode = db.execute(SqlCode)

				do until RsCode.eof
					j = j + 1
					SelectedValue = ""
					if j = 1 then
						sReplyHtml = "<input type='RADIO' value='" & RsCode("cclass") & "' name='SOSOKGB3' class='none' " & SelectedValue & "  onClick=""fn_putetc3('"&j&"');"">" & RsCode("classname")	
					else
						sReplyHtml = sReplyHtml & "&nbsp;<input type='RADIO' value='" & RsCode("cclass") & "' name='SOSOKGB3' class='none' " & SelectedValue & "  onClick=""fn_putetc3('"&j&"');"">" & RsCode("classname")	
					end if
					RsCode.movenext
				loop
				RsCode.close
				response.write sReplyHtml

			end if
		%>		
<select name="level3_1" style="display:none">		
		<option value="" <% if SOSOKETCGB2 = "" then%>selected<%end if%>>3차분류</option>
		<%

				'======= 처리구분 코드 가져오기 ==================================================
				SqlCode = "select * from tb_armyinfo where aclass = '"&SOSOKGB&"' and bclass = '"&SOSOKETCGB&"' and cclass is not null order by cclass"
				set RsCode = db.execute(SqlCode)

				do until RsCode.eof

					CODE = RsCode("counseloryn")
					if RsCode("counseloryn") = "Y" then
						CODENAME = "배치"
					else
						CODENAME = "미배치"
					end if
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &SOSOKETCGB2& "")%>
							<%
					RsCode.MoveNext
				LOOP

				RsCode.Close
				set RsCode = NOTHING

		%></select>
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/Include/Bottom.asp" -->
