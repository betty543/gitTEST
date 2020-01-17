<!-- #include virtual="/Include/Top_Frame.asp" -->

<!-- 프레임1 시작 -->


<%

	JUBSEQ = request("JUBSEQ")
	Keyword = request("Keyword")
	if JUBSEQ = "" then
		response.end

	end if
%>
<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0">

<div name="ifr" id="ifr">
<table cellspacing="0" cellpadding="0" width="100%" >
	<tr>
		<td>

<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT QUESTION, REPLY, REMARK FROM TB_CRIMECALLHISTORY "
							SqlCode = SqlCode& " WHERE JUBSEQ = '" & JUBSEQ & "'"
							set RsCode = db.execute(SqlCode)

							IF RsCode.EOF = FALSE THEN

								response.write "<b><font color='#ff000'>[문의내용]</font></b>" & vbcrlf
								response.write replace(RsCode("QUESTION"),Keyword,"<b><font color='#000ff'>"&Keyword&"</font></b>")


					
%>

		</td>
	</tr>


	<tr>
		<td>

<%



								response.write "<b><font color='#ff000'>[조치내용]</font></b>" & vbcrlf
								response.write replace(RsCode("REPLY"),Keyword,"<b><font color='#0000ff'>"&Keyword&"</font></b>")



					
%>

		</td>
	</tr>
	<tr>
		<td>

<%



								response.write "<b><font color='#ff000'>[특이사항]</font></b>" & vbcrlf
								response.write replace(RsCode("REMARK"),Keyword,"<b><font color='#0000ff'>"&Keyword&"</font></b>")


							END IF
							RsCode.close
					
%>

		</td>
	</tr>
</table>
</div>

<!-- #include virtual="/Include/Bottom.asp" -->
