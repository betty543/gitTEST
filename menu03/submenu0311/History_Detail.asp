<!-- #include virtual="/Include/Top_Frame.asp" -->

<!-- ������1 ���� -->


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
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT QUESTION, REPLY, REMARK FROM TB_LIFECALLHISTORY_OB "
							SqlCode = SqlCode& " WHERE JUBSEQ = '" & JUBSEQ & "'"

							'response.write SqlCode
							set RsCode = db.execute(SqlCode)

							IF RsCode.EOF = FALSE THEN

								response.write "<b><font color='#ff000'>[��㳻��]</font></b>" & vbcrlf
								response.write replace(RsCode("QUESTION"),Keyword,"<b><font color='#000ff'>"&Keyword&"</font></b>")


					
%>

		</td>
	</tr>


	<tr>
		<td>

<%



								response.write "<b><font color='#ff000'>[Ư�̻���]</font></b>" & vbcrlf
								response.write replace(RsCode("REMARK"),Keyword,"<b><font color='#0000ff'>"&Keyword&"</font></b>")


							END IF
							RsCode.close
					
%>

		</td>
	</tr>
</table>
</div>

<!-- #include virtual="/Include/Bottom.asp" -->