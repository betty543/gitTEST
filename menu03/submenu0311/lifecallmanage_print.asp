<!-- #include virtual="/Include/Top_PopUp.asp" -->
<%
	guboon = request("guboon")
	JUBSEQ = request("JUBSEQ")
	InType = request("InType")

	'response.write JUBSEQ

	SQL = "	SELECT *, CONVERT(CHAR(19),JUBTIME,121) AS JUBTIME1 FROM TB_LIFECALLHISTORY_OB"
	SQL = SQL & "		WHERE	JUBSEQ = '" & JUBSEQ & "'"

	Set Rs = server.createObject("ADODB.Recordset")
	Rs.open SQL,db

	if rs.eof = false then

		db_IDX	= RS("IDX")
		db_JUBSEQ	= RS("JUBSEQ")
		db_JUBDATE	= RS("JUBDATE")
		db_JUBTIME	= RS("JUBTIME")
		db_JUBTIME1	= RS("JUBTIME1")

		db_IOFLAG	= RS("IOFLAG")
		db_EMERYN	= RS("EMERYN")
		db_CUSTNO	= RS("CUSTNO")
		db_CUSTNAME	= RS("CUSTNAME")
		db_TELNO	= RS("TELNO")
		db_TELNO2	= RS("TELNO2")
		db_CID	= RS("CID")
		db_SEXGB	= RS("SEXGB")
		db_CHANNELGB	= RS("CHANNELGB")
		db_REQUESTERGB	= RS("REQUESTERGB")
		db_CONSULTGB	= RS("CONSULTGB")
		db_CONSULTETCGB	= RS("CONSULTETCGB")
		db_SOSOKGB_A	= RS("SOSOKGB_A")
		db_SOSOKGB_B	= RS("SOSOKGB_B")
		db_SOSOKGB_C	= RS("SOSOKGB_C")
		db_SOSOKGB_D	= RS("SOSOKGB_D")
		db_SOSOKGB_E	= RS("SOSOKGB_E")
		db_LEVEL_B	= RS("LEVEL_B")
		db_LEVEL_C	= RS("LEVEL_C")
		db_LEVEL_D	= RS("LEVEL_D")
		db_FAMILYGB	= RS("FAMILYGB")
		db_CALLCLASS_B	= RS("CALLCLASS_B")
		db_CALLCLASS_C	= RS("CALLCLASS_C")
		db_CHANNELGB_B	= RS("CHANNELGB_B")
		db_CHANNELGB_C	= RS("CHANNELGB_C")
		db_CALLFLAG	= RS("CALLFLAG")
		db_CALLKIND_B	= RS("CALLKIND_B")
		db_CALLKIND_C	= RS("CALLKIND_C")
		db_QUESTION	= RS("QUESTION")
		db_REPLY	= RS("REPLY")
		db_REMARK	= RS("REMARK")
		db_RESULTGB	= RS("RESULTGB")
		db_RESERVEDATE	= RS("RESERVEDATE")
		db_RESERVETIME	= RS("RESERVETIME")
		db_PROCESSGB	= RS("PROCESSGB")
		db_WEATHER	= RS("WEATHER")
		db_CALLID	= RS("CALLID")
		db_RECORDFILE	= RS("RECORDFILE")
		db_CALLTIMEDP	= RS("CALLTIMEDP")
		db_CALLTIME	= RS("CALLTIME")
		db_CB_SEQ	= RS("CB_SEQ")
		db_REFERJUBSEQ	= RS("REFERJUBSEQ")
		db_REFCNT	= RS("REFCNT")
		db_FILENAME	= RS("FILENAME")
		db_INCODE	= RS("INCODE")
		db_INDATE	= RS("INDATE")

	IF WEEKDAY(db_JUBTIME1)=1 THEN

		JUBDAY="��"
	ELSEIF WEEKDAY(db_JUBTIME1)=2 THEN
		JUBDAY="��"
	ELSEIF WEEKDAY(db_JUBTIME1)=3 THEN
		JUBDAY="ȭ"
	ELSEIF WEEKDAY(db_JUBTIME1)=4 THEN
		JUBDAY="��"
	ELSEIF WEEKDAY(db_JUBTIME1)=5 THEN
		JUBDAY="��"
	ELSEIF WEEKDAY(db_JUBTIME1)=6 THEN
		JUBDAY="��"
	ELSEIF WEEKDAY(db_JUBTIME1)=7 THEN
		JUBDAY="��"
	END IF

	end if
%>
<!--<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" border="1">-->


<table width="760" border="0" cellspacing="0" cellpadding="0" bgcolor="#ffffff" align='center'>
<tr bgcolor="#ffffff">
<td align='center'>
<img src="/Images/Btn/BtnPrint.gif" style="cursor:hand;" onClick="javascript:print_info();" title="�����ͷ� ���">
</td>
</tr>
</table>
<div id="A" style="OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 760;  HEIGHT: 400;">
<table width="600" height="10" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table width="600" cellspacing="0" align="center" border="1" bordercolor="black" bordercolordark="white" bordercolorlight="black">
	<tr bgcolor="#FFFFFF">
		<td>	
	

			<table width="600" height="80" border="1" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff" bordercolor="black" bordercolordark="white" bordercolorlight="black">
			    <tr height="80">
					<td align='center' height="80">
						<b><font color="#000000" size="5px">���Ļ������</font></b>
					</td>
					<!--<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="8">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;<b><font color="#0000ff" size=15px>�������</font></td>-->
				</tr>
			</table>

			<table width="600" border="1" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC" bordercolor="black" bordercolordark="white" bordercolorlight="black">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">����</td>
					<td bgcolor="#FFFFFF" width=100 nowrap>&nbsp;<b><%=db_JUBSEQ%></b>
					</td>
					<td bgcolor="#FFFFFF" width=100><%if db_EMERYN = "Y" then%><font color="#0000ff">&nbsp;���</font><%else%>&nbsp;<%end if%></td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">����Ͻ�</td>
					<td bgcolor="#FFFFFF" colspan='2' >&nbsp;<b><%=db_JUBTIME1%>(<%=JUBDAY%>)</b>
					</td>
				</tr>

			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��    ��</td>
					<td bgcolor="#FFFFFF">&nbsp;<b><% if db_SEXGB = "1" or SEXGB = "" then %>��<% else %>��<%end if%></b>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��    ��</td>
					<td bgcolor="#FFFFFF" >&nbsp;<b><%=db_CUSTNAME%></b>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��ȭ�ð�</td>
					<td bgcolor="#FFFFFF">&nbsp;<b><%=db_CALLTIMEDP%></b>
					</td>

				</tr>
				<tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">����ó1</td>
					<td bgcolor="#FFFFFF"  >&nbsp;<b><%=db_TELNO%></b></td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">����ó2</td>
					<td bgcolor="#FFFFFF"  >&nbsp;<b><%=db_TELNO2%></b></td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">�߽Ź�ȣ</td>
					<td bgcolor="#FFFFFF" width=100  >&nbsp;<b><%=db_CID%></b></td>
				</tr>

			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��    ��</td>
					<td bgcolor="#FFFFFF" colspan='3' nowrap>&nbsp;<b><%=db_getCateNameA_(db_SOSOKGB_A)%><%if db_getCateNameB_(db_SOSOKGB_A,db_SOSOKGB_B) <> "" then %>><%end if%><%=db_getCateNameB_(db_SOSOKGB_A,db_SOSOKGB_B)%><%if db_getCateNameC_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C) <> "" then %>><%end if%><%=db_getCateNameC_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C)%><%if db_getCateNameD_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C,db_SOSOKGB_D) <> "" then %>><%end if%><%=db_getCateNameD_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C,db_SOSOKGB_D)%><%if db_getCateNameE_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C,db_SOSOKGB_D,db_SOSOKGB_E) <> "" then %>><%end if%><%=db_getCateNameE_(db_SOSOKGB_A,db_SOSOKGB_B,db_SOSOKGB_C,db_SOSOKGB_D,db_SOSOKGB_E)%></td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��    ��</td>
					<td bgcolor="#FFFFFF" height=20 nowrap>&nbsp;<b><%=db_getCateNameB_("P",db_LEVEL_B)%><%if db_getCateNameC_("P",db_LEVEL_B,db_LEVEL_C) <> "" then %>><%end if%> <%=db_getCateNameC_("P",db_LEVEL_B,db_LEVEL_C)%><%if db_getCateNameD_("P",db_LEVEL_B,db_LEVEL_C,db_LEVEL_D) <> "" then %>><%end if%><%=db_getCateNameD_("P",db_LEVEL_B,db_LEVEL_C,db_LEVEL_D)%></b>
					</td>

				</tr>

			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center" >�������</td>
					<td bgcolor="#FFFFFF" ><%=db_GetCustNameA(db_REFERJUBSEQ)%></b>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center" colspan='1' nowrap>������ڿ��ǰ���</td>
					<td bgcolor="#FFFFFF" colspan='1'>&nbsp;<b><%=db_GetCodeName("C14",db_REQUESTERGB)%></b>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center" >�ļ�Ȯ��1</td>
					<td bgcolor="#FFFFFF" colspan='1'>&nbsp;<b><%=db_GetCodeName("C13",db_CALLKIND_B)%></b>
					</td>
				</tr>

			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">��㳻��</td>
					<td bgcolor="#FFFFFF" colspan=5 width=500>&nbsp;<b><%=db_QUESTION%></b>		
					</td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">Ư�̻���</td>
					<td bgcolor="#FFFFFF" colspan=5 width=500>&nbsp;<b><%=db_REMARK%></b>	
					</td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align="center">����</td>
					<td bgcolor="#FFFFFF" colspan=5 width=500>&nbsp;<b><%=db_getUserName(db_INCODE)%></b>&nbsp;&nbsp;(��)	
					</td>
				</tr>

			</table>

		</td>
	</tr>
</table>
			</div>




<form name="pf">
<input type=hidden name="printzone">
</form>
<script>
	
	
function print_info()
{
	document.pf.printzone.value=A.innerHTML;		
	window.open("/print_page.html","print_open","width=800,height=700,top=0,left=0,noresizable,toolbar=no,status=no,scrollbars=yes,directory=no");	
}
</script>