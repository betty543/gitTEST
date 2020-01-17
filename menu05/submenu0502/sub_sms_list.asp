<!-- #include virtual="/Include/Top_Frame.asp" -->
<%
	INCODE = SESSION("SS_LoginID")

	idx = request("idx")
	IsType = request("IsType")
	CellPhone = request("CellPhone")
	if IsType = "DEL" then
		SQL = "DELETE	FROM	temp_conference where userid = '" & INCODE &"' and datagb = '2' and idx = " & idx
		db.Execute(SQL)

	elseif IsType = "INS" then
		strSQL = "INSERT INTO temp_conference ( addr_idx, userid, cellphone, gunphone, datagb)" &_
						" values (0,'"& INCODE	& "', " &_
								"'" & CellPhone		& "','" & gunphone		& "','2')"
		db.Execute(strSQL)
	end if

%>

											  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" valign=top >
												<FORM name="Dwrite">
													<input type="hidden" name="IsType" value="">
													<input type="hidden" name="idx" value="">
<!-- 												 <tr>
												   <td>
													  <table width="100%" border="0" cellpadding="0" cellspacing="0"> -->

<%
strSQL2="select idx, cellphone from temp_conference where userid = '" & INCODE &"' and datagb = '2'"
set rs2 = db.Execute(strSQL2)
if rs2.EOF or rs2.BOF then
	NoData = True
Else	
	NoData = False
end if	

	if NoData = False Then
	arrTable = rs2.GetRows()
	for i=0 to Ubound(arrTable,2)
		idx		= arrTable(0,i)
		cellphone		= arrTable(1,i)

%>

														<tr bgcolor="D1D0CB">
														  <td height="1" colspan="6" align="center"></td>
														</tr>
														<tr bgcolor=#e9e9e9>
														  <td width="10%" height="20" align="center"><%=i+1%></td>
														  <td width="70%" align="center"><%=cellphone%></td>
														  <td width="20%" align="center">
							<img src="/Images/Btn/BtnIconDel.gif" title="전송대상에서삭제" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_del('<%=idx%>','D');"></td>
														</tr>



	<%
		 next 
	 Else 
	%>
														<tr bgcolor="D1D0CB">
														  <td height="1" align="center" colspan="6" ></td>
														</tr>
														<tr bgcolor="#e9e9e9">
														  <td width="100%" height="30" align="center" colspan="6" > 전송대상자가 없습니다.</td>
														</tr>
														<tr bgcolor="D1D0CB">
														  <td height="1" align="center" colspan="6" ></td>
														</tr>
	<%
	 end If
	 rs2.Close    
	 set rs2 = nothing   	
	%>
<!-- 													  </table>

												   </td>
												 </tr> -->
											   </form>
<%
strSql2="Select count(idx) as total_count from temp_conference where userid = '" & INCODE &"' and datagb = '2'"
set rs2 = db.Execute(strSQL2)
	total_count = rs2("total_count")
rs2.close
Set rs2 = Nothing %>
														<tr bgcolor="D1D0CB">
														  <td height="1" colspan="6" align="center"></td>
														</tr>
														<tr>
														  <td width="100%" height="20" align="center" colspan="6"> 총 데이타 수 : <%=total_count%></td>
														</tr>
											   </table>
											   <form name="total_count_form">
											   <input type="hidden" name="total_count" value="<%=total_count%>">
											   </form>


<script>

	function fn_del(arg0,arg1){
		document.Dwrite.idx.value = arg0;
		document.Dwrite.IsType.value = 'DEL';
		document.Dwrite.submit();
	}
</script>