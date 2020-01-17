<!-- #include virtual="/include/top_frame.asp" -->
<%

'1. 파라미터 얻어오기
curPage = request("curPage")

SS_Login_Secgroup = SESSION("SS_Login_Secgroup")
SS_Login_Grade = SESSION("SS_Login_Grade")
SS_LoginID = SESSION("SS_LoginID")

'2. 쿼리조건절 셋팅
pageSize = 15
pageSector = 10
if curPage = "" then curPage = 1 end If
where1 = "a=a"
where2 = "curPage=" & curPage & "&" & where1

sql_tb = "TB_USERINFO"
'sql_index = "index_desc(" & sql_tb & " IDX_TB_MANUAL_MANUALSEQ)"
sql_field = "USERID, USERNAME, SECGROUP, GRADE, USEYN, SOSOK, [LEVEL], CTIID,EXTNO, PASSWORD"
sql_orderby = "USEYN desc,USERID"

sql_where = " 1=1 "
if SS_Login_Grade <> "A" then
	sql_where = sql_where & "	and GRADE = '" & SS_Login_Grade &"'"
end if
if SS_Login_Secgroup = "A" then
	sql_where = sql_where & "	and USERID = '" & SS_LoginID &"'"
end if

'3. 쿼리 실행
sql = db_getSqlWithPage(sql_tb, sql_index, sql_field, sql_where, sql_orderby, pageSize, curPage)
set rs = db.execute(sql)

'4. Paging HTML 작성
totalCount = db_getCount(db, sql_tb, sql_where)
startRow = totalCount - pageSize * (curPage - 1)
pageHtml = getPageHtml(pageSector, pageSize, totalCount, curPage, currentURL & "?" & where1)

%>

			
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk">◈ <b>사용자 리스트</b></td></tr>
        	</table>
        	<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr height="20" bgcolor="#EEF6FF" align="center">
        			<td>NO</td>
        			<td>아이디</td>
        			<td>소속</td>
        			<td>계급</td>

        			<td>성명</td>
        			<td>비밀번호</td>
        			<td>운용업무</td>
        			<td>권한</td>
        			<td>CTIID</td>
        			<td>사용여부</td>
        		</tr>
        		<tr><td colspan="10" height="1" bgcolor="#FFFFFF"></td></tr>

<% 
	if rs.EOF and rs.BOF then 
%>

	<tr>
		<td height="50" colspan="50" bgcolor="#FFFFFF">
			<p align="center">검색된 자료가 없습니다.</p>
		</td>
	</tr>

<%
	else

		do until rs.EOF
		
%>
        		<tr height="20" bgcolor="#FFFFFF" onClick="parent.DetailFrame.location.href='User_Detail.asp?guboon=UP&userid=<%=rs("USERID")%>';" onmouseover="setSelectColor(this);" onmouseout="setOutColor(this);" style="cursor:hand">
        			<td align="center"><%=startRow%></td>
        			<td align="center"><%=rs("USERID")%></td>
        			<td align="center" width="100"><%=db_getCodeName("C04", rs("SOSOK"))%></td>
        			<td align="center"><%=db_getCodeName("Z05", rs("LEVEL"))%></td>

        			<td align="center"><%=rs("USERNAME")%></td>
        			<td align="center"><%=rs("PASSWORD")%></td>
        			<td align="center"><%=db_getCodeName("Z04", rs("GRADE"))%></td>
        			<td align="center"><%=db_getCodeName("Z02", rs("SECGROUP"))%></td>

        			<td align="center"><%=rs("CTIID")%>-<%=rs("EXTNO")%></td>

        			<td align="center"><%=rs("USEYN")%></td>
        		</tr>

<%
			startRow = startRow - 1
			rs.MoveNext 
		Loop
		
		rs.close 
		set rs = Nothing
		
	end if
%>  

        	</table>
        	


			<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
				<tr><td height="5"></td></tr>
				<tr><td height="1" bgcolor="#D6D6D6"></td></tr>
				<tr height="22" bgcolor="#EEF6FF">
					<td align="center"><%=pageHtml%></td>
				</tr>
				<tr><td height="1" bgcolor="#D6D6D6"></td></tr>
			</table>
        	
			<table border="0" cellspacing="0" width="100%" align="center">
				<tr height="30">
					<td align="right"><%if SS_Login_Secgroup ="A" or SS_Login_Secgroup ="B" then%><%else%><img src="/Images/Btn/BtnUserAdd.gif" style="cursor:hand;" align="absmiddle" onClick="parent.DetailFrame.location.href='User_Detail.asp?guboon=INS';"><%end if%></td>
				</tr>
			</table>