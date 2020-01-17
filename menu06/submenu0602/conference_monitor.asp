<!-- #include virtual="/Include/Top_Frame.asp" -->
<%
			SS_LoginID = SESSION("SS_LoginID")

			SQL = "	select a.idx as idx,sosok_name, class, name, c.cellphone,c.gunphone,processstep,successflag,datediff(second,stepdate,getdate()) as processseconds"
			SQL = SQL & "	from	temp_conference c, TB_SMSADDR a"
			SQL = SQL & "	where	addr_idx = a.idx and userid = '" & SS_LoginID & "' and datagb = '1' order by c.idx"

i = 0
j = 0
k = 0
l = 0
			set RS2 = db.execute(SQL)

			do until RS2.eof
				i = i + 1
				idx = RS2("idx")
				processstep = RS2("processstep")
				processseconds = RS2("processseconds")
				successflag = RS2("successflag")

				if successflag = "0" then
					successflagname = "대기"		
					processstep = ""
					forecolor ="#000000"
				elseif successflag = "1" then		'성공	
					successflagname = "성공"
					j = j + 1	
					forecolor ="#000000"
				elseif successflag = "2" then		'진행중		
					successflagname = "진행중"						
					k = k + 1
					forecolor ="#0000ff"
				else				'실패
					successflagname = "실패"
					l = l + 1			
					processseconds = ""
					forecolor ="#ff000"
				end if			
				if processstep = "종료" then
					processseconds = "" 
				end if
%>
		<script>
			eval("parent.document.all.panresult_<%=idx%>").innerHTML ="<b><font color='<%=forecolor%>' size='3px'><%=successflagname%></font></b>";
			eval("parent.document.all.successflag_<%=idx%>").value = "<%=successflagname%>"; // 상태		
			eval("parent.document.all.result_<%=idx%>").value = "<%=processstep%>"; // 상태
			eval("parent.document.all.time_<%=idx%>").value = "<%=processseconds%>"; //초
		</script>
<%
				RS2.movenext

			loop

%>
<script>
	parent.document.all.cnt1.value = "<%=i%>";	// 총건수
	parent.document.all.cnt2.value = "<%=j%>";	// 총건수
	parent.document.all.cnt3.value = "<%=k%>";	// 총건수
	parent.document.all.cnt4.value = "<%=l%>";	// 총건수
</script>
