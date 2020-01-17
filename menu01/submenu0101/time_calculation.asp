<!-- #include virtual="/Include/Top2.asp" -->

<%
	On Error Resume next

	'선택한 값 찾아서   계산하기
	sDateControlName = request("DateControlName")
	sHourControlName = request("HourControlName")
	sMinControlName = request("MinControlName")
	RESERVETIME = request("RESERVETIME")

	if RESERVETIME = "1" then
		SQL = "select CONVERT(CHAR(19),DATEADD(n,10,GETDATE()),121)"
	elseif RESERVETIME = "2" then
		SQL = "select CONVERT(CHAR(19),DATEADD(n,30,GETDATE()),121)"
	elseif RESERVETIME = "3" then
		SQL = "select CONVERT(CHAR(19),DATEADD(n,60,GETDATE()),121)"
	elseif RESERVETIME = "4" then
		SQL = "select CONVERT(CHAR(19),DATEADD(n,120,GETDATE()),121)"
	end if

	set rs = db.execute(SQL)

	response.write SQL

	sDate = left(rs(0),10)
	sHOUR = mid(rs(0),12,2)
	sMIN = mid(rs(0),15,2)
		response.write "-----------------------"
		response.write rs(0)

%>
	
	<script>
		eval("<%=sDateControlName%>").value = "<%=sDate%>";
		eval("<%=sHourControlName%>").value = "<%=sHOUR%>";
		eval("<%=sMinControlName%>").value = "<%=sMIN%>";
	</script>
<!-- #include virtual="/Include/Bottom.asp" -->