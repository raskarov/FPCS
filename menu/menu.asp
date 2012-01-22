<%@ Language=VBScript %>
<?xml version='1.0'?>
		<menu site="FPCS" subsite="Main">
			<submenu handle="student">
				<item href="<% = Application.Value("strSSLWebRoot") %>Forms/SIS/default.asp" label="SIS" />
				<item href="<% = Application.Value("strSSLWebRoot") %>Forms/ILP/viewILP.asp" label="ILP" />
			</submenu>			<submenu handle="teacher">
				<item href="<% = Application.Value("strSSLWebRoot") %>Forms/Teachers/addTeacher.asp" label="Add a Teacher" />
			</submenu>			<submenu handle="vendor">
				<item href="<% = Application.Value("strSSLWebRoot") %>Forms/VIS/vendorAdmin.asp" label="Vendor Admin" />
			</submenu>			<submenu handle="reports">
				<item href="<% = Application.Value("strSSLWebRoot") %>Reports/teacherPerDiem.asp" label="Per Diem" />
			</submenu>
			<submenu handle="admin">
				<item href="javascript:jfURLopen('<% = Application.Value("strSSLWebRoot") %>admin/reset.asp', 400, 210, 1, 1, 'lock', 'refreshDD');" label="Reset Drop Downs" />
				<item href="javascript:jfURLopen('<% = Application.Value("strSSLWebRoot") %>dev/debugAppVar.asp', 800, 700, 1, 1, 'loose', 'debugAppVar');" label="Show Session and App Variables" />
			</submenu>		
		</menu>