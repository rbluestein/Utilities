<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ErrorPage.aspx.vb" Inherits="ErrorPage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>BVI Directory</title>
    <link href="css/BVI.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
			<table class="PrimaryTbl" style="Z-INDEX: 100; LEFT: 0px; POSITION: absolute; TOP: 0px"
				cellspacing="0" cellpadding="0" width="650" border="0">
				<tr>
					<td class="PrimaryTblTitle">An error has occurred:</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td class="Cell9Reg">
						<asp:literal id="litError" runat="server" enableviewstate="false"></asp:literal>
					</td>
				</tr>
				<tr>
					<td><asp:placeholder id="plClock" runat="server">
							<asp:image id="imgClock" runat="server" height="100px" imageurl="img/Clock.png"></asp:image>
						</asp:placeholder></td>
				</tr>
			</table>      
    </div>
    </form>
</body>
</html>
