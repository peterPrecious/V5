<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="certLogs.aspx.cs" Inherits="ecomLogs._default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
  <title>certLogs</title>
  <script src="//code.jquery.com/jquery-1.12.0.min.js"></script>
  <script src="//code.jquery.com/jquery-migrate-1.2.1.min.js"></script>
  <script>
    $(function () {
      $("#edit,#hist").hide();
    });

    function edit(svMembNo, sesId, membNo, modsNo, progNo) {
      $("#edit,#hist").hide();
      $("#ifrEdit")[0].src = "//vubiz.com/Gold/vuSCORMAdmin/SessionQuickEdit.aspx?MembNo=" + svMembNo + "&SessionID=" + sesId + "&memberID=" + membNo + "&moduleID=" + modsNo + "&programID=" + progNo;
      $("#edit").show();
    };

    function hist(svMembNo, membNo) {
      $("#edit,#hist").hide();
      $("#ifrHist")[0].src = "//vubiz.com/Gold/vuSCORMAdmin/Default.aspx?MembNo=" + svMembNo + "&memberID=" + membNo;
      $("#hist").show();
    };

  </script>

  <style>
    #edit { position: absolute; top:50px; left: 50px; background-color: white; text-align: right; border: 1px solid navy; width: 275px; height: 470px; }

    #hist { position: absolute; top:50px; left: 50px; background-color: white; text-align: right; border: 1px solid navy; width: 90%; height: 1000px; }
  </style>
</head>
<body>
  <form id="form1" runat="server">
    <div>
      <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:appsConnectionString %>" SelectCommand="SELECT * FROM [vEcomLogs]"></asp:SqlDataSource>
      <h3 style="margin: 50px 0 30px 0; text-align: center; color: #2868c6">Certificates Generated without Module Completion</h3>

      <asp:GridView
        ID="GridView1" runat="server"
        HorizontalAlign="Center"
        AutoGenerateColumns="False"
        DataSourceID="SqlDataSource1"
        CellPadding="4"
        ForeColor="#333333"
        GridLines="None" AllowSorting="True">
        <AlternatingRowStyle BackColor="White" />
        <Columns>
          <asp:BoundField DataField="issued" HeaderText="Issued" SortExpression="issued" />
          <asp:BoundField DataField="custId" HeaderText="Cust Id" SortExpression="custId" />
          <asp:BoundField DataField="membId" HeaderText="Memb Id" SortExpression="membId" />
          <asp:BoundField DataField="membName" HeaderText="Name" SortExpression="membName" />
          <asp:BoundField DataField="progId" HeaderText="Prog Id" SortExpression="progId" />
          <asp:BoundField DataField="modsId" HeaderText="Mods Id" SortExpression="modsId" />
          <asp:BoundField DataField="score" HeaderText="Score" SortExpression="score" />
          <asp:BoundField DataField="ip" HeaderText="IP" SortExpression="ip" />
          <asp:TemplateField>
            <ItemTemplate><a href="#" onclick="edit(<%=svMembNo %>,<%#Eval("sesId")%>,<%#Eval("membNo")%>,<%#Eval("modsNo")%>,<%#Eval("progNo")%>)">Edit</a></ItemTemplate>
          </asp:TemplateField>
          <asp:TemplateField>
            <ItemTemplate><a href="#" onclick="hist(<%=svMembNo %>,<%#Eval("membNo")%>)">Hist</a></ItemTemplate>
          </asp:TemplateField>
        </Columns>
        <EditRowStyle BackColor="#2461BF" />
        <FooterStyle BackColor="#4c8be8" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#4c8be8" Font-Bold="True" ForeColor="White" Font-Size="Smaller" HorizontalAlign="Left" />
        <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
        <RowStyle BackColor="#EFF3FB" Font-Size="Smaller" />
        <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
        <SortedAscendingCellStyle BackColor="#F5F7FB" />
        <SortedAscendingHeaderStyle BackColor="#6D95E1" />
        <SortedDescendingCellStyle BackColor="#E9EBEF" />
        <SortedDescendingHeaderStyle BackColor="#4870BE" />
      </asp:GridView>

      <div id="edit">
        <input style="float: right; margin: 5px; color: red; font-weight: bold;" type="button" onclick="$('#edit').hide()" value="X" name="bClose" class="button" />
        <iframe class="div" id="ifrEdit" name="ifrEdit" style="width: 100%; height: 100%; border: 0"></iframe>
      </div>

      <div id="hist">
        <input style="float: right; margin: 5px; color: red; font-weight: bold;" type="button" onclick="$('#hist').hide()" value="X" name="bClose" class="button" />
        <iframe class="div" id="ifrHist" name="ifrHist" style="width: 100%; height: 100%; border: 0"></iframe>
      </div>


    </div>
  </form>
</body>
</html>

