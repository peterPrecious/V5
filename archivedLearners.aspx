<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="archivedLearners.aspx.cs" Inherits="V5.archivedLearners" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
  <title></title>
</head>
<body>
  <form id="form1" runat="server">
    <div>

      <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="no" DataSourceID="SqlDataSource1" ForeColor="#333333" GridLines="None">
        <AlternatingRowStyle BackColor="White" />
        <Columns>
          <asp:BoundField DataField="archived" HeaderText="archived" SortExpression="archived" />
          <asp:BoundField DataField="custId" HeaderText="custId" SortExpression="custId" />
          <asp:BoundField DataField="custAcctId" HeaderText="custAcctId" SortExpression="custAcctId" />
          <asp:BoundField DataField="custTitle" HeaderText="custTitle" SortExpression="custTitle" />
          <asp:BoundField DataField="custOrganization" HeaderText="custOrganization" SortExpression="custOrganization" />
          <asp:BoundField DataField="custParentId" HeaderText="custParentId" SortExpression="custParentId" />
          <asp:BoundField DataField="custCreated" HeaderText="custCreated" SortExpression="custCreated" />
          <asp:BoundField DataField="custExpired" HeaderText="custExpired" SortExpression="custExpired" />
          <asp:BoundField DataField="membId" HeaderText="membId" SortExpression="membId" />
          <asp:BoundField DataField="membFirstName" HeaderText="membFirstName" SortExpression="membFirstName" />
          <asp:BoundField DataField="membLastName" HeaderText="membLastName" SortExpression="membLastName" />
          <asp:BoundField DataField="membOrganization" HeaderText="membOrganization" SortExpression="membOrganization" />
          <asp:BoundField DataField="programId" HeaderText="programId" SortExpression="programId" />
          <asp:BoundField DataField="programTitle" HeaderText="programTitle" SortExpression="programTitle" />
          <asp:BoundField DataField="moduleId" HeaderText="moduleId" SortExpression="moduleId" />
          <asp:BoundField DataField="moduleTitle" HeaderText="moduleTitle" SortExpression="moduleTitle" />
          <asp:BoundField DataField="sessionTimeSpent" HeaderText="sessionTimeSpent" SortExpression="sessionTimeSpent" />
          <asp:BoundField DataField="sessionLastScore" HeaderText="sessionLastScore" SortExpression="sessionLastScore" />
          <asp:BoundField DataField="sessionLastAccessed" HeaderText="sessionLastAccessed" SortExpression="sessionLastAccessed" />
          <asp:BoundField DataField="sessionCompleted" HeaderText="sessionCompleted" SortExpression="sessionCompleted" />
        </Columns>
        <EditRowStyle BackColor="#2461BF" />
        <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
        <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
        <RowStyle BackColor="#EFF3FB" />
        <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
        <SortedAscendingCellStyle BackColor="#F5F7FB" />
        <SortedAscendingHeaderStyle BackColor="#6D95E1" />
        <SortedDescendingCellStyle BackColor="#E9EBEF" />
        <SortedDescendingHeaderStyle BackColor="#4870BE" />
      </asp:GridView>
      <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:appsConnectionString %>" SelectCommand="sp5archivedLearners" SelectCommandType="StoredProcedure">
        <SelectParameters>
          <asp:FormParameter DefaultValue="He" FormField="firstName" Name="firstName" Type="String" />
          <asp:FormParameter DefaultValue="Eg" FormField="lastName" Name="lastName" Type="String" />
        </SelectParameters>
      </asp:SqlDataSource>
    </div>

  </form>
</body>
</html>
