<%@ Page Language="VB" %>
<script runat="server">

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        response.redirect("http://yahoo.com.hk")
    End Sub

</script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
        <div>&nbsp;<asp:TextBox id="TextBox1" runat="server"></asp:TextBox>
            <asp:Button id="Button1" runat="server" Text="Search"></asp:Button>
            <br />
            <asp:FormView id="FormView1" runat="server" GridLines="Horizontal" DataSourceID="SqlDataSource1" DataKeyNames="memNo" CellPadding="4" BorderWidth="3px" BorderStyle="Double" BorderColor="#336666" BackColor="White" Height="671px" Width="1553px">
                <FooterStyle backcolor="White" forecolor="#333333"></FooterStyle>
                <EditRowStyle backcolor="#339966" forecolor="White" font-bold="True"></EditRowStyle>
                <EditItemTemplate>
                    memNo: <asp:Label id="memNoLabel1" runat="server" text='<%# Eval("memNo") %>'></asp:Label> memName: 
                    <asp:TextBox ID="memNameTextBox" runat="server" Text='<%# Bind("memName") %>'></asp:TextBox>
                    <br />
                    memAddr1: 
                    <asp:TextBox ID="memAddr1TextBox" runat="server" Text='<%# Bind("memAddr1") %>'></asp:TextBox>
                    <br />
                    memAddr2: 
                    <asp:TextBox ID="memAddr2TextBox" runat="server" Text='<%# Bind("memAddr2") %>'></asp:TextBox>
                    <br />
                    memAddr3: 
                    <asp:TextBox ID="memAddr3TextBox" runat="server" Text='<%# Bind("memAddr3") %>'></asp:TextBox>
                    <br />
                    memContactTel: 
                    <asp:TextBox ID="memContactTelTextBox" runat="server" Text='<%# Bind("memContactTel") %>'></asp:TextBox>
                    <br />
                    memMobile: 
                    <asp:TextBox ID="memMobileTextBox" runat="server" Text='<%# Bind("memMobile") %>'></asp:TextBox>
                    <br />
                    memHKID: 
                    <asp:TextBox ID="memHKIDTextBox" runat="server" Text='<%# Bind("memHKID") %>'></asp:TextBox>
                    <br />
                    memGender: 
                    <asp:TextBox ID="memGenderTextBox" runat="server" Text='<%# Bind("memGender") %>'></asp:TextBox>
                    <br />
                    memBday: 
                    <asp:TextBox ID="memBdayTextBox" runat="server" Text='<%# Bind("memBday") %>'></asp:TextBox>
                    <br />
                    memGrade: 
                    <asp:TextBox ID="memGradeTextBox" runat="server" Text='<%# Bind("memGrade") %>'></asp:TextBox>
                    <br />
                    memSection: 
                    <asp:TextBox ID="memSectionTextBox" runat="server" Text='<%# Bind("memSection") %>'></asp:TextBox>
                    <br />
                    memGuarantorNo: 
                    <asp:TextBox ID="memGuarantorNoTextBox" runat="server" Text='<%# Bind("memGuarantorNo") %>'></asp:TextBox>
                    <br />
                    memGuarantor4Others: 
                    <asp:TextBox ID="memGuarantor4OthersTextBox" runat="server" Text='<%# Bind("memGuarantor4Others") %>'></asp:TextBox>
                    <br />
                    treasRefNo: 
                    <asp:TextBox ID="treasRefNoTextBox" runat="server" Text='<%# Bind("treasRefNo") %>'></asp:TextBox>
                    <br />
                    employCond: 
                    <asp:TextBox ID="employCondTextBox" runat="server" Text='<%# Bind("employCond") %>'></asp:TextBox>
                    <br />
                    firstAppointDate: 
                    <asp:TextBox ID="firstAppointDateTextBox" runat="server" Text='<%# Bind("firstAppointDate") %>'></asp:TextBox>
                    <br />
                    memDate: 
                    <asp:TextBox ID="memDateTextBox" runat="server" Text='<%# Bind("memDate") %>'></asp:TextBox>
                    <br />
                    autopayAmt: 
                    <asp:TextBox ID="autopayAmtTextBox" runat="server" Text='<%# Bind("autopayAmt") %>'></asp:TextBox>
                    <br />
                    autopayPerm: 
                    <asp:TextBox ID="autopayPermTextBox" runat="server" Text='<%# Bind("autopayPerm") %>'></asp:TextBox>
                    <br />
                    salaryDedut: 
                    <asp:TextBox ID="salaryDedutTextBox" runat="server" Text='<%# Bind("salaryDedut") %>'></asp:TextBox>
                    <br />
                    loanRepaid: 
                    <asp:TextBox ID="loanRepaidTextBox" runat="server" Text='<%# Bind("loanRepaid") %>'></asp:TextBox>
                    <br />
                    calcInterest: 
                    <asp:CheckBox ID="calcInterestCheckBox" runat="server" Checked='<%# Bind("calcInterest") %>' />
                    <br />
                    OSInterest: 
                    <asp:TextBox ID="OSInterestTextBox" runat="server" Text='<%# Bind("OSInterest") %>'></asp:TextBox>
                    <br />
                    thisInterest: 
                    <asp:TextBox ID="thisInterestTextBox" runat="server" Text='<%# Bind("thisInterest") %>'></asp:TextBox>
                    <br />
                    leagueDue: 
                    <asp:CheckBox ID="leagueDueCheckBox" runat="server" Checked='<%# Bind("leagueDue") %>' />
                    <br />
                    bankAccNo: 
                    <asp:TextBox ID="bankAccNoTextBox" runat="server" Text='<%# Bind("bankAccNo") %>'></asp:TextBox>
                    <br />
                    personEntitled: 
                    <asp:TextBox ID="personEntitledTextBox" runat="server" Text='<%# Bind("personEntitled") %>'></asp:TextBox>
                    <br />
                    overdue: 
                    <asp:TextBox ID="overdueTextBox" runat="server" Text='<%# Bind("overdue") %>'></asp:TextBox>
                    <br />
                    <br />
                    deleted: 
                    <asp:CheckBox ID="deletedCheckBox" runat="server" Checked='<%# Bind("deleted") %>' />
                    <br />
                    <asp:LinkButton ID="UpdateButton" runat="server" CausesValidation="True" CommandName="Update" Text="Update"></asp:LinkButton>
                    <asp:LinkButton ID="UpdateCancelButton" runat="server" CausesValidation="False" CommandName="Cancel" Text="Cancel"></asp:LinkButton>
                </EditItemTemplate>
                <RowStyle backcolor="White" forecolor="#333333"></RowStyle>
                <PagerStyle backcolor="#336666" forecolor="White" horizontalalign="Center"></PagerStyle>
                <InsertItemTemplate>
                    memNo: 
                    <asp:TextBox ID="memNoTextBox" runat="server" Text='<%# Bind("memNo") %>'></asp:TextBox>< />
                    memName: 
                    <asp:TextBox ID="memNameTextBox" runat="server" Text='<%# Bind("memName") %>'></asp:TextBox>
                    <br />
                    memAddr1: 
                    <asp:TextBox ID="memAddr1TextBox" runat="server" Text='<%# Bind("memAddr1") %>'></asp:TextBox>< />
                    memAddr2: 
                    <asp:TextBox ID="memAddr2TextBox" runat="server" Text='<%# Bind("memAddr2") %>'></asp:TextBox>
                    <br />
                    memAddr3: 
                    <asp:TextBox ID="memAddr3TextBox" runat="server" Text='<%# Bind("memAddr3") %>'></asp:TextBox>< />
                    memContactTel: 
                    <asp:TextBox ID="memContactTelTextBox" runat="server" Text='<%# Bind("memContactTel") %>'></asp:TextBox>
                    <br />
                    memMobile: 
                    <asp:TextBox ID="memMobileTextBox" runat="server" Text='<%# Bind("memMobile") %>'></asp:TextBox>< />
                    memHKID: 
                    <asp:TextBox ID="memHKIDTextBox" runat="server" Text='<%# Bind("memHKID") %>'></asp:TextBox>
                    <br />
                    memGender: 
                    <asp:TextBox ID="memGenderTextBox" runat="server" Text='<%# Bind("memGender") %>'></asp:TextBox>< />
                    memBday: 
                    <asp:TextBox ID="memBdayTextBox" runat="server" Text='<%# Bind("memBday") %>'></asp:TextBox>
                    <br />
                    memGrade: 
                    <asp:TextBox ID="memGradeTextBox" runat="server" Text='<%# Bind("memGrade") %>'></asp:TextBox>< />
                    memSection: 
                    <asp:TextBox ID="memSectionTextBox" runat="server" Text='<%# Bind("memSection") %>'></asp:TextBox>
                    <br />
                    memGuarantorNo: 
                    <asp:TextBox ID="memGuarantorNoTextBox" runat="server" Text='<%# Bind("memGuarantorNo") %>'></asp:TextBox>< />
                    memGuarantor4Others: 
                    <asp:TextBox ID="memGuarantor4OthersTextBox" runat="server" Text='<%# Bind("memGuarantor4Others") %>'></asp:TextBox>
                    <br />
                    treasRefNo: 
                    <asp:TextBox ID="treasRefNoTextBox" runat="server" Text='<%# Bind("treasRefNo") %>'></asp:TextBox>< />
                    employCond: 
                    <asp:TextBox ID="employCondTextBox" runat="server" Text='<%# Bind("employCond") %>'></asp:TextBox>
                    <br />
                    firstAppointDate: 
                    <asp:TextBox ID="firstAppointDateTextBox" runat="server" Text='<%# Bind("firstAppointDate") %>'></asp:TextBox>< />
                    memDate: 
                    <asp:TextBox ID="memDateTextBox" runat="server" Text='<%# Bind("memDate") %>'></asp:TextBox>
                    <br />
                    autopayAmt: 
                    <asp:TextBox ID="autopayAmtTextBox" runat="server" Text='<%# Bind("autopayAmt") %>'></asp:TextBox>< />
                    autopayPerm: 
                    <asp:TextBox ID="autopayPermTextBox" runat="server" Text='<%# Bind("autopayPerm") %>'></asp:TextBox>
                    <br />
                    salaryDedut: 
                    <asp:TextBox ID="salaryDedutTextBox" runat="server" Text='<%# Bind("salaryDedut") %>'></asp:TextBox>< />
                    loanRepaid: 
                    <asp:TextBox ID="loanRepaidTextBox" runat="server" Text='<%# Bind("loanRepaid") %>'></asp:TextBox>
                    <br />
                    calcInterest: 
                    <asp:CheckBox ID="calcInterestCheckBox" runat="server" Checked='<%# Bind("calcInterest") %>' />
                    <br />
                    OSInterest: 
                    <asp:TextBox ID="OSInterestTextBox" runat="server" Text='<%# Bind("OSInterest") %>'></asp:TextBox>
                    <br />
                    thisInterest: 
                    <asp:TextBox ID="thisInterestTextBox" runat="server" Text='<%# Bind("thisInterest") %>'></asp:TextBox>
                    <br />
                    leagueDue: 
                    <asp:CheckBox ID="leagueDueCheckBox" runat="server" Checked='<%# Bind("leagueDue") %>' />
                    <br />
                    bankAccNo: 
                    <asp:TextBox ID="bankAccNoTextBox" runat="server" Text='<%# Bind("bankAccNo") %>'></asp:TextBox>
                    <br />
                    personEntitled: 
                    <asp:TextBox ID="personEntitledTextBox" runat="server" Text='<%# Bind("personEntitled") %>'></asp:TextBox>
                    <br />
                    overdue: 
                    <asp:TextBox ID="overdueTextBox" runat="server" Text='<%# Bind("overdue") %>'></asp:TextBox>
                    <br />
                    ttlShare: 
                    <asp:TextBox ID="ttlShareTextBox" runat="server" Text='<%# Bind("ttlShare") %>'></asp:TextBox>
                    <br />
                    ttlLastShare: 
                    <asp:TextBox ID="ttlLastShareTextBox" runat="server" Text='<%# Bind("ttlLastShare") %>'></asp:TextBox>
                    <br />
                    deleted: 
                    <asp:CheckBox ID="deletedCheckBox" runat="server" Checked='<%# Bind("deleted") %>' />
                    <br />
                    <asp:LinkButton ID="InsertButton" runat="server" CausesValidation="True" CommandName="Insert" Text="Insert"></asp:LinkButton>
                    <asp:LinkButton ID="InsertCancelButton" runat="server" CausesValidation="False" CommandName="Cancel" Text="Cancel"></asp:LinkButton>
                </InsertItemTemplate>
                <ItemTemplate>
                    memNo: <asp:Label id="memNoLabel" runat="server" text='<%# Eval("memNo") %>'></asp:Label> 
                    <br />
                    memName: <asp:Label id="memNameLabel" runat="server" text='<%# Bind("memName") %>'></asp:Label> 
                    <br />
                    memAddr1: <asp:Label id="memAddr1Label" runat="server" text='<%# Bind("memAddr1") %>'></asp:Label> 
                    <br />
                    memAddr2: <asp:Label id="memAddr2Label" runat="server" text='<%# Bind("memAddr2") %>'></asp:Label> 
                    <br />
                    memAddr3: <asp:Label id="memAddr3Label" runat="server" text='<%# Bind("memAddr3") %>'></asp:Label> 
                    <br />
                    memContactTel: <asp:Label id="memContactTelLabel" runat="server" text='<%# Bind("memContactTel") %>'></asp:Label> 
                    <br />
                    memMobile: <asp:Label id="memMobileLabel" runat="server" text='<%# Bind("memMobile") %>'></asp:Label> 
                    <br />
                    memHKID: <asp:Label id="memHKIDLabel" runat="server" text='<%# Bind("memHKID") %>'></asp:Label> 
                    <br />
                    memGender: <asp:Label id="memGenderLabel" runat="server" text='<%# Bind("memGender") %>'></asp:Label> 
                    <br />
                    memBday: <asp:Label id="memBdayLabel" runat="server" text='<%# Bind("memBday") %>'></asp:Label> 
                    <br />
                    memGrade: <asp:Label id="memGradeLabel" runat="server" text='<%# Bind("memGrade") %>'></asp:Label> 
                    <br />
                    memSection: <asp:Label id="memSectionLabel" runat="server" text='<%# Bind("memSection") %>'></asp:Label> 
                    <br />
                    memGuarantorNo: <asp:Label id="memGuarantorNoLabel" runat="server" text='<%# Bind("memGuarantorNo") %>'></asp:Label> 
                    <br />
                    memGuarantor4Others: <asp:Label id="memGuarantor4OthersLabel" runat="server" text='<%# Bind("memGuarantor4Others") %>'></asp:Label> 
                    <br />
                    treasRefNo: <asp:Label id="treasRefNoLabel" runat="server" text='<%# Bind("treasRefNo") %>'></asp:Label> 
                    <br />
                    employCond: <asp:Label id="employCondLabel" runat="server" text='<%# Bind("employCond") %>'></asp:Label> 
                    <br />
                    firstAppointDate: <asp:Label id="firstAppointDateLabel" runat="server" text='<%# Bind("firstAppointDate") %>'></asp:Label> 
                    <br />
                    memDate: <asp:Label id="memDateLabel" runat="server" text='<%# Bind("memDate") %>'></asp:Label> 
                    <br />
                    autopayAmt: <asp:Label id="autopayAmtLabel" runat="server" text='<%# Bind("autopayAmt") %>'></asp:Label> 
                    <br />
                    autopayPerm: <asp:Label id="autopayPermLabel" runat="server" text='<%# Bind("autopayPerm") %>'></asp:Label> 
                    <br />
                    salaryDedut: <asp:Label id="salaryDedutLabel" runat="server" text='<%# Bind("salaryDedut") %>'></asp:Label> 
                    <br />
                    loanRepaid: <asp:Label id="loanRepaidLabel" runat="server" text='<%# Bind("loanRepaid") %>'></asp:Label> 
                    <br />
                    calcInterest: 
                    <asp:CheckBox ID="calcInterestCheckBox" runat="server" Checked='<%# Bind("calcInterest") %>' Enabled="false" />
                    <br />
                    OSInterest: <asp:Label id="OSInterestLabel" runat="server" text='<%# Bind("OSInterest") %>'></asp:Label> 
                    <br />
                    thisInterest: <asp:Label id="thisInterestLabel" runat="server" text='<%# Bind("thisInterest") %>'></asp:Label> 
                    <br />
                    leagueDue: 
                    <asp:CheckBox ID="leagueDueCheckBox" runat="server" Checked='<%# Bind("leagueDue") %>' Enabled="false" />
                    <br />
                    bankAccNo: <asp:Label id="bankAccNoLabel" runat="server" text='<%# Bind("bankAccNo") %>'></asp:Label> 
                    <br />
                    personEntitled: <asp:Label id="personEntitledLabel" runat="server" text='<%# Bind("personEntitled") %>'></asp:Label> 
                    <br />
                    overdue: <asp:Label id="overdueLabel" runat="server" text='<%# Bind("overdue") %>'></asp:Label> 
                    <br />
                    ttlShare: <asp:Label id="ttlShareLabel" runat="server" text='<%# Bind("ttlShare") %>'></asp:Label> 
                    <br />
                    ttlLastShare: <asp:Label id="ttlLastShareLabel" runat="server" text='<%# Bind("ttlLastShare") %>'></asp:Label> 
                    <br />
                    dividend: <asp:Label id="dividendLabel" runat="server" text='<%# Bind("dividend") %>'></asp:Label> 
                    <br />
                    deleted: 
                    <asp:CheckBox ID="deletedCheckBox" runat="server" Checked='<%# Bind("deleted") %>' Enabled="false" />
                    <br />
                </ItemTemplate>
                <HeaderStyle backcolor="#336666" forecolor="White" font-bold="True"></HeaderStyle>
                <HeaderTemplate>
                    Member Maintenance 
                </HeaderTemplate>
            </asp:FormView>
            <br />
            <asp:SqlDataSource id="SqlDataSource1" runat="server" SelectCommand="SELECT * FROM [memMaster] WHERE ([memNo] = @memNo)" ConnectionString="<%$ ConnectionStrings:emsdcuConnectionString %>">
                <SelectParameters>
                    <asp:FormParameter DefaultValue="0" FormField="textbox1" Name="memNo" Type="Int32" />
                </SelectParameters>
            </asp:SqlDataSource>
            <asp:GridView id="GridView1" runat="server" GridLines="None" DataSourceID="SqlDataSource2" CellPadding="4" ForeColor="#333333" AutoGenerateColumns="False" AllowPaging="True">
                <FooterStyle backcolor="#990000" font-bold="True" forecolor="White" />
                <Columns>
                    <asp:BoundField DataField="txDate" HeaderText="txDate" SortExpression="txDate" />
                    <asp:BoundField DataField="sharePaid" HeaderText="sharePaid" SortExpression="sharePaid" />
                    <asp:BoundField DataField="monthlyRepaid" HeaderText="monthlyRepaid" SortExpression="monthlyRepaid" />
                    <asp:BoundField DataField="interestPaid" HeaderText="interestPaid" SortExpression="interestPaid" />
                    <asp:BoundField DataField="txAmt" HeaderText="txAmt" SortExpression="txAmt" />
                    <asp:CheckBoxField DataField="deleted" HeaderText="deleted" SortExpression="deleted" />
                </Columns>
                <RowStyle backcolor="#FFFBD6" forecolor="#333333" />
                <SelectedRowStyle backcolor="#FFCC66" font-bold="True" forecolor="Navy" />
                <PagerStyle backcolor="#FFCC66" forecolor="#333333" horizontalalign="Center" />
                <HeaderStyle backcolor="#990000" font-bold="True" forecolor="White" />
                <AlternatingRowStyle backcolor="White" />
            </asp:GridView>
            <asp:Button id="Submit" onclick="Button2_Click" runat="server" Text="Button"></asp:Button>
            <asp:SqlDataSource id="SqlDataSource2" runat="server" SelectCommand="SELECT txDate, sharePaid, monthlyRepaid, interestPaid, txAmt, deleted FROM dbo.memTx WHERE (memNo = @memNo)" ConnectionString="<%$ ConnectionStrings:emsdcuConnectionString2 %>">
                <SelectParameters>
                    <asp:FormParameter DefaultValue="0" FormField="textbox1" Name="memNo" Type="Int32" />
                </SelectParameters>
            </asp:SqlDataSource>
        </div>
    </form>
</body>
</html>
