<%@ Page Language="VB" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim conn As New OleDbConnection
        conn.ConnectionString = "Provider='sqloledb' ;" & _
                                           "Data Source='XP-NOTEBOOK1'; " & _
                                             "Initial Catalog='wsdscu';Integrated Security='SSPI';"
        conn.Open()

        Dim objCmd As New OleDbCommand()
        objCmd.Connection = conn
        objCmd.CommandText = "Select * from memmaster  where memno = '" & memno.Text & "' "
        
        Dim objReader As OleDbDataReader = objCmd.ExecuteReader()
        If objReader.Read() = True Then
            memname.Text = objReader.Item(1)
            monthsave.Text = objReader.Item(41)
            monthssave.Text = objReader.Item(42)
        End If
        

    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <div style="z-index: 101; left: 10px; width: 411px; position: absolute; top: 6px;
            height: 27px">
            社員編號 : &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
            <asp:TextBox ID="memno" runat="server" Width="117px"></asp:TextBox>
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Search" />
            <div style="z-index: 101; left: 0px; width: 596px; position: absolute; top: 27px;
                height: 222px">
                <table style="width: 594px">
                    <tr>
                        <td style="width: 100px">
                            姓名&nbsp;</td>
                        <td style="width: 370px">
        <asp:TextBox ID="memname" runat="server" Width="362px"></asp:TextBox></td>
                    </tr>
                    <tr>
                        <td style="width: 100px">
                            自動轉賬</td>
                        <td style="width: 370px">
                            <asp:TextBox ID="monthsave" runat="server"></asp:TextBox></td>
                    </tr>
                    <tr>
                        <td style="width: 100px">
                            庫房扣息</td>
                        <td style="width: 370px">
                            <asp:TextBox ID="monthssave" runat="server"></asp:TextBox></td>
                    </tr>
                    <tr>
                        <td style="width: 100px">
                        </td>
                        <td style="width: 370px">
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 100px">
                        </td>
                        <td style="width: 370px">
                        </td>
                    </tr>
                </table>
            </div>
        </div>
        <br />
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
    </div>
    </form>
</body>
</html>
