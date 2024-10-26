<%@ Import Namespace="System.Web.Security " %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>

<Html>
<Body>
<Form runat="server">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1" height="253">
  <tr>
    <td width="27%" height="27">　</td>
    <td width="44%" height="27">

<Html>
<Body>
<h3></h3>
<p></p>
<p></p>
<p></p>
<p>　</p>
</Body>
</Html>

    </td>
    <td width="29%" height="27">　</td>
  </tr>
  <tr>
    <td width="27%" height="231">　</td>
    <td width="44%" height="231">

<Html>
<Body>
<p>&nbsp;&nbsp;&nbsp;&nbsp;
<img border="0" src="images/經銷商帳號密碼2.gif"></p>
<Blockquote>

<p>帳號：<asp:TextBox runat="server" id="Account" /></p>
<p>
密碼：<asp:TextBox runat="server" id="Password" 
                  TextMode="Password" /><p>
記得我：<asp:CheckBox runat="server" id="RememberMe" /><p>

<asp:Button runat="server" text="登  入" OnClick="Login_Click" /><p>

<p>

<p>

<p>
<asp:Label runat="server" id="Msg" ForeColor="Red" />

</Blockquote><HR>
<p style="margin-top: 0; margin-bottom: 0">
<Font Size=-1 Color=Blue>登入網頁前, 請確定瀏覽器的 Cookie 是開啟的.</Font> </p>
</Body>
</Html>

<script language="VB" runat="server">

   Sub Login_Click( sender As Object, e As EventArgs )

      If Verify( Account.Text, Password.Text ) Then
         'FormsAuthentication.RedirectFromLoginPage(Account.Text, RememberMe.Checked)
         Server.Transfer("Sale-password.aspx")

      Else
         Msg.Text = "帳號或密碼錯誤, 請重新輸入!"
      End If
   End Sub

   Function Verify( 帳號 As String, 密碼 As String) As Boolean
      Dim Conn As OleDbConnection, Cmd As OleDbCommand
      Dim Rd As OleDbDataReader, SQL  As String

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "Sample.mdb" )
      Dim DbPass = "Jet OLEDB:Database Password=dyitcom33"
      Conn = New OleDbConnection( Provider & ";" & DataBase & ";" & DbPass )
      'Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      SQL = "Select * From SaleMe Where " & _
                    "UserID='" & 帳號 & "'" & _
                    " And Password='" & 密碼 & "'"
      Cmd = New OleDbCommand( SQL, Conn )
      Rd = Cmd.ExecuteReader()
      If Rd.Read() Then ' 表示有找到 UserID 及 Password, 通過驗證
         Conn.Close()
         Return True
      Else
         Conn.Close()
         Msg.Text = "帳號或密碼錯誤, 請重新輸入!"
         Return False
      End If
   End Function

    </script>
    <p>　</td>
    <td width="29%" height="231">　</td>
  </tr>
</table>
</Form>
</Body>
</Html>