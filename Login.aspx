<%@ Import Namespace="System.Web.Security " %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>

<Html>
<Body>
<Form runat="server">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1" height="253">
  <tr>
    <td width="27%" height="27">�@</td>
    <td width="44%" height="27">

<Html>
<Body>
<h3></h3>
<p></p>
<p></p>
<p></p>
<p>�@</p>
</Body>
</Html>

    </td>
    <td width="29%" height="27">�@</td>
  </tr>
  <tr>
    <td width="27%" height="231">�@</td>
    <td width="44%" height="231">

<Html>
<Body>
<p>&nbsp;&nbsp;&nbsp;&nbsp;
<img border="0" src="images/�g�P�ӱb���K�X2.gif"></p>
<Blockquote>

<p>�b���G<asp:TextBox runat="server" id="Account" /></p>
<p>
�K�X�G<asp:TextBox runat="server" id="Password" 
                  TextMode="Password" /><p>
�O�o�ڡG<asp:CheckBox runat="server" id="RememberMe" /><p>

<asp:Button runat="server" text="�n  �J" OnClick="Login_Click" /><p>

<p>

<p>

<p>
<asp:Label runat="server" id="Msg" ForeColor="Red" />

</Blockquote><HR>
<p style="margin-top: 0; margin-bottom: 0">
<Font Size=-1 Color=Blue>�n�J�����e, �нT�w�s������ Cookie �O�}�Ҫ�.</Font> </p>
</Body>
</Html>

<script language="VB" runat="server">

   Sub Login_Click( sender As Object, e As EventArgs )

      If Verify( Account.Text, Password.Text ) Then
         'FormsAuthentication.RedirectFromLoginPage(Account.Text, RememberMe.Checked)
         Server.Transfer("Sale-password.aspx")

      Else
         Msg.Text = "�b���αK�X���~, �Э��s��J!"
      End If
   End Sub

   Function Verify( �b�� As String, �K�X As String) As Boolean
      Dim Conn As OleDbConnection, Cmd As OleDbCommand
      Dim Rd As OleDbDataReader, SQL  As String

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "Sample.mdb" )
      Dim DbPass = "Jet OLEDB:Database Password=dyitcom33"
      Conn = New OleDbConnection( Provider & ";" & DataBase & ";" & DbPass )
      'Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      SQL = "Select * From SaleMe Where " & _
                    "UserID='" & �b�� & "'" & _
                    " And Password='" & �K�X & "'"
      Cmd = New OleDbCommand( SQL, Conn )
      Rd = Cmd.ExecuteReader()
      If Rd.Read() Then ' ��ܦ���� UserID �� Password, �q�L����
         Conn.Close()
         Return True
      Else
         Conn.Close()
         Msg.Text = "�b���αK�X���~, �Э��s��J!"
         Return False
      End If
   End Function

    </script>
    <p>�@</td>
    <td width="29%" height="231">�@</td>
  </tr>
</table>
</Form>
</Body>
</Html>