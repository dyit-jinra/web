<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>



<Html>
<Body BgColor="White">
<H3>             �ж�g�H�U��ơA�ڭ̱N���t�N�n��e��z�⤤�C���¡I <HR></H3>

<Form runat="server">
<Blockquote>
                ���q�W�� :  <asp:TextBox runat="server" id="���q�W��" /><br>
                �p �� �H :  <asp:TextBox runat="server" id="�p���H" /><br>
                �l���ϸ� :  <asp:TextBox id="�l���ϸ�" runat="server"/>
                            <asp:RegularExpressionValidator runat="server"
                             ControlToValidate="�l���ϸ�" Text="(�п�J 3 ��Ʀr)"
                             ValidationExpression="[0-9]{3}" /><p>
                �a    �} :  <asp:TextBox runat="server" id="�a�}" /><br>
                �q    �� :  <asp:TextBox id="�q��" runat="server"/>
                            <asp:RegularExpressionValidator runat="server"
                             ControlToValidate="�q��"
                             Text="(XX)XXXX-XXXX, (XX)XXX-XXXX, (XXX)XXX-XXXX, (XXX)XX-XXXX"
                             ValidationExpression="\([0-9]{2,3}\)[0-9]{2,4}-[0-9]{4}" /><p>
                 ��   �u :  <asp:TextBox id="�ǯu" runat="server"/>
                            <asp:RegularExpressionValidator runat="server"
                             ControlToValidate="�ǯu"
                             Text="(XX)XXXX-XXXX, (XX)XXX-XXXX, (XXX)XXX-XXXX, (XXX)XX-XXXX"
                             ValidationExpression="\([0-9]{2,3}\)[0-9]{2,4}-[0-9]{4}" /><p>
                E-mail�G    <asp:TextBox id="Email" runat="server"/>
                            <asp:RegularExpressionValidator runat="server"
                             ControlToValidate="Email" Text="(Email ���t�� @ �Ÿ�)"
                             ValidationExpression=".{1,}@.{3,}" /><p>
                 �� �~ �O�G <asp:ListBox runat="server" id="��~�O" size="1">
                            <asp:ListItem>�q���g��������~</asp:ListItem>
                            <asp:ListItem>�i�X�f�T���~</asp:ListItem>
                            <asp:ListItem>���ʲ�,�򤶬����~</asp:ListItem>
                            <asp:ListItem>�s�i��P�~</asp:ListItem>
                            <asp:ListItem>��L</asp:ListItem>
                            </asp:ListBox></td>
                �եγn�� :  <asp:ListBox runat="server" id="test" SelectionMode="Multiple"                                size=4>
                            <asp:ListItem>�ǿ�� TOP-Link</asp:ListItem>
                            <asp:ListItem>�i�P�s�n��</asp:ListItem>
                            <asp:ListItem>�|�p�n��</asp:ListItem>
                            <asp:ListItem>�ͺ޳n��</asp:ListItem>
                            </asp:ListBox><p>
                �d   �� :   <asp:TextBox runat="server" id="�d��" /><p>
                            <asp:Button runat="server" Text="�e�X" OnClick="InsertData" />
                            </Blockquote>
<HR><asp:Label runat="server" id="Msg" ForeColor="Red" />
</Form>
</Body>
</Html>

<script Language="VB" runat="server">
Sub InsertData(sender As Object, e As EventArgs) 
     
      
      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "Sample.mdb" )
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      Dim SQL As String
SQL = "Insert Into �եΪ� (���q�W��,�p���H,�l���ϸ�,�a�},�q��,�ǯu,Email,��~�O,test,�d��) Values( ?, ?, ?, ?,?, ?,?,?,?,?)"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("���q�W��", OleDbType.Char, 50))
      Cmd.Parameters.Add( New OleDbParameter("�p���H", OleDbType.Char, 50))
      Cmd.Parameters.Add( New OleDbParameter("�l���ϸ�", OleDbType.Char, 50))
      Cmd.Parameters.Add( New OleDbParameter("�a�}", OleDbType.Char, 50))
      Cmd.Parameters.Add( New OleDbParameter("�q��", OleDbType.Char, 50))
      cmd.parameters.add( New OleDbParameter("�ǯu", OleDbType.Char, 50))
      cmd.parameters.add( New OleDbParameter("Email", OleDbType.Char, 50)) 
      cmd.parameters.add( New OleDbParameter("��~�O", OleDbType.Char, 50))
      cmd.parameters.add( New OleDbParameter("test", OleDbType.Char, 50))
      cmd.parameters.add( New OleDbParameter("�d��", OleDbType.Char, 50))

      Cmd.Parameters("���q�W��").Value = ���q�W��.Text
      Cmd.Parameters("�p���H").Value = �p���H.Text
      Cmd.Parameters("�l���ϸ�").Value = �l���ϸ�.Text
      Cmd.Parameters("�a�}").Value = �a�}.Text
      Cmd.Parameters("�q��").Value = �q��.Text
      Cmd.Parameters("�ǯu").Value = �ǯu.Text
      Cmd.Parameters("Email").Value = Email.Text
      Cmd.Parameters("��~�O").Value = ��~�O.selectedItem.Text
      Cmd.Parameters("test").Value = test.selectedItem.Text
      Cmd.Parameters("�d��").Value = �d��.Text
      Cmd.ExecuteNonQuery()
           Conn.Close()

     If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
 Dim mail As MailMessage = New MailMessage

      mail.To         = "dyit01@dyit.com.tw"
      mail.From       = ���q�W��.Text
      mail.Subject    = "�M�D�եγn��"
      mail.BodyFormat = MailFormat.Text 
      mail.Body       = ���q�W��.Text & �a�}.text & �p���H.text

      On Error Resume Next
      SmtpMail.SmtpServer = "dyit.com.tw"
      SmtpMail.Send(mail)




       Dim URL, testURL, I
       URL  = "Show_Insert.aspx"
       URL &= "?���q�W��=" & Server.URLEncode( ���q�W��.Text )
       URL &= "&�p���H=" & Server.URLEncode( �p���H.Text )
       URL &= "&�l���ϸ�=" & Server.URLEncode( �l���ϸ�.Text )
       URL &= "&�a�}=" & Server.URLEncode( �a�}.Text )
       URL &= "&�q��=" & Server.URLEncode( �q��.Text )
       URL &= "&�ǯu=" & Server.URLEncode( �ǯu.Text )
       URL &= "&Email=" & Server.URLEncode( Email.Text )
       URL &= "&��~�O=" & Server.URLEncode( ��~�O.SelectedItem.Text )
       URL &= "&�d��=" & Server.URLEncode( �d��.Text )
       For I = 0 To test.Items.Count - 1
          If test.Items(I).Selected Then
             testURL &= test.Items(I).Text 
          End If
       Next
       URL &= "&test=" & testURL

       Server.Transfer( URL )

      

 End If

   End Sub

</script>
