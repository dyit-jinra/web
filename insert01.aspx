<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Web.Mail" %> 


<Html>
<HEAD>
      <TITLE> �����եΪ� </TITLE>
</HEAD>
<Body BgColor="#FFFFCC">
<H3> <font color="#3333CC"> �ж�g�H�U��ơA�ڭ̱N���t�N�n��e��z�⤤�C���¡I</font> <HR></H3>

<Form runat="server">
<Blockquote>

<font color="#008000">���q�W�� :</font> <asp:TextBox size="35" runat="server" id="���q�W��"                                          /><br>
<font color="#008000">�p �� �H :</font> <asp:TextBox runat="server" id="�p���H" /><br>
<font color="#008000">�l���ϸ� :</font> <asp:TextBox size="7" id="�l���ϸ�" runat="server"/>
                         <font size=-1> <asp:RegularExpressionValidator runat="server"
                                         ControlToValidate="�l���ϸ�" Text="(�п�J 3 ��Ʀr)"
                                         ValidationExpression="[0-9]{3}" /></font><p>
<font color="#008000">�a    �} :</font> <asp:TextBox size="35" runat="server" id="�a�}" /><br>
<font color="#008000">�q    �� :</font> <asp:TextBox id="�q��" runat="server"/>
                             <font size=-1> <asp:RegularExpressionValidator runat="server"
                                         ControlToValidate="�q��" 
                                         Text="�п�J�P�k��d�ҬۦP�����@�榡(XX)XXXX-XXXX, (XX)                                                 XXX-XXXX"
                                         ValidationExpression="\([0-9]{2,3}\)[0-9]{2,4}-[0-9]{4}"                                           /></font><p>
<font color="#008000">��   �u :</font> <asp:TextBox id="�ǯu" runat="server"/>
                              <font size=-1><asp:RegularExpressionValidator runat="server"
                                        ControlToValidate="�ǯu"
                                        Text="�п�J�P�k��d�ҬۦP�����@�榡(XX)XXXX-XXXX, (XX)                                              XXX-XXXX"
                                        ValidationExpression="\([0-9]{2,3}\)[0-9]{2,4}-[0-9]{4}"                                          /></font><p>
<font color="#008000">E-mail�G</font>  <asp:TextBox id="Email" runat="server"/>
                          <font size=-1> <asp:RegularExpressionValidator runat="server"
                                       ControlToValidate="Email" Text="(Email ���t�� @ �Ÿ�)"
                                       ValidationExpression=".{1,}@.{3,}" /></font><p>
<font color="#008000">�� �~ �O�G</font> <asp:ListBox runat="server" id="��~�O" size="1">
                                        <asp:ListItem>�q���g��������~</asp:ListItem>
                                        <asp:ListItem>�i�X�f�T���~</asp:ListItem>
                                        <asp:ListItem>���ʲ�,�򤶬����~</asp:ListItem>
                                        <asp:ListItem>�s�i��P�~</asp:ListItem>
                                        <asp:ListItem>��L</asp:ListItem>
                                        </asp:ListBox><br></td><br>
<font color="#008000">�եγn�� :</font><font color="red" size=-1>(�i�ƿ�A�Ы�Ctrl+�ƹ�����) </font><asp:ListBox runat="server" id="test"                                          SelectionMode="Multiple" size=4>
                                        <asp:ListItem>�ǿ�� TOP-Link</asp:ListItem>
                                        <asp:ListItem>�i�P�s�n��</asp:ListItem>
                                        <asp:ListItem>�|�p�n��</asp:ListItem>
                                        <asp:ListItem>�ͺ޳n��</asp:ListItem>
                                        <asp:ListItem>�i�P�s�n��</asp:ListItem>
                                        <asp:ListItem>�|�p�n��</asp:ListItem>
                                        <asp:ListItem>�ͺ޳n��</asp:ListItem>
                                        </asp:ListBox><br><p>
<font color="#008000">�d   �� : </font> <asp:TextBox runat="server" id="�d��" /><p>
                                        <asp:Button runat="server" Text="�e�X"                                              OnClick="InsertData" />
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

      Dim mail As MailMessage = New MailMessage
      mail.To         = "dyit01@dyit.com.tw"
      mail.From       = Email.text
      mail.Subject    = "�M�D�եγn��"
      mail.BodyFormat = MailFormat.Text 
      mail.Body       = "���q�W�١G" & ���q�W��.text & chr(13) & chr(10) & _
     "�p���H�G" & �p���H.text  & chr(13) & chr(10) & "�l���ϸ��G" & �l���ϸ�.text & _
      chr(13) & chr(10) & "�a�}�G" & �a�}.text & chr(13) & chr(10) & "�q�ܡG" & _
      �q��.text & chr(13) & chr(10) & "�ǯu�G" & �ǯu.text & chr(13) & chr(10) & _
      "E-mail :" & Email.text & chr(13) & chr(10) & "��~�O�G" & ��~�O.selectedItem.Text & _
      chr(13) & chr(10) & "�եγn��G"  & test.selectedItem.Text  & chr(13) & chr(10) & _
      "�d���G" & �d��.text     
      SmtpMail.SmtpServer = "dyit.com.tw"
      SmtpMail.Send(mail)

     If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
 
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
