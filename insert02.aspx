<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Web.Mail" %> 


<Html>
<HEAD>
      <TITLE> �����եΪ� </TITLE>
</HEAD>
<Body BgColor="#FFFFCC">
<H3 align="center" style="margin-bottom: 0; margin-top:0"> �@</h3>
<H3 align="center" style="margin-bottom: 0; margin-top:0"> <font color="#3333CC"> �ж�g�H�U��ơA�ڭ̱N���t�P�z�p���C���¡I</font> </h3>
<hr>
    
<Form runat="server">
<Table Border=0>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber1" width="711" height="235">
  <tr>
    <td width="174" height="20">�@</td>
    <td width="88" height="20">
<font color="#008000">���q�W�� :</font>
    <td width="353" height="20">
    <asp:TextBox size="35" runat="server" id="���q�W��"/></td>
    <td width="96" height="20">�@</td>
  </tr>
  <tr>
    <td width="174" height="20">�@</td>
    <td width="88" height="20">

<font color="#008000">�p���H :</font>
    <td width="353" height="20"><asp:TextBox runat="server" id="�p���H"/></td>
    <td width="96" height="20">�@</td>
  </tr>
  <tr>
    <td width="174" height="21">�@</td>
    <td width="88" height="21">


<font color="#008000">�l���ϸ� :</font> 
    <td width="353" height="21"><asp:TextBox size="5" id="�l���ϸ�" runat="server"/></td>
    <td width="96" height="21">�@</td>
  </tr>
  <tr>
    <td width="174" height="17"></td>
    <td width="88" height="17">
<font color="#008000">�a�} :</font> 
    <td width="353" height="17"><asp:TextBox size="35" runat="server" id="�a�}" /></td>
    <td width="96" height="17"></td>
  </tr>
  <tr>
    <td width="174" height="17"></td>
    <td width="88" height="17">
<font color="#008000">�q�� :</font> <br>
�@<td width="353" height="17"><asp:TextBox id="�q��" runat="server"/>
                             <font size=-1> 
<asp:RegularExpressionValidator runat="server"
                                         ControlToValidate="�q��" 
                                         Text="�п�J�P�k��d�ҬۦP�����@�榡 06-2088808�A069-2888808�A0911-xxxxxx"
                                         ValidationExpression="[0-9]{2,4}-[0-9]{6,8}"  /></font>
    </td>
    <td width="96" height="17"></td>
  </tr>
  <tr>
    <td width="174" height="20">�@</td>
    <td width="88" height="20">
<font color="#008000">�ǯu :</font> 
                       

    </td>
    <td width="353" height="20"><asp:TextBox id="�ǯu" runat="server"/>      
    <td width="96" height="20">�@</td>
  </tr>
  <tr>
    <td width="174" height="16"></td>
    <td width="88" height="16">
<font color="#008000">E-mail:</font> 

    </td>
    <td width="353" height="16"><asp:TextBox id="Email" runat="server"/>
    <td width="96" height="16"></td>
  </tr>
  <tr>
    <td width="174" height="15"></td>
    <td width="88" height="15">
<font color="#008000">��~�O�G</font></td>
    <td width="353" height="15"><asp:ListBox runat="server" id="��~�O" rows=1 >
                                        <asp:ListItem selected>�[�u�B�Ҩ�~</asp:ListItem>
                                        <asp:ListItem>�q�l�s��B�q���g������~</asp:ListItem>
                                        <asp:ListItem>����B�T�����B�����]�ƻs�y�~</asp:ListItem>
                                        <asp:ListItem>�q���t�ΡB�q�l�B��T�B�q�T�s�y�~</asp:ListItem>
                                        <asp:ListItem>�q���q�u�����D�����D��L��������~</asp:ListItem>
                                        <asp:ListItem>���q�q�H���������~�B��L��K�u�~</asp:ListItem>     
                                        <asp:ListItem>��㪱��§�~�~</asp:ListItem>
                                        <asp:ListItem>����B�s�c�B��|�Ϋ~�s�y�~</asp:ListItem>
                                        <asp:ListItem>�q�T�l�F�T�q���ΨƷ~</asp:ListItem>
                                        <asp:ListItem>�ǮդΤ�о��c</asp:ListItem>     
                                        <asp:ListItem>��L</asp:ListItem>
                                        </asp:ListBox></td>    
    <td width="96" height="15"></td>
  </tr>
  
  <tr>
    <td width="174" height="15"></td>
    <td width="88" height="15">
                                        
<font color="#008000">�եγn��:<br>
</font><font size="2" color="#FF0000">(�i�ƿ�A�Ы�Ctrl+�ƹ�����)</font></td>
    <td width="353" height="15"></font><font color="red" size=-1>&nbsp;</font><asp:ListBox runat="server" id="test"  SelectionMode="Multiple" size=4>
                      <asp:ListItem selected>�޲zĹ�aERP��</asp:ListItem>
                      <asp:ListItem >�޲zĹ�a���ت�</asp:ListItem>
                                        <asp:ListItem>�ͺ޳n��t��</asp:ListItem>
                                        <asp:ListItem>TOP-Link �ǿ���t��</asp:ListItem>
                                        <asp:ListItem>SGTM CNC���ɼ����t��</asp:ListItem>
                                        <asp:ListItem>CNC�ѧɵ{���s�@�t��</asp:ListItem>
                                        <asp:ListItem>CNC�ѧɵ{���۰��ഫ�t��</asp:ListItem>
                                        <asp:ListItem>SGTM CNC���ɻs�@�t��</asp:ListItem>
                                        </asp:ListBox>  </td>
    <td width="96" height="15"></td>
  </tr>
  <tr>
    <td width="174" height="22">�@</td>
    <td width="88" height="22">
    <font color="#008000">�d�� : </font> 
 
 
    </td>
    <td width="353" height="22">
    <asp:TextBox runat="server" id="�d��" textmode="MultiLine" rows="5" Columns="50" />�@</td>
    <td width="96" height="22">�@</td>
  </tr>
  <tr>
    <td width="174" height="31">�@</td>
    <td width="88" height="31">�@</td>
    <td width="353" height="31"> 
    <asp:Button runat="server" Text="�e   �X" OnClick="InsertData" />
                               <asp:Label runat="server" id="Msg" ForeColor="Red" /></td>
    <td width="96" height="31">�@</td>
  </tr>
</table>


</p>

</Table>
</Form>


</Body>
</Html>

<script Language="VB" runat="server">
Sub InsertData(sender As Object, e As EventArgs) 
     
      
      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "Sample.mdb" )
      Dim DbPass = "Jet OLEDB:Database Password=dyitcom33"
      Conn = New OleDbConnection( Provider & ";" & DataBase & ";" & DbPass )
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
      
      
      dim testmail,mailURL ,J
       For J = 0 To test.Items.Count - 1
          If test.Items(J).Selected Then
             testmail &= test.Items(J).Text & "�A"
          End If
       Next
       mailURL &=  testmail



      Dim mail As MailMessage = New MailMessage
      mail.To         = "dyit@dyit.com.tw"
      mail.From       = "dyit@dyit.com.tw"
      mail.Subject    = "�M�D�եγn��" & �q��.text
      mail.BodyFormat = MailFormat.Text 
      mail.Body       = "���q�W�١G" & ���q�W��.text & chr(13) & chr(10) & _
     "�p���H�G" & �p���H.text  & chr(13) & chr(10) & "�l���ϸ��G" & �l���ϸ�.text & _
      chr(13) & chr(10) & "�a�}�G" & �a�}.text & chr(13) & chr(10) & "�q�ܡG" & _
      �q��.text & chr(13) & chr(10) & "�ǯu�G" & �ǯu.text & chr(13) & chr(10) & _
      "E-mail :" & Email.text & chr(13) & chr(10) & "��~�O�G" & ��~�O.selectedItem.Text & _
      chr(13) & chr(10) & "�եγn��G"  & mailURL  & chr(13) & chr(10) & _
      "�d���G" & �d��.text     
      SmtpMail.SmtpServer = "dyit.com.tw"
      SmtpMail.Send(mail)

     If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
 
       Dim URL, testURL, I
       URL  = "Show_Insert2.aspx"
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
             testURL &= test.Items(I).Text & "�A"
          End If
       Next
       URL &= "&test=" & testURL

       Server.Transfer( URL )

      

 End If

   End Sub

</script>