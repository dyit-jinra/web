<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Web.Mail" %> 

<Html>
<Body BgColor="White">
<H2>�ж�g�H�U��ơA�ڭ̱N���t�N�n��e��z�⤤�C���¡I<Hr></H2>

<Form runat="server">
<Table Border=1>
  <Tr><Td>���q�W�١G</Td>
  <Td><asp:TextBox id="���q�W��" Size=40 runat="server"/></Td></Tr>
  <Tr><Td>�a�}�G</Td>
  <Td><asp:TextBox id="�a�}" Size=40 runat="server"/></Td></Tr>
  <Tr><Td>�D���G</Td>
  <Td><asp:TextBox id="mailSubject" Size=40 runat="server"/></Td></Tr>
  <Tr><Td>����G</Td>
  <Td><asp:TextBox runat="server" id="mailBody" TextMode="MultiLine" Rows=8 Cols=60 />
  </Td></Tr>
</Table>
<asp:Button runat="server" Text="�e �X" OnClick="Button_Click" />
</Form>

<Hr><asp:Label id="Msg" runat="server" ForeColor="Red" /><p>

</Body>
</Html>

<script Language="VB" runat="server">

   Sub Button_Click(sender As Object, e As EventArgs) 
      Dim mail As MailMessage = New MailMessage

      mail.To         = "dyit01@dyit.com.tw"
      mail.From       = ���q�W��.Text
      mail.Subject    = mailSubject.Text
      mail.BodyFormat = MailFormat.Text 
      mail.Body       = mailBody.Text

      On Error Resume Next
      SmtpMail.SmtpServer = "dyit.com.tw"
      SmtpMail.Send(mail) 

      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
         Msg.Text = "�l��w�g�e�X!"
      End If
   End Sub

</script>
