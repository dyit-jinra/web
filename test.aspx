<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Web.Mail" %> 

<Html>
<Body BgColor="White">
<H2>請填寫以下資料，我們將儘速將軟體送到您手中。謝謝！<Hr></H2>

<Form runat="server">
<Table Border=1>
  <Tr><Td>公司名稱：</Td>
  <Td><asp:TextBox id="公司名稱" Size=40 runat="server"/></Td></Tr>
  <Tr><Td>地址：</Td>
  <Td><asp:TextBox id="地址" Size=40 runat="server"/></Td></Tr>
  <Tr><Td>主旨：</Td>
  <Td><asp:TextBox id="mailSubject" Size=40 runat="server"/></Td></Tr>
  <Tr><Td>內文：</Td>
  <Td><asp:TextBox runat="server" id="mailBody" TextMode="MultiLine" Rows=8 Cols=60 />
  </Td></Tr>
</Table>
<asp:Button runat="server" Text="送 出" OnClick="Button_Click" />
</Form>

<Hr><asp:Label id="Msg" runat="server" ForeColor="Red" /><p>

</Body>
</Html>

<script Language="VB" runat="server">

   Sub Button_Click(sender As Object, e As EventArgs) 
      Dim mail As MailMessage = New MailMessage

      mail.To         = "dyit01@dyit.com.tw"
      mail.From       = 公司名稱.Text
      mail.Subject    = mailSubject.Text
      mail.BodyFormat = MailFormat.Text 
      mail.Body       = mailBody.Text

      On Error Resume Next
      SmtpMail.SmtpServer = "dyit.com.tw"
      SmtpMail.Send(mail) 

      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
         Msg.Text = "郵件已經送出!"
      End If
   End Sub

</script>
