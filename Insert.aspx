<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>



<Html>
<Body BgColor="White">
<H3>             請填寫以下資料，我們將儘速將軟體送到您手中。謝謝！ <HR></H3>

<Form runat="server">
<Blockquote>
                公司名稱 :  <asp:TextBox runat="server" id="公司名稱" /><br>
                聯 絡 人 :  <asp:TextBox runat="server" id="聯絡人" /><br>
                郵遞區號 :  <asp:TextBox id="郵遞區號" runat="server"/>
                            <asp:RegularExpressionValidator runat="server"
                             ControlToValidate="郵遞區號" Text="(請輸入 3 位數字)"
                             ValidationExpression="[0-9]{3}" /><p>
                地    址 :  <asp:TextBox runat="server" id="地址" /><br>
                電    話 :  <asp:TextBox id="電話" runat="server"/>
                            <asp:RegularExpressionValidator runat="server"
                             ControlToValidate="電話"
                             Text="(XX)XXXX-XXXX, (XX)XXX-XXXX, (XXX)XXX-XXXX, (XXX)XX-XXXX"
                             ValidationExpression="\([0-9]{2,3}\)[0-9]{2,4}-[0-9]{4}" /><p>
                 傳   真 :  <asp:TextBox id="傳真" runat="server"/>
                            <asp:RegularExpressionValidator runat="server"
                             ControlToValidate="傳真"
                             Text="(XX)XXXX-XXXX, (XX)XXX-XXXX, (XXX)XXX-XXXX, (XXX)XX-XXXX"
                             ValidationExpression="\([0-9]{2,3}\)[0-9]{2,4}-[0-9]{4}" /><p>
                E-mail：    <asp:TextBox id="Email" runat="server"/>
                            <asp:RegularExpressionValidator runat="server"
                             ControlToValidate="Email" Text="(Email 應含有 @ 符號)"
                             ValidationExpression=".{1,}@.{3,}" /><p>
                 行 業 別： <asp:ListBox runat="server" id="行業別" size="1">
                            <asp:ListItem>電腦週邊相關產業</asp:ListItem>
                            <asp:ListItem>進出口貿易業</asp:ListItem>
                            <asp:ListItem>不動產,仲介相關業</asp:ListItem>
                            <asp:ListItem>廣告行銷業</asp:ListItem>
                            <asp:ListItem>其他</asp:ListItem>
                            </asp:ListBox></td>
                試用軟體 :  <asp:ListBox runat="server" id="test" SelectionMode="Multiple"                                size=4>
                            <asp:ListItem>傳輸王 TOP-Link</asp:ListItem>
                            <asp:ListItem>進銷存軟體</asp:ListItem>
                            <asp:ListItem>會計軟體</asp:ListItem>
                            <asp:ListItem>生管軟體</asp:ListItem>
                            </asp:ListBox><p>
                留   言 :   <asp:TextBox runat="server" id="留言" /><p>
                            <asp:Button runat="server" Text="送出" OnClick="InsertData" />
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
SQL = "Insert Into 試用版 (公司名稱,聯絡人,郵遞區號,地址,電話,傳真,Email,行業別,test,留言) Values( ?, ?, ?, ?,?, ?,?,?,?,?)"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("公司名稱", OleDbType.Char, 50))
      Cmd.Parameters.Add( New OleDbParameter("聯絡人", OleDbType.Char, 50))
      Cmd.Parameters.Add( New OleDbParameter("郵遞區號", OleDbType.Char, 50))
      Cmd.Parameters.Add( New OleDbParameter("地址", OleDbType.Char, 50))
      Cmd.Parameters.Add( New OleDbParameter("電話", OleDbType.Char, 50))
      cmd.parameters.add( New OleDbParameter("傳真", OleDbType.Char, 50))
      cmd.parameters.add( New OleDbParameter("Email", OleDbType.Char, 50)) 
      cmd.parameters.add( New OleDbParameter("行業別", OleDbType.Char, 50))
      cmd.parameters.add( New OleDbParameter("test", OleDbType.Char, 50))
      cmd.parameters.add( New OleDbParameter("留言", OleDbType.Char, 50))

      Cmd.Parameters("公司名稱").Value = 公司名稱.Text
      Cmd.Parameters("聯絡人").Value = 聯絡人.Text
      Cmd.Parameters("郵遞區號").Value = 郵遞區號.Text
      Cmd.Parameters("地址").Value = 地址.Text
      Cmd.Parameters("電話").Value = 電話.Text
      Cmd.Parameters("傳真").Value = 傳真.Text
      Cmd.Parameters("Email").Value = Email.Text
      Cmd.Parameters("行業別").Value = 行業別.selectedItem.Text
      Cmd.Parameters("test").Value = test.selectedItem.Text
      Cmd.Parameters("留言").Value = 留言.Text
      Cmd.ExecuteNonQuery()
           Conn.Close()

     If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
 Dim mail As MailMessage = New MailMessage

      mail.To         = "dyit01@dyit.com.tw"
      mail.From       = 公司名稱.Text
      mail.Subject    = "尋求試用軟體"
      mail.BodyFormat = MailFormat.Text 
      mail.Body       = 公司名稱.Text & 地址.text & 聯絡人.text

      On Error Resume Next
      SmtpMail.SmtpServer = "dyit.com.tw"
      SmtpMail.Send(mail)




       Dim URL, testURL, I
       URL  = "Show_Insert.aspx"
       URL &= "?公司名稱=" & Server.URLEncode( 公司名稱.Text )
       URL &= "&聯絡人=" & Server.URLEncode( 聯絡人.Text )
       URL &= "&郵遞區號=" & Server.URLEncode( 郵遞區號.Text )
       URL &= "&地址=" & Server.URLEncode( 地址.Text )
       URL &= "&電話=" & Server.URLEncode( 電話.Text )
       URL &= "&傳真=" & Server.URLEncode( 傳真.Text )
       URL &= "&Email=" & Server.URLEncode( Email.Text )
       URL &= "&行業別=" & Server.URLEncode( 行業別.SelectedItem.Text )
       URL &= "&留言=" & Server.URLEncode( 留言.Text )
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
