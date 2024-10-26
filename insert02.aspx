<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Web.Mail" %> 


<Html>
<HEAD>
      <TITLE> 索取試用版 </TITLE>
</HEAD>
<Body BgColor="#FFFFCC">
<H3 align="center" style="margin-bottom: 0; margin-top:0"> 　</h3>
<H3 align="center" style="margin-bottom: 0; margin-top:0"> <font color="#3333CC"> 請填寫以下資料，我們將儘速與您聯絡。謝謝！</font> </h3>
<hr>
    
<Form runat="server">
<Table Border=0>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber1" width="711" height="235">
  <tr>
    <td width="174" height="20">　</td>
    <td width="88" height="20">
<font color="#008000">公司名稱 :</font>
    <td width="353" height="20">
    <asp:TextBox size="35" runat="server" id="公司名稱"/></td>
    <td width="96" height="20">　</td>
  </tr>
  <tr>
    <td width="174" height="20">　</td>
    <td width="88" height="20">

<font color="#008000">聯絡人 :</font>
    <td width="353" height="20"><asp:TextBox runat="server" id="聯絡人"/></td>
    <td width="96" height="20">　</td>
  </tr>
  <tr>
    <td width="174" height="21">　</td>
    <td width="88" height="21">


<font color="#008000">郵遞區號 :</font> 
    <td width="353" height="21"><asp:TextBox size="5" id="郵遞區號" runat="server"/></td>
    <td width="96" height="21">　</td>
  </tr>
  <tr>
    <td width="174" height="17"></td>
    <td width="88" height="17">
<font color="#008000">地址 :</font> 
    <td width="353" height="17"><asp:TextBox size="35" runat="server" id="地址" /></td>
    <td width="96" height="17"></td>
  </tr>
  <tr>
    <td width="174" height="17"></td>
    <td width="88" height="17">
<font color="#008000">電話 :</font> <br>
　<td width="353" height="17"><asp:TextBox id="電話" runat="server"/>
                             <font size=-1> 
<asp:RegularExpressionValidator runat="server"
                                         ControlToValidate="電話" 
                                         Text="請輸入與右邊範例相同的任一格式 06-2088808，069-2888808，0911-xxxxxx"
                                         ValidationExpression="[0-9]{2,4}-[0-9]{6,8}"  /></font>
    </td>
    <td width="96" height="17"></td>
  </tr>
  <tr>
    <td width="174" height="20">　</td>
    <td width="88" height="20">
<font color="#008000">傳真 :</font> 
                       

    </td>
    <td width="353" height="20"><asp:TextBox id="傳真" runat="server"/>      
    <td width="96" height="20">　</td>
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
<font color="#008000">行業別：</font></td>
    <td width="353" height="15"><asp:ListBox runat="server" id="行業別" rows=1 >
                                        <asp:ListItem selected>加工、模具業</asp:ListItem>
                                        <asp:ListItem>電子零件、電腦週邊相關業</asp:ListItem>
                                        <asp:ListItem>機械、汽機車、儀器設備製造業</asp:ListItem>
                                        <asp:ListItem>電腦系統、電子、資訊、通訊製造業</asp:ListItem>
                                        <asp:ListItem>電機電工器材．五金．其他機械相關業</asp:ListItem>
                                        <asp:ListItem>光電通信器材相關業、其他精密工業</asp:ListItem>     
                                        <asp:ListItem>文具玩具禮品業</asp:ListItem>
                                        <asp:ListItem>成衣、製鞋、體育用品製造業</asp:ListItem>
                                        <asp:ListItem>電訊郵政汽電公用事業</asp:ListItem>
                                        <asp:ListItem>學校及文教機構</asp:ListItem>     
                                        <asp:ListItem>其他</asp:ListItem>
                                        </asp:ListBox></td>    
    <td width="96" height="15"></td>
  </tr>
  
  <tr>
    <td width="174" height="15"></td>
    <td width="88" height="15">
                                        
<font color="#008000">試用軟體:<br>
</font><font size="2" color="#FF0000">(可複選，請按Ctrl+滑鼠左鍵)</font></td>
    <td width="353" height="15"></font><font color="red" size=-1>&nbsp;</font><asp:ListBox runat="server" id="test"  SelectionMode="Multiple" size=4>
                      <asp:ListItem selected>管理贏家ERP版</asp:ListItem>
                      <asp:ListItem >管理贏家豪華版</asp:ListItem>
                                        <asp:ListItem>生管軟體系統</asp:ListItem>
                                        <asp:ListItem>TOP-Link 傳輸王系統</asp:ListItem>
                                        <asp:ListItem>SGTM CNC車床模擬系統</asp:ListItem>
                                        <asp:ListItem>CNC銑床程式製作系統</asp:ListItem>
                                        <asp:ListItem>CNC銑床程式自動轉換系統</asp:ListItem>
                                        <asp:ListItem>SGTM CNC車床製作系統</asp:ListItem>
                                        </asp:ListBox>  </td>
    <td width="96" height="15"></td>
  </tr>
  <tr>
    <td width="174" height="22">　</td>
    <td width="88" height="22">
    <font color="#008000">留言 : </font> 
 
 
    </td>
    <td width="353" height="22">
    <asp:TextBox runat="server" id="留言" textmode="MultiLine" rows="5" Columns="50" />　</td>
    <td width="96" height="22">　</td>
  </tr>
  <tr>
    <td width="174" height="31">　</td>
    <td width="88" height="31">　</td>
    <td width="353" height="31"> 
    <asp:Button runat="server" Text="送   出" OnClick="InsertData" />
                               <asp:Label runat="server" id="Msg" ForeColor="Red" /></td>
    <td width="96" height="31">　</td>
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
      
      
      dim testmail,mailURL ,J
       For J = 0 To test.Items.Count - 1
          If test.Items(J).Selected Then
             testmail &= test.Items(J).Text & "，"
          End If
       Next
       mailURL &=  testmail



      Dim mail As MailMessage = New MailMessage
      mail.To         = "dyit@dyit.com.tw"
      mail.From       = "dyit@dyit.com.tw"
      mail.Subject    = "尋求試用軟體" & 電話.text
      mail.BodyFormat = MailFormat.Text 
      mail.Body       = "公司名稱：" & 公司名稱.text & chr(13) & chr(10) & _
     "聯絡人：" & 聯絡人.text  & chr(13) & chr(10) & "郵遞區號：" & 郵遞區號.text & _
      chr(13) & chr(10) & "地址：" & 地址.text & chr(13) & chr(10) & "電話：" & _
      電話.text & chr(13) & chr(10) & "傳真：" & 傳真.text & chr(13) & chr(10) & _
      "E-mail :" & Email.text & chr(13) & chr(10) & "行業別：" & 行業別.selectedItem.Text & _
      chr(13) & chr(10) & "試用軟體："  & mailURL  & chr(13) & chr(10) & _
      "留言：" & 留言.text     
      SmtpMail.SmtpServer = "dyit.com.tw"
      SmtpMail.Send(mail)

     If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
 
       Dim URL, testURL, I
       URL  = "Show_Insert2.aspx"
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
             testURL &= test.Items(I).Text & "，"
          End If
       Next
       URL &= "&test=" & testURL

       Server.Transfer( URL )

      

 End If

   End Sub

</script>