<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Web.Mail" %> 


<Html>
<HEAD>
      <TITLE> 網路版出勤系統 </TITLE>
</HEAD>
<Body BgColor="#FFFFCC">
<p align="center" style="margin-bottom: 0; margin-top:0"> 　</p>
<H3 align="center" style="margin-bottom: 0; margin-top:0"> 
<font color="#3333CC" size="5">東怡科技網路版出勤系統</font> </h3>
<hr>
    
<Form runat="server">
<Table Border=0>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber1" width="714" height="44">
 
 
  <tr>
    <td width="101" height="7"></td>
    <td width="24" height="7">
    </td>
    <td width="61" height="7">
    </td>
    <td width="139" height="7"></td>    
                                        
    <td width="277" height="4"></td>
    <td width="112" height="4"></td>
  </tr>
 
 
  <tr>
    <td width="101" height="26"></td>
    <td width="24" height="26">
<p align="right">
　</td>
    <td width="61" height="26">
<font color="#008000">姓名：</font></td>
    <td width="139" height="26">
    <asp:ListBox runat="server" id="姓名" rows=1 Width="104px" Height="24px" Font-Size="Small">
                                        <asp:ListItem selected>蕭文豪</asp:ListItem>
                                        <asp:ListItem>楊中良</asp:ListItem>
                                        <asp:ListItem>林政達</asp:ListItem>
                                        <asp:ListItem>楊謙驊</asp:ListItem>
                                        <asp:ListItem>陳麗真</asp:ListItem>
                                        <asp:ListItem>謝建成</asp:ListItem>
                                        </asp:ListBox></td>    
                                        
 </td>
    <td width="277" height="11" rowspan="2">　 
    <asp:Button runat="server" Text=" 登錄出勤時間 " OnClick="InsertData" />
                               <asp:Label runat="server" id="Msg" ForeColor="Red" /></td></td>
    <td width="112" height="3"></td>
  </tr>
  <tr>
    <td width="101" height="11">　</td>
    <td width="24" height="11">
    </td>
    <td width="61" height="11">
    <font color="#008000">備註：</font></td>
    <td width="359" height="11">
    <asp:TextBox runat="server" id="備註" textmode="MultiLine" rows="4" Columns="38" />      
    <td width="112" height="11">　</td>
  </tr>
  <tr>
                                        
                                        
    <td width="243" height="17">
   
    <td width="128" height="3" colspan="2"></td>
  </tr></font>
 
  <tr>
    <td width="101" height="1"></td>
    <td width="25" height="1" colspan="2"></td>
    <td width="139" height="1"> 
    
    
    </td>
  </tr>
</table>


<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2" height="47">
  <tr>
    <td width="100%" height="10" colspan="4"><hr></td>
  </tr>
  <tr>
    <td width="11%" height="10"></td>
    <td width="36%" height="10">
    
    <font color="#800080">上午</font></td>
    <td width="31%" height="10">
    

    <font color="#800080">下午</font></td>
    <td width="22%" height="10"></td>
  </tr>
  <tr>
    <td width="11%" height="14"><font color="#800080">登錄人員：</font></td>
    <td width="36%" height="14"><asp:DataGrid runat = "Server" id = "MyGrid" /></td>
    <td width="31%" height="14"><asp:DataGrid runat = "Server" id = "MyGrid2" /></td>
    <td width="22%" height="14"></td>
  </tr>
  </table>

</Table>
</Form>


</Body>
</Html>

<script Language="VB" runat="server">
  Sub Page_Load(sender As Object, e As EventArgs)
      Dim Conn As OleDbConnection
      Dim Adpt  As OleDbDataAdapter
      Dim Ds As DataSet
      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "Sample.mdb" )
      Dim DbPass = "Jet OLEDB:Database Password=dyitcom33"
      Conn = New OleDbConnection( Provider & ";" & DataBase & ";" & DbPass )
      Conn.Open()
      
      dim mDate As date 
      Dim SQL As String
      mDate = now()
      mDate =Format(mDate ,"Long Date")
      SQL = "Select 姓名,時間 From 出勤 Where Format(時間," & """" & "yyyy/m/d" & """" & _
            ") = #" & mDate & "# AND Hour(時間) <= 13"  
      Adpt  = New OleDbDataAdapter( SQL, Conn )
      Ds = New Dataset()
      Adpt.Fill(Ds,"出勤")
      
      Dim Table1 As DataTable = Ds.Tables("出勤")
      MyGrid.DataSource = Table1.DefaultView
      MyGrid.DataBind()
      
      SQL = "Select 姓名,時間 From 出勤 Where Format(時間," & """" & "yyyy/m/d" & """" & _
            ") = #" & mDate & "# AND Hour(時間) > 13" 
      Adpt  = New OleDbDataAdapter( SQL, Conn )
      Ds = New Dataset()
      Adpt.Fill(Ds,"出勤")
      
    
      Dim Table2 As DataTable = Ds.Tables("出勤")
      MyGrid2.DataSource = Table2.DefaultView
      MyGrid2.DataBind()
     
      conn.Close()
      
  End Sub
  
  Sub InsertData(sender As Object, e As EventArgs) 
     
      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "Sample.mdb" )
      Dim DbPass = "Jet OLEDB:Database Password=dyitcom33"
      Conn = New OleDbConnection( Provider & ";" & DataBase & ";" & DbPass )
      Conn.Open()

      Dim SQL As String
      SQL = "Insert Into 出勤 (姓名,備註) Values( ?, ?)"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("姓名", OleDbType.Char, 5))
      Cmd.parameters.add( New OleDbParameter("備註", OleDbType.Char, 50))

      Cmd.Parameters("姓名").Value = 姓名.selectedItem.Text
      Cmd.Parameters("備註").Value = 備註.Text
      Cmd.ExecuteNonQuery()
      Conn.Close()
      
      
      dim mailURL
      Dim mNote as string 
      dim mTime as date
      mtime = Format(now,"Long Time")
      mailURL = "出勤" &  姓名.selectedItem.Text & mTime
      mNote = "備註：" &  備註.Text

      Dim mail As MailMessage = New MailMessage
      mail.To         = "dyit@dyit.com.tw"
      mail.From       = "dyit@dyit.com.tw"
      mail.Subject    = mailURL 
      mail.BodyFormat = MailFormat.Text 
      mail.Body       = mNote
       
      SmtpMail.SmtpServer = "dyit.com.tw"
      SmtpMail.Send(mail)

     If Err.Number <> 0 Then
         Msg.Text = Err.Description
     Else
 
       Dim URL
       Dim mDate as string 
       dim mYear as long 
       mYear = Format(now,"yyyy")-1911
       mDate= mYear  & "/" & format(now,"MM/dd")
       URL  = "Show_work.aspx"
              
       URL &= "?姓名=" & Server.URLEncode( 姓名.SelectedItem.Text )
       URL &= "&日期=" & Server.URLEncode( mDate)
       URL &= "&時間=" & Server.URLEncode( mTime )
       URL &= "&備註=" & Server.URLEncode( 備註.Text )

       Server.Transfer( URL )

     End If

   End Sub

</script>