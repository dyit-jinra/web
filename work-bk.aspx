<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Web.Mail" %> 


<Html>
<HEAD>
      <TITLE> �������X�Ԩt�� </TITLE>
</HEAD>
<Body BgColor="#FFFFCC">
<p align="center" style="margin-bottom: 0; margin-top:0"> �@</p>
<H3 align="center" style="margin-bottom: 0; margin-top:0"> 
<font color="#3333CC" size="5">�F�ɬ�޺������X�Ԩt��</font> </h3>
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
�@</td>
    <td width="61" height="26">
<font color="#008000">�m�W�G</font></td>
    <td width="139" height="26">
    <asp:ListBox runat="server" id="�m�W" rows=1 Width="104px" Height="24px" Font-Size="Small">
                                        <asp:ListItem selected>���廨</asp:ListItem>
                                        <asp:ListItem>�����}</asp:ListItem>
                                        <asp:ListItem>�L�F�F</asp:ListItem>
                                        <asp:ListItem>�����~</asp:ListItem>
                                        <asp:ListItem>���R�u</asp:ListItem>
                                        <asp:ListItem>�«ئ�</asp:ListItem>
                                        </asp:ListBox></td>    
                                        
 </td>
    <td width="277" height="11" rowspan="2">�@ 
    <asp:Button runat="server" Text=" �n���X�Ԯɶ� " OnClick="InsertData" />
                               <asp:Label runat="server" id="Msg" ForeColor="Red" /></td></td>
    <td width="112" height="3"></td>
  </tr>
  <tr>
    <td width="101" height="11">�@</td>
    <td width="24" height="11">
    </td>
    <td width="61" height="11">
    <font color="#008000">�Ƶ��G</font></td>
    <td width="359" height="11">
    <asp:TextBox runat="server" id="�Ƶ�" textmode="MultiLine" rows="4" Columns="38" />      
    <td width="112" height="11">�@</td>
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
    
    <font color="#800080">�W��</font></td>
    <td width="31%" height="10">
    

    <font color="#800080">�U��</font></td>
    <td width="22%" height="10"></td>
  </tr>
  <tr>
    <td width="11%" height="14"><font color="#800080">�n���H���G</font></td>
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
      SQL = "Select �m�W,�ɶ� From �X�� Where Format(�ɶ�," & """" & "yyyy/m/d" & """" & _
            ") = #" & mDate & "# AND Hour(�ɶ�) <= 13"  
      Adpt  = New OleDbDataAdapter( SQL, Conn )
      Ds = New Dataset()
      Adpt.Fill(Ds,"�X��")
      
      Dim Table1 As DataTable = Ds.Tables("�X��")
      MyGrid.DataSource = Table1.DefaultView
      MyGrid.DataBind()
      
      SQL = "Select �m�W,�ɶ� From �X�� Where Format(�ɶ�," & """" & "yyyy/m/d" & """" & _
            ") = #" & mDate & "# AND Hour(�ɶ�) > 13" 
      Adpt  = New OleDbDataAdapter( SQL, Conn )
      Ds = New Dataset()
      Adpt.Fill(Ds,"�X��")
      
    
      Dim Table2 As DataTable = Ds.Tables("�X��")
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
      SQL = "Insert Into �X�� (�m�W,�Ƶ�) Values( ?, ?)"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("�m�W", OleDbType.Char, 5))
      Cmd.parameters.add( New OleDbParameter("�Ƶ�", OleDbType.Char, 50))

      Cmd.Parameters("�m�W").Value = �m�W.selectedItem.Text
      Cmd.Parameters("�Ƶ�").Value = �Ƶ�.Text
      Cmd.ExecuteNonQuery()
      Conn.Close()
      
      
      dim mailURL
      Dim mNote as string 
      dim mTime as date
      mtime = Format(now,"Long Time")
      mailURL = "�X��" &  �m�W.selectedItem.Text & mTime
      mNote = "�Ƶ��G" &  �Ƶ�.Text

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
              
       URL &= "?�m�W=" & Server.URLEncode( �m�W.SelectedItem.Text )
       URL &= "&���=" & Server.URLEncode( mDate)
       URL &= "&�ɶ�=" & Server.URLEncode( mTime )
       URL &= "&�Ƶ�=" & Server.URLEncode( �Ƶ�.Text )

       Server.Transfer( URL )

     End If

   End Sub

</script>