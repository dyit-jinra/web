
<HTML>
<HEAD>
      <TITLE> �������X�Ԩt�ε��G��ܵe�� </TITLE>
</HEAD>

<BODY BgColor="#CCFFCC">
<p align="center" style="margin-top: 0; margin-bottom: 2">�@</p>
<p align="center" style="margin-top: 0; margin-bottom: 2"><font color="#0099FF">
<b>�z�ҿ�J���X�Ը�Ʀp�U�G</b></font></p>
<HR>
<div align="center">
  <center>
  <Form runat="server">

<TABLE Border="1" Cellpadding="2" Cellspacing="0" style="border-collapse: collapse; border: 3px outset #808080" bordercolor="#111111">
 
  <TR>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080">
     <font color="#800000">�m�W�G</font></TD>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080"><%=Request("�m�W")%>�@</TD>
  </TR>
  
  <TR>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080">
     <font color="#800000">����G</font></TD>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080"><%=Request("���")%>�@</TD>
  </TR>
  

  <TR>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080">
     <font color="#800000">�ɶ��G</font></TD>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080"><%=Request("�ɶ�")%>�@</TD>
  </TR>

 <TR>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080">
     <font color="#800000">�Ƶ��G</font></TD>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080"><%=Request("�Ƶ�")%>�@</TD>
  </TR>
  <TR>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080">
     <font color="#800000"></font></TD>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080"></TD>
  </TR>

</TABLE>


  <asp:Button runat="server" Text=" �^ �W �� " OnClick="InsertData" />
                               <asp:Label runat="server" id="Msg" ForeColor="Red" />
                               
</Form>

  </center>
</div>
<hr>
<p align="right">�F�ɬ�޺������X�Ԩt�� V1.1</p>
</BODY>
</HTML>
<script Language="VB" runat="server">
   Sub InsertData(sender As Object, e As EventArgs) 
       Response.Redirect("work.aspx")
   End Sub
   
   Sub Page_Load(sender As Object, e As EventArgs)
       Dim mTime As Date
       mTime = Format(now,"Long Time")
       Dim mH as Integer
       mH = Hour(now)
       'CDate("13:01:02")
       
       Response.Write("<align=" & "center" & ">")
       Response.Write("<font size=6 color=#808000 >")

       If mH > 13 then
          Response.Write("�P�±z�@�Ѫ����� ! ���� !")
       Else
          Response.Write("���W�n ! ���z���ӥR�����O���@�� !")
       End If
       
       Response.Write("</font>")
        Response.Write("</center>")
   End Sub

</script>