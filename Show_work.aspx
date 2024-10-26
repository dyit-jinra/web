
<HTML>
<HEAD>
      <TITLE> 網路版出勤系統結果顯示畫面 </TITLE>
</HEAD>

<BODY BgColor="#CCFFCC">
<p align="center" style="margin-top: 0; margin-bottom: 2">　</p>
<p align="center" style="margin-top: 0; margin-bottom: 2"><font color="#0099FF">
<b>您所輸入的出勤資料如下：</b></font></p>
<HR>
<div align="center">
  <center>
  <Form runat="server">

<TABLE Border="1" Cellpadding="2" Cellspacing="0" style="border-collapse: collapse; border: 3px outset #808080" bordercolor="#111111">
 
  <TR>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080">
     <font color="#800000">姓名：</font></TD>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080"><%=Request("姓名")%>　</TD>
  </TR>
  
  <TR>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080">
     <font color="#800000">日期：</font></TD>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080"><%=Request("日期")%>　</TD>
  </TR>
  

  <TR>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080">
     <font color="#800000">時間：</font></TD>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080"><%=Request("時間")%>　</TD>
  </TR>

 <TR>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080">
     <font color="#800000">備註：</font></TD>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080"><%=Request("備註")%>　</TD>
  </TR>
  <TR>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080">
     <font color="#800000"></font></TD>
     <TD style="border-left: 1px outset #808080; border-right: 1px outset #808080; border-top: 1px solid #808080; border-bottom: 1px outset #808080"></TD>
  </TR>

</TABLE>


  <asp:Button runat="server" Text=" 回 上 頁 " OnClick="InsertData" />
                               <asp:Label runat="server" id="Msg" ForeColor="Red" />
                               
</Form>

  </center>
</div>
<hr>
<p align="right">東怡科技網路版出勤系統 V1.1</p>
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
          Response.Write("感謝您一天的辛勞 ! 謝謝 !")
       Else
          Response.Write("早上好 ! 祝您有個充滿活力的一天 !")
       End If
       
       Response.Write("</font>")
        Response.Write("</center>")
   End Sub

</script>