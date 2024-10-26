<script language="VBScript" RunAt="Server">
Function check_Pass()
  '取得表單欄位內容
  id= Request.Form("id")
  
  SQLstr = "select * from 經銷商 where id='" & id & "'"
  '---------------------------------------------
  'Step 1：建立ADO Connection，然後開啟SQL或Access資料庫
  set adocon=Server.CreateObject("ADODB.Connection")
  adocon.Open "Driver={Microsoft Access Driver(*.mdb)}; DBQ="& Server.MapPath("sample.mdb")
  
  set sa1=adocon.Execute(SQLstr)
  '---------------------------------------------
 
  If sa1.EOF Then  'id找不到，表示尚未註冊
     check_Pass = False
  Else
     check_Pass = True
  End If
  '---------------------------------------------
   sa1.Close
  adocon.Close
End Function


Sub Register()
 '如果已經註冊過的話，則顯示提示訊息
 If check_Pass() Then
    Show_Warning()
 Else  
    response.write "輸入資料不正確!"
 End If
End Sub

Sub Show_Warning()
  Response.Write("<h1><hr>")
  Response.Write("<center>A~HA...<p>")
  Response.Write("該會員已註冊，請返回前一畫面！")
  Response.Write("</center>")
  Response.Write("</h1><hr>")
end sub
</script>
