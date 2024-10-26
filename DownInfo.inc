<script language="VBScript" RunAt="Server">
Function check_Pass()
  dim SQLstr,sa1,adocon,param 
  company = Request.Form("company")
   SQLstr = "select * from DownInfo where company='" & company & "'"
  '---------------------------------------------
  'Step 1：建立ADO Connection，然後開啟SQL或Access資料庫
   set adocon=Server.CreateObject("ADODB.Connection")
   adocon.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & server.MapPath("DownInfo")
  '---------------------------------------------
  'Step 2：執行SQL指令，並將結果儲存於Recordse中
  set sa1=adocon.Execute(SQLstr)
  '---------------------------------------------
  'Step 4：關閉RecordSet及資料庫
  sa1.Close
  adocon.Close
End Function

Function adjust_sql(adj_str)
  '目的：轉換查詢字串中的特殊字元至SQL指令
  Dim final_str, i
  '
  adj_str = Trim(adj_str)
  final_str = ""
  '每次從傳進來的字串中取出一個字元，並判斷是否為下述三個
  '字元之一：左中括號"["、DOS的Filter符號"|"以及單引號"'"
  If Len(adj_str) > 0 Then
     For i = 1 To Len(adj_str)
       Select Case Mid(adj_str, i, 1)
          Case "[":
            final_str = final_str & "[[]"
          Case "|":
            final_str = final_str & "[{-}]"
          Case "'":
            final_str = final_str & "[&-(]"
          Case Else:  '以上皆非則什麼事都不做
            final_str = final_str & Mid(adj_str, i, 1)
       End Select
     Next
  End If
  adjust_sql = final_str
End Function

Sub Add_Reg()
     
   Dim company, name, idno, email, address
  
   Dim phone, fax,  sdate, chk_info
   Dim SQLstr
   '-----------------------------------------------
   On Error Resume Next
   '取出表單欄位到變數中
   company = Request.Form("company")
   name = Request.Form("name")
   address = Request.Form("address")
   no = Request.Form("no")
   phone = Request.Form("phone")
   fax = Request.Form("fax")
   'mail = Request.Form("mail")
      
   sdate = Request("syear") & Request("smonth") & Request("sday")
   '-----------------------------------------------
   '調整SQL指令中的特定字元
   company = adjust_sql(company)
  
   name = adjust_sql(name)
   no = adjust_sql(no)
   address = adjust_sql(address)
   phone = adjust_sql(phone)
   fax = adjust_sql(fax)
   '-----------------------------------------------
   SQLstr = "insert into DownInfo values ('" & company
   SQLstr = SQLstr & "','" & fax & "','" & no
   SQLstr = SQLstr & "','" & name & "','" & address & "','"
   SQLstr = SQLstr & "','" & phone & "','" & sdate & "')"
   '-----------------------------------------------
   'Step 1：建立ADO Connection，然後開啟Access或SQL資料庫
   set adocon=Server.CreateObject("ADODB.Connection")
   adocon.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & server.MapPath("DownInfo")
   '-----------------------------------------------
 
  
   '若有錯誤發生
   If Err.Number <> 0 Then
      Response.Write "錯誤號碼：" & Err.Number & "<br>"
      Response.Write "錯誤程式：" & Err.Source & "<br>"
      Response.Write "錯誤訊息：" & Err.Description & "<br>"
   Else  '將成功訊息輸出到瀏覽器中
      Response.Write("<hr><h1>")
      Response.Write("<center>恭禧！<p>")
      Response.Write("註冊成功，請回前一畫面！")
      Response.Write("</center>")
      Response.Write("</h1><hr>")
   End If
   '-----------------------------------------------
   'Step 3：關閉資料庫
   adocon.Close
   '-----------------------------------------------
End Sub


</script>
