<!-- #include file="DownInfo.inc" -->
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>新增網頁1</title>
</head>

<body bgcolor="#FFFFFF">


<blockquote>
  <blockquote>
    <blockquote>
      <blockquote>


<p style="margin-top:10; margin-bottom:10; line-height:150%" align="left">
<font size="4" color="#800000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 資料已送出 ， 我 們 將 儘 快 
與 您 聯 絡! 或請來電 (06)208-8808  (04)2260-0001</font></p>


<%
   dim buf 
	On error resume next

	Dim mySmartMail
	Set mySmartMail = Server.CreateObject("aspSmartMail.SmartMail")

'	Mail Server
'	***********
	mySmartMail.Server = "dyit.com.tw"

'	From
'	****
	mySmartMail.SenderName = request("company") & "-" &  request("name") & " TEL:" & request("phone")
	mySmartMail.SenderAddress = "dyit01@dyit.com.tw"

'	To
'	**
	mySmartMail.Recipients.Add "dyit01@dyit.com.tw", "東怡科技有限公司"

'	Message
'	*******
	mySmartMail.Subject = "尋求試用" & request("product") & request("address")
'	buf =  "公司:" & Company & "姓名:" & Name  & "地址:" & Address & "電話:" & phone & "行業:" & Item & "索取軟體:" & product
 	mySmartMail.Body = "132sss" 
 	
'	Send the message
'	****************
	mySmartMail.SendMail

	if Err.Number<>0 then

		Response.write "Error: " & Err.description

	else

		
	end if

%>


</p>


      </blockquote>
    </blockquote>
  </blockquote>
</blockquote>


</body>
</body>

</html>