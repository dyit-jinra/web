<script language="VBScript" RunAt="Server">
Function check_Pass()
  '���o������줺�e
  id= Request.Form("id")
  
  SQLstr = "select * from �g�P�� where id='" & id & "'"
  '---------------------------------------------
  'Step 1�G�إ�ADO Connection�A�M��}��SQL��Access��Ʈw
  set adocon=Server.CreateObject("ADODB.Connection")
  adocon.Open "Driver={Microsoft Access Driver(*.mdb)}; DBQ="& Server.MapPath("sample.mdb")
  
  set sa1=adocon.Execute(SQLstr)
  '---------------------------------------------
 
  If sa1.EOF Then  'id�䤣��A���ܩ|�����U
     check_Pass = False
  Else
     check_Pass = True
  End If
  '---------------------------------------------
   sa1.Close
  adocon.Close
End Function


Sub Register()
 '�p�G�w�g���U�L���ܡA�h��ܴ��ܰT��
 If check_Pass() Then
    Show_Warning()
 Else  
    response.write "��J��Ƥ����T!"
 End If
End Sub

Sub Show_Warning()
  Response.Write("<h1><hr>")
  Response.Write("<center>A~HA...<p>")
  Response.Write("�ӷ|���w���U�A�Ъ�^�e�@�e���I")
  Response.Write("</center>")
  Response.Write("</h1><hr>")
end sub
</script>
