//將SDateForm輸入之參數先檢查一次

function strcmp(item1,item2){
  eval('obj1=document.forms.SDateForm.'+item1);
  eval('obj2=document.forms.SDateForm.'+item2);
  if(obj1.value==obj2.value) return true;
  else return false
}

function strlen(item){
  eval('obj=document.forms.SDateForm.'+item);
  return obj.value.length;
}

function lt(item,min){
  eval('obj=document.forms.SDateForm.'+item);
  if(obj.value.length<min) return false;
  else return true;
}

function gt(item,max){
  eval('obj=document.forms.SDateForm.'+item);
  if(obj.value.length>max) return false;
  else return true;
}

function is_alpha(item){
  var cm="abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
  eval('obj=document.forms.SDateForm.'+item);
  if(obj.value.length<=0) return false;
  var chkStr=obj.value;
  for(var i=0;i<chkStr.length;i++){
     cmChar=chkStr.substring(i,i+1);
     if(cm.indexOf(cmChar)<0){
      return false;
     }
  }
  return true;  
}

function is_num(item){
  var cm="0123456789";
  eval('obj=document.forms.SDateForm.'+item);
  if(obj.value.length<=0) return false;
  var chkStr=obj.value;
  for(var i=0;i<chkStr.length;i++){
     cmChar=chkStr.substring(i,i+1);
     if(cm.indexOf(cmChar)<0){
      return false;
     }
  }
  return true;  
}

function is_alphanumeric(item){
  var cm="0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
  eval('obj=document.forms.SDateForm.'+item);
  if(obj.value.length<=0) return false;
  var chkStr=obj.value;
  for(var i=0;i<chkStr.length;i++){
     cmChar=chkStr.substring(i,i+1);
     if(cm.indexOf(cmChar)<0){
      return false;
     }
  }
  return true;
}       

function has_text(item) {
  eval('obj=document.forms.SDateForm.'+item);
  if(obj.value.length>0) return true;
  else return false;
} 

function checked(item) {
  eval('obj=document.forms.SDateForm.'+item);
  for(i=0;i<obj.length;i++){
    if(obj[i].checked) return true;
  }
  return false;
}


function selected(item,index){
  eval('obj=document.forms.SDateForm.'+item);
  if(index.length>0){
    if(obj.options.selectedIndex>index) return true;
    else return false;
  }else{
    if(obj.value.length>0)return true;
    else return false;
  }
}    

function valid_email(item) {
  eval('obj=document.forms.SDateForm.'+item);
  if (obj.value.indexOf("@")<1) return false;
  if (obj.value.indexOf(".")<3)return false;
  return true;
}

function ids() {
  this.A=1;this.B=0;this.C=9;this.D=8;this.E=7;
  this.F=6;this.G=5;this.H=4;this.J=3;this.K=2;
  this.L=2;this.M=1;this.N=0;this.P=9;this.Q=8;
  this.R=7;this.S=6;this.T=5;this.U=4;this.V=3;
  this.X=3;this.Y=2;this.W=1;this.Z=0;this.I=9;this.O=8;
}

function checkid(item) {
  eval('obj=document.forms.SDateForm.'+item);
  var cm="abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
  var num="0123456789";
  var chkStr=obj.value;

  if(obj.value.length != 10)  {
    return false;
  }
  for(var i=1;i<chkStr.length;i++){
     cmChar=chkStr.substring(i,i+1);
     if(num.indexOf(cmChar)<0){
      return false;
     }
  }
  cmChar=chkStr.substring(0,1);
  if(num.indexOf(cmChar)<0){
  return false;
  }
return true;
}
