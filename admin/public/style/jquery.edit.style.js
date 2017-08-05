//定义鼠标移过样式
$(function(){
	
  //全选或取消列表项
  $("#delsel").click(function(){
	 var thischeck = $(this).attr("checked");
	 thischeck = thischeck ? true : false;
	 $(".delitem").attr("checked",thischeck);
  });
	 
  $("#del_checkbox").click(function(){
	 var thischeck = $(this).attr("checked");
	 $(".del_id").attr("checked",thischeck);
  });

  $("#Submit_delsel").click(function(){
	 if(ischecked()){
		 if(confirm("确定要删除选中项?")){
			 return true;
		 }else{
			 return false;	
		 }
	 }else{
		 return false;
	 }
  });
  
  //提示是否操作按钮
  $(".cmbtn").click(function(){
	  var title = $(this).attr("title");
	  var url = $(this).attr("url");
	  if(title!="" && url!=""){
		  if( confirm(title) ){ window.location.href = url; }
	  }
	  return false;
  });
  
  //加载验证码
  $("#verifycodeImg").click(function(){
	  var thisImg = $(this).find("img");
	  var src = thisImg.attr("src");
	  var tempSrc = thisImg.attr("tempSrc");
	  if(tempSrc==''||tempSrc==null){
		  thisImg.attr("tempSrc",src); 
		  tempSrc = src;
	  }
	  thisImg.attr("src",tempSrc + "?T=" + Math.random() ); 
  });


	
});


/*判断是否选中项*/
function ischecked(){
	var thisis=false;
	$(".del_id").each(function(){
		if(thisis==false){
			var thischecked=$(this).attr("checked");
			if(thischecked){thisis=true;}
			}
		});
	if(thisis==false){alert("至少要选择一项!");}
	return thisis;
	}