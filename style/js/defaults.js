$(function() {
	var body_height = parseInt($(window).height());
	$('.body-left').css({"height": body_height + "px"});
	
	$('.btu-to-top').click(function(){
		$('body').animate({ scrollTop: 0 }, 120);
	});
	
	$('.over-img').hover( function(){ changeImg( $(this) ,1); }, function(){ changeImg( $(this) ,0); });
	function changeImg(obj,ty){
		   var src  = obj.attr('src');
		   var src1 = obj.attr('src1');
		   var src2 = obj.attr('src2');
		   if(src1==undefined||src1==null){ obj.attr('src1',src); }
		   if(ty==0){ obj.attr('src',src1); }else{ obj.attr('src',src2); }
	}
});