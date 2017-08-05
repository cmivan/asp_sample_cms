/* 彦磊基于jQuery浮动对联广告插件(2009.04)
 * 在IE6/IE7/Mozilla 5.0（Firefox 3.0.5）中测试通过
 * 此插件要求运行在jQuery v1.3 或更高版本上
 * 使用方法：
 * 
 * 1、在页面中为您要浮动的内容设置容器，如<div id="adl">左侧内容</div>或<div id="adr">右侧内容</div>或<div id="ad">您的内容</div>
 * 2、在页面中添加jQuery语句，调用此插件，并对相关参数进行设置,如：
 * <script type="text/javascript">
 * $(document).ready(function(){
 *  $("#adl").jFloat({
 *     position:"left",
 *     top:0,
 *     height:200,
 *     width:100,
 *     left:20
 *   });//将页面中id为adl的容器中的内容设置为左浮动广告，广告距窗口顶部0px,距左侧20px。
 *  $("#adr").jFloat({
 *     position:"right",
 *     top:0,
 *     height:200,
 *     width:100,
 *     right:20
 *  });//将页面中id为adr的容器中的内容设置为右浮动广告，广告距窗口顶部0px，距右侧20px。
 *  $("#ad").jFloat({
 *     position:"left",
 *     top:260,
 *     width:100,
 *     height:50,
 *     allowClose:false
 *  });//将面面中id为ad的容器中的内容设置为左右浮动（即左右两侧内容样式等一样），广告中容器顶部260px，距左侧0,不允许关闭。
 *});
 * </script>
 * 上面的例子最终结果，将出现四个浮动内容，上面两个分别是adl和adr中的内容,下面两个是ad中的内容。如果您愿意，您可以加无数个这样的广告。
 * 3、插件相关参数：
 *     top－广告距页面顶部距离，默认为60
 *     left－广告距页面左侧距离，默认为0
 *     right－广告距页面右侧距离，默认为0
 *     width－广告容器的宽度，默认为100
 *     height－广告容器的高度，默认为360
*minScreenW－出现广告的最小屏幕宽度，当屏幕分辨率小于此，将不出现对联广告，默认为800，即在800×600分辨率下不会显示广告内容
 *     position－对联广告的位置,left-在左侧出现,right-在右侧出现，默认为"left"。注意要加英文单或双引号。
 *     allowClose－是否允许关闭，如果为true，则会在广告内容上方添加“关闭”，单击时将关闭所在广告内容。值为true或false 
 */


var base_url=base_url;
if(base_url==''||base_url==null){base_url='/';}

(function($) { 
    $.fn.jFloat = function(o) {
    
        o = $.extend({
            top:60,  //广告距页面顶部距离
            left:0,//广告左侧距离
            right:10,//广告右侧距离
            width:100,  //广告容器的宽度
            height:360, //广告容器的高度
            minScreenW:800,//出现广告的最小屏幕宽度，当屏幕分辨率小于此，将不出现对联广告
            position:"left", //对联广告的位置left-在左侧出现,right-在右侧出现
            allowClose:false //是否允许关闭 
        }, o || {});
		var h=o.height;
      var showAd=true;
      var fDiv=$(this);
      if(o.minScreenW>=$(window).width()){
          fDiv.hide();
          showAd=false;
       }
       else{
		   fDiv.css("display","block")
           var closeHtml='<div align="right" style="padding:2px;z-index:2000;font-size:12px;cursor:pointer;height:20px;" class="closeFloat"><span style="border:1px solid #000;height:12px;display:block;width:12px;display:none;">×</span></div>';
           switch(o.position){
               case "left":
                    if(o.allowClose){
                       fDiv.prepend(closeHtml);
					   $(".closeFloat",fDiv).click(function(){$(this).hide();fDiv.hide();showAd=false;})
					   h+=20;
					}
                    fDiv.css({display:"block",position:"absolute",left:o.left+"px",top:o.top+"px",width:o.width+"px",height:h+"px",overflow:"hidden"});
                    break;
               case "right":
                    if(o.allowClose){
                       fDiv.prepend(closeHtml)
					   $(".closeFloat",fDiv).click(function(){$(this).hide();fDiv.hide();showAd=false;})
					   h+=20;
					}
               fDiv.css({display:"block",position:"absolute",left:"auto",right:o.right+"px",top:o.top+"px",width:o.width+"px",height:h+"px",overflow:"hidden"});
               break;
            };
        };
        function ylFloat(){
            if(!showAd){return}
            var windowTop=$(window).scrollTop();
            if(fDiv.css("display")!="none")
				fDiv.css("top",o.top+windowTop+"px");
			   //fDiv.animate({top:o.top+windowTop+"px"}, "slow");
        };
      $(window).scroll(ylFloat);
      $(document).ready(ylFloat);
    }; 
})(jQuery);


//创建浮动框
$(function(){
	$(document.body).append('<div id=\"right_talk\" style=\"display:none;z-index:9999\"><\/div>');
	$("#right_talk").fadeOut(100).fadeIn(1000).load('qq.asp').jFloat({position:"right",top:160,height:275,width:65,right:5});
});