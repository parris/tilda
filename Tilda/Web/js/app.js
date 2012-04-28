function resize() {
    var maxWidth = $(window).width();
    var maxHeight = $(window).height();
    var width = 1;
    var height = 1;
    if (maxWidth < 435)
        return;
    if (maxHeight < 330)
        return;

    if (maxWidth > maxHeight) {
        height = maxHeight;
        width = Math.round((4*height)/3);
    } else {
        var width = maxWidth;
        var height = Math.round((3*width)/4);
    }
    $("#"+window.element).width(width).height(height);
    window.paper.setSize(width, height);
    window.paper.setViewBox(0, 0, 1024, 768, true);
}

$(function(){
    $(window).resize(function() {
        resize();
    });
    window.element = "holder";
    window.paper = Raphael(window.element, "1024", "768");
    window.paper.setStart();

    window.shapes = new Array();window.animations = new Array();
    
    runSlide();
    
    function checkNextAnimations(){
        if(window.animations.length!=0){
            setTimeout(playNextAnimation,window.animations[0].delay);
        }
    }
    function playNextAnimation(){
        var anim = window.animations.shift();
        for(var i=0;i<anim.ids.length;i++)
            if(i==anim.ids.length-1)
                shapes[anim.ids[i]].animate({'fill-opacity':1,'stroke-opacity':1},anim.dur,checkNextAnimations);
            else 
                shapes[anim.ids[i]].animate({'fill-opacity':1,'stroke-opacity':1},anim.dur);
    }
    checkNextAnimations();

    window.paper.setFinish();
    resize();
});