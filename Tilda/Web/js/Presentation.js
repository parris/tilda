function Presentation(element, width, height) {
    this.element = element;
    this.paper = Raphael(this.element, width, height);
    this.shapes = new Array();
    this.animations = new Array();
    this.slides = new Array();
    this.currentSlide = 0;
}

Presentation.prototype.checkNextAnimations = function() {
    if (preso.animations.length != 0) {
        setTimeout(preso.playNextAnimation, preso.animations[0].delay);
    }
}

Presentation.prototype.playNextAnimation = function() {
    var anim = preso.animations.shift();
    for (var i = 0; i < anim.ids.length; i++)
        if (i == anim.ids.length - 1)
            preso.shapes[anim.ids[i]].animate(anim.animate, anim.dur, preso.checkNextAnimations);
        else
            preso.shapes[anim.ids[i]].animate(anim.animate, anim.dur);
}

Presentation.prototype.play = function() {
    this.slides[0]();
    this.currentSlide = 0;
}

Presentation.prototype.next = function() {
    if (this.slides.length > currentSlide + 1) {
        this.currentSlide++;
        this.paper.clear();
        this.slides[this.currentSlide]();
    }
}

Presentation.prototype.prev = function() {
    if (currentSlide - 1 > 0) {
        this.currentSlide--;
        this.paper.clear();
        this.slides[this.currentSlide]();
    }
}

/**
* Static function to resize the window
*/
Presentation.resize = function() {
    var maxWidth = $(window).width();
    var maxHeight = $(window).height();
    var width = 1;
    var height = 1;

    if (maxWidth > maxHeight) {
        height = maxHeight;
        width = Math.round((4 * height) / 3);
    } else {
        var width = maxWidth;
        var height = Math.round((3 * width) / 4);
    }
    $("#" + window.element).width(width).height(height);
    preso.paper.setSize(width, height);
    preso.paper.setViewBox(0, 0, 1024, 768, true);
}

$(function() {
    $(window).resize(function() {
        Presentation.resize();
    });
    window.preso = new Presentation("holder", "1024", "768");
    $.getScript("content.js", function() {
        preso.paper.setStart();
        //start up first slide
        preso.play();
        //run animations
        preso.checkNextAnimations();
        //wrap up the paper
        preso.paper.setFinish();
        Presentation.resize();
    });
});