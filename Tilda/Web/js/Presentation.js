function Presentation(element, width, height) {
    this.element = element;
    this.paper = Raphael(this.element, width, height);
    this.shapes = new Array();
    this.animations = new Array();
    this.slides = new Array();
    this.currentSlideSettings = new Object();
    this.currentSlide = 0;
}

Presentation.prototype.setUpAnimations = function() {
    var timingMode = this.currentSlideSettings.timingMode;
    if (typeof timingMode === "undefined") {
        return false;
    }

    if (this.currentSlideSettings.timingMode == "audio") {
        //first change the timings
        var timeTotal = 0;
        for (var i = 0; i < preso.animations.length; i++) {
            preso.animations[i].delay = preso.animations[i].delay / 1000;
        }

        //attach timeupdate handler to jplayer
        $("#audio-player").bind($.jPlayer.event.timeupdate + ".jp-player", function(event) {
            if (event.jPlayer.status.currentTime > preso.animations[0].delay) {
                var anim = preso.animations.shift();
                for (var i = 0; i < anim.ids.length; i++) {
                    preso.shapes[anim.ids[i]].animate(anim.animate, anim.dur);
                }
            }
        });
    } else {
        //otherwise use internal timer
        this.checkNextAnimations();
    }
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

Presentation.prototype.play = function(slideNum) {
    this.clearSlide();
    if (typeof slideNum !== "undefined")
        this.currentSlideSettings = this.slides[parseInt(slideNum)]();
    else
        this.currentSlideSettings = this.slides[0]();
    this.currentSlide = 0;
    this.setUpAnimations();
}

Presentation.prototype.next = function() {
    if (this.slides.length > this.currentSlide + 1) {
        this.currentSlide++;
        this.clearSlide();
        this.currentSlideSettings = this.slides[this.currentSlide]();
        this.setUpAnimations();

    }
}

Presentation.prototype.prev = function() {
    if (this.currentSlide - 1 > 0) {
        this.currentSlide--;
        this.clearSlide();
        this.currentSlideSettings = this.slides[this.currentSlide]();
        this.setUpAnimations();
    }
}

Presentation.prototype.clearSlide = function() {
    this.paper.clear();
    this.shapes = new Array();
    this.animations = new Array();
}

/**
* Static function to resize the window
*/
Presentation.resize = function() {
    var maxWidth = $(window).width();
    var maxHeight = $(window).height();
    var width = 1;
    var height = 1;

    //aspect ratio
    var ratio = Presentation.gcd(window.settings.width, window.settings.height);
    var windowRatio = Presentation.gcd(maxWidth, maxHeight);
    var xRatio = window.settings.width / ratio;
    var yRatio = window.settings.height / ratio;

    var widthCalc = Math.round((xRatio * maxHeight) / yRatio);

    if (maxWidth > maxHeight && widthCalc <= maxWidth) {
        height = maxHeight;
        width = widthCalc;
    } else {
        width = maxWidth;
        height = Math.round((yRatio * width) / xRatio);
    }

    $("#" + preso.element).width(width).height(height);
    preso.paper.setSize(width, height);
    preso.paper.setViewBox(0, 0, window.settings.width, window.settings.height, true);
}

Presentation.gcd = function(a, b) {
    return (b == 0) ? a : Presentation.gcd(b, a % b);
}

Presentation.getFromUrl = function(key, queryStr, delim, equal) {
    if (queryStr == null)
        queryStr = window.location.search;
    if (queryStr.indexOf("?") == 0)
        queryStr = queryStr.substring(1, queryStr.length);
    if (delim == null)
        delim = "&"
    if (equal == null)
        equal = "="

    if (queryStr.indexOf(key) != -1) {
        var ary1 = queryStr.split(delim);
        for (var i = 0; i < ary1.length; i++) {
            var ary2 = ary1[i].split(equal)
            if (ary2[0] == key) {
                return ary1[i].substring((ary1[i].indexOf(equal) + 1), ary1[i].length);
            }
        }
    }
    return null;
}

$(function() {
    window.preso = new Presentation("holder", window.settings.width, window.settings.height);
    $.getScript("content.js", function() {
        preso.paper.setStart();
        var startOn = Presentation.getFromUrl("slide");
        if (startOn == null) //start up first slide
            preso.play();
        else
            preso.play(startOn);
        //wrap up the paper
        preso.paper.setFinish();
        Presentation.resize();
        $(window).resize(function() {
            Presentation.resize();
        });
    });
});