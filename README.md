What works:
- Text Export, with some alignment features
- All other shapes export as images 
- Opacity based animations work, but only in "after previous" mode and all are simplified to a simple linear transition   
- Z-ordering  
- Presentation export  
- Single Level Bulleting (bullet char not preserved yet)  

What needs work:
- https://www.pivotaltracker.com/projects/536345/stories 

Issues:
- http://code.google.com/p/chromium/issues/detail?id=125608
   - In chrome, resizing is a little funky for some reason. This does not occur in firefox or IE. I believe their implementation of svg is having some issues.

Tilda
=====

A powerpoint to SVG/VML/JS/CSS/HTML &quot;compiler&quot;. 
Tilda has multiple components, the Powerpoint(PPT) add-in itself which allows you to grab a specific slide or all slides. RaphaelJS and the resizing script to make the slide appear "fullbrowser" much like how swfs are rendered fullbrowser. A test project to test all the functionality of C# functionality of Tilda. We will most likely in the future add Jasmine to perform JS testing. 


FAQs
=====
Q: What will Tilda do for me?  
A: Tilda aims to collect all the data from a slide, then, as accurately as possible, will convert that slide data to javascript code that interacts with RaphaelJS. RaphaelJS is a library that creates cross-browser vector graphics and animations using SVG and VML. 

Q: Can I work on Tilda?  
A: Yes PLEASE!!! I need help :).

Q: How far along are you?  
A: Text is able to be grabbed, line breaks/wraps preserved. We have a script to pull out any mp4s, mp3s, jpgs from a ppt. We can do "Fade" animations on various shapes. We have unit tests with some mocks.

Q: What do you need help with?  
A: https://www.pivotaltracker.com/projects/536345/stories , We will be using pivotaltracker to keep schedule features. This should help us see what is complete, what needs to be done, and track our ideas as well!
  
Q: Why is it called Tilda?  
A: Throughout the code before completely converting to JS we use '~' as a form of simple markup various shapes/animation. Text conversion from PPT to Raphael is actually not very easy; especially when you are trying to maintain animations across lines in a textbox. Also tilda is only 5 characters, the most undervalued typable character, and works great as a delimeter since no one even knows what it is.

