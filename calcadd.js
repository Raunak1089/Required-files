var mydiv = document.createElement('div'); 
mydiv.style.height = '50px';
mydiv.style.width = 'fit-content';
mydiv.style.background = 'grey';
mydiv.style.borderRadius = '10px';
mydiv.style.padding = '10px';
mydiv.style.position = 'absolute';
mydiv.style.opacity = 0.7;
mydiv.style.flex = '1';
mydiv.style.zIndex = '1000';
mydiv.style.transition = 'opacity 0.3s';
mydiv.style.top = '0px';


var myin = document.createElement('input'); 
myin.style.width = '70px';
myin.style.background = '#34ebd5';
var mybtn = document.createElement('input'); 
mybtn.type = 'submit';
mybtn.value = '=';

var myans = document.createElement('span');
myans.style.background = 'white';
myans.style.padding = '2px 10px';


mybtn.onclick = function (){myans.innerHTML=eval(myin.value)};



mydiv.appendChild(myin);
mydiv.appendChild(mybtn);
mydiv.appendChild(myans);
document.body.appendChild(mydiv);




mydiv.ontouchmove = function(ev) {
           var e = ev.targetTouches[0];
           mydiv.style.left = e.pageX - 50 + 'px'; 
           mydiv.style.top = e.pageY - 30 + 'px';
           
           mydiv.style.transition = 'opacity 0.3s';
           window.scrollTo(scrollLeft, scrollTop);
           }



