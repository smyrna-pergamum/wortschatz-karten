/* Wortschatz Karten tanıtım sayfası. Dile özgü metinler sayfadaki window.T nesnesinden gelir (build.py üretir). */
document.documentElement.classList.add('js');
const T=window.T;
const $=s=>document.querySelector(s);
const reduce=matchMedia('(prefers-reduced-motion: reduce)').matches;
const base=document.documentElement.dataset.assets||'assets/';
const screens={light:base+'app-light.png',dark:base+'app-dark.png',stats:base+'stats.png'};
let currentTheme='light',flipped=false,previousFocus=null;

/* Seçilen dili hatırla (kök sayfadaki yönlendirme için) */
document.querySelectorAll('.lang a').forEach(a=>a.addEventListener('click',()=>{try{localStorage.setItem('wk-lang',a.hreflang)}catch(e){}}));

/* Kaynayan çizgi: çizgi film karelerindeki titreşim */
if(!reduce){const noise=$('#boil-noise');let seed=1;setInterval(()=>{seed=seed%4+1;noise.setAttribute('seed',seed)},120)}

/* Kayan şerit (Almanca kelimeler, her dilde aynı) */
const words=['DER DIALEKT','DIE MUNDART','DAS WORT','WUSCH!','DER WORTSCHATZ','DIE KARTE','ZACK!','DAS BEISPIEL','DIE WIEDERHOLUNG','PENG!'];
$('#ticker').innerHTML=[...words,...words].map(w=>'<span>'+w+' ★</span>').join('');

/* TV kanal değiştirme */
function setTheme(theme){if(theme===currentTheme)return;currentTheme=theme;const dark=theme==='dark';const scr=$('#main-zoom');scr.classList.remove('switching');void scr.offsetWidth;scr.classList.add('switching');
  $('#main-screenshot').src=screens[theme];$('#main-screenshot').alt=dark?T.altDark:T.altLight;
  $('#light-btn').setAttribute('aria-pressed',String(!dark));$('#dark-btn').setAttribute('aria-pressed',String(dark));
  $('#screen-caption').textContent=dark?T.capDark:T.capLight}
$('#light-btn').onclick=()=>setTheme('light');$('#dark-btn').onclick=()=>setTheme('dark');

/* Kart çevirme + toz bulutu */
function puffs(el){if(reduce)return;const r=el.getBoundingClientRect(),host=el.parentElement,hr=host.getBoundingClientRect();
  for(let i=0;i<7;i++){const p=document.createElement('span');p.className='puff';const a=Math.PI*2*i/7;
    p.style.left=(r.left-hr.left+r.width/2+Math.cos(a)*r.width*.42)+'px';p.style.top=(r.top-hr.top+r.height/2+Math.sin(a)*r.height*.42)+'px';
    p.style.setProperty('--dx',Math.cos(a)*50+'px');p.style.setProperty('--dy',Math.sin(a)*40+'px');host.appendChild(p);setTimeout(()=>p.remove(),650)}}
$('#flip-card').onclick=()=>{flipped=!flipped;const f=$('#flipper');f.classList.toggle('flipped',flipped);f.classList.remove('squish');void f.offsetWidth;f.classList.add('squish');puffs(f);
  $('#card-front').setAttribute('aria-hidden',String(flipped));$('#card-back').setAttribute('aria-hidden',String(!flipped));
  $('#card-side-label').textContent=flipped?T.sideBack:T.sideFront;
  $('#flip-card').textContent=flipped?T.flipBack:T.flip;$('#flip-card').setAttribute('aria-pressed',String(flipped));
  $('#card-live').textContent=flipped?T.liveBack:T.liveFront};

/* Leitner kutuları */
document.querySelectorAll('[data-box]').forEach(button=>button.onclick=()=>{const n=Number(button.dataset.box);
  document.querySelectorAll('[data-box]').forEach(b=>b.setAttribute('aria-pressed',String(b===button)));
  $('#box-title').textContent=T.box+' '+n+' · '+T.boxes[n-1][0];$('#box-description').textContent=T.boxes[n-1][1]});

/* Görsel penceresi */
const dialog=$('#image-dialog');
function openImage(key){previousFocus=document.activeElement;$('#dialog-image').src=screens[key];
  $('#dialog-image').alt=key==='stats'?T.altStats:key==='dark'?T.altDark:T.altLight;
  $('#dialog-title').textContent=key==='stats'?T.dlgStats:key==='dark'?T.dlgDark:T.dlgLight;
  dialog.showModal();document.body.style.overflow='hidden';$('#close-dialog').focus()}
function closeImage(){dialog.close();document.body.style.overflow='';if(previousFocus?.isConnected)previousFocus.focus()}
$('#main-zoom').onclick=()=>openImage(currentTheme);$('#stats-zoom').onclick=()=>openImage('stats');$('#close-dialog').onclick=closeImage;
dialog.addEventListener('cancel',e=>{e.preventDefault();closeImage()});
dialog.addEventListener('click',e=>{if(e.target===dialog){const r=dialog.getBoundingClientRect();if(e.clientX<r.left||e.clientX>r.right||e.clientY<r.top||e.clientY>r.bottom)closeImage()}});

/* İris açılışı */
const iris=$('#iris');
function checkIris(){const r=iris.getBoundingClientRect();if(r.top<innerHeight*.8&&r.bottom>0){iris.classList.add('in');removeEventListener('scroll',checkIris)}}
if(reduce)iris.classList.add('in');else{addEventListener('scroll',checkIris,{passive:true});checkIris()}

/* Tıklama efektleri: Almanca çizgi roman sesleri */
const zaps=['ZACK!','PENG!','WUSCH!','BOING!','KLICK!','ZOING!','HUI!'];
if(!reduce)addEventListener('pointerdown',e=>{if(e.target.closest('dialog'))return;const z=document.createElement('span');z.className='zap';z.textContent=zaps[Math.random()*zaps.length|0];
  z.style.left=e.clientX+'px';z.style.top=e.clientY+'px';z.style.setProperty('--r',(Math.random()*30-15)+'deg');document.body.appendChild(z);setTimeout(()=>z.remove(),720)});
