/* PPF Smart Farm — Service Worker (ใช้งานออฟไลน์หน้าฟาร์มเน็ตไม่เสถียร)
   - vendor/รูป = cache-first (ไฟล์ไม่เปลี่ยน โหลดเร็ว)
   - html/js อื่น = network-first (ออนไลน์ได้ของใหม่เสมอ · ออฟไลน์ใช้แคช)
   ⚠️ เวลาแก้โค้ดแล้ว deploy ใหม่ ให้เพิ่มเลขเวอร์ชัน CACHE ด้านล่าง เพื่อล้างแคชเก่า */
const CACHE='ppf-farmfan-v6';
const ASSETS=[
  './','./index.html','./analytics.html','./mortality.html','./fan-plan.html','./parse-worker.js',
  './vendor/xlsx.full.min.js','./vendor/chart.umd.min.js','./vendor/html2canvas.min.js',
  './home-render.jpg','./manifest.webmanifest','./icon.svg'
];
self.addEventListener('install',e=>{
  e.waitUntil(caches.open(CACHE).then(c=>Promise.allSettled(ASSETS.map(u=>c.add(u)))).then(()=>self.skipWaiting()));
});
self.addEventListener('activate',e=>{
  e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));
});
self.addEventListener('fetch',e=>{
  const req=e.request;
  if(req.method!=='GET')return;                          // POST (ซิงค์ขึ้นชีท) → ปล่อยผ่าน
  const url=new URL(req.url);
  if(url.origin!==location.origin)return;                // Apps Script/ฟอนต์/CDN → ปล่อยผ่าน
  const cacheFirst=/\/vendor\/|\.(jpg|jpeg|png|svg|webp|woff2?|ttf)$/i.test(url.pathname);
  if(cacheFirst){
    e.respondWith(caches.match(req).then(hit=>hit||fetch(req).then(res=>{const cp=res.clone();caches.open(CACHE).then(c=>c.put(req,cp));return res;})));
  }else{
    e.respondWith(fetch(req).then(res=>{if(res&&res.ok){const cp=res.clone();caches.open(CACHE).then(c=>c.put(req,cp));}return res;}).catch(()=>caches.match(req).then(hit=>hit||caches.match('./index.html'))));
  }
});
