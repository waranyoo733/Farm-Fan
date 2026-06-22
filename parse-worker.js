/* Web Worker: อ่าน+แปลงไฟล์ใบหน้าเล้า .xlsx นอก main thread (กันหน้าค้าง)
   mode 'mortality' → คืนทั้งฟาร์ม (dead/cull + detail จากชีตสรุป)
   mode 'face'      → คืนรายเล้า (num/birds/dead/age) สำหรับนำเข้าในแอปหลัก */
importScripts('https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js');
var RD={type:'array',cellFormula:false,cellHTML:false,cellText:false,cellNF:false,cellStyles:false,sheetStubs:false};
function num(c){return c&&c.t!=='e'&&c.v!=null&&!isNaN(+c.v)?+c.v:0;}  // c.t==='e' = เซลล์ error (#VALUE! ฯลฯ) → ข้าม
function zeros(){return new Array(42).fill(0);}
function nameFromKeywords(txt){
  if(/ชัยนาท/.test(txt))return'ฟาร์ม ชัยนาท';
  if(/ยิ่งรวย/.test(txt))return'ฟาร์ม ยิ่งรวย';
  if(/รุ่งเรือง/.test(txt))return'ฟาร์ม หนึ่งรุ่งเรือง';
  if(/ขวัญใจ/.test(txt))return'ฟาร์ม ขวัญใจ';
  if(/พุ่มวงศ์/.test(txt))return'ฟาร์ม พุ่มวงศ์';
  if(/ดอนทอง/.test(txt))return'ฟาร์ม PPF ดอนทอง';
  return null;
}
function farmTexts(wb){var t=[];wb.SheetNames.forEach(function(sn){var ws=wb.Sheets[sn];if(!ws)return;for(var k in ws){if(k.charAt(0)==='!')continue;var c=ws[k];if(c&&c.t==='s'&&typeof c.v==='string'&&c.v.indexOf('ฟาร์ม')>=0)t.push(c.v);}});return t;}
function smartFarmName(wb,fileName){
  var byFile=nameFromKeywords(fileName||''); if(byFile)return byFile;        // 1) ชื่อไฟล์
  var texts=farmTexts(wb);
  var byContent=nameFromKeywords(texts.join(' ')); if(byContent)return byContent;  // 2) คีย์เวิร์ดในเนื้อไฟล์
  for(var i=0;i<texts.length;i++){var m=texts[i].match(/ฟาร์ม[\s.…]*([ก-๙a-zA-Z]{2,})/);if(m){var nm=m[1].replace(/รุ่น.*$/,'').trim();if(nm.length>=2)return'ฟาร์ม '+nm;}}  // 3) จับชื่อจาก title
  return 'ฟาร์ม '+((fileName||'').replace(/\.xlsx?$/i,'').replace(/ใบหน้าเล้า|ฟาร์ม|รุ่นที่|รุ่น/g,'').replace(/[0-9.()_\-]+/g,' ').replace(/\s+/g,' ').trim()||'นำเข้า');  // 4) จากชื่อไฟล์
}
/* อ่านยอดตายรายวันจากชีต H — รองรับ 2 เทมเพลต:
   (1) แบบเก่า: คอลัมน์ B = วันอายุ 1..42, W = สูญเสีย/วัน (1 แถว/วัน)
   (2) แบบบล็อกรายสัปดาห์ (ไฟล์ PPF ปัจจุบัน ทุกฟาร์ม): วันอายุไม่ได้อยู่คอลัมน์ B
       (อยู่คอลัมน์ A แบบเยื้อง + วันที่คอลัมน์ C เป็นข้อความ/serial) → อ่านยอดตายคอลัมน์ W
       ของ "แถวข้อมูล" เรียงตามลำดับเป็นวัน 1,2,3,...  แถวข้อมูล = W เป็นตัวเลข,
       ไม่ใช่หัวตาราง, ข้ามแถวหัวไฟล์ (rr<6), ข้ามแถวสรุป/รวม (A>=200 หรือมีค่าในคอลัมน์ Z) */
function readHouseDead(ws){
  var dead=zeros(),last=0;
  var rng=XLSX.utils.decode_range(ws['!ref']); var maxR=Math.min(rng.e.r,rng.s.r+1000);
  // วิธี (1) คอลัมน์ B = วันอายุ
  var usedB=false;
  for(var r=rng.s.r;r<=maxR;r++){
    var rr=r+1, bc=ws['B'+rr]; if(!bc||bc.t==='e'||typeof bc.v!=='number')continue;
    var day=bc.v; if(!(Number.isInteger(day)&&day>=1&&day<=42))continue;
    var wc=ws['W'+rr]; if(!wc||wc.t==='e'||typeof wc.v!=='number')continue;
    dead[day-1]=Math.round(wc.v); if(wc.v>0)last=Math.max(last,day); usedB=true;
  }
  if(usedB)return {dead:dead,last:last};
  // วิธี (2) บล็อกรายสัปดาห์ — อ่าน W ตามลำดับแถวข้อมูล
  var di=0;
  for(var r2=rng.s.r;r2<=maxR;r2++){
    var rr2=r2+1; if(rr2<6)continue;
    var w=ws['W'+rr2]; if(!w||w.t==='s'||w.t==='e'||typeof w.v!=='number')continue;  // ข้ามหัวตาราง(ข้อความ)
    var a=ws['A'+rr2]; if(a&&typeof a.v==='number'&&a.v>=200)continue;               // ข้ามแถวรวม (A=260,261,...)
    var z=ws['Z'+rr2]; if(z&&z.v!=null&&z.v!=='')continue;                           // ข้ามแถวสะสม (มี ยอดคงเหลือ/ratio)
    if(di<42){dead[di]=Math.round(w.v); if(w.v>0)last=di+1; di++;}
  }
  return {dead:dead,last:last};
}
function parseWB(wb,name){
  var houses=[];
  wb.SheetNames.forEach(function(sn){
    if(!/^H\s*\d+$/.test(sn))return;
    var ws=wb.Sheets[sn]; if(!ws||!ws['!ref'])return;
    var birds=Math.round(num(ws['W4'])); if(birds<=1000)return;
    var d=readHouseDead(ws);
    houses.push({name:'เล้า '+sn.replace(/[^0-9]/g,''),birds:birds,dead:d.dead,cull:zeros(),age:d.last});
  });
  return houses.length?{name:name,houses:houses}:null;
}
function parseSummary(wb){
  var sn=wb.SheetNames.find(function(s){return /ประมาณการ|คงเหลือจับ/.test(s);}); if(!sn)return {};
  var ws=wb.Sheets[sn]; if(!ws||!ws['!ref'])return {};
  var rng=XLSX.utils.decode_range(ws['!ref']);
  var get=function(col,r){var c=ws[col+r];return c&&c.t!=='e'&&c.v!=null?c.v:null;};
  var out={}, maxR=Math.min(rng.e.r,rng.s.r+1000);
  for(var r=rng.s.r;r<=maxR;r++){
    var rr=r+1, bv=get('B',rr);
    if(!(Number.isInteger(bv)&&bv>=1&&bv<=60))continue;
    var birds=Math.round(+get('G',rr)||0); if(birds<=1000)continue;
    var cumDead=Math.round(+get('N',rr)||0);
    // ยอดคงเหลือ: บางฟาร์มอยู่คอลัมน์ R (ชัยนาท) บางฟาร์มอยู่ Q (พุ่มวงศ์) → เลือกตัวที่ใกล้ (ลงไก่-ตายสะสม)
    var aliveR=Math.round(+get('R',rr)||0), aliveQ=Math.round(+get('Q',rr)||0), exp=birds-cumDead;
    var alive=Math.abs(aliveQ-exp)<Math.abs(aliveR-exp)?aliveQ:aliveR;
    out[bv]={house:bv,birds:birds,sex:String(get('D',rr)||''),mDeath:Math.round(+get('J',rr)||0),mCull:Math.round(+get('K',rr)||0),eDeath:Math.round(+get('L',rr)||0),eCull:Math.round(+get('M',rr)||0),today:Math.round(+get('I',rr)||0),cumDead:cumDead,alive:alive,feedDeliv:+get('S',rr)||0,feedEaten:+get('T',rr)||0,feedPerDay:Math.round(+get('W',rr)||0),water:+get('X',rr)||0,wtKg:+get('Y',rr)||0};
  }
  return out;
}
function parseFace(wb){
  var out=[];
  wb.SheetNames.forEach(function(sn){
    if(!/^H\s*\d+$/.test(sn))return;
    var ws=wb.Sheets[sn]; if(!ws||!ws['!ref'])return;
    var birds=Math.round(num(ws['W4'])); if(!(birds>1000))return;
    var d=readHouseDead(ws);
    out.push({num:+sn.replace(/[^0-9]/g,''),birds:birds,dead:d.dead,age:Math.max(1,d.last)});
  });
  return out.sort(function(a,b){return a.num-b.num;});
}
onmessage=function(e){
  var d=e.data;
  try{
    var wb=XLSX.read(new Uint8Array(d.buf),RD);
    var fname=smartFarmName(wb,d.fileName);
    if(d.mode==='face'){
      postMessage({id:d.id,ok:true,faceName:fname,houses:parseFace(wb)});
    }else{
      var f=parseWB(wb,fname);
      if(f){var sum=parseSummary(wb);f.houses.forEach(function(h){var k=+h.name.replace(/[^0-9]/g,'');var dd=sum[k];if(dd&&Math.abs(dd.birds-h.birds)<2){var loss=h.dead.reduce(function(a,b){return a+b;},0);var okD=Math.abs(dd.cumDead-loss)<=Math.max(5,loss*0.03),okA=Math.abs(dd.alive-(h.birds-loss))<=Math.max(5,(h.birds-loss)*0.02);if(okD&&okA)h.detail=dd;}});}
      postMessage({id:d.id,ok:true,farm:f});
    }
  }catch(err){postMessage({id:d.id,ok:false,error:String(err&&err.message||err)});}
};
