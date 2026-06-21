/* Web Worker: อ่าน+แปลงไฟล์ใบหน้าเล้า .xlsx นอก main thread (กันหน้าค้าง)
   mode 'mortality' → คืนทั้งฟาร์ม (dead/cull + detail จากชีตสรุป)
   mode 'face'      → คืนรายเล้า (num/birds/dead/age) สำหรับนำเข้าในแอปหลัก */
importScripts('https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js');
var RD={type:'array',cellFormula:false,cellHTML:false,cellText:false,cellNF:false,cellStyles:false,sheetStubs:false};
function num(c){return c&&c.t!=='e'&&c.v!=null&&!isNaN(+c.v)?+c.v:0;}  // c.t==='e' = เซลล์ error (#VALUE! ฯลฯ) → ข้าม
function zeros(){return new Array(42).fill(0);}
function farmInfo(fn){
  if(/ชัยนาท/.test(fn))return'ฟาร์ม ชัยนาท';
  if(/รุ่งเรือง|หนึ่ง/.test(fn))return'ฟาร์ม หนึ่งรุ่งเรือง';
  if(/ยิ่งรวย/.test(fn))return'ฟาร์ม ยิ่งรวย';
  if(/ขวัญใจ/.test(fn))return'ฟาร์ม ขวัญใจ';
  return fn.replace(/\.xlsx?$/i,'').replace(/ใบหน้าเล้า|ฟาร์ม/g,'').trim()||'ฟาร์ม';
}
function faceFarmName(fn){
  if(/ชัยนาท/.test(fn))return'ฟาร์ม ชัยนาท';
  if(/รุ่งเรือง|หนึ่ง/.test(fn))return'ฟาร์ม หนึ่งรุ่งเรือง';
  if(/ยิ่งรวย/.test(fn))return'ฟาร์ม ยิ่งรวย';
  if(/ขวัญใจ/.test(fn))return'ฟาร์ม ขวัญใจ';
  return 'ฟาร์ม '+(fn.replace(/\.xlsx?$/i,'').replace(/ใบหน้าเล้า|ฟาร์ม|รุ่นที่|รุ่น/g,'').replace(/[0-9.()_\-]+/g,' ').replace(/\s+/g,' ').trim()||'นำเข้า');
}
function parseWB(wb,fileName){
  var name=farmInfo(fileName), houses=[];
  wb.SheetNames.forEach(function(sn){
    if(!/^H\s*\d+$/.test(sn))return;
    var ws=wb.Sheets[sn]; if(!ws||!ws['!ref'])return;
    var birds=Math.round(num(ws['W4'])); if(birds<=1000)return;
    var dead=zeros(),cull=zeros();
    var rng=XLSX.utils.decode_range(ws['!ref']); var last=0; var maxR=Math.min(rng.e.r,rng.s.r+1000);
    for(var r=rng.s.r;r<=maxR;r++){
      var rr=r+1, bc=ws['B'+rr]; if(!bc||bc.t==='e')continue;
      var day=+bc.v; if(!(Number.isInteger(day)&&day>=1&&day<=42))continue;
      var loss=Math.round(num(ws['W'+rr]));
      dead[day-1]=loss; cull[day-1]=0; if(loss>0)last=Math.max(last,day);
    }
    houses.push({name:'เล้า '+sn.replace(/[^0-9]/g,''),birds:birds,dead:dead,cull:cull,age:last});
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
    out[bv]={house:bv,birds:birds,sex:String(get('D',rr)||''),mDeath:Math.round(+get('J',rr)||0),mCull:Math.round(+get('K',rr)||0),eDeath:Math.round(+get('L',rr)||0),eCull:Math.round(+get('M',rr)||0),today:Math.round(+get('I',rr)||0),cumDead:Math.round(+get('N',rr)||0),alive:Math.round(+get('R',rr)||0),feedDeliv:+get('S',rr)||0,feedEaten:+get('T',rr)||0,feedPerDay:Math.round(+get('W',rr)||0),water:+get('X',rr)||0,wtKg:+get('Y',rr)||0};
  }
  return out;
}
function parseFace(wb){
  var out=[];
  wb.SheetNames.forEach(function(sn){
    if(!/^H\s*\d+$/.test(sn))return;
    var ws=wb.Sheets[sn]; if(!ws||!ws['!ref'])return;
    var birds=Math.round(num(ws['W4'])); if(!(birds>1000))return;
    var dead=zeros(); var rng=XLSX.utils.decode_range(ws['!ref']); var last=0; var maxR=Math.min(rng.e.r,rng.s.r+1000);
    for(var r=rng.s.r;r<=maxR;r++){var rr=r+1,b=ws['B'+rr];if(!b||b.t==='e')continue;var day=+b.v;if(!(Number.isInteger(day)&&day>=1&&day<=42))continue;var loss=Math.round(num(ws['W'+rr]));dead[day-1]=loss;if(loss>0)last=Math.max(last,day);}
    out.push({num:+sn.replace(/[^0-9]/g,''),birds:birds,dead:dead,age:Math.max(1,last)});
  });
  return out.sort(function(a,b){return a.num-b.num;});
}
onmessage=function(e){
  var d=e.data;
  try{
    var wb=XLSX.read(new Uint8Array(d.buf),RD);
    if(d.mode==='face'){
      postMessage({id:d.id,ok:true,faceName:faceFarmName(d.fileName),houses:parseFace(wb)});
    }else{
      var f=parseWB(wb,d.fileName);
      if(f){var sum=parseSummary(wb);f.houses.forEach(function(h){var k=+h.name.replace(/[^0-9]/g,'');var dd=sum[k];if(dd&&Math.abs(dd.birds-h.birds)<2){var loss=h.dead.reduce(function(a,b){return a+b;},0);var okD=Math.abs(dd.cumDead-loss)<=Math.max(5,loss*0.03),okA=Math.abs(dd.alive-(h.birds-loss))<=Math.max(5,(h.birds-loss)*0.02);if(okD&&okA)h.detail=dd;}});}
      postMessage({id:d.id,ok:true,farm:f});
    }
  }catch(err){postMessage({id:d.id,ok:false,error:String(err&&err.message||err)});}
};
