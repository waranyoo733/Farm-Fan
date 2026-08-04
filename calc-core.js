/* ============================================================
 * PPF Smart Farm — calc-core.js
 * แหล่งความจริงเดียวของ: ตาราง STD (นน./อาหาร/FCR) · สูตรระบายอากาศ
 * (IN2CM=2.514 + ชิ่งลม baffle) · แผนพัดลม FANPLANS · ตัวจับคู่แผน resolvePlan
 * ใช้ร่วมกันโดย index.html / fan-plan.html / view.html — แก้สูตรที่นี่ที่เดียว
 * ⚠️ ห้ามแก้ตัวเลขตาราง/สูตรโดยไม่ได้รับยืนยันจากหมอวรัญญู
 * ============================================================ */
const BW=[42,55,71,90,112,138,168,202,240,283,330,382,440,503,570,639,711,786,864,945,1029,1116,1205,1296,1390,1486,1583,1682,1783,1886,1989,2094,2200,2306,2413,2521,2629,2738,2846,2954,3062,3170,3278];
const FEED=[0,14,18,21,24,27,31,34,38,42,47,51,56,61,66,71,76,82,87,93,98,104,109,115,120,125,130,135,140,145,150,154,159,163,167,171,175,178,182,185,188,192,194];
// %อาหาร ตามตาราง STD (As-Hatched) = อาหารสะสมต่อตัว(ก.) ÷ ถุงอาหาร 30 กก./กระสอบ (÷30000×100 = ÷300) · วัน16=2.33, วัน42=15.95
const FPCT=[0.00,0.04,0.10,0.17,0.25,0.34,0.44,0.56,0.69,0.83,0.99,1.17,1.37,1.58,1.81,2.06,2.33,2.62,2.93,3.25,3.60,3.97,4.36,4.77,5.21,5.65,6.12,6.62,7.12,7.65,8.19,8.76,9.34,9.94,10.54,11.17,11.82,12.47,13.14,13.83,14.52,15.23,15.95];
// FCR & อาหารสะสม ตามตาราง STD (As-Hatched Performance) — วัน42: FCR 1.68, อาหารสะสม 4,786 g (หมอยืนยันใช้ 1.68)
const FCR_STD=[0,0.23,0.42,0.57,0.68,0.77,0.84,0.90,0.96,1.01,1.05,1.10,1.13,1.15,1.17,1.18,1.19,1.21,1.22,1.24,1.25,1.26,1.28,1.30,1.32,1.34,1.36,1.38,1.40,1.42,1.44,1.46,1.48,1.50,1.51,1.53,1.55,1.57,1.59,1.61,1.63,1.65,1.68];
const CUMF_STD=[0,13,30,51,74,101,132,167,206,250,298,352,410,474,542,617,698,785,878,976,1081,1192,1309,1432,1562,1696,1837,1985,2137,2295,2458,2627,2801,2981,3163,3352,3545,3742,3943,4148,4357,4570,4786];
const D=a=>Math.max(0,Math.min(42,Math.round(a)));
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const bw=a=>BW[D(a)],feedG=d=>FEED[D(d)];
/* อุณหภูมิเป้าหมายในเล้า / อุณหภูมิเปิดปั๊มแพด — ลดวันละ 0.3° จนถึงพื้น 27.6 / 30.6
   ⚠️ หมอชี้ขาด 4 ส.ค. 2569 ว่า "ให้ยึด 30.6 และ 27.6 อากาศช่วงนี้ร้อนจัด"
   🔴 ห้ามแก้ตามไฟล์ Excel รุ่นหลัง: ไฟล์ที่หมอส่ง 4 ส.ค. รอบบ่าย เปลี่ยนเป็นลดวันละ 0.4°
   แล้วหยุดที่พื้น 30.2 / 31.2 ตั้งแต่วันที่ 8 (ปั๊มก็เลิกเป็น เป้า+3 กลายเป็นเส้นแยก)
   หมอดูแล้วสั่ง "คง 27.6 ไว้ก่อน" — ถ้าเจอว่าโค้ดไม่ตรง Excel นี่คือเหตุผล อย่าเพิ่งแก้ ให้ถามหมอก่อน */
const targetT=a=>{const x=D(a);return x<1?33:Math.max(27.6,33-0.3*(x-1));};
const pumpT=a=>{const x=D(a);return x<1?36:Math.max(30.6,36-0.3*(x-1));};
const cfmkg=a=>a<=13?3:(a<=27?4:5);
/* ลมสูงสุดที่ไก่ทนได้ (m/s) — ใช้ตัดสินว่าต้องเปิด Cooling pad แทนการเพิ่มลมหรือยัง
   ⚠️ ช่วง >35 วัน แก้ 2.8 → 3.0 ตามที่หมอชี้ขาด 2 ส.ค. 2569
   เหตุผล: เดิมขัดกับ targetWind ที่บอกให้ทำถึง 3.0 ทั้งที่เพดานเขียนไว้ 2.8
   ผลกระทบที่วัดแล้ว: กระทบเฉพาะไก่ 36+ วัน ในช่วงลม 2.8–3.0 (ตัวอย่างชัยนาท = ตอนเปิดพัดลม 17 ตัว)
   ⚠️ ช่วง 29–35 วัน แก้ 2.3 → 2.5 ตามที่หมอชี้ขาด 4 ส.ค. 2569 (เหตุผลเดียวกัน: เป้า 2.5 > เพดาน 2.3)
   ผลกระทบที่วัดแล้ว (52 เล้า): พลิกผลแค่ 2 เล้า เล้าละ 1 วัน — ธนวัฒน์ เล้า 1 วัน 29 (2.45) ·
   เต็งหนึ่ง เล้า 1 วัน 34 (2.46) และเฉพาะตอนพัดลมอยู่จำนวนขั้นต่ำ · วัน 35 ไม่เปลี่ยนสักเล้า
   (พัดลมขึ้นเต็มตามกฎหมอ ลมจริง 2.67–4.65 m/s เกินทั้ง 2.3 และ 2.5 อยู่แล้ว) */
const maxWind=a=>a<=7?0.3:a<=14?0.7:a<=21?1.2:a<=28?1.8:a<=35?2.5:3.0;
// ความเร็วลมเป้าหมายที่ไก่ต้องการต่ออายุ (UGA Poultry Extension, m/s) — แสดงผลเป็นไกด์ ไม่กระทบ logic เปิดแพด
const targetWind=a=>a<=7?0.25:a<=14?0.5:a<=21?1.0:a<=28?1.75:a<=35?2.5:3.0;
function vmaxRaw(fc,birds,fans,age){const reqAir=cfmkg(age)*(BW[D(age)]/1000)*birds;return Math.min(fans,Math.max(1,Math.ceil(fc>0?reqAir/fc:0)));}
function vmaxCurve(fc,birds,fans,age){const MS=[7,14,21,28,35,42];const vals=[4];for(let k=1;k<6;k++)vals.push(Math.max(vals[k-1],vmaxRaw(fc,birds,fans,MS[k])));if(age<=7)return 4;if(age>=42)return vals[5];for(let k=0;k<5;k++)if(age<=MS[k+1]){const t=(age-MS[k])/(MS[k+1]-MS[k]);const ts=t*t*(3-2*t);return Math.round(vals[k]+ts*(vals[k+1]-vals[k]));}return vals[5];}
/* ระยะห่างอุณหภูมิระหว่าง Step ออโต้ (°C) — ถอดจาก Excel ชีต Temp1200/Temp716 คอลัมน์ Diff (สองชีตตรงกัน)
   ⚠️ หมอชี้ขาด 4 ส.ค. 2569: เดิมแอปใช้ +1/+2/+3 ตายตัว → ตอนไก่โตสั่งช้ากว่าตู้จริงเกือบ 2°
   (วัน 21 แอปบอก Step4 ขึ้นที่ 30.6° แต่ตู้จริง 28.7°) ไก่โตทนร้อนน้อยลง สเต็ปจึงไล่ถี่ขึ้น
   Step n เปิดที่ เป้า + (n-1)×diff · ปั๊ม/แพดยังเป็น เป้า+3 คงที่ (ตรวจแล้วทุกแถวในทั้งสองชีต) */
const stepDiff=a=>{const x=D(a);return x<=6?1:x<=13?0.8:x<=20?0.6:x<=27?0.38:0.4;};
const stepTemp=(a,n)=>+(targetT(a)+(n-1)*stepDiff(a)).toFixed(1);
/* เวลาโช๊ค ทำ/พัก (นาที) · index = อายุไก่เป็นวัน
   ⚠️ หมอชี้ขาด 4 ส.ค. 2569: แยกตารางตามขนาดเล้า — เดิมใช้ตารางของ 34×120 กับทุกฟาร์ม
   เล้า 28×120 ไต่ช้ากว่า 1 วัน (ขึ้นเดินต่อเนื่อง 3/2 ที่วัน 8 ไม่ใช่วัน 7) · ที่มา Temp1200 / Temp716
   มีข้อมูลจริงแค่ 2 ขนาดนี้ — ขนาดอื่นคงตารางเดิมไว้ก่อน ห้ามเดาแทนหมอ
   หมายเหตุหมอ: โช๊คต้องดูขนาดเล้า+จำนวนไก่ด้วย → ส่วนนั้นคำนวณอยู่แล้วใน minFan() ด้านล่าง */
const CK28={on:[0.5,0.5,0.7,0.7,0.8,0.8,0.8,0.6],off:[10,10,9,8,7,6,5,3]};
const CK34={on:[0.5,0.6,0.7,0.7,0.8,0.8,0.6],off:[10,9,8,7,6,5,3]};
const choke=(a,h)=>{const x=D(a),T=(h&&Math.round(+h.W||0)===28)?CK28:CK34,last=T.on.length-1;
  if(x<=last){const on=T.on[x];return{on:on>0.6?1:on,off:T.off[x]};}
  return{on:3,off:2};};
// แสดงเวลาโช้ค (รอบเปิด-ปิดพัดลม) มีหน่วย — ถ้าเวลาทำ < 1 นาที แปลงเป็นวินาที
const chokeOnU=c=>c.on>=1?(c.on+' นาที'):(Math.round(c.on*60)+' วินาที');
const chokeText=c=>'ทำ '+chokeOnU(c)+' หยุด '+c.off+' นาที';
const chokeShort=c=>(c.on>=1?c.on+'น.':Math.round(c.on*60)+'วิ.')+' / '+c.off+'น.';
/* ขนาดช่องลมพัดลมที่ "วัดได้จริง" — พัดลม 52 นิ้ว = 130.728 ซม. (ไม่ใช่ 132.08 ตามทฤษฎี 2.54)
   ที่มา: Excel ของหมอ ชีต "หาประสิทธิภาพพัดลม.ต่อตัว" (ฟาร์มมั่งมี) → 4.289 ฟุต → 18.395 ตร.ฟุต
   หมอยืนยันให้ใช้ค่าที่วัดจริง 31 ก.ค. 2026 · ค่าเดิม 2.54 ให้พื้นที่มากเกินจริง 2.1% */
const IN2CM=2.514;
function fanCFMof(h){const side=h.fan*IN2CM/30.48,area=side*side,avg=h.pts.reduce((a,b)=>a+(+b||0),0)/9;return area*avg;}
const fmtMinT=m=>m<=0?'0':(m<1?Math.round(m*60)+' วินาที':Math.round(m)+' นาที'); // ปัดเป็นนาทีเต็ม (60 วินาที = 1 นาที)
const fmtMc=m=>m<=0?'0':(m<1?Math.round(m*60)+'วิ':Math.round(m)+'น');
// พัดลมขั้นต่ำ + เวลาทำงาน (สูตรหมอตั้ม): I = อากาศต้องการ ÷ ประสิทธิภาพพัดลม/ตัว · เดินโช๊คเป็นรอบ 5 นาที
function minFan(h,nck){
  const fc=fanCFMof(h),wkg=bw(h.age)/1000,q=cfmkg(h.age)*wkg*(+h.birds||0),I=fc>0?q/fc:0,TC=5;
  let on,off,cont=false,duty;
  if(nck>0&&I>=nck){cont=true;duty=1;on=TC;off=0;}
  else if(nck>0){duty=I/nck;on=Math.max(1,Math.min(TC-1,Math.round(TC*duty)));off=TC-on;} // ปัดเวลาทำเป็นนาทีเต็ม · on+off = TC (5 นาที) เสมอ
  else{duty=0;on=0;off=0;}
  return {q:q,fanCFM:fc,I:I,nck:nck,duty:duty,on:on,off:off,cont:cont};
}
const lux=a=>{const x=D(a);return x<=1?40:x<=14?30:x<=21?20:10;};
const lightHr=a=>{const x=D(a);return x<=1?24:x<=7?20:x<=14?18:x<=21?16:x<=28?10:12;};
/* ===== แผนการทำงานพัดลม (ฤดูร้อน) — ถอดจาก Excel "แผนการระบายอากาศ หมอวรัญญู" คงลำดับเดิมของแต่ละฟาร์ม =====
   choke = พัดลมโช๊ค (วิ่ง Cycle ตอน≤เป้า / ต่อเนื่องที่เป้า+diff) · steps[].n = Step Auto (เปิดที่ เป้า+(n-1)×stepDiff) · manual[] = เปิดมือตามอายุ (สะสม)
   manual[].off = ปิดพัดลมที่เคยเปิดไว้ — หมอสลับตัวเปิดเพื่อ "บาลานซ์ลม" ไม่ใช่เปิดเพิ่มอย่างเดียว
   ⚠️ phum (พุ่มวงศ์/ยิ่งรวย) แก้ทั้งชุด 4 ส.ค. 2569 ให้ตรง PDF "แผนการระบายอากาศฟาร์มพุ่มวงษ์"
   ของเดิมเป็นแบบบวกอย่างเดียว → ผิด 18 วันจาก 42 · ตัวที่ 11 ของจริงเปิด-ปิดสลับ 6 รอบ
   วิธีถอด: อ่านสีจากตัว PDF โดยตรง (เขียว = เปิดมือ) แล้วจับคู่กับคอลัมน์ "Manual (ตัว)" หน้า 1 → N = ปัดลง
   หมอชี้ขาดว่า "ยึดสี" ตรงแถว 10/11/12 ตัว ที่สีเป็นตัวที่ 5 แต่ตัวเลขในไฟล์พิมพ์ 6 (ชุดที่ใช้ 5 สมมาตรพอดี)
   สเต็ปออโต้ + โช๊ค ของ phum ตรวจแล้วถูกอยู่แล้ว ไม่ได้แก้ (โช๊ค 2+21 · T3 4+19 · T8 6+17 · T6 8+15)
   ⚠️ mk12 (มั้งกี้1 = เล้า 1 ของมั่งมี/ดอนทอง) แก้ 4 ส.ค. 2569 ให้ตรง Excel "แผนการระบายอากาศมั้งกี้1"
   ชีต "แผนการใช้พัดลม" — ที่นี่ผังพัดลมอยู่แถวเดียวกับตารางรายวัน (แถว = อายุตรง ๆ ไม่ต้องแปลงผ่าน N)
   ของเดิมบวกอย่างเดียว → ผิด 6 วัน (16–19, 22–23) ที่ต้องปิดพัดลมตัวที่ 6 ชั่วคราวเพื่อบาลานซ์ลม */
const FANPLANS={"chai34":{"farms":["ชัยนาท"],"size":"34×120","nf":24,"choke":[7,18],"steps":[{"n":2,"fans":[]},{"n":3,"fans":[5,20]},{"n":4,"fans":[3,22]}],"manual":[{"age":8,"fans":[2,23]},{"age":12,"fans":[9,16]},{"age":22,"fans":[11,14]},{"age":25,"fans":[6,19]},{"age":29,"fans":[4,21]},{"age":33,"fans":[1,24]},{"age":36,"fans":[8,17]},{"age":39,"fans":[10,12,13,15]}]},"chai28":{"farms":["ชัยนาท"],"size":"28×120","nf":18,"choke":[4,15],"steps":[{"n":2,"fans":[]},{"n":3,"fans":[6,13]},{"n":4,"fans":[2,17]}],"manual":[{"age":8,"fans":[8,11]},{"age":15,"fans":[3,16]},{"age":22,"fans":[1,18]},{"age":29,"fans":[5,14]},{"age":33,"fans":[7,12]},{"age":36,"fans":[9,10]}]},"thana":{"farms":["ธนวัฒน์"],"size":"28×120","nf":18,"choke":[2,17],"steps":[{"n":2,"fans":[4,15]},{"n":3,"fans":[6,13]},{"n":4,"fans":[8,11]}],"manual":[{"age":10,"fans":[9]},{"age":15,"fans":[1,18]},{"age":19,"fans":[10]},{"age":21,"fans":[7,12]},{"age":23,"fans":[5]},{"age":25,"fans":[14]},{"age":26,"fans":[3]},{"age":27,"fans":[16]},{"age":30,"fans":[8]},{"age":31,"fans":[11]},{"age":33,"fans":[6]},{"age":34,"fans":[13]},{"age":36,"fans":[2,4,15,17]}]},"phum":{"farms":["พุ่มวงษ์","พุ่มวงศ์","ยิ่งรวย"],"size":"28×120","nf":22,"choke":[2,21],"steps":[{"n":2,"fans":[4,19]},{"n":3,"fans":[6,17]},{"n":4,"fans":[8,15]}],"manual":[{"age":8,"fans":[11]},{"age":13,"fans":[3,20],"off":[11]},{"age":15,"fans":[1,11,22],"off":[3,20]},{"age":16,"fans":[9,14],"off":[11]},{"age":18,"fans":[11]},{"age":20,"fans":[5,10,13,18],"off":[9,11,14]},{"age":22,"fans":[7,9,14,16],"off":[10,13]},{"age":23,"fans":[11]},{"age":25,"fans":[3,20],"off":[11]},{"age":26,"fans":[11]},{"age":27,"fans":[13]},{"age":29,"fans":[8,12,15]},{"age":31,"fans":[10]},{"age":32,"fans":[6,17],"off":[12]},{"age":33,"fans":[4,19],"off":[11]},{"age":34,"fans":[11]},{"age":35,"fans":[12]}]},"money":{"farms":["มันนี่","สาธิต"],"size":"21×120","nf":16,"choke":[2,15],"steps":[{"n":2,"fans":[4,13]},{"n":3,"fans":[6,11]},{"n":4,"fans":[7,10]}],"manual":[{"age":12,"fans":[8]},{"age":15,"fans":[1,16]},{"age":21,"fans":[5,12]},{"age":24,"fans":[3,14]},{"age":28,"fans":[9]},{"age":30,"fans":[10]},{"age":32,"fans":[7]},{"age":34,"fans":[6]},{"age":35,"fans":[11]},{"age":37,"fans":[2,4,13,15]}]},"mk12":{"farms":["มั้งกี้1","มอส"],"size":"16×100","nf":12,"choke":[5],"steps":[{"n":2,"fans":[2,11]},{"n":3,"fans":[7]}],"manual":[{"age":13,"fans":[6]},{"age":16,"fans":[1,12],"off":[6]},{"age":20,"fans":[6]},{"age":22,"fans":[4,8],"off":[6]},{"age":24,"fans":[6]},{"age":26,"fans":[10]},{"age":29,"fans":[3]},{"age":31,"fans":[9]},{"age":34,"fans":[7]},{"age":36,"fans":[2,11]},{"age":38,"fans":[5]}]},"mk10":{"farms":["มั้งกี้2","มอส"],"size":"16×100","nf":10,"choke":[4],"steps":[{"n":2,"fans":[1,10]},{"n":3,"fans":[6]}],"manual":[{"age":13,"fans":[5]},{"age":22,"fans":[3,7]},{"age":26,"fans":[9]},{"age":29,"fans":[2]},{"age":31,"fans":[8]},{"age":34,"fans":[6]},{"age":36,"fans":[1,4,10]}]},"mos25":{"farms":["มอส","มั้งกี้3"],"size":"25×120","nf":14,"choke":[2,13],"steps":[{"n":2,"fans":[4,11]},{"n":3,"fans":[6,9]}],"manual":[{"age":10,"fans":[7]},{"age":14,"fans":[1,14]},{"age":18,"fans":[8]},{"age":21,"fans":[3,12]},{"age":23,"fans":[5,10]},{"age":26,"fans":[9]},{"age":27,"fans":[6]},{"age":29,"fans":[11]},{"age":31,"fans":[4]},{"age":33,"fans":[2,13]}]},"sriha":{"farms":["ศรีเหรา"],"size":"40×128","nf":26,"choke":[2,25],"steps":[{"n":2,"fans":[4,23]},{"n":3,"fans":[5,22]},{"n":4,"fans":[7,20]}],"manual":[{"age":7,"fans":[13]},{"age":11,"fans":[3,24]},{"age":16,"fans":[1,9,14,18,26]},{"age":20,"fans":[15]},{"age":21,"fans":[11]},{"age":22,"fans":[6,19,21]},{"age":25,"fans":[8,17]},{"age":26,"fans":[12]},{"age":27,"fans":[16]},{"age":29,"fans":[10]},{"age":30,"fans":[5,22]},{"age":31,"fans":[4]},{"age":32,"fans":[23]},{"age":33,"fans":[2,25]}]},"teng":{"farms":["เต็งหนึ่ง"],"size":"28×84","nf":20,"choke":[6,15],"steps":[{"n":2,"fans":[4,17]},{"n":3,"fans":[9,12]},{"n":4,"fans":[2,19]}],"manual":[{"age":7,"fans":[10]},{"age":12,"fans":[3,18]},{"age":17,"fans":[8,13]},{"age":22,"fans":[11]},{"age":24,"fans":[5,7,14,16]},{"age":32,"fans":[1,20]}]},"rr36":{"farms":["รุ่งเรือง3-6","หนึ่งรุ่งเรือง"],"size":"32×120","nf":21,"choke":[2,20],"steps":[{"n":2,"fans":[4,18]},{"n":3,"fans":[6,16]},{"n":4,"fans":[8,14]}],"manual":[{"age":8,"fans":[11]},{"age":13,"fans":[1,21],"off":[11]},{"age":15,"fans":[11]},{"age":17,"fans":[7,15],"off":[11]},{"age":19,"fans":[5,11,17],"off":[7,15]},{"age":21,"fans":[9,13],"off":[11]},{"age":23,"fans":[7,11,15],"off":[9,13]},{"age":25,"fans":[9,13],"off":[11]},{"age":27,"fans":[3]},{"age":29,"fans":[11]},{"age":30,"fans":[10,12],"off":[11]},{"age":31,"fans":[8,11,14],"off":[10,12]},{"age":32,"fans":[10,12],"off":[11]},{"age":33,"fans":[11]}]},"rr2":{"farms":["รุ่งเรือง2","หนึ่งรุ่งเรือง"],"size":"16×100","nf":11,"choke":[],"steps":[],"manual":[{"age":13,"fans":[6]},{"age":18,"fans":[1,11],"off":[6]},{"age":21,"fans":[6]},{"age":22,"fans":[3]},{"age":25,"fans":[9]},{"age":28,"fans":[4,8],"off":[6]},{"age":30,"fans":[6]},{"age":33,"fans":[7]},{"age":35,"fans":[10]},{"age":37,"fans":[2,5]}]},"rr1":{"farms":["รุ่งเรือง1","หนึ่งรุ่งเรือง"],"size":"14×72","nf":8,"choke":[4],"steps":[],"manual":[{"age":15,"fans":[5]},{"age":22,"fans":[1,8],"off":[5]},{"age":24,"fans":[5]},{"age":28,"fans":[3]},{"age":33,"fans":[6]},{"age":37,"fans":[4]},{"age":38,"fans":[2,7]}]}};
/* ฟาร์มเดียวกัน คนละชื่อในเอกสารแต่ละยุค — หมอยืนยัน 31 ก.ค. 2026
   "PPF ดอนทอง" (ชื่อในแอป) = "มั่งมี" (ชื่อในไฟล์ Excel) = "มั้งกี้/มอส" (ชื่อในผังพัดลม)
   เล้า 1 = 12 พัดลม (mk12) · เล้า 2 = 10 พัดลม (mk10) · เล้า 3-11 = 14 พัดลม (mos25) */
['mk12','mk10','mos25'].forEach(function(k){ if(FANPLANS[k])FANPLANS[k].farms=FANPLANS[k].farms.concat(['ดอนทอง','มั่งมี']); });
function resolvePlan(h,farmName){
  // เลือกแผนพัดลมให้ถูกอัตโนมัติ: 1) ที่หมอเลือกเอง 2) ชื่อฟาร์ม+ขนาดเล้า 3) ชื่อฟาร์ม+จำนวนพัดลม 4) ชื่อฟาร์มตัวเดียว
  if(h&&h.planKey&&FANPLANS[h.planKey])return{key:h.planKey,p:FANPLANS[h.planKey],byName:true};
  const fn=(farmName||'').replace(/ฟาร์ม|\s/g,'');
  const ent=Object.entries(FANPLANS);
  const named=ent.filter(([k,p])=>p.farms.some(f=>fn&&(fn.indexOf(f)>=0||f.indexOf(fn)>=0)));
  const sz=h?(Math.round(h.W)+'×'+Math.round(h.L)):'';
  if(named.length){                                    // จับด้วยชื่อฟาร์มก่อนเสมอ (กันจับข้ามฟาร์ม เช่น ศรีเหรา→ชัยนาท)
    // ลำดับสำคัญ: "จำนวนพัดลม" มาก่อน "ขนาดเล้า" — เพราะแผนคือใบสั่งว่าเปิดพัดลม "ตัวที่เท่าไร"
    // ถ้าจำนวนพัดลมไม่เท่ากัน เบอร์ในแผนจะใช้ไม่ได้เลย (เช่น มั่งมี เล้า3-6 = 16×100 แต่ 14 พัดลม → ต้องได้แผน 14 ตัว ไม่ใช่แผน 12 ตัวที่ขนาดตรง)
    const m=named.find(([k,p])=>p.size===sz&&p.nf===(h&&h.fans))   // ตรงทั้งขนาดและจำนวนพัดลม = แม่นที่สุด
         || named.find(([k,p])=>p.nf===(h&&h.fans))    // ตรงจำนวนพัดลม
         || named.find(([k,p])=>p.size===sz)           // ตรงขนาดเล้า
         || named[0];                                  // ชื่อฟาร์มตรงตัวเดียว/ตัวแรก
    return{key:m[0],p:m[1],byName:true};
  }
  const m=ent.filter(([k,p])=>p.nf===(h&&h.fans))[0];  // ไม่รู้จักชื่อ → เดาตามจำนวนพัดลม (ไม่มั่นใจ ไม่ซ่อมค่า)
  return m?{key:m[0],p:m[1],byName:false}:null;
}
function planFor(h,farmName){const r=resolvePlan(h,farmName);return r?{key:r.key,p:r.p,byName:r.byName}:null;}
function manualUpTo(p,a){const out=new Set();for(const s of p.manual)if(s.age<=a){(s.fans||[]).forEach(f=>out.add(f));(s.off||[]).forEach(f=>out.delete(f));}return[...out].sort((x,y)=>x-y);}
function fanStateAt(p,age){
  const a=D(age),tgt=targetT(a);
  const man=manualUpTo(p,a);
  const dif=stepDiff(a);
  const steps=p.steps.filter(s=>s.fans.length).map(s=>({n:s.n,fans:s.fans,temp:+(tgt+(s.n-1)*dif).toFixed(1)}));
  let next=null;for(const s of p.manual)if(s.age>a){next=s;break;}
  return{man,steps,tgt:+tgt.toFixed(1),diff:dif,chokeTemp:+(tgt+dif).toFixed(1),pumpTemp:+pumpT(a).toFixed(1),next};
}
const CFM_TO_M3S=0.00047195, MINRATE=1, GAP=4;
function cumDead(h,upto){let s=0;for(let d=1;d<=upto;d++)s+=(+h.dead[d-1]||0);return s;}
function aliveAt(h,age){return Math.max(0,h.birds-cumDead(h,Math.max(0,age-1)));}

/* ── ตัวช่วยคำนวณราย "วัน" (ใช้โดย fan-plan.html / view.html — กติกาเดียวกับ calc() ของแอปหลัก) ── */
function ventAt(h,farmName,d){
  const fc=fanCFMof(h);
  const pm=resolvePlan(h,farmName);
  let vmax,vmin;
  if(pm){
    const mc=manualUpTo(pm.p,D(d)).length,ck=pm.p.choke.length,au=ck+pm.p.steps.reduce((s,x)=>s+x.fans.length,0);
    vmax=Math.min(h.fans,mc+au);vmin=Math.min(h.fans,mc+ck);
  }else if(d<=7){vmax=Math.min(h.fans,4);vmin=Math.min(1,vmax);}
  else{
    const needMin=fc>0?(MINRATE*(BW[D(d)]/1000)*h.birds)/fc:0;
    vmax=Math.max(vmaxRaw(fc,h.birds,h.fans,d),Math.min(h.fans,4));
    vmin=Math.min(vmax,Math.max(1,Math.ceil(needMin)));vmin=Math.max(vmin,vmax-GAP);
  }
  const planMaxRaw=vmax, needFans=vmaxRaw(fc,h.birds,h.fans,d);
  if(needFans>vmax)vmax=needFans;                 // กฎหมอ 31 ก.ค. 2569: ความต้องการลมของไก่มาก่อนแผนเสมอ
  if(D(d)>=35)vmax=h.fans;                        // กฎหมอ: อายุ 35 วันขึ้นไป เปิดครบทุกตัว
  if(vmin>vmax)vmin=vmax;
  return {fc,pm,vmin,vmax,planMaxRaw,needFans};
}
function crossOf(h){return h.W*Math.max(0.1,h.H-(+h.baffle||0)/100);}   // หน้าตัดลม = กว้าง × (สูง − ชิ่งลม)
function velOf(h,fc,n){const cx=crossOf(h);return cx>0?n*fc*CFM_TO_M3S/cx:0;}
function curtainCm(h,n){return h.fans>0?(+h.padH||0)*(n/h.fans)*100:0;}
