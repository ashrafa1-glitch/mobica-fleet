/* Tracking Agent — Excel Worker
   يشتغل في thread منفصل عشان المتصفح ميتجمدش */
importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');

const TEAM_MAP = {
  'هشام جمال':'109962','هشام':'109962',
  'سامى فؤاد':'P1','سامي فؤاد':'P1','سامى':'P1',
  'حازم قاعود':'P4','حازم':'P4',
  'أحمد حسن':'ahmed','احمد حسن':'ahmed',
  'محمد وسيم':'A4','وسيم':'A4',
  'الارضيات':'A4','ارضيات':'A4','أرضيات':'A4',
};
// أرقام الأعمدة من ملف SAP
const CI = { order:33, customer:1, address:2, phone:4, group:23,
             delivery:25, planned:18, material:7, mainGrp:32,
             qty:8, unitMeasure:26, instTime:31, notes:20 };

function g(row, idx){ return (row[idx]===undefined||row[idx]===null)?'':row[idx]; }

self.onmessage = function(e){
  try{
    const { arrayBuffer, fileName } = e.data;
    const wb = XLSX.read(new Uint8Array(arrayBuffer), {type:'array', cellDates:true});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const allRows = XLSX.utils.sheet_to_json(ws, {defval:'', header:1});

    if(allRows.length < 2){ self.postMessage({error:'الملف فارغ'}); return; }

    const headers = allRows[0];
    const dataRows = allRows.slice(1);

    // تحديد أعمدة رقم الأمر والمجموعة (مرونة لو تغير الترتيب)
    let orderCol = CI.order, groupCol = CI.group;
    for(let i=0;i<headers.length;i++){
      const h = String(headers[i]||'').replace(/\u0640/g,'').trim();
      if(h==='رقم امر التركيب'||h==='رقم أمر التركيب') orderCol=i;
      if(h==='مجموعة التركيب') groupCol=i;
    }

    const ordersMap={};
    dataRows.forEach(row=>{
      const no = String(g(row,orderCol)||'').trim();
      if(!no||no==='0') return;
      const teamName = String(g(row,groupCol)||'').trim();
      const teamCode = TEAM_MAP[teamName]||teamName;
      if(!ordersMap[no]){
        const pr = g(row,CI.planned);
        ordersMap[no]={
          order_no:no, customer:String(g(row,CI.customer)),
          address:String(g(row,CI.address)), phone:String(g(row,CI.phone)),
          team_code:teamCode, team_name:teamName, status:'REL',
          delivery:String(g(row,CI.delivery)),
          planned: pr instanceof Date ? pr.toLocaleDateString('ar-EG') : String(pr||''),
          main_group:String(g(row,CI.mainGrp)), notes:String(g(row,CI.notes)||''),
          items:[], total_units:0, file_name:fileName
        };
      }
      const units    = parseFloat(g(row,CI.unitMeasure))||0; // col 26 = Unit Measure مباشرةً
      const instTime = parseFloat(g(row,CI.instTime))||0;   // col 31 = الدقايق
      ordersMap[no].items.push({material:String(g(row,CI.material)), qty:parseFloat(g(row,CI.qty))||0, instTime, units});
      ordersMap[no].total_units += units;
    });

    self.postMessage({orders: Object.values(ordersMap), total_rows: dataRows.length});
  } catch(err){
    self.postMessage({error: err.message});
  }
};
