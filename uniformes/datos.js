// Data de la APP de uniformes: lectura de DATA ICG y listas de precios, coste y precios por tarifa.
// Lo usa la página y también el armado de datos.json en GitHub (tools/datos-uniformes.js), para no leer
// ~60 MB de Excel cada vez que se abre la APP.

const N=s=>String(s==null?'':s).normalize('NFD').replace(/[̀-ͯ]/g,'').trim().toUpperCase();
const r2=x=>Math.round(parseFloat(((+x||0)*100).toFixed(6)))/100;
const dg=v=>String(v==null?'':v).replace(/\.0+$/,'').replace(/\D/g,'');
const CBCOL=/^COD(IGO)?\.? ?(DE )?BARRAS( \d)?$/;
const COSTCOL=/^(ULT(IMO)?\.? ?COSTE?|ULTIMO COSTO|ULT\. COSTO|COSTE ULTIMO|COSTO ULTIMO)$/;
const PRICECOL=/^(PRECIO( DE)? VENTA|PVP|P\.V\.P\.?|PRECIO|PRECIO TIENDA|PRECIO NETO|TARIFA|PRECIO VENTA SIN ITBMS|PVP SIN ITBMS)$/;
const DESCCOL=/^(DESC(UENTO)?\.?( ?%)?|% ?DESC(UENTO)?\.?|DTO\.?( ?%)?)$/;

// ---------- data ----------
// Almacenes de ICG: el de la tienda es el almacén del albarán de venta. Tarifa NB = tiendas New Balance, RB = Running Balboa (multimarca)
const TIENDAS=[{code:'A2',name:'NEW BALANCE MULTIPLAZA',tarifa:'NB'},{code:'A3',name:'NEW BALANCE ALBROOK',tarifa:'NB'},{code:'A4',name:'NEW BALANCE DORADO MALL',tarifa:'NB'},{code:'A7',name:'NEW BALANCE METRO MALL',tarifa:'NB'},
  {code:'B1',name:'RUNNING BALBOA MULTIPLAZA',tarifa:'RB'},{code:'B2',name:'RUNNING BALBOA ALBROOK',tarifa:'RB'}];
const OFICINA={code:'C99',name:'OFICINA',tarifa:''};
const tiendaDe=code=>TIENDAS.concat(OFICINA).find(t=>t.code===code)||null;
const TARIFA={NB:'New Balance',RB:'Running Balboa'};
const nuevaTarifa=()=>({cb:new Map(),reft:new Map(),ref:new Map()});
const D={byCB:new Map(),byRef:new Map(),hasCost:false,precios:{NB:nuevaTarifa(),RB:nuevaTarifa(),'*':nuevaTarifa(),n:0}};
function headerRow(A){let h=A.findIndex((r,i)=>i<15&&r.some(c=>CBCOL.test(N(c))||N(c)==='REFERENCIA'));return h<0?0:h;}
function indexIcg(wb){
  const ws=wb.Sheets[wb.SheetNames[0]],A=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true});
  const h=headerRow(A),H=(A[h]||[]).map(N),up=A[h-1]||[];
  const ix=k=>H.indexOf(k),iR=ix('REFERENCIA'),iT=ix('TALLA'),iD=ix('DESCRIPCION'),iC=ix('COLOR'),iDp=ix('DEPARTAMENTO'),iM=ix('MARCA'),iS=ix('SECCION');
  const iInv=H.map((k,i)=>/^INVENTARIO\b/.test(k)?i:-1).filter(i=>i>=0),iVal=H.map((k,i)=>/^VALORACION INVENTARIO\b/.test(k)?i:-1).filter(i=>i>=0);
  const iCost=H.findIndex(k=>COSTCOL.test(k)),cbs=H.map((k,i)=>CBCOL.test(k)?i:-1).filter(i=>i>=0);
  if(iR<0||!cbs.length)throw new Error('no tiene columnas Referencia y Cód. Barras');
  if(iCost>=0||iVal.length)D.hasCost=true;
  let n=0;
  for(let j=h+1;j<A.length;j++){const r=A[j];if(!r[iR])continue;
    const a={ref:String(r[iR]).trim(),talla:iT>=0?String(r[iT]).trim():'',desc:iD>=0?String(r[iD]).trim():'',color:iC>=0?String(r[iC]).trim():'',depto:iDp>=0?N(r[iDp]):'',seccion:iS>=0?N(r[iS]):'',marca:iM>=0?String(r[iM]).trim():'',coste:iCost>=0?num(r[iCost]):null,cb:'',
      inv:iInv.reduce((s,i)=>s+(num(r[i])||0),0),val:iVal.reduce((s,i)=>s+(num(r[i])||0),0)};
    for(const i of cbs){const b=dg(r[i]);if(b){if(!a.cb)a.cb=b;D.byCB.set(b,a);}}
    D.byRef.set(N(a.ref)+'|'+N(a.talla),a);n++;}
  return n;
}
function num(v){if(v===''||v==null)return null;const x=+String(v).replace(/[^0-9.,-]/g,'').replace(/,(?=\d{3}\b)/g,'').replace(',','.');return isFinite(x)?x:null;}
// Listas de precios de ICG (un archivo por marca): fila de grupos "Precios de Venta: NEW BALANCE" y "Precios de Venta: RUNNING BALBOA",
// cada grupo con Neto, Bruto y %Dto. También acepta una lista simple con una sola columna de precio.
function tarifaDeGrupo(t){t=N(t);return /NEW BALANCE|\bNB\b/.test(t)?'NB':/RUNNING|BALBOA|\bRB\b/.test(t)?'RB':null;}
function indexPrecios(wb,nombre){
  let n=0;
  for(const sn of wb.SheetNames){
    const A=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:true});
    const h=A.findIndex((r,i)=>i<20&&r.some(c=>/^(BRUTO|NETO|%? ?DTO\.?)$/.test(N(c))||PRICECOL.test(N(c))));if(h<0)continue;
    const H=A[h].map(N),up=A[h-1]||[],iR=H.indexOf('REFERENCIA'),iT=H.indexOf('TALLA'),cbs=H.map((k,i)=>CBCOL.test(k)?i:-1).filter(i=>i>=0);
    if(iR<0&&!cbs.length)continue;
    // columnas por tarifa: el grupo es el último título de la fila de arriba a la izquierda de la columna
    const G={};let g=null;
    H.forEach((k,i)=>{const t=String(up[i]||'');if(/precio/i.test(t))g=tarifaDeGrupo(t)||'*';const key=g||'*';G[key]=G[key]||{};
      if(k==='BRUTO')G[key].bruto=i;else if(k==='NETO')G[key].neto=i;else if(/^%? ?DTO\.?( ?%)?$|^DESC/.test(k))G[key].dto=i;else if(PRICECOL.test(k)&&G[key].bruto==null)G[key].bruto=i;});
    for(let j=h+1;j<A.length;j++){const r=A[j];let any=false;
      for(const [t,c] of Object.entries(G)){if(c.bruto==null)continue;const bruto=num(r[c.bruto]);if(!(bruto>0))continue;
        let dto=c.dto!=null?(num(r[c.dto])||0):0;if(dto>0&&dto<1)dto*=100;
        const o={precio:bruto,desc:dto,neto:c.neto!=null?num(r[c.neto]):null,origen:nombre||sn},P=D.precios[t];any=true;
        for(const i of cbs){const b=dg(r[i]);if(b)P.cb.set(b,o);}
        if(iR>=0&&r[iR]){const k=N(r[iR]);if(iT>=0&&r[iT]!=='')P.reft.set(k+'|'+N(r[iT]),o);if(!P.ref.has(k))P.ref.set(k,o);}}
      if(any)n++;}
  }
  D.precios.n+=n;return n;
}
function precioDe(a,tarifa){
  if(a.pre)return a.pre[tarifa]||null;
  for(const t of [tarifa,'*']){const P=D.precios[t];if(!P)continue;const o=P.cb.get(a.cb)||P.reft.get(N(a.ref)+'|'+N(a.talla))||P.ref.get(N(a.ref));if(o)return o;}
  return null;}
function buscar(code){
  const b=dg(code);if(b&&D.byCB.has(b))return D.byCB.get(b);
  if(b&&b.length<13){for(const [k,a] of D.byCB)if(k.endsWith(b)&&k.replace(/^0+/,'')===b.replace(/^0+/,''))return a;}
  return null;
}
// Coste = Valoración inventario ÷ Inventario (como lo calcula el reporte de ICG). Si la talla no tiene stock,
// se usa el de la referencia (suma de valoración ÷ suma de inventario de sus tallas con stock).
function calcCostes(){
  const ref=new Map();
  for(const a of D.byRef.values())if(a.inv>0&&a.val>0){const k=N(a.ref),o=ref.get(k)||{inv:0,val:0};o.inv+=a.inv;o.val+=a.val;ref.set(k,o);}
  let n=0;
  for(const a of D.byRef.values()){if(a.costeCol==null)a.costeCol=a.coste;
    if(a.costeCol!=null&&a.costeCol>0){a.coste=a.costeCol;a.costeDe='ICG';}
    else if(a.inv>0&&a.val>0){a.coste=r2(a.val/a.inv);a.costeDe='talla';}
    else{const o=ref.get(N(a.ref));if(o){a.coste=r2(o.val/o.inv);a.costeDe='referencia';}else{a.coste=null;a.costeDe='';}}
    if(a.coste!=null)n++;}
  D.conCoste=n;
}

// Firma de los archivos de data (nombre + fecha) para saber si datos.json corresponde a los Excel publicados
function firmaDe(lists){return ['icg-files','precio-files'].map(k=>(lists[k]||[]).map(f=>f.name+'@'+f.modified).sort().join('|')).join('#');}
// datos.json: un artículo por fila con su coste y su precio en cada tarifa ya resueltos
const COLS=['cbs','ref','talla','desc','color','depto','seccion','marca','coste','costeDe','nb','nbDto','rb','rbDto','origen'];
function compactar(firma){
  const origenes=[],oi=o=>{if(!o)return -1;let i=origenes.indexOf(o.origen);if(i<0){origenes.push(o.origen);i=origenes.length-1;}return i;};
  const cbs=new Map();for(const [b,a] of D.byCB){if(!cbs.has(a))cbs.set(a,[]);cbs.get(a).push(b);}
  const filas=[];
  for(const a of D.byRef.values()){const nb=precioDe(a,'NB'),rb=precioDe(a,'RB');
    filas.push([cbs.get(a)||[],a.ref,a.talla,a.desc,a.color,a.depto,a.seccion,a.marca,a.coste,a.costeDe,nb?nb.precio:null,nb?nb.desc:null,rb?rb.precio:null,rb?rb.desc:null,[oi(nb),oi(rb)]]);}
  return {firma,cols:COLS,origenes,filas};
}
function cargarCompacto(J){
  for(const f of J.filas){const [cbs,ref,talla,desc,color,depto,seccion,marca,coste,costeDe,nb,nbDto,rb,rbDto,[o1,o2]]=f;
    const a={ref,talla,desc,color,depto,seccion,marca,coste,costeDe,cb:cbs[0]||'',
      pre:{NB:nb!=null?{precio:nb,desc:nbDto||0,origen:J.origenes[o1]||''}:null,RB:rb!=null?{precio:rb,desc:rbDto||0,origen:J.origenes[o2]||''}:null}};
    for(const b of cbs)D.byCB.set(b,a);D.byRef.set(N(ref)+'|'+N(talla),a);}
  D.precios.n=J.filas.filter(f=>f[10]!=null||f[12]!=null).length;
}
