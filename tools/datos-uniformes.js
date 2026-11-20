// Arma uniformes/datos.json con la data que publica publicar.ps1 (lista.json, DATA ICG y PRECIOS RETAL ICG),
// usando la misma lectura que la APP (uniformes/datos.js). Así la APP baja un solo archivo en vez de ~60 MB de Excel.
// Lo corre GitHub Actions (.github/workflows/datos-uniformes.yml) cada vez que cambia la data.
const fs = require('fs'), path = require('path'), vm = require('vm'), XLSX = require('xlsx');
const root = path.join(__dirname, '..');
const lists = JSON.parse(fs.readFileSync(path.join(root, 'lista.json'), 'utf8')).lists;
const ctx = vm.createContext({ XLSX, console });
vm.runInContext(fs.readFileSync(path.join(root, 'uniformes', 'datos.js'), 'utf8') +
  '\n;this.api={D,indexIcg,indexPrecios,calcCostes,compactar,firmaDe};', ctx);
const { D, indexIcg, indexPrecios, calcCostes, compactar, firmaDe } = ctx.api;
const leer = name => XLSX.read(fs.readFileSync(path.join(root, name)), { type: 'buffer' });

for (const f of lists['icg-files'] || []) indexIcg(leer(f.name));
calcCostes();
for (const f of lists['precio-files'] || []) indexPrecios(leer(f.name), path.basename(f.name));
const J = compactar(firmaDe(lists));

// Promociones con fecha (tools/promos/*.json), p. ej. Black Friday. Mientras están vigentes (fecha de Panamá),
// el %Dto de la tarifa se reemplaza por el de la promo, código por código. Al vencer, la corrida diaria
// de la Action vuelve a armar datos.json con el descuento regular de la lista de ICG.
// Formato: {"nombre":"...","tarifa":"NB","desde":"AAAA-MM-DD","hasta":"AAAA-MM-DD","dto":{"<cód. barras>":70,...},"bruto":{"<cód. barras>":99.95,...}}
const hoy = new Date(Date.now() - 5 * 3600e3).toISOString().slice(0, 10);   // Panamá, UTC-5
const pdir = path.join(__dirname, 'promos');
const COL = { NB: [10, 11, 0], RB: [12, 13, 1] };
for (const f of fs.existsSync(pdir) ? fs.readdirSync(pdir).filter(n => n.endsWith('.json')).sort() : []) {
  const P = JSON.parse(fs.readFileSync(path.join(pdir, f), 'utf8')), c = COL[P.tarifa];
  if (!c) { console.log(f + ': tarifa desconocida ' + P.tarifa); continue; }
  if (hoy < P.desde || hoy > P.hasta) { console.log(f + ': no vigente hoy (' + hoy + '), se ignora'); continue; }
  let oi = J.origenes.indexOf(P.nombre); if (oi < 0) { J.origenes.push(P.nombre); oi = J.origenes.length - 1; }
  let n = 0;
  for (const fila of J.filas) {
    const b = fila[0].find(x => P.dto[x] != null); if (b == null) continue;
    if (fila[c[0]] == null) { if (!P.bruto || P.bruto[b] == null) continue; fila[c[0]] = P.bruto[b]; }   // sin precio en la lista: usa el bruto de la promo
    fila[c[1]] = P.dto[b]; fila[14][c[2]] = oi; n++;
  }
  console.log(f + ': ' + n + ' artículos con el descuento de ' + P.nombre);
}
fs.writeFileSync(path.join(root, 'uniformes', 'datos.json'), JSON.stringify(J));
console.log(J.filas.length + ' artículos, ' + D.conCoste + ' con coste, ' + D.precios.n + ' con precio');
