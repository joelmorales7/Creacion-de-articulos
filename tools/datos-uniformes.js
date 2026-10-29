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
fs.writeFileSync(path.join(root, 'uniformes', 'datos.json'), JSON.stringify(J));
console.log(J.filas.length + ' artículos, ' + D.conCoste + ' con coste, ' + D.precios.n + ' con precio');
