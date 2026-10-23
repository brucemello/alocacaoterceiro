/**
 * Verifica a sintaxe de todos os fontes antes de publicar no Apps Script.
 *  - .gs: compila todos juntos (como o Apps Script faz), o que também
 *    acusa nomes globais duplicados entre arquivos (const/let/function).
 *  - .html: compila cada bloco <script> sem src.
 * Uso: node scripts/verificar-sintaxe.js src
 */
const fs = require('fs');
const path = require('path');

const dir = process.argv[2] || 'src';
const arquivos = fs.readdirSync(dir).sort();
let erros = 0;

const gs = arquivos.filter(f => f.endsWith('.gs'));
const fonte = gs.map(f => `// ===== ${f} =====\n` + fs.readFileSync(path.join(dir, f), 'utf8')).join('\n');
try {
  new Function(fonte);
  console.log(`✔ ${gs.length} arquivo(s) .gs compilados juntos sem erro`);
} catch (e) {
  erros++;
  console.error(`✖ Erro nos arquivos .gs: ${e.message}`);
}

arquivos.filter(f => f.endsWith('.html')).forEach(f => {
  const html = fs.readFileSync(path.join(dir, f), 'utf8');
  const blocos = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi)].map(m => m[1]);
  blocos.forEach((js, i) => {
    try { new Function(js); }
    catch (e) { erros++; console.error(`✖ ${f} (script ${i + 1}): ${e.message}`); }
  });
  console.log(`✔ ${f}: ${blocos.length} bloco(s) <script>`);
});

if (!fs.existsSync(path.join(dir, 'appsscript.json'))) {
  erros++;
  console.error('✖ src/appsscript.json não encontrado');
}

if (erros) {
  console.error(`\n${erros} erro(s) encontrado(s). Nada foi publicado.`);
  process.exit(1);
}
console.log('\nTudo certo.');
