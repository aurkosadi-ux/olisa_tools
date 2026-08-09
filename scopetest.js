/* scopetest.js — the permanent guard against the escHtml-class regression.
   Parses the real script block and proves that every function CALLED is DECLARED in a scope that
   is visible from the call site. Catches a helper that has drifted inside an IIFE. */
const fs = require('fs');
const acorn = require('acorn');

const html = fs.readFileSync(process.argv[2] || 'olisa.html', 'utf8');
const blocks = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const src = blocks.sort((a, b) => b.length - a.length)[0];

let ast;
try {
  ast = acorn.parse(src, { ecmaVersion: 2022, sourceType: 'script', locations: true });
} catch (e) {
  console.error('FAIL parse: ' + e.message);
  process.exit(1);
}

// ---- build scope tree ----
let idc = 0;
function newScope(parent, node) { return { id: ++idc, parent, node, names: new Set(), children: [] }; }
const root = newScope(null, ast);

function declarePattern(pat, scope) {
  if (!pat) return;
  switch (pat.type) {
    case 'Identifier': scope.names.add(pat.name); break;
    case 'ObjectPattern': pat.properties.forEach(p => declarePattern(p.value || p.argument, scope)); break;
    case 'ArrayPattern': pat.elements.forEach(e => declarePattern(e, scope)); break;
    case 'AssignmentPattern': declarePattern(pat.left, scope); break;
    case 'RestElement': declarePattern(pat.argument, scope); break;
  }
}

const calls = []; // {name, scope, line}
function isFn(n) { return n.type === 'FunctionDeclaration' || n.type === 'FunctionExpression' || n.type === 'ArrowFunctionExpression'; }

function walk(node, scope) {
  if (!node || typeof node.type !== 'string') return;

  // `window.foo = foo` genuinely publishes foo as a global, so a bare foo() elsewhere resolves.
  // The file uses this deliberately for helpers that must stay near their own scoped dependencies.
  if (node.type === 'AssignmentExpression' && node.left.type === 'MemberExpression' &&
      !node.left.computed && node.left.object.type === 'Identifier' &&
      (node.left.object.name === 'window' || node.left.object.name === 'globalThis') &&
      node.left.property.type === 'Identifier') {
    root.names.add(node.left.property.name);
  }

  if (node.type === 'CallExpression' && node.callee.type === 'Identifier') {
    calls.push({ name: node.callee.name, scope, line: node.loc.start.line });
  }

  if (isFn(node)) {
    if (node.type === 'FunctionDeclaration' && node.id) scope.names.add(node.id.name);
    const inner = newScope(scope, node);
    scope.children.push(inner);
    if (node.id && node.type !== 'FunctionDeclaration') inner.names.add(node.id.name);
    node.params.forEach(p => declarePattern(p, inner));
    inner.names.add('arguments');
    walk(node.body, inner);
    return;
  }
  if (node.type === 'VariableDeclaration') {
    node.declarations.forEach(d => { declarePattern(d.id, scope); walk(d.init, scope); });
    return;
  }
  if (node.type === 'ClassDeclaration' && node.id) scope.names.add(node.id.name);

  for (const k of Object.keys(node)) {
    if (k === 'loc' || k === 'start' || k === 'end') continue;
    const v = node[k];
    if (Array.isArray(v)) v.forEach(c => c && typeof c.type === 'string' && walk(c, scope));
    else if (v && typeof v.type === 'string') walk(v, scope);
  }
}
walk(ast, root);

function visible(name, scope) {
  for (let s = scope; s; s = s.parent) if (s.names.has(name)) return true;
  return false;
}

const GLOBALS = new Set(['parseInt','parseFloat','isNaN','isFinite','String','Number','Boolean','Array','Object','Date','Math','JSON','Promise','Set','Map','WeakMap','RegExp','Error','TypeError','Symbol','BigInt','fetch','setTimeout','setInterval','clearTimeout','clearInterval','requestAnimationFrame','cancelAnimationFrame','alert','confirm','prompt','encodeURIComponent','decodeURIComponent','encodeURI','decodeURI','btoa','atob','structuredClone','queueMicrotask','Blob','File','FileReader','URL','URLSearchParams','FormData','Uint8Array','ArrayBuffer','DataView','TextDecoder','TextEncoder','Intl','XLSX','ExcelJS','mammoth','pdfjsLib','google','gapi','Proxy','Reflect','AbortController','Response','Request','Headers','Image','CustomEvent','Event','MutationObserver','IntersectionObserver','indexedDB','crypto','importScripts','getComputedStyle','matchMedia','navigator','window','document','location','history','screen','localStorage','sessionStorage','caches','require','eval','Function','WeakSet','Int8Array','Uint16Array','Uint32Array','Float32Array','Float64Array']);

const bad = [];
for (const c of calls) {
  if (GLOBALS.has(c.name)) continue;
  if (!visible(c.name, c.scope)) bad.push(c);
}

if (bad.length) {
  console.error('FAIL scopetest — ' + bad.length + ' call(s) to a function not visible from the call site:');
  const seen = new Set();
  bad.forEach(b => { const k = b.name; if (!seen.has(k)) { seen.add(k); console.error('  ' + b.name + '()  first at script line ' + b.line); } });
  process.exit(1);
}
console.log('PASS scopetest — ' + calls.length + ' identifier calls, all resolvable. Functions declared: ' + root.names.size + ' at top level.');
