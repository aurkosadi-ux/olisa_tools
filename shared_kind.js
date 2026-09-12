function itemKind(desc) {
  const d = String(desc || '');
  if (/cross\s*divider/i.test(d)) return 'cross';
  if (/punch|perforat/i.test(d)) return 'punch';
  return 'master';
}
const KIND_LABEL = { master: 'Master Carton', punch: 'Chip Box (Punch)', cross: 'Cross Divider' };
function itemLabel(r) {
  return KIND_LABEL[itemKind(String((r && r.itemDesc) || '') + ' ' + String((r && r.item) || ''))];
}
module.exports = { itemKind, KIND_LABEL, itemLabel };