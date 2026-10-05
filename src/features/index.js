// ==================== FEATURE REGISTRY ====================
// Every `src/features/<id>/index.jsx` that `export default`s a feature object is
// picked up automatically (Vite import.meta.glob), so adding a capability never
// requires editing App.jsx or this file.
//
// A feature is a plain object:
// {
//   id: 'qa',                                   // unique, also used as the tab id
//   order: 50,                                  // tab position; core tabs use 10..40
//   tab: {                                      // optional: adds a tab
//     label: 'ตรวจคุณภาพ',
//     Icon: Icons.Check,                        // component from components/Icons.jsx
//     enabled: (ctx) => !!ctx.state.processedData,
//     tooltip: 'ต้องประมวลผลไฟล์ก่อน',           // shown while disabled
//     badge: (ctx) => 3,                        // optional small counter on the tab
//     Component,                                // receives { state, actions, monthLabels }
//   },
//   headerActions: (ctx) => <button .../>,      // optional: controls in the header bar
//   exportSheets: (ctx) => [{ name, aoa }],     // optional: extra sheets in the export
//   onProcessed: (ctx) => void,                 // optional: runs after processing finishes
// }
//
// ctx = { state, actions, monthLabels }  (see state/useConsolidator.js)
const modules = import.meta.glob('./*/index.jsx', { eager: true });

export const FEATURES = Object.entries(modules)
  .map(([file, mod]) => {
    const feature = mod?.default;
    if (!feature?.id) {
      console.warn(`features: ${file} has no default export with an id, skipped`);
      return null;
    }
    return feature;
  })
  .filter(Boolean)
  .sort((a, b) => (a.order ?? 100) - (b.order ?? 100));

const seen = new Set();
FEATURES.forEach((f) => {
  if (seen.has(f.id)) console.warn(`features: duplicate feature id "${f.id}"`);
  seen.add(f.id);
});
