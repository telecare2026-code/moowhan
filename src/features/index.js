// ==================== FEATURE REGISTRY ====================
// Optional capabilities plug in here without touching App.jsx.
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
//     Component,                                // receives { state, actions, monthLabels }
//   },
//   headerActions: (ctx) => <button .../>,      // optional: controls in the header bar
//   exportSheets: (ctx) => [{ name, aoa }],     // optional: extra sheets in the export
//   onProcessed: (ctx) => void,                 // optional: runs after processing finishes
// }
//
// ctx = { state, actions, monthLabels }  (see state/useConsolidator.js)
export const FEATURES = [];

export const registerFeature = (feature) => {
  if (!feature?.id) throw new Error('feature.id is required');
  const idx = FEATURES.findIndex((f) => f.id === feature.id);
  if (idx === -1) FEATURES.push(feature); else FEATURES[idx] = feature;
  FEATURES.sort((a, b) => (a.order ?? 100) - (b.order ?? 100));
  return feature;
};
