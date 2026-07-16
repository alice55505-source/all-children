import { jsonResponse, isValidRoomId, parseJsonColumn } from "../../../_lib.js";

// Reads the current stats_json, applies one congregation's change, and
// writes the whole row back - all within this single request, so two
// people uploading different congregations around the same time each get
// merged into whatever is currently on the server instead of one
// clobbering the other with a stale full-state snapshot taken at their
// last page load.
async function withStats(db, id, mutate) {
  var row = await db.prepare("SELECT stats_json FROM rooms WHERE id = ?").bind(id).first();
  if (!row) return null;

  var stats = parseJsonColumn(row.stats_json, {});
  mutate(stats);

  var now = new Date().toISOString();
  await db.prepare("UPDATE rooms SET stats_json = ?, updated_at = ? WHERE id = ?")
    .bind(JSON.stringify(stats), now, id).run();

  return now;
}

export async function onRequestPut(context) {
  var id = String(context.params.id || "").toUpperCase();
  if (!isValidRoomId(id)) return jsonResponse({ error: "房間代碼格式錯誤" }, 400);

  var name = decodeURIComponent(context.params.name || "");
  if (!name) return jsonResponse({ error: "缺少召會名稱" }, 400);

  var body;
  try {
    body = await context.request.json();
  } catch (e) {
    return jsonResponse({ error: "請求格式錯誤" }, 400);
  }

  if (typeof body !== "object" || body === null) {
    return jsonResponse({ error: "資料格式錯誤" }, 400);
  }

  var db = context.env.DB;
  var updatedAt = await withStats(db, id, function (stats) {
    stats[name] = body;
  });

  if (updatedAt == null) return jsonResponse({ error: "找不到這個房間" }, 404);
  return jsonResponse({ ok: true, updatedAt: updatedAt });
}

export async function onRequestDelete(context) {
  var id = String(context.params.id || "").toUpperCase();
  if (!isValidRoomId(id)) return jsonResponse({ error: "房間代碼格式錯誤" }, 400);

  var name = decodeURIComponent(context.params.name || "");
  if (!name) return jsonResponse({ error: "缺少召會名稱" }, 400);

  var db = context.env.DB;
  var updatedAt = await withStats(db, id, function (stats) {
    delete stats[name];
  });

  if (updatedAt == null) return jsonResponse({ error: "找不到這個房間" }, 404);
  return jsonResponse({ ok: true, updatedAt: updatedAt });
}
