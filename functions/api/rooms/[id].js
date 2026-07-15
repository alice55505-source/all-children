function jsonResponse(data, status) {
  return new Response(JSON.stringify(data), {
    status: status || 200,
    headers: { "Content-Type": "application/json" }
  });
}

function isValidRoomId(id) {
  return typeof id === "string" && /^[A-Z0-9]{4,12}$/.test(id);
}

export async function onRequestGet(context) {
  var id = String(context.params.id || "").toUpperCase();
  if (!isValidRoomId(id)) return jsonResponse({ error: "房間代碼格式錯誤" }, 400);

  var db = context.env.DB;
  var row = await db.prepare("SELECT groups_json, stats_json FROM rooms WHERE id = ?").bind(id).first();
  if (!row) return jsonResponse({ error: "找不到這個房間" }, 404);

  var groups, stats;
  try {
    groups = JSON.parse(row.groups_json);
  } catch (e) {
    groups = [];
  }
  try {
    stats = JSON.parse(row.stats_json);
  } catch (e) {
    stats = {};
  }

  return jsonResponse({ id: id, groups: groups, stats: stats });
}

export async function onRequestPut(context) {
  var id = String(context.params.id || "").toUpperCase();
  if (!isValidRoomId(id)) return jsonResponse({ error: "房間代碼格式錯誤" }, 400);

  var db = context.env.DB;
  var existing = await db.prepare("SELECT id FROM rooms WHERE id = ?").bind(id).first();
  if (!existing) return jsonResponse({ error: "找不到這個房間" }, 404);

  var body;
  try {
    body = await context.request.json();
  } catch (e) {
    return jsonResponse({ error: "請求格式錯誤" }, 400);
  }

  if (!Array.isArray(body.groups) || typeof body.stats !== "object" || body.stats === null) {
    return jsonResponse({ error: "資料格式錯誤" }, 400);
  }

  var now = new Date().toISOString();
  await db.prepare(
    "UPDATE rooms SET groups_json = ?, stats_json = ?, updated_at = ? WHERE id = ?"
  ).bind(JSON.stringify(body.groups), JSON.stringify(body.stats), now, id).run();

  return jsonResponse({ ok: true, updatedAt: now });
}
