import { jsonResponse, isValidRoomId } from "../../_lib.js";

export async function onRequestPatch(context) {
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

  var name = typeof body.name === "string" ? body.name.trim().slice(0, 60) : "";
  var now = new Date().toISOString();

  await db.prepare("UPDATE rooms SET name = ?, updated_at = ? WHERE id = ?")
    .bind(name, now, id).run();

  return jsonResponse({ ok: true, name: name });
}
