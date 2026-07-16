export function jsonResponse(data, status) {
  return new Response(JSON.stringify(data), {
    status: status || 200,
    headers: { "Content-Type": "application/json" }
  });
}

export function isValidRoomId(id) {
  return typeof id === "string" && /^[A-Z0-9]{4,12}$/.test(id);
}

export function parseJsonColumn(text, fallback) {
  try {
    return JSON.parse(text);
  } catch (e) {
    return fallback;
  }
}
