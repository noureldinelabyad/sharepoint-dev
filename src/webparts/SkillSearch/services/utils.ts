
/** Two-letter initials from a display name */
export function getInitials(name: string): string {
  const parts = (name || "").trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return "?";
  const first = parts[0][0] || "";
  const last = (parts[parts.length - 1] || "")[0] || "";
  return (first + last).toUpperCase();
}

/** Deterministic pleasing color from a string (used for avatar background) */
export function colorFromString(s: string): string {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) % 360;
  return `hsl(${h}, 60%, 55%)`;
}

/** Create a PNG dataURL avatar with initials (client-side canvas) */
export function makeInitialsAvatar(name: string, size = 72): string {
  const canvas = document.createElement("canvas");
  canvas.width = size; canvas.height = size;
  const ctx = canvas.getContext("2d")!;
  ctx.fillStyle = colorFromString(name || " ");
  ctx.fillRect(0, 0, size, size);

  const initials = getInitials(name);
  ctx.fillStyle = "#fff";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.font = `bold ${Math.round(size * 0.46)}px system-ui, Segoe UI, Arial, sans-serif`;
  ctx.fillText(initials, size / 2, Math.round(size * 0.57));

  return canvas.toDataURL("image/png");
}

/** Split an array into equally-sized chunks (last chunk may be shorter). */
export function chunk<T>(arr: T[], size: number): T[][] {
  const n = Math.max(1, (size | 0)); // coerce to int, min 1
  const out: T[][] = [];
  for (let i = 0; i < arr.length; i += n) {
    out.push(arr.slice(i, i + n));
  }
  return out;
}
