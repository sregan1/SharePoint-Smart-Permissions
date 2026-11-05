// Parses #rgb / #rrggbb. Returns null for anything else (e.g. rgb(), named colors).
function parseHex(hex: string): [number, number, number] | null {
  const m = /^#?([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(hex.trim());
  if (!m) return null;
  let h = m[1];
  if (h.length === 3) h = h.split('').map((c) => c + c).join('');
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
}

function relativeLuminance([r, g, b]: [number, number, number]): number {
  const lin = (v: number): number => {
    const s = v / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
}

function contrastRatio(l1: number, l2: number): number {
  const [hi, lo] = l1 >= l2 ? [l1, l2] : [l2, l1];
  return (hi + 0.05) / (lo + 0.05);
}

const LIGHT_TEXT = '#ffffff';
const DARK_TEXT = '#242424';

/**
 * Text color that stays readable on top of `background`. The site theme's primary color
 * can be light (yellow, light green), where white text fails contrast — pick whichever of
 * white or near-black contrasts more. Falls back to white when the color can't be parsed.
 */
export function readableTextOn(background: string): string {
  const rgb = parseHex(background);
  if (!rgb) return LIGHT_TEXT;
  const lum = relativeLuminance(rgb);
  const withLight = contrastRatio(1, lum);
  const withDark = contrastRatio(relativeLuminance([0x24, 0x24, 0x24]), lum);
  return withLight >= withDark ? LIGHT_TEXT : DARK_TEXT;
}
