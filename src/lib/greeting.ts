/** Time-based greeting for the homepage lounge (see docs/02). Icon is an Icon-component name — never an emoji. */
export function timeGreeting(
  d = new Date(),
): { text: string; icon: 'sunrise' | 'sun' | 'moon' | 'moonStars' } {
  const h = d.getHours();
  if (h >= 5 && h < 12) return { text: 'Good morning', icon: 'sunrise' };
  if (h >= 12 && h < 17) return { text: 'Good afternoon', icon: 'sun' };
  if (h >= 17 && h < 21) return { text: 'Good evening', icon: 'moon' };
  return { text: 'Good night', icon: 'moonStars' };
}
