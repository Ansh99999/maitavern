/** Time-based greeting for the homepage lounge (see docs/02). */
export function timeGreeting(d = new Date()): { text: string; emoji: string } {
  const h = d.getHours();
  if (h >= 5 && h < 12) return { text: 'Good morning', emoji: '☀️' };
  if (h >= 12 && h < 17) return { text: 'Good afternoon', emoji: '🌤️' };
  if (h >= 17 && h < 21) return { text: 'Good evening', emoji: '🌙' };
  return { text: 'Good night', emoji: '🌌' };
}
