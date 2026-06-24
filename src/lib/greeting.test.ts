import { describe, it, expect } from 'vitest';
import { timeGreeting } from './greeting';

describe('timeGreeting', () => {
  it('returns morning before noon', () => {
    expect(timeGreeting(new Date(2026, 0, 1, 8)).text).toBe('Good morning');
  });
  it('returns evening in the early night', () => {
    expect(timeGreeting(new Date(2026, 0, 1, 19)).text).toBe('Good evening');
  });
  it('returns night in the small hours', () => {
    expect(timeGreeting(new Date(2026, 0, 1, 2)).text).toBe('Good night');
  });
});
