type TimedCue = {
  id: string;
  atMs: number;
  endMs: number;
  nativeAnimation?: unknown;
  effect: { type: string; effect?: unknown; elementId?: unknown };
};

// Keep a cue inside its verified audio interval. Never invent a new trigger for a collision.
export function sequenceSpeechFocus<T extends TimedCue>(cues: T[]): T[] {
  const rank = (cue: T) => {
    if (cue.effect.type === 'spotlight') return 3;
    if (cue.effect.type !== 'element_animation') return 4;
    if (cue.nativeAnimation) return 0;
    return ['pulse', 'gentleRotate'].includes(String(cue.effect.effect)) ? 2 : 1;
  };
  const kept = new Map<T, T>();
  for (const cue of [...cues].sort((a, b) => rank(a) - rank(b))) {
    const motion = cue.effect.type === 'element_animation';
    const occupied = [...kept.values()].filter((other) => other.atMs < cue.endMs && cue.atMs < other.endMs);
    if (motion) {
      // An intentional group reveal may animate several different targets together.
      if (
        occupied.every(
          (other) =>
            other.effect.elementId !== cue.effect.elementId &&
            other.effect.effect === cue.effect.effect &&
            other.atMs === cue.atMs &&
            other.endMs === cue.endMs,
        )
      )
        kept.set(cue, cue);
      continue;
    }
    let start = cue.atMs,
      bestStart = start,
      bestLength = 0;
    for (const other of occupied.sort((a, b) => a.atMs - b.atMs)) {
      if (other.atMs - start > bestLength) {
        bestStart = start;
        bestLength = other.atMs - start;
      }
      start = Math.max(start, other.endMs);
    }
    if (cue.endMs - start > bestLength) {
      bestStart = start;
      bestLength = cue.endMs - start;
    }
    if (bestLength > 0) kept.set(cue, { ...cue, atMs: bestStart, endMs: bestStart + bestLength });
  }
  return cues.flatMap((c) => kept.get(c) ?? []);
}
