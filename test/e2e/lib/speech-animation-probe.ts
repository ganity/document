export { prepareSpeechNativeSchedule } from '../../../lib/deckwise-teaching-native-schedule';
export { createTeachingSlideshow } from '../../../lib/deckwise-teaching-slideshow';
export function nativeMotionPreset(effect: string) {
  return effect === 'fadeOut'
    ? { presetClass: 2, presetId: 10, subtype: 0 }
    : { presetClass: 1, presetId: 10, subtype: 0 };
}

export { teachingObjectChoices } from '../../../lib/deckwise-teaching-object-choice';
