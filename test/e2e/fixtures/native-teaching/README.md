# Synthetic native teaching media fixtures

These local fixtures contain no provider output or user content. They exercise
actual browser decoding, natural completion, pause/resume and sequential playback.
They do not prove provider generation or Host persistence.

Created with FFmpeg:

```sh
ffmpeg -f lavfi -i color=c=steelblue:s=320x180:r=20:d=3 \
  -f lavfi -i sine=frequency=440:sample_rate=24000:duration=3 \
  -c:v libx264 -pix_fmt yuv420p -c:a aac -movflags +faststart -shortest video.mp4
ffmpeg -f lavfi -i sine=frequency=330:sample_rate=24000:duration=1 \
  -c:a libmp3lame -b:a 64k music.mp3
```
