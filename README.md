# cutie-natives

Native modules I use with some of my applications. This is mainly for me but if you find it useful also go ahead. I don't generally know or work with a bunch of C++ so these might be pretty bad, sorry.

- **FocusAssist (`focus`, win32)** - Determines whether Focus Assist is currently enabled under Windows 11. It exposes 1 method, `FocusAssist.isEnabled`.
- **WinMedia (`winmedia`, win32)** - Plays media - only exposes an AudioPlayer that can play/pause/stop sounds.

## Install

Make sure you have the Windows build tools and Windows RT headers installed, then you can `npm build`.

## License

MIT