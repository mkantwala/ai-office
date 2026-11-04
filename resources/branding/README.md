# Sangam One branding

The app name is **Sangam One**. `sangam-one-original.jpg` is the artwork supplied by the user. `sangam-one.png` is its 512-pixel PNG conversion. All logo assets retain the complete artwork and its aspect ratio.

Product display names live in `product.json`; the installed browser app also uses `resources/server/manifest.json`. Desktop, favicon, Windows tile and installer icons use the existing `code.*` resource paths consumed by the upstream build. Those filenames are implementation details. macOS `code.icns` contains PNG representations from 16 through 1024 pixels. Workbench titlebar/watermark SVGs embed PNG data so the existing CSS and asset packaging remain valid. Extension and chat logos use `media/sangam-one.png`.

Existing extension/command IDs, settings and secret-storage keys, profile folders, platform application IDs and URL protocol remain stable so this branding update can reuse the current installation's data. The executable name is `sangam-one`. Upstream license and copyright attribution remains in place.
