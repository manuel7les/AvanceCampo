# Avance de Campo

App Android para que los auxiliares reporten el avance de las labores (labor, lote, cantidad o %) y puedan revisar y corregir sus propios reportes. Los reportes llegan a una Hoja de Google.

- `app/src/main/assets/index.html`: toda la app.
- `hoja/Codigo.gs`: código para pegar en la Hoja de Google (Extensiones → Apps Script) y publicar como aplicación web.
- Las labores y los lotes se editan en las pestañas **Labores** y **Lotes** de la hoja.

Cada vez que se sube un cambio, se genera `AvanceCampo.apk` en *Releases*.
