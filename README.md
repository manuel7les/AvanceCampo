# Avance de Campo

App Android para que los auxiliares reporten el avance de las labores (labor, lote, cantidad o %) y puedan revisar y corregir sus propios reportes. Los reportes llegan a una Hoja de Google.

- `app/src/main/assets/index.html`: toda la app.
- `hoja/Codigo.gs`: código para pegar en la Hoja de Google (Extensiones → Apps Script) y publicar como aplicación web.
- Las labores se editan en la pestaña **Labores** (Labor | Unidad | Por válvulas Sí/No).
- Los lotes y sus válvulas en la pestaña **Lotes** (Lote | Válvulas, ej. `35-41` o `35,36,40`).
- En las labores por válvulas, cada válvula se marca como Terminada, A medias o Sin tocar; en la hoja queda en **Observaciones (válvulas)**.

Cada vez que se sube un cambio, se genera `AvanceCampo.apk` en *Releases*.
