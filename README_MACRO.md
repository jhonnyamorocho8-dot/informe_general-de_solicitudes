# Estado de la macro corregida

**No reemplaces el módulo completo del libro con este archivo.** El borrador anterior `macro_corregida_general.txt` no es una versión completa validada: su lógica de actualización de Word fue reducida y no debe usarse como sustituto del módulo original.

Para corregir el error 1004 de celdas combinadas en el paso de ajuste de la tabla:

1. Haz una copia de seguridad de `informe general.xlsm`.
2. En Excel pulsa `Alt+F11` y abre el módulo original que contiene `AjustarTabla`.
3. Reemplaza únicamente ese procedimiento por el contenido de [`macro_fix_1004.bas`](macro_fix_1004.bas). No importes el `.bas` como módulo adicional, ya que el procedimiento tiene el mismo nombre que el original.
4. En el procedimiento principal, asegúrate de que `ActualizarDashboardVBA` se ejecute antes de `ThisWorkbook.Save`.
5. En el editor VBA usa **Depurar > Compilar VBAProject** antes de ejecutar la macro.

La carga general debe continuar sin filtros (`FILTRAR_SOLO_GESTANTES = False`). Este parche evita borrar filas completas al reducir la tabla y evita rellenar hacia abajo las columnas de datos. No modifica el libro binario ni se ha podido probar en Excel de escritorio desde aquí; guarda una copia antes de ejecutar y verifica el resultado.
