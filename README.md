**WorldEditor**

- Carga de Graphics.AO (graficos comprimidos) en Recursos\Graficos. Si NO existe, carga los .PNG's/.BMP's descomprimidos. (jopiortiz)
- Ahora en el Dialog para abrir/guardar mapa se muestra primero los ".map" (jopiortiz)
- Carga de Graficos.ini, si no existe, carga Graficos.ind (jopiortiz-Wyr0X)
- Tabule y aprolije bastante el codigo. (jopiortiz)
- Los recursos del cliente se cargan desde la carpeta Recursos. (jopiortiz)
- Saqué la carga de recursos comprimida en archivos .DRAG (jopiortiz)
- Si no existe MiniMap.dat, carga los .BMP's de los MiniMapas de Recursos\Graficos\MiniMapa. (jopiortiz)
- Lectura y Escritura de .INI via clsIniManager. (jopiortiz)
- Se aplicó el [parche aportado por Mufarety](https://www.gs-zone.org/temas/world-editor-de-lorwik-parcheado.99336/).
- Nuevo botón corre traslados de los mapas. Modificar `ClienteWidth` y `ClienteHeight` por `Tamaño del Render[X e Y] / 32`

![imagen](https://cdn.discordapp.com/attachments/668202050743435265/670359756040437812/WE_Demo.png)
