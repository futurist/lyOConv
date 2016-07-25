@rem HIDE CMD WINDOW if you don't want to see cmd
title lyoconv
@ping /n 2 127.1>nul
@nircmd.exe win hide title "lyoconv"


set DEBUG=*
supervisor -i . index.js
