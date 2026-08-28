# -*- coding: utf-8 -*-
"""Servidor estatico del tablero para el panel de vista previa.

Lee el puerto de la variable PORT, que es lo que el harness le pasa cuando
autoPort esta activo. Si no viene, cae al 8740 -- el mismo que usa
Abrir_Dashboard_COCO.bat, para que a mano se comporte igual que siempre.

http.server a secas no sirve aqui: toma el puerto como argumento posicional y
no mira PORT, asi que chocaba contra el servidor que el .bat ya tiene abierto.
"""
import http.server
import os
import socketserver

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main():
    puerto = int(os.environ.get("PORT") or 8740)
    os.chdir(RAIZ)
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("127.0.0.1", puerto),
                                http.server.SimpleHTTPRequestHandler) as srv:
        print("tablero en http://127.0.0.1:%d/Dashboard_COCO.html" % puerto, flush=True)
        srv.serve_forever()


if __name__ == "__main__":
    main()
