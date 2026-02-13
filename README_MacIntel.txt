===============================
Posada en marxa a Mac Intel (sense canvis de codi)
===============================

1) Doble clic a: setup_mac_intel.command
   - Crea l'entorn virtual .venv i instal·la dependències (si hi ha requirements.txt).

2) Doble clic a: start_mac.command
   - Allibera el port 5051 i arrenca l'aplicació (python app.py).

3) Obre el navegador i entra a:
   http://127.0.0.1:5051

Notes:
- Aquests scripts NO modifiquen cap arxiu de codi del projecte.
- Si manca requirements.txt, instal·la manualment les llibreries necessàries.
- Si macOS bloqueja l'execució, fes clic dret > Obre, o executa:
    chmod +x setup_mac_intel.command start_mac.command
- Qualsevol error de port: assegura't que no hi ha cap altre procés usant el port 5051.
