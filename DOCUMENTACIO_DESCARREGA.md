# Documentació del Mètode de Descàrrega (INFOFOTO VECTOR)

Aquest document detalla el funcionament intern del sistema de descàrrega de documents Word (DOCX) per evitar modificacions accidentals en el futur.

## Mecanisme Actual (v2.32 i posteriors)

1. **Generació en Memòria/Temporal**: El document es genera utilitzant la llibreria `python-docx` i es guarda temporalment a la carpeta `WORK_DIR` (habitualment `/tmp/uploads/work` en producció o `./container/uploads/work` en local).
2. **Noms de Fitxer**:
   - **Intern**: `temp_[TIMESTAMP].docx` (per evitar col·lisions entre usuaris).
   - **Descàrrega**: `Informe_[NAT].docx` (neteat de caràcters especials).
3. **Enviament al Navegador**: S'utilitza la funció `send_file` de Flask amb els següents paràmetres crítics:
   - `as_attachment=True`: Obliga al navegador a descarregar el fitxer en lloc d'intentar obrir-lo.
   - `download_name`: Estableix el nom que veurà l'usuari.
   - `mimetype`: `application/vnd.openxmlformats-officedocument.wordprocessingml.document` (standard DOCX).
4. **Resistència a Timeouts**: El sistema és asíncron en la generació de descripcions (IA), però la generació del Word final és sincrònica i ràpida (menys de 5 segons per a 50 fotos), evitant els timeouts de 30 segons de Google Cloud Run.

## Per què no canviar-ho?
Aquest mètode ha demostrat ser el més compatible amb dispositius mòbils, navegadors d'escriptori i sistemes operatius variats, evitant problemes de bloqueig de finestres emergents (pop-ups) i garantint que el fitxer no estigui corrupte.

---
**Recordatori per a futures versions**: Si s'ha de modificar la descàrrega, s'ha de demanar permís explícit a l'usuari referenciant aquest document.
