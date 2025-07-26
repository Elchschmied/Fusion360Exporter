# TotalExport: Automatischer Projekt-Export für Fusion 360
# Autor: Justin Nesselrotte (modifiziert & erweitert)
# Beschreibung: Exportiert alle Fusion 360-Projekte automatisiert mit Fortschrittsanzeige, 
# Fehlerbehandlung, Protokollierung und Fortsetzungsfunktion.

import adsk.core, adsk.fusion, adsk.cam, traceback
import os
import time
import threading

# Speicherorte definieren
EXPORT_DIR = os.path.expanduser("~/Fusion360Exports")  # Exportverzeichnis
PROGRESS_LOG = os.path.join(EXPORT_DIR, "project_progress.tsv")  # Fortschrittsdatei
EXPORT_LOG = os.path.join(EXPORT_DIR, "exported_files.log")      # Exportierte Dateien

# Projektfortschritt laden oder neu beginnen
def load_export_progress():
    if not os.path.exists(PROGRESS_LOG):
        return set()
    with open(PROGRESS_LOG, "r", encoding="utf-8") as f:
        return set(line.strip() for line in f)

# Fortschritt abspeichern
def save_export_progress(entry):
    with open(PROGRESS_LOG, "a", encoding="utf-8") as f:
        f.write(entry + "\n")

# Exportierte Datei in Log schreiben
def log_exported_file(path):
    with open(EXPORT_LOG, "a", encoding="utf-8") as f:
        f.write(path + "\n")

# Hauptklasse für den Exportprozess
class FusionExporter:
    def __init__(self):
        self.app = adsk.core.Application.get()
        self.ui = self.app.userInterface
        self.data = self.app.data

    def export_all_projects(self):
        try:
            # Fortschritt laden
            exported_projects = load_export_progress()

            # Alle Projekte durchlaufen
            for project in self.data.dataProjects:
                project_name = project.name

                # Projekt überspringen, wenn bereits exportiert
                if project_name in exported_projects:
                    continue

                # UI-Status anzeigen
                self.ui.messageBox(f"Exportiere Projekt: {project_name}")

                # Hauptordner im Projekt öffnen
                root_folder = project.rootFolder
                self._export_folder(project_name, root_folder)

                # Projekt als exportiert markieren
                save_export_progress(project_name)

        except Exception as e:
            self.ui.messageBox(f"Fehler beim Exportieren: {str(e)}\n{traceback.format_exc()}")

    def _export_folder(self, project_name, folder):
        # Durchlaufe alle Zeichnungen in diesem Ordner
        for item in folder.dataFiles:
            if item.fileExtension.lower() == "f3d":  # Nur Fusion-Dateien
                self._export_design(project_name, item)

        # Auch Unterordner exportieren
        for subfolder in folder.dataFolders:
            self._export_folder(project_name, subfolder)

    def _export_design(self, project_name, item):
        doc = item.open("ReadOnly")  # Zeichnung öffnen
        design = adsk.fusion.Design.cast(doc.products.itemByProductType("DesignProductType"))
        name = item.name
        drawing_name = design.rootComponent.name if design else "Unbenannt"

        # Exportpfad erstellen
        filename = f"{project_name}_{name}_{drawing_name}".replace(" ", "_")
        filename = "".join(c if c.isalnum() or c in "-_." else "_" for c in filename)  # Dateisystemsicher
        filepath = os.path.join(EXPORT_DIR, filename + ".f3d")

        # Export durchführen
        item.download(filepath)
        log_exported_file(filepath)

        # UI-Status aktualisieren
        self.ui.messageBox(f"Gespeichert: {filename}", "Exportiert")

        doc.close(False)

# Starte Export im Hintergrundthread (für UI-Reaktionsfähigkeit)
def run(context):
    def threaded_export():
        try:
            exporter = FusionExporter()
            exporter.export_all_projects()
        except:
            print("Exportfehler:\n" + traceback.format_exc())

    thread = threading.Thread(target=threaded_export)
    thread.start()
