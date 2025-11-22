import time
import os
import shutil
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

from ler_notas import ler_xml
from atualizar_excel import atualizar_excel_compras


class NotaHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory or not event.src_path.lower().endswith(".xml"):
            return

        nome = os.path.basename(event.src_path)
        print(f"[INFO] Nova nota detectada: {nome}")

        try:
            produtos, chave = ler_xml(event.src_path)

            if chave:
                atualizar_excel_compras(produtos, chave)
                print(f"[INFO] Nota processada: {nome}")
            else:
                print(f"[ERRO] Chave da nota não encontrada: {nome}")

            destino = "notas_processadas"
            os.makedirs(destino, exist_ok=True)
            shutil.move(event.src_path, os.path.join(destino, nome))

        except Exception as e:
            print(f"[ERRO] Falha ao processar {nome}: {e}")


def iniciar_monitoramento(pasta="notas"):
    os.makedirs(pasta, exist_ok=True)

    observer = Observer()
    handler = NotaHandler()

    observer.schedule(handler, pasta, recursive=False)
    observer.start()
    print(f"[INFO] Monitorando a pasta: {pasta}")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()

    observer.join()


if __name__ == "__main__":
    iniciar_monitoramento()
