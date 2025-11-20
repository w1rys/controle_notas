import time
import os
import shutil
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

from ler_notas import ler_xml
from atualizar_excel import atualizar_excel_compras


class NotaHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory or not event.src_path.endswith(".xml"):
            return

        nome = os.path.basename(event.src_path)
        print(f"Nova nota detectada: {nome}")

        try:
            produtos, chave = ler_xml(event.src_path)

            if chave:
                atualizar_excel_compras(produtos, chave)
            else:
                print("Chave não encontrada. Nota pulada.")

            # mover nota para notas_processadas/
            destino = "notas_processadas"
            os.makedirs(destino, exist_ok=True)
            shutil.move(event.src_path, os.path.join(destino, nome))
            print(f"Nota movida para: {destino}/{nome}")

        except Exception as e:
            print("Erro ao processar nota:", nome)
            print(e)


def iniciar_monitoramento(pasta="notas"):
    observer = Observer()
    handler = NotaHandler()

    observer.schedule(handler, pasta, recursive=False)
    observer.start()

    print(f"Monitoramento iniciado na pasta: {pasta}")
    print("Aguardando novas notas XML...")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("Monitoramento encerrado.")

    observer.join()


if __name__ == "__main__":
    iniciar_monitoramento()
