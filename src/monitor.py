import time
import os
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from ler_notas import ler_xml, atualizar_excel_compras


class NotaHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory and event.src_path.endswith(".xml"):
            print("Nova nota detectada:", event.src_path)

            try:
                produtos, chave = ler_xml(event.src_path)

                if chave is None:
                    print("Chave da nota não encontrada. Nota ignorada.")
                    return

                atualizar_excel_compras(produtos, chave)

            except Exception as e:
                print("Erro ao processar nota:", event.src_path)
                print(e)


def iniciar_monitoramento(pasta="notas"):
    observer = Observer()
    handler = NotaHandler()

    observer.schedule(handler, pasta, recursive=False)
    observer.start()

    print("Monitoramento iniciado na pasta:", pasta)
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
