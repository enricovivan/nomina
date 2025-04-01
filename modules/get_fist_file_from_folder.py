import os

class GetFirstFileFromFolder:
    
    def __init__(self, folder_path: str, extension: str):
        self.folder_path = folder_path
        self.extension = extension

    def getFilePath(self) -> str | None:
        with os.scandir(self.folder_path) as entradas:
            for entrada in entradas:
                if entrada.is_file() and entrada.name.lower().endswith(f'.{self.extension}'):
                    print(f'Arquivo encontrado: {entrada.path}')
                    return entrada.path
                elif entrada.is_dir():
                    result = self.getFilePath(entrada.path)
                    if result:
                        print(f'Arquivo encontrado: {result}')
                        return result
        print(f'Nenhum arquivo com extensão "{self.extension}" encontrada.')            
        return None


    