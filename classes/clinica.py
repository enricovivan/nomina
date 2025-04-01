class Clinica:
    def __init__(self, nome: str, endereco: str):
        self.nome = nome
        self.endereco = endereco

    def getNome(self) -> str:
        return self.nome
    
    def getEndereco(self) -> str:
        return self.endereco