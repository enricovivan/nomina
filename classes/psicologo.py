class Psicologo:
    def __init__(self, nome: str, crp: str):
        self.nome = nome
        self.crp = crp

    def getNome(self):
        return self.nome
    
    def getCrp(self):
        return self.crp