import json
import os

class LoadConfig:
    def __init__(self, filename="config.json"):
        self.filename = filename
        self.config = self._load_defaults()
        
    def _load_defaults(self):
        # Valores padrão caso o arquivo não exista
        return {
            "planilha": {
                "entrada": {
                    "posicao": 1,

                    "celula_nome": "F6",
                    "celula_processo": "X6",

                    "celula_dia": "I43",
                    "nome_combobox_mes": "Drop-down 726",
                    "celula_ano": "R43"
                },
                "cadastro": {
                    "posicao": 15,

                    "celula_x_dmtran": "B9",
                    "celula_x_transitar": "B11",

                    "celula_x_meire": "B23",
                    "celula_x_onilce": "B25"
                }
            },

            "gerador": {

            }
        }
        
    def load(self):
        # carrega os valores do arquivo json
        if os.path.exists(self.filename):
            try:
                with open(self.filename, 'r') as file:
                    self.config = json.load(file)
            except Exception as e:
                print(f"Erro ao abrir arquivo {self.filename}. Erro: {e}")
                
        return self.config
    

    def save(self, new_data) -> bool:
        # salva as configurações no arquivo json
        try:
            with open(self.filename, 'w') as file:
                json.dump(new_data, file, indent=4)
            self.config = new_data
            return True
        except Exception as e:
            print(f"Erro ao salvar arquivo {self.filename} com dados: {new_data}. Erro: {e}")
            return False
        