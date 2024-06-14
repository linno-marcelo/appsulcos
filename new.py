import openpyxl
from kivy.lang import Builder
from kivymd.app import MDApp




KV = '''
MDFloatLayout:
    MDFloatLayout:
        md_bg_color: 1,1,1,1  # Altere para a cor de fundo desejada (R, G, B, A)

    MDLabel:
        text: "FORMULÁRIO DE PROFUNDIDADE DE SULCO"
        halign: "center"
        pos_hint: {"center_x": 0.5, "center_y": 0.97}
        font_size: 18
        theme_text_color: "Custom"
        text_color: 0.73, 0.55, 0.73, 1    

    BoxLayout:
        orientation: 'horizontal'
        pos_hint: {"center_x": 0.5, "center_y": 0.92}
        size_hint: 0.6, 0.25

        Image:
            source: "imagens/29775.png"  # Insira o caminho para a sua imagem
            size_hint: 0.5, 0.6
        
        Widget:
            size_hint_x: 0.1  # Espaçamento entre as imagens
        
        Image:
            source: "imagens/logo_raizen.png"  # Insira o caminho para a sua imagem
            size_hint: 0.5, 0.9

    BoxLayout:
        orientation: 'horizontal'
        pos_hint: {"center_x": 0.5, "center_y": 0.79}
        size_hint_x: 0.8
        size_hint_y: 0.1

        MDTextField:
            id: equipamento
            hint_text: "Equipamento:"
            max_text_length: 10
            font_size: 12
            text_color: 1, 1, 1, 1
            size_hint_x: 0.45

        Widget:
            size_hint_x: 0.1 # espaçamento entre os campos

        MDTextField:
            id: cs
            hint_text: "CS:"
            max_text_length: 6
            font_size: 12
            text_color: 1, 1, 1, 1
            size_hint_x: 0.45



    MDTextField:
        id: data
        pos_hint: {"center_x": 0.5, "center_y": 0.71}
        size_hint_x: 0.4
        size_hint_y: 0.1
        hint_text: "Data:"
        max_text_length: 10      
        font_size: 12

    MDTextField:
        id: codigo_pneu
        pos_hint: {"center_x": 0.5, "center_y": 0.63}
        size_hint_x: 0.4
        size_hint_y: 0.1
        mode: "rectangle"
        hint_text: "Código Pneu:"
        max_text_length: 10
        font_size: 12
        text_color: 1, 1, 1, 1

    MDTextField:
        id: sulco1
        pos_hint: {"center_x": 0.5, "center_y": 0.52}
        size_hint_x: 0.4
        size_hint_y: 0.1
        mode: "rectangle"
        hint_text: "Sulco 1:"
        max_text_length: 4
        font_size: 12
        text_color: 1, 1, 1, 1

    MDTextField:
        id: sulco2   
        pos_hint: {"center_x": 0.5, "center_y": 0.41}
        size_hint_x: 0.4
        size_hint_y: 0.1
        mode: "rectangle"
        hint_text: "Sulco 2:"
        max_text_length: 4
        font_size: 12
        text_color: 1, 1, 1, 1

    MDTextField:
        id: sulco3   
        pos_hint: {"center_x": 0.5, "center_y": 0.30}
        size_hint_x: 0.4
        size_hint_y: 0.1
        mode: "rectangle"
        hint_text: "Sulco 3:"
        max_text_length: 4
        font_size: 12
        text_color: 1, 1, 1, 1

    BoxLayout:
        orientation: 'horizontal'
        pos_hint: {"center_x": 0.5, "center_y": 0.20}
        size_hint_x: 0.8
        size_hint_y: 0.1

        MDTextField:
            id: medida
            hint_text: "Medida:"
            max_text_length: 5
            font_size: 12
            text_color: 1, 1, 1, 1
            size_hint_x: 0.45
            mode: "rectangle"
        Widget:
            size_hint_x: 0.1  # Espaçamento entre os campos

        MDTextField:
            id: calibrada
            hint_text: "Calibrada:"
            max_text_length: 5
            font_size: 12
            text_color: 1, 1, 1, 1
            size_hint_x: 0.45
            mode: "rectangle"
    MDRaisedButton:
        text: "Salvar"
        size_hint_x: 0.4        
        size_hint_y: 0.05
        pos_hint: {"center_x": 0.5, "center_y": 0.05}
        font_size: 12
        md_bg_color: 0.47,0.11,0.46, 1
        on_release: app.salvar_informacoes()
'''


class MyApp(MDApp):
    def build(self):
        return Builder.load_string(KV)
    def salvar_informacoes(self):
        dados = {
            "equipamento": self.root.ids.equipamento.text,
            "cs": self.root.ids.equipamento.text,
            "data": self.root.ids.data.text,
            "codigo_pneu": self.root.ids.codigo_pneu.text,
            "sulco1": self.root.ids.sulco1.text,
            "sulco2": self.root.ids.sulco2.text,
            "sulco3": self.root.ids.sulco3.text,
            "medida": self.root.ids.medida.text,
            "calibrada": self.root.ids.calibrada.text
        }

        try:
            wb = openpyxl.load_workbook("dados.xlsx")
            ws = wb.active
        except FileNotFoundError:
            wb = openpyxl.Workbook()
            ws = wb.active         
            ws.append(["Equipamento", "Data", "Código Pneu", "Sulco1", "Sulco2", "Sulco3", "Medida", "Calibrada"])

        ws.append([dados["equipamento"], dados["data"], dados["codigo_pneu"], dados["sulco1"], dados["sulco2"], dados["sulco3"], dados["medida"], dados["calibrada"]])
        wb.save("dados.xlsx")
        print("Informações salvas com sucesso!")

MyApp().run()


