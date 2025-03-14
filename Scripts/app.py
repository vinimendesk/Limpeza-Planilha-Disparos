import flet as ft
import pandas as pd
import openpyxl

# Variável global que verifica se foi feito o upload do arquivo.
is_uploaded = False
# Variável que armazena a planilha lida.
global planilha
# Variável que armazena o caminho da planilha.
global file_path

# (Executar_button) Função remover o hífen e o 9, caso necessário.
def formatar_coluna(telefones):
    telefones = str(telefones)

    telefones = telefones.replace("-", "")

    if len(telefones) > 10:
        telefones = telefones[:2] + telefones[3:]
                
    return telefones

# Componentes.
# Título do app.
titulo_text = ft.Text(
    "FORMATAÇÃO TELEFONES PLANILHA DISPAROS",
    size=18,
    weight=ft.FontWeight.BOLD,
    font_family="Roboto",
    text_align=ft.TextAlign.CENTER
)

# Título do tutorial.
tutorial_text = ft.Text(
    "TUTORIAL",
    size=18,
    weight=ft.FontWeight.BOLD,
    font_family="Roboto",
    text_align=ft.TextAlign.CENTER
)

# Botões.
upload_button = ft.ElevatedButton(
    "Upload",
    icon=ft.icons.UPLOAD_FILE,
)

executar_button = ft.ElevatedButton(
    "Executar",
    icon=ft.icons.PLAY_CIRCLE_FILLED
)

tutorial_button = ft.ElevatedButton(
    "Tutorial",
    icon=ft.icons.ARROW_FORWARD,
)

back_button = ft.ElevatedButton(
    "Voltar",
    icon=ft.icons.ARROW_BACK
)

# Status inicial.
status = ft.Text("Esperando planilha disparos...", text_align=ft.TextAlign.CENTER)

# Tela principal.
def main_page():
    return ft.View(
        "/main",
        controls=[
            ft.Column(
                [
                    # Texto do título.
                    ft.Row(
                        [titulo_text],
                        alignment=ft.MainAxisAlignment.CENTER,
                    ),
                    ft.Container(height=200),
                    # Faixa dos botões.
                    ft.Row(
                        [upload_button, executar_button],
                        alignment=ft.MainAxisAlignment.CENTER,
                    ),
                    ft.Container(height=26),
                    # Card com informações.
                    ft.Card(
                        content=ft.Container(
                            content=ft.Column(
                                [
                                    ft.Text("Status:"),  # Rótulo "Status"
                                    status,  # Texto de status centralizado
                                ]
                            ),
                            width=500,
                            padding=10,
                        )
                    ),
                    ft.Container(height=200),
                    # Botão tutorial.
                    ft.Row(
                        [tutorial_button],
                        alignment=ft.MainAxisAlignment.END,
                    )
                ],
                alignment=ft.MainAxisAlignment.CENTER,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER
            )
        ]
    )

# Tela de tutorial.
def tutorial_page():
    return ft.View(
        "/tutorial",
        controls=[
            ft.Column(
                [
                    ft.Row(
                        [back_button, ft.Container(width=500), tutorial_text],
                        alignment=ft.MainAxisAlignment.START
                    )
                ]
            )
        ]
    )

# Função principal para gerenciar as views.
def main(page: ft.Page):

    # Função do botão "Tutorial".
    def go_to_tutorial(e):
        page.views.append(tutorial_page())
        page.update()

    # Função para voltar à tela principal.
    def go_back(e):
        page.views.pop()
        page.update()

    # Função para upload do arquivo.
    def upload_file_result(e: ft.FilePickerResultEvent):
        global is_uploaded
        global planilha
        global file_path
        # Se arquivo encontrado.
        if e.files:
            # Obtém o caminho do arquivo selecionado.
            file_path = e.files[0].path

            try:
                planilha = pd.ExcelFile(file_path)
                is_uploaded = True
                status.value = "Arquivo carregado com sucesso!"
                status.update()
            except Exception as err:
                status.value = f"Erro ao processar o arquivo: {err}"
                status.update()
    
    # Função para fazer a limpeza da planilha.
    def execute_file(e):
        global is_uploaded
        if is_uploaded:
            status.value = "Processando arquivo..."
            status.update()

            # Dicionário para armazena os dados formatados.
            dados_processados = {}

            # Percorre todos os estados.
            for estado in planilha.sheet_names:
                
                # Abre a aba de cada estado.
                dados = planilha.parse(estado)

                if "Telefones" in dados.columns:
                    # Faz a formatação.
                    dados["Telefones"] = dados["Telefones"].apply(formatar_coluna)
                else:
                    dados["Numero"] = dados["Numero"].apply(formatar_coluna)

                # Armazena os dados em um dicionário.
                dados_processados[estado] = dados

            with pd.ExcelWriter(file_path+r"Atualizada.xlsx") as writer:
                for estado, dados in dados_processados.items():
                    dados.to_excel(writer, sheet_name = estado, index = False)
                    print("Gravação completa.")
            status.value = "Planilha processada com sucesso!"
            status.update()

        else:
            status.value = "Faça upload da planilha primeiro."
            status.update()

    #Cria o FilePicker para selecionar os arquivos.
    file_picker = ft.FilePicker(on_result=upload_file_result)

    # Associando funções aos botões.
    tutorial_button.on_click = go_to_tutorial
    back_button.on_click = go_back
    upload_button.on_click = lambda e: file_picker.pick_files(
        allow_multiple = False,
        allowed_extensions = ["xlsx", "xls"]
        )
    executar_button.on_click = execute_file

    # Configurando a página inicial.
    page.overlay.append(file_picker)
    page.views.append(main_page())
    page.update()

ft.app(target=main)
