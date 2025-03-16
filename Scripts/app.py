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

# Textos Tutoriais.
passo1_text = ft.Text(
    "1. Clique no botão upload e carregue a planilha disparos.", 
    size = 18,
    weight=ft.FontWeight.BOLD,
    font_family="Roboto",)

# Textos Tutoriais.
passo2_text = ft.Text(
    "2. Procure pelo arquivo da planilha disparos e clique em 'abrir'.", 
    size = 18,
    weight=ft.FontWeight.BOLD,
    font_family="Roboto",)

# Textos Tutoriais.
passo3_text = ft.Text(
    "3. Clique no botão executar e espere o processamento da planilha.", 
    size = 18,
    weight=ft.FontWeight.BOLD,
    font_family="Roboto",)

# Textos Tutoriais.
passo4_text = ft.Text(
    "4. Procure pela planilha disparos atualizada na mesma pasta da original.", 
    size = 18,
    weight=ft.FontWeight.BOLD,
    font_family="Roboto",)

# Imagens.
passo1_img = ft.Image(
    src="img/passo1.png",
    width=600,
    height=300,
    fit=ft.ImageFit.FILL
)

passo2_img = ft.Image(
    src="img/passo2.png",
    width=600,
    height=300,
    fit=ft.ImageFit.FILL
)

passo3_img = ft.Image(
    src="img/passo3.png",
    width=600,
    height=300,
    fit=ft.ImageFit.FILL
)

passo4_img = ft.Image(
    src="img/passo4.png",
    width=600,
    height=300,
    fit=ft.ImageFit.FILL
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

# Tela de tutorial
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
                    ),
                    ft.Container(height=20),  # Espaçamento entre os controles
                    ft.ListView(
                        controls=[
                            ft.Column(
                                controls=[
                                    ft.Column(
                                        [
                                            passo1_img,
                                            passo1_text,
                                        ],
                                        spacing=10,
                                    ),
                                    ft.Container(height=20),  # Espaçamento entre passos
                                    ft.Column(
                                        [
                                            passo2_img,
                                            passo2_text,
                                        ],
                                        spacing=10,
                                    ),
                                    ft.Container(height=20),
                                    ft.Column(
                                        [
                                            passo3_img,
                                            passo3_text,
                                        ],
                                        spacing=10,
                                    ),
                                    ft.Container(height=20),
                                    ft.Column(
                                        [
                                            passo4_img,
                                            passo4_text,
                                        ],
                                        spacing=10,
                                    ),
                                ],
                                spacing=20,
                                expand=True,
                                alignment=ft.MainAxisAlignment.START,
                                horizontal_alignment=ft.CrossAxisAlignment.START,
                            )
                        ],
                        expand=True,
                        spacing=10,
                        padding=10,
                        auto_scroll=True, 
                    ),
                ],
                spacing=10,
                scroll=ft.ScrollMode.AUTO,
                alignment=ft.MainAxisAlignment.START,
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
