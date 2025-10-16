from PIL import Image, ImageDraw, ImageFont
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.utils import ImageReader

# Definir o caminho da pasta de recursos
resource_path = os.path.join(os.path.dirname(__file__), "Resources")

# Função auxiliar para verificar a existência dos arquivos
def check_file_exists(file_path):
    if not os.path.isfile(file_path):
        print(f"Arquivo não encontrado: {file_path}")
        return False
    return True

def gerar_certificado(aluno, ls_stars, rw_stars, sp_stars, output_dir, mock_level):
    # Carregar a imagem de fundo
    if mock_level == "Starters":
        certificate_front = "Starters_P1.PNG"
        certificate_back = "Starters_P2.PNG"

    if mock_level == "Movers":
        certificate_front = "Movers_P1.PNG"
        certificate_back = "Movers_P2.PNG"

    if mock_level == "Flyers":
        certificate_front = "Flyers_P1.PNG"
        certificate_back = "Flyers_P2.PNG"
    imagem_fundo_path = os.path.join(resource_path, certificate_front)
    if not check_file_exists(imagem_fundo_path):
        return None
    imagem_fundo = Image.open(imagem_fundo_path)

    # Preparar para desenhar na imagem
    draw = ImageDraw.Draw(imagem_fundo)

    # Definir a fonte e tamanho do texto
    fonte_path = os.path.join(resource_path, "dTBommerSans_Rg.otf")
    if not check_file_exists(fonte_path):
        return None
    fonte = ImageFont.truetype(fonte_path, 24)

    # Adicionar texto
    texto = aluno
    cor_do_texto = "black" # A cor pode ser uma tupla (R, G, B) também
    posicao_do_texto = (97, 275) # Ajuste conforme necessário
    draw.text(posicao_do_texto, texto, fill=cor_do_texto, font=fonte)

    # Adicionar uma imagem (exemplo: uma estrela)
    estrela_cheia_path = os.path.join(resource_path, "Estrela.png")
    estrela_vazia_path = os.path.join(resource_path, "estrela_eb.png")
    
    if not check_file_exists(estrela_cheia_path) or not check_file_exists(estrela_vazia_path):
        return None
        
    estrela_cheia = Image.open(estrela_cheia_path)
    estrela_vazia = Image.open(estrela_vazia_path)
    
    # Posições x e y iniciais
    pos_x = 121
    ls_stars_y = 374
    rw_stars_y = 491
    sp_stars_y = 610
    espaco = 85 # É o espaço em pixels de uma estrela para outra
    for i in range(5):
        ls_estrela = estrela_vazia if ls_stars - i <= 0 else estrela_cheia
        rw_estrela = estrela_vazia if rw_stars - i <= 0 else estrela_cheia
        sp_estrela = estrela_vazia if sp_stars - i <= 0 else estrela_cheia
        pos_estrela = pos_x + 85 * i
        imagem_fundo.paste(ls_estrela, (pos_estrela, ls_stars_y), ls_estrela)
        imagem_fundo.paste(rw_estrela, (pos_estrela, rw_stars_y), rw_estrela)
        imagem_fundo.paste(sp_estrela, (pos_estrela, sp_stars_y), sp_estrela)

    # Salvar a imagem resultante como PNG temporário
    certificado_path = os.path.join(output_dir, f"certificado_temp_{aluno}.png")
    imagem_fundo.save(certificado_path)

    # Carregar a segunda página
    slide2_path = os.path.join(resource_path, certificate_back)
    if not check_file_exists(slide2_path):
        return None
    slide2 = Image.open(slide2_path)

    # Criar um PDF com as duas páginas em formato paisagem (A4)
    pdf_filename = os.path.join(output_dir, f"Relatório_{aluno}.pdf")
    c = canvas.Canvas(pdf_filename, pagesize=landscape(A4))

    # Adicionar a primeira página
    c.drawImage(ImageReader(certificado_path), 0, 0, width=landscape(A4)[0], height=landscape(A4)[1])

    # Adicionar a segunda página
    c.showPage()
    c.drawImage(ImageReader(slide2_path), 0, 0, width=landscape(A4)[0], height=landscape(A4)[1])

    # Salvar o PDF
    c.save()

    # Remover a imagem temporária
    os.remove(certificado_path)

    return pdf_filename
