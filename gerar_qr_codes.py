
import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

# Função para adicionar texto abaixo do QR code
def add_text_below_image(image, text, font_path='arial.ttf', font_size=20):
    width, height = image.size
    new_height = height + font_size + 10  # 10 pixels de margem
    new_image = Image.new('RGB', (width, new_height), 'white')
    new_image.paste(image, (0, 0))
    draw = ImageDraw.Draw(new_image)
    font = ImageFont.truetype(font_path, font_size)
    text_width, text_height = draw.textbbox((0, 0), text, font=font)[2:4]
    text_x = (width - text_width) / 2  # Centralizar o texto
    text_y = height - 15  # 5 pixels abaixo do QR code
    draw.text((text_x, text_y), text, fill='black', font=font)
    return new_image

# Passo 1: Leitura dos dados do Excel
file_path = 'dados.xlsx'  # Caminho para o arquivo Excel
df = pd.read_excel(file_path, header=0)

# Configurações da página A4
page_width, page_height = A4
margin = 10
qr_code_size = 200
space_between_qr = 20
columns = 2
rows = (page_height - 2 * margin) // (qr_code_size + space_between_qr)

# Lista para armazenar as imagens dos QR codes
qr_code_images = []

# Passo 2: Iteração célula por célula e geração de QR codes
for col_index in range(0, len(df.columns), 2):  # Itera pelas colunas ímpares (nomes)
    name_col = df.columns[col_index]
    qr_col = df.columns[col_index + 1]
    for name, value in zip(df[name_col], df[qr_col]):
        if pd.notna(value):  # Verifica se a célula do QR code não é NaN
            value_str = str(value)
            name_str = str(name)
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=4,
            )
            qr.add_data(value_str)
            qr.make(fit=True)

            # Cria uma imagem do QR code
            img = qr.make_image(fill_color="black", back_color="white")
            img = img.convert("RGB")

            # Adiciona o texto abaixo do QR code
            img_with_text = add_text_below_image(img, name_str)
            img_with_text = img_with_text.resize((qr_code_size, qr_code_size + 30))  # Redimensionar
            qr_code_images.append(img_with_text)
            print(value_str, name_str)

# Passo 3: Criação do PDF
pdf_file_path = 'qr_codes.pdf'
c = canvas.Canvas(pdf_file_path, pagesize=A4)

x, y = margin, page_height - margin - qr_code_size - 30  # Inicializar a posição de desenho

for i, qr_img in enumerate(qr_code_images):
    if i > 0 and i % columns == 0:
        y -= qr_code_size + 30 + space_between_qr
        x = margin
    if y < margin:
        c.showPage()
        y = page_height - margin - qr_code_size - 30
    c.drawInlineImage(qr_img, x, y, width=qr_code_size, height=qr_code_size + 30)
    x += qr_code_size + space_between_qr

c.save()

print("QR codes gerados e salvos no arquivo PDF com sucesso!")
