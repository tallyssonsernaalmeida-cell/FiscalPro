"""
Gera icone_3c.ico com a logo ATHOS (triangulo azul + cinza)
Rode: python gerar_icone.py
Requer: pip install pillow
"""
from PIL import Image, ImageDraw
import os

ATHOS_AZUL  = (27, 43, 138)    # #1B2B8A
ATHOS_CINZA = (160, 168, 184)  # #A0A8B8
BG          = (13, 15, 26)     # #0d0f1a

def desenhar_logo(draw, size):
    s  = size
    cx = s // 2

    # Triângulo azul grande
    tri_azul = [
        (cx - 2,             int(s * 0.08)),
        (int(s * 0.05),      int(s * 0.88)),
        (cx + 6,             int(s * 0.88)),
    ]
    draw.polygon(tri_azul, fill=ATHOS_AZUL)

    # Triângulo cinza pequeno
    tri_cinza = [
        (cx + 14,            int(s * 0.48)),
        (int(s * 0.92),      int(s * 0.88)),
        (cx + 6,             int(s * 0.88)),
    ]
    draw.polygon(tri_cinza, fill=ATHOS_CINZA)

def gerar_icone():
    tamanhos = [16, 32, 48, 64, 128, 256]
    imagens  = []

    for size in tamanhos:
        img  = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Fundo arredondado azul escuro
        r = max(2, size // 8)
        draw.rounded_rectangle([0, 0, size-1, size-1], radius=r, fill=BG)

        desenhar_logo(draw, size)
        imagens.append(img)

    caminho = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icone_3c.ico")
    imagens[0].save(
        caminho, format="ICO",
        sizes=[(s, s) for s in tamanhos],
        append_images=imagens[1:]
    )
    print(f"Ícone ATHOS gerado: {caminho}")

if __name__ == "__main__":
    gerar_icone()
