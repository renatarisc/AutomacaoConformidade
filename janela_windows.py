"""Geometria de janela pywebview ajustada à área de trabalho do Windows.

Uma janela mais alta que a tela (descontada a barra de tarefas) joga o rodapé pra fora da área
visível - o botão principal (ex: Salvar) some atrás da barra de tarefas e só reaparece arrastando
a janela pra cima. `geometria_para_tela` corta a altura pra caber, centraliza na horizontal e
encosta a janela mais pro topo na vertical (deixa a folga embaixo). Só Windows; em qualquer outro
lugar (ou se a consulta falhar) devolve x/y = None e o pywebview centraliza a janela sozinho.
"""

import ctypes

_SPI_GETWORKAREA = 0x0030

def geometria_para_tela(largura=980, altura=760, folga_borda=64):
    # devolve (x, y, largura, altura) pra passar direto pro webview.create_window - x/y None quando
    # não deu pra medir a tela (pywebview então centraliza). folga_borda desconta a barra de título
    # + bordas da janela, que o pywebview soma por fora da altura do conteúdo.
    x = y = None
    try:
        import ctypes.wintypes
        rect = ctypes.wintypes.RECT()
        ctypes.windll.user32.SystemParametersInfoW(_SPI_GETWORKAREA, 0, ctypes.byref(rect), 0)
        area_largura = rect.right - rect.left
        area_altura = rect.bottom - rect.top
        if area_altura > 200:
            altura = min(altura, area_altura - folga_borda)
        if area_largura > 200:
            largura = min(largura, area_largura - 40)
        x = rect.left + max((area_largura - largura) // 2, 0)
        y = rect.top + min(max(area_altura - altura, 0) // 4, 40)  # encosta pro topo, folga embaixo
    except Exception:
        pass
    return x, y, largura, altura
