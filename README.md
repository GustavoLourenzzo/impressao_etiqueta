# impressao_etiqueta

Esse projeto tem como objetivo imprimir etiquetas em impressoras tipo zebra.

Funcionamento:
No arquivo de configuração[localizado dentro da pasta confg, se não houver ela e criada na primeira execução]
 encontra-se as 2 informações pertinentes, a impressora padrão, e o local onde se encontrar as etiquetas,
todas em formatos validos de imagem [png, jpg ...], ao finalizar a impressão a etiqueta e movida para uma 
pasta na raiz do programa.


Requerimentos:
Ambiemte Python 3.6+ com as seguintes bibliotecas:

wx, win32ui, win32print, win32con, os, sys, configparser, shutil, win32gui, PIL   