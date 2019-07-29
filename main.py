import wx, win32ui, win32print, win32con, os, sys, configparser, shutil, win32gui
from PIL import Image, ImageWin

The_program_to_hide = win32gui.GetForegroundWindow()
win32gui.ShowWindow(The_program_to_hide , win32con.SW_HIDE)

#globais
listaImpressoras = None
dirBase = os.getcwd()
config = configparser.ConfigParser()

#função de criar diretorios apartir da raiz
def createDir(text, local = True):
    try:
        if local:
            os.makedirs(dirBase+"/"+text)
        else:
            os.makedirs(text)
        return True
    except:
        pass
    return False
#fim **

# Impressora
class impressao():
    def __init__(self):
        #constantes
        # HORZRES / VERTRES = printable area
        self.HORZRES = 8
        self.VERTRES = 10
        #
        # LOGPIXELS = dots per inch
        #
        self.LOGPIXELSX = 88
        self.LOGPIXELSY = 90
        #
        # PHYSICALWIDTH/HEIGHT = total area
        #
        self.PHYSICALWIDTH = 110
        self.PHYSICALHEIGHT = 111
        #
        # PHYSICALOFFSETX/Y = left / top margin
        #
        self.PHYSICALOFFSETX = 112
        self.PHYSICALOFFSETY = 113
        pass

    def geraListaImpressoras(self):
        global listaImpressoras
        lista = win32print.EnumPrinters ( win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        listaImpressoras = [name for (flags, description, name, comment) in lista]
        listaImpressoras = sorted(listaImpressoras)

    def listarImpressoras(self):
        global listaImpressoras
        if listaImpressoras is None:
            self.geraListaImpressoras()
        return listaImpressoras
    
    def getDefaultImpressora(self):
        return  win32print.GetDefaultPrinter()

    def imprimirEtiqueta(self,impressora_name, caminho, etiqueta, move = False, destino= ''):
        hDC = win32ui.CreateDC ()
        hDC.CreatePrinterDC (impressora_name)
        printable_area = hDC.GetDeviceCaps (self.HORZRES), hDC.GetDeviceCaps (self.VERTRES)
        printer_size = hDC.GetDeviceCaps (self.PHYSICALWIDTH), hDC.GetDeviceCaps (self.PHYSICALHEIGHT)
        printer_margins = hDC.GetDeviceCaps (self.PHYSICALOFFSETX), hDC.GetDeviceCaps (self.PHYSICALOFFSETY)

        bmp = Image.open(caminho+"/"+etiqueta).convert("1")
        bmp = bmp.rotate (180)
        hDC.StartDoc (etiqueta)
        hDC.StartPage ()
        dib = ImageWin.Dib(bmp)
        dib.draw(hDC.GetHandleOutput (), ((0+70), (0), (639), (300-50))) #ajustar a area imprimivel do papel
        hDC.EndPage()
        hDC.EndDoc()
        hDC.DeleteDC()
        if(move):
            self.movEtiqueta(caminho,etiqueta,destino)
        pass
    
    def movEtiqueta(self, caminho,etiqueta, destino):
        try:
            createDir(destino, False)
            shutil.move(caminho+"/"+etiqueta,destino)
            return True
        except:
            if os.path.isfile(destino+"/"+etiqueta):
                os.remove(destino+"/"+etiqueta)
                try:
                    shutil.move(caminho+"/"+etiqueta,destino)
                    return True
                except:
                    return False
            return False
        pass

 
#interface 
class componentePrincipal(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        self.parente = parent       
        #pastas e diretorios
        createDir("confg")
        #configurações iniciais
        self.compImpr = impressao()
        self.args = self.leConfig(config) 
        if self.args is None:
            self.args = self.leConfig(config)
        #opções da impressora
        wx.StaticText(self, label="Selecione a impressora: ", pos=(20, 20))
        self.impressoras = wx.ComboBox(self, pos=(190, 20), size=(300, -1),value=(self.args['impressora']) , choices=self.compImpr.listarImpressoras(), style=wx.CB_DROPDOWN|wx.CB_READONLY)
        self.Bind(wx.EVT_COMBOBOX, self.selecaoImpressora, self.impressoras)
        
        #opções da pasta
        wx.StaticText(self, label="Selecione a pasta de origem: ", pos=(20, 60))
        self.end_local = wx.TextCtrl(self, pos=(190,60), size=(300,-1), style=wx.TE_READONLY, value=self.args['local'])
        self.button_local =wx.Button(self, label="Origem", pos=(495, 59))
        self.Bind(wx.EVT_BUTTON, self.OnButton_local,self.button_local)
        
        #lista de itens para impressao
        wx.StaticLine(self, pos=(15,95), style=wx.LI_HORIZONTAL, size=(570,-1)) #ajustar apos o size
        wx.StaticText(self, label="Selecione a etiqueta para impressao: ", pos=(20, 100))
        self.etiquetaAtiva = ''
        self.lista_etiq = self.gerarListaEtiquetas()
        self.impress_list = wx.ListBox ( self,  pos = (20,120), size=(470,108), choices=self.lista_etiq, style=wx.LB_SORT|wx.LB_NEEDED_SB|wx.LB_SINGLE)# ajustar
        self.Bind(wx.EVT_LISTBOX, self.defineEtiquetaAlvo, self.impress_list)
        self.Bind(wx.EVT_LISTBOX_DCLICK, self.defineEtiqueta_imprime, self.impress_list) # imprime tb
        
        #botoes associados a eventos da lista
        self.button_imprimir = wx.Button(self, label="Imprimir", pos=(495, 120), size=(80,-1))
        self.Bind(wx.EVT_BUTTON, self.OnButton_imprimir,self.button_imprimir)
        
        self.button_attList = wx.Button(self, label="Atualizar Lista", pos=(495, 160), size=(80,-1))
        self.Bind(wx.EVT_BUTTON, self.OnButton_attList,self.button_attList)
        
        self.button_fechar = wx.Button(self, label="Fechar", pos=(495, 200), size=(80,-1))
        self.Bind(wx.EVT_BUTTON, self.OnButton_fechar,self.button_fechar)

        
       
    
    #evento acionado quando o botao imprimir e acionado
    def OnButton_imprimir(self, event):
        if self.impressoras.GetSelection() != wx.NOT_FOUND:
            if self.etiquetaAtiva != '':
                self.imprimeEtiqueta()
                msg = wx.MessageBox("A etiqueta foi imprimida corretamente ?", "Confirmar Impressão", wx.YES_NO , self)
                if msg == wx.YES:
                    if self.compImpr.movEtiqueta(self.end_local.GetLineText(0),self.etiquetaAtiva,self.end_local.GetLineText(0)+"/impresso"):
                        self.attListEvent()
                        self.etiquetaAtiva = ''
                    else:
                        self.mensagemUser("Não foi possivel mover a etiqueta para a pasta de destino.","Erro ao mover etiqueta", "erro")
                    
            else:
                self.mensagemUser("Nenhuma etiqueta foi selecionada", "Seleção vazia")
        else:
            self.mensagemUser("Nenhuma impressora foi selecionada", "Seleção vazia")
        pass


    #função que realiza a impressao da etiqueta
    def imprimeEtiqueta(self):
        self.compImpr.imprimirEtiqueta(self.impressoras.GetString(self.impressoras.GetSelection()),self.end_local.GetLineText(0),self.etiquetaAtiva)
        pass
    
    #mensagens de erro ou informação ao usuario
    def mensagemUser(self, msg, rotulo, tipo = 'info'):
        if tipo == 'info':
            wx.MessageBox(msg, rotulo, wx.OK | wx.ICON_INFORMATION)
        elif tipo == 'erro':
            wx.MessageBox(msg, rotulo, wx.OK | wx.ICON_ERROR)
        pass

    #evento acionado quando o botao fechar e pressionado
    def OnButton_fechar(self, event):
        wx.CallAfter(self.Destroy)
        self.Close()
        self.parente.Close()

    #evento acionado quando o botao attList e pressionado
    def OnButton_attList(self, event):
        self.attListEvent()
        pass

    def attListEvent(self):
        self.impress_list.Clear()
        self.lista_etiq = self.gerarListaEtiquetas()
        for etiqueta in self.lista_etiq:
            self.impress_list.Append(etiqueta)
        pass

    #marca a etiqueta que sera impressa
    def defineEtiquetaAlvo(self, event):
        self.etiquetaAtiva = event.GetString()
        #print(self.etiquetaAtiva)
        pass

    #marca a etiqueta que sera impressa e a imprime
    def defineEtiqueta_imprime(self, event):
         self.etiquetaAtiva = event.GetString()
         if self.impressoras.GetSelection() != wx.NOT_FOUND:
            self.imprimeEtiqueta()
            msg = wx.MessageBox("A etiqueta foi imprimida corretamente ?", "Confirmar Impressão", wx.YES_NO , self)
            if msg == wx.YES:
                if self.compImpr.movEtiqueta(self.end_local.GetLineText(0),self.etiquetaAtiva,self.end_local.GetLineText(0)+"/impressas"):
                    self.attListEvent()
                    self.etiquetaAtiva = ''
                else:
                    self.mensagemUser("Não foi possivel mover a etiqueta para a pasta de destino.","Erro ao mover etiqueta", "erro")   
         else:
             self.mensagemUser("Nenhuma impressora foi selecionada", "Seleção vazia")
         pass

        

    #gera a lista atualizada de etiquetas
    def gerarListaEtiquetas(self):
        listaEt = []
        lista_extencoes = ['.png','.jpg','.tif', '.bmp', ".jpeg", ".gif"]
        dir = os.listdir(str(self.end_local.GetValue()))
        for d in dir:
            if any(d.endswith(e) for e in lista_extencoes):
                listaEt.append(d)
            pass
        return listaEt

    #evento associado ao clique do botao button_local
    def OnButton_local(self, event):
        _local = wx.DirDialog(self,pos=(160,40),size=(300,-1), defaultPath=self.args['local'],style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
        if _local.ShowModal() == wx.ID_OK:
           self.end_local.SetValue(_local.GetPath())
        else:
           pass
        self.attListEvent()
        self.salvaConfig('origem', str(self.end_local.GetValue()))
        _local.Destroy()

    # le as configurações do arquivo ou cria
    def leConfig(self,config):
        try:
            config.read(dirBase+"/confg/confg.ini")
            local = (config['origem']['value'] if  config['origem']['value'] != '$' else dirBase )
            impr = (config['impressora']['value'] if config['impressora']['value'] != '$' else self.compImpr.getDefaultImpressora())
            ret = {
                 'impressora': impr,
                 'local': local
            }
            return ret
        except:
            self.criaCampos(config)


    #cria os campos do arquivo de config
    def criaCampos(self, config):
        config.add_section('impressora')
        config.set('impressora', 'value', '$')
        config.add_section('origem')
        config.set('origem', 'value', '$')
        
        with open(dirBase+"/confg/confg.ini", 'w') as configfile:
            config.write(configfile)

    #salva as configurações no arquivo .ini
    def salvaConfig(self, tropico, valor):
        config = configparser.ConfigParser()
        config.read(dirBase+"/confg/confg.ini")
        config.set(tropico, "value", valor)
        file = open(dirBase+"/confg/confg.ini", "w")
        config.write(file)
        pass

    #evento disparado toda vez que o valor padrao da impressora muda.
    def selecaoImpressora(self,event):
        self.salvaConfig('impressora',event.GetString())


#metodo para remover ambiguidade na execução de processos no windows
if __name__ == '__main__':
    app = wx.App(False)
    frame = wx.Frame(None,size=(615,280),title="Impressão de Etiqueta",style=wx.DEFAULT_FRAME_STYLE ^ wx.RESIZE_BORDER)
    panel = componentePrincipal(frame)
    frame.Show()
    app.MainLoop()
