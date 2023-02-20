import pyodbc, time, mysql.connector
from tkinter import *
import customtkinter, requests, base64
from tkinter import ttk, messagebox
from tkcalendar import *
from tkcalendar import Calendar
from PIL import ImageTk, Image
from datetime import datetime, date
from dateutil.relativedelta import relativedelta

#///////////////////////// FONTES UTILIZADAS
fonte_padrao = ("Calibri",12)
fonte_padrao_bold = ("Calibri Bold",12)
fonte_padrao_titulo = ("Calibri Bold",20)
fonte_padrao_titulo_janelas = ("Calibri Bold",16)

#/////VARIAVEIS GLOBAIS
estilo_entry_padrao = {'justify':'center','font':fonte_padrao, 'fg_color':'#ffffff', 'bg_color':'#ffffff', 'text_color':'#2a2d2e', 'border_color':'#0275d8'}
data = time.strftime('%Y%m%d', time.localtime())
hora = time.strftime('%H:%M:%S', time.localtime())
hora_calculo_pedido = time.strftime('%d/%m/%Y %H:%M:%S', time.localtime())
usuario_logado = ''
titulos = 'Programação de Carregamento - Versão: Omega 1'
primeiroDiaMes = data[:4]+data[4:6]+'01'
data_ultimo_dia_mes = datetime(int(data[:4]),int(data[4:6]),int(data[6:8]))
ultimoDiaMes = data_ultimo_dia_mes + relativedelta(day=31)
ultimoDiaMes = str(data[:4])+str(data[4:6])+str(ultimoDiaMes.day)
filtro_ativo = 0
dt_ini_filtro = ''
dt_fim_filtro = ''

bt_icone = {'background':"#ffffff", 'borderwidth':0, 'highlightthickness':0, 'relief':RIDGE, 'activebackground':"#ffffff", 'activeforeground':"#7c7c7c", 'cursor':"hand2"}


#/////FUNÇÕES
def filtro_principal():
    obs = ''
    img0 = Image.open('img\\00.png')
    resize_img0 = img0.resize((15, 15))
    nova_img0 = ImageTk.PhotoImage(resize_img0)

    img1 = Image.open('img\\01.png')
    resize_img1 = img1.resize((15, 15))
    nova_img1 = ImageTk.PhotoImage(resize_img1)
    
    img2 = Image.open('img\\02.png')
    resize_img2 = img2.resize((15, 15))
    nova_img2 = ImageTk.PhotoImage(resize_img2)

    img3 = Image.open('img\\03.png')
    resize_img3 = img3.resize((15, 15))
    nova_img3 = ImageTk.PhotoImage(resize_img3)

    img4 = Image.open('img\\04.png')
    resize_img4 = img4.resize((15, 15))
    nova_img4 = ImageTk.PhotoImage(resize_img4)

    img5 = Image.open('img\\05.png')
    resize_img5 = img5.resize((15, 15))
    nova_img5 = ImageTk.PhotoImage(resize_img5)

    img6 = Image.open('img\\06.png')
    resize_img6 = img6.resize((15, 15))
    nova_img6 = ImageTk.PhotoImage(resize_img6)

    img7 = Image.open('img\\07.png')
    resize_img7 = img7.resize((15, 15))
    nova_img7 = ImageTk.PhotoImage(resize_img7)

    img8 = Image.open('img\\08.png')
    resize_img8 = img8.resize((15, 15))
    nova_img8 = ImageTk.PhotoImage(resize_img8)

    img9 = Image.open('img\\09.png')
    resize_img9 = img9.resize((15, 15))
    nova_img9 = ImageTk.PhotoImage(resize_img9)

    img10 = Image.open('img\\10.png')
    resize_img10 = img10.resize((15, 15))
    nova_img10 = ImageTk.PhotoImage(resize_img10)

    img11 = Image.open('img\\11.png')
    resize_img11 = img11.resize((15, 15))
    nova_img11 = ImageTk.PhotoImage(resize_img11)

    img12 = Image.open('img\\12.png')
    resize_img12 = img12.resize((15, 15))
    nova_img12 = ImageTk.PhotoImage(resize_img12)

    img13 = Image.open('img\\13.png')
    resize_img13 = img13.resize((15, 15))
    nova_img13 = ImageTk.PhotoImage(resize_img13)

    db.cmd_reset_connection()
    cursor.execute("SELECT * FROM carregamento_temp1 where usuario = %s order by\
         (case\
            when status = 10 then 1\
            when status = 11 then 2\
            when status = 7 then 3\
            when status = 6 then 4\
            when status = 5 then 5\
            when status = 4 then 6\
            when status = 3 then 7\
            when info != '' then 8\
        else 100 end),\
        (case when\
        status = 10 or\
        status = 11 or\
        status = 7 or\
        status = 6 or\
        status = 5 or\
        status = 4 or\
        status = 3 or\
        info != ''\
        then dtcarreg else dtprog end)",(usuario_logado,))
    tabela_temp1 = cursor.fetchall()
    tree_principal.delete(*tree_principal.get_children())
    cont = 0
    for i in tabela_temp1:
        if i[2] == '0':
            status_ot = 'Sem OT'
            img_status = nova_img0
        elif i[2] == '1':
            status_ot = 'Aguard. Transp'
            img_status = nova_img1
        elif i[2] == '2':
            status_ot = 'Aguard. Caminhão'
            img_status = nova_img2
        elif i[2] == '3':
            status_ot = 'Lib. Embarque'
            img_status = nova_img3
        elif i[2] == '4':
            status_ot = 'Embarque Iniciado'
            img_status = nova_img4
        elif i[2] == '5':
            status_ot = 'Embarque Finalizado'
            img_status = nova_img5
        elif i[2] == '6':
            status_ot = 'Expedido'
            img_status = nova_img6
        elif i[2] == '7':
            status_ot = 'Faturado'
            img_status = nova_img7
        elif i[2] == '8':
            status_ot = 'Devolvido'
            img_status = nova_img8
        elif i[2] == '9':
            status_ot = 'Refaturado'
            img_status = nova_img9            
        elif i[2] == '10':
            status_ot = 'Refat. Expedido'
            img_status = nova_img10
        elif i[2] == '11':
            status_ot = 'Refat. Faturado'
            img_status = nova_img11            
        elif i[2] == '12':
            status_ot = 'Cancelado'
            img_status = nova_img12            
        elif i[2] == '13':
            status_ot = 'Bloqueado Crédito'
            img_status = nova_img13            

        if cont % 2 == 0:
            if i[42] == None or i[42] == '':
                info = ''
                tags_cor = 'par'
            else:
                if i[2] == '3':
                    info = i[42]
                    tags_cor = 'par'
                else:
                    info = i[42]
                    tags_cor = 'contem'                

            if i[43] == None or i[43] == ' ':
                obs = ''

            if i[42] == '' and i[43] != '':
                obs = i[43]
                tags_cor = 'contem_obs'


            dt_prog = i[4][6:8]+'/'+i[4][4:6]+'/'+i[4][:4]
            dt_carregamento = i[6][6:8]+'/'+i[6][4:6]+'/'+i[6][:4]
            dt_entrega = i[20][6:8]+'/'+i[20][4:6]+'/'+i[20][:4]
            tree_principal.insert('', 'end', text=' ', image=img_status,
                                    values=(
                                    info,i[1].lstrip('0'), i[22].strip(), i[3], i[19], dt_entrega, str(i[23]).replace('.0', ''), str(i[24]).replace('.0', ''), str(i[25]).replace('.0', ''), str(i[26]).replace('.0', ''), str(i[27]).replace('.0', ''), str(i[28]).replace('.0', ''), str(i[29]).replace('.0', ''), str(i[30]).replace('.0', ''), str(i[31]).replace('.0', ''), str(i[32]).replace('.0', ''), str(i[33]).replace('.0', ''), str(i[34]).replace('.0', ''), str(i[35]).replace('.0', ''), str(i[36]).replace('.0', ''), str(i[37]).replace('.0', ''), str(i[38]).replace('.0', ''), str(i[39]).replace('.0', ''), str(i[40]).replace('.0', ''), dt_prog, i[5], dt_carregamento, i[7].strip(), i[8], i[21], obs),
                                    tags=(tags_cor,))
        else:
            if i[42] == None or i[42] == '':
                info = ''
                tags_cor = 'impar'
            else:
                if i[2] == '3':
                    info = i[42]
                    tags_cor = 'impar'
                else:
                    info = i[42]
                    tags_cor = 'contem'                


            if i[43] == None or i[43] == ' ':
                obs = ''

            if i[42] == '' and i[43] != '':
                obs = i[43]
                tags_cor = 'contem_obs'
          
            tree_principal.insert('', 'end', text=' ', image=img_status,
                                    values=(
                                    info,i[1].lstrip('0'), i[22].strip(), i[3], i[19], dt_entrega, str(i[23]).replace('.0', ''), str(i[24]).replace('.0', ''), str(i[25]).replace('.0', ''), str(i[26]).replace('.0', ''), str(i[27]).replace('.0', ''), str(i[28]).replace('.0', ''), str(i[29]).replace('.0', ''), str(i[30]).replace('.0', ''), str(i[31]).replace('.0', ''), str(i[32]).replace('.0', ''), str(i[33]).replace('.0', ''), str(i[34]).replace('.0', ''), str(i[35]).replace('.0', ''), str(i[36]).replace('.0', ''), str(i[37]).replace('.0', ''), str(i[38]).replace('.0', ''), str(i[39]).replace('.0', ''), str(i[40]).replace('.0', ''), dt_prog, i[5], dt_carregamento, i[7].strip(), i[8], i[21], obs),
                                    tags=(tags_cor,))
        cont += 1
    tree_principal.yview_moveto(1)        
    root.mainloop()

def atualizar_lista():
    versao()
    #// QUERY DIRETA DO PROTHEUS - TROCAR A LOGICA AQUI POR UMA API 
    cursor2.execute("SELECT \
            SZ2.Z2_NUMERO, SZ2.Z2_STATUS, SZ2.Z2_CLIENTE, (CASE WHEN SZ2.Z2_STATUS = '0' THEN SZ2.Z2_EMISSAO ELSE SZ2.Z2_DTCARR END) Z2_DTCARR, \
			(CASE WHEN SZ2.Z2_STATUS = '0' THEN '  :  ' ELSE SZ2.Z2_HORA END) Z2_HRCARR, SZ2.Z2_INIEMB, SZ2.Z2_MUNENT, SZ2.Z2_UFENT, \
			SZ2.Z2_DTLBEMB, SZ2.Z2_HRLBEMB, SZ2.Z2_INIEMB, SZ2.Z2_HINIEMB, SZ2.Z2_FIMEMB, SZ2.Z2_HFIMEMB, SZ2.Z2_DTEXP, SZ2.Z2_HREXP, \
			SZ2.Z2_DTNF, SZ2.Z2_NOTA, SZ3.Z3_PEDIDO, SC6.C6_ENTREG, ISNULL(SA4.A4_NREDUZ,'') A4_NREDUZ, SA1.A1_NREDUZ, \
			SUM(CASE WHEN SB1.B1__TPPA = 'F' AND SB1.B1__BITOLA = 5.5  THEN SZ3.Z3_QUANT ELSE 0 END) QTDFM055, \
			SUM(CASE WHEN SB1.B1__TPPA = 'F' AND SB1.B1__BITOLA = 6.5  THEN SZ3.Z3_QUANT ELSE 0 END) QTDFM065, \
			SUM(CASE WHEN SB1.B1__TPPA = 'F' AND SB1.B1__BITOLA = 7    THEN SZ3.Z3_QUANT ELSE 0 END) QTDFM070, \
			SUM(CASE WHEN SB1.B1__TPPA = 'R' AND SB1.B1__BITOLA = 6.3  THEN SZ3.Z3_QUANT ELSE 0 END) QTDRL063, \
			SUM(CASE WHEN SB1.B1__TPPA = 'R' AND SB1.B1__BITOLA = 8    THEN SZ3.Z3_QUANT ELSE 0 END) QTDRL080, \
			SUM(CASE WHEN SB1.B1__TPPA = 'R' AND SB1.B1__BITOLA = 10   THEN SZ3.Z3_QUANT ELSE 0 END) QTDRL100, \
			SUM(CASE WHEN SB1.B1__TPPA = 'R' AND SB1.B1__BITOLA = 12.5 THEN SZ3.Z3_QUANT ELSE 0 END) QTDRL125, \
			SUM(CASE WHEN SB1.B1__TPPA = 'R' AND SB1.B1__BITOLA = 16   THEN SZ3.Z3_QUANT ELSE 0 END) QTDRL160, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 6    THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR060, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 6.3  THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR063, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 8    THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR080, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 10   THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR100, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 12.5 THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR125, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 16   THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR160, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 20   THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR200, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 25   THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR250, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 32   THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR320, \
			SUM(SZ3.Z3_QUANT) QTDTOT \
            FROM SZ2010 SZ2 \
                INNER JOIN SA1010 SA1 \
                    ON SA1.A1_FILIAL  = '  ' \
                    AND SA1.A1_COD     = SZ2.Z2_CLIENTE \
                    AND SA1.A1_LOJA    = SZ2.Z2_LOJA \
                    AND SA1.D_E_L_E_T_ = ' ' \
                INNER JOIN SZ3010 SZ3 \
                    ON SZ3.Z3_FILIAL  = SZ2.Z2_FILIAL \
                    AND SZ3.Z3_NUMERO  = SZ2.Z2_NUMERO \
                    AND SZ3.D_E_L_E_T_ = ' ' \
                INNER JOIN SC6010 SC6 \
                    ON SC6.C6_FILIAL  = SZ3.Z3_FILIAL \
                    AND SC6.C6_NUM     = SZ3.Z3_PEDIDO \
                    AND SC6.C6_ITEM    = SZ3.Z3_ITEM \
                    AND SC6.D_E_L_E_T_ = ' ' \
                INNER JOIN SB1010 SB1 \
                    ON SB1.B1_FILIAL  = '01' \
                    AND SB1.B1_COD     = SZ3.Z3_PRODUTO \
                    AND SB1.D_E_L_E_T_ = ' ' \
                LEFT JOIN SA4010 SA4 \
                    ON SA4.A4_FILIAL  = '  ' \
                    AND SA4.A4_COD     = SZ2.Z2_TRANSP \
                    AND SA4.D_E_L_E_T_ = ' ' \
            WHERE   SZ2.Z2_FILIAL  = '01' \
                AND ((SZ2.Z2_STATUS IN ('1','2','3','4','5','6','7','8') \
                AND   SZ2.Z2_DTCARR BETWEEN ? AND ?) \
                OR  SZ2.Z2_STATUS  =  '0') \
                AND SZ2.D_E_L_E_T_ = ' ' \
            GROUP BY SZ2.Z2_NUMERO, SZ2.Z2_STATUS, SZ2.Z2_CLIENTE, (CASE WHEN SZ2.Z2_STATUS = '0' THEN SZ2.Z2_EMISSAO ELSE SZ2.Z2_DTCARR END), \
                        (CASE WHEN SZ2.Z2_STATUS = '0' THEN '  :  ' ELSE SZ2.Z2_HORA END), SZ2.Z2_INIEMB, SZ2.Z2_MUNENT, SZ2.Z2_UFENT, SZ2.Z2_DTLBEMB, \
                        SZ2.Z2_HRLBEMB, SZ2.Z2_INIEMB, SZ2.Z2_HINIEMB, SZ2.Z2_FIMEMB, SZ2.Z2_HFIMEMB, SZ2.Z2_DTEXP, SZ2.Z2_HREXP, SZ2.Z2_DTNF, \
                        SZ2.Z2_NOTA, SZ3.Z3_PEDIDO, SC6.C6_ENTREG, ISNULL(SA4.A4_NREDUZ,''), SA1.A1_NREDUZ \
            ORDER BY (CASE WHEN SZ2.Z2_STATUS = '0' THEN SZ2.Z2_EMISSAO ELSE SZ2.Z2_DTCARR END), SZ2.Z2_NUMERO", (primeiroDiaMes, ultimoDiaMes))

    #// APAGA A TABELA TEMP1 DO USUARIO LOGADO 
    cursor.execute("delete from carregamento_temp1 where usuario = %s",(usuario_logado,))
    #cursor.execute("ALTER TABLE carregamento_temp1 AUTO_INCREMENT = 0")
    db.commit()

    cursor.execute('select * from carregamento')
    tabela_fixa = cursor.fetchall()

    #// VERIFICA SE EXISTE REGISTRO(ALTERAÇÕES DE INFO E OBS ADICIONADOS PELO USUARIO) E ADICIONA NA TABELA TEMP1 CASO TENHA
    for i in cursor2:
        entrou = 0
        for t in tabela_fixa:
            if i[0].lstrip('0') == t[1] and str(i[39]) == t[40]: #// VERIFICA SE A OT É IGUAL NA TABELA CARREGAMENTO - CASO SEJA IGUAL, A LINHA SERA ADICIONADA NA TABELA TEMP
                if i[1] == t[2]: #// VERIFICA SE O STATUS DE AMBOS SÃO IGUAIS. CASO NÃO SEJA, A TABELA TEMP1 SERA PREENCHIDA COM O STATUS VINDO DO PROTHEUS
                    cursor.execute("INSERT INTO carregamento_temp1 (\
                        ot,\
                        status,\
                        ncliente,\
                        dtprog,\
                        hrprog,\
                        dtcarreg,\
                        cidade,\
                        estado,\
                        dtlibemb,\
                        hrlibemb,\
                        dtiniemb,\
                        hriniemb,\
                        dtfimemb,\
                        hrfimemb,\
                        dtexped,\
                        hrexped,\
                        dtfat,\
                        nf,\
                        pedido,\
                        dtentrega,\
                        transp,\
                        cliente,\
                        fm055,\
                        fm065,\
                        fm070,\
                        rl063,\
                        rl080,\
                        rl100,\
                        rl125,\
                        rl160,\
                        br060,\
                        br063,\
                        br080,\
                        br100,\
                        br125,\
                        br160,\
                        br200,\
                        br250,\
                        br320,\
                        totaltn,\
                        info,\
                        obs,\
                        usuario)\
                        values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (t[1],t[2],t[3],t[4],t[5],t[6],t[7],t[8],t[9],t[10],t[11],t[12],t[13],t[14],t[15],t[16],t[17],t[18],t[19],t[20],t[21],t[22],t[23],t[24],t[25],t[26],t[27],t[28],t[29],t[30],t[31],t[32],t[33],t[34],t[35],t[36],t[37],t[38],t[39],t[40],t[42],t[43],usuario_logado,))
                else:
                    cursor.execute("INSERT INTO carregamento_temp1 (\
                        ot,\
                        status,\
                        ncliente,\
                        dtprog,\
                        hrprog,\
                        dtcarreg,\
                        cidade,\
                        estado,\
                        dtlibemb,\
                        hrlibemb,\
                        dtiniemb,\
                        hriniemb,\
                        dtfimemb,\
                        hrfimemb,\
                        dtexped,\
                        hrexped,\
                        dtfat,\
                        nf,\
                        pedido,\
                        dtentrega,\
                        transp,\
                        cliente,\
                        fm055,\
                        fm065,\
                        fm070,\
                        rl063,\
                        rl080,\
                        rl100,\
                        rl125,\
                        rl160,\
                        br060,\
                        br063,\
                        br080,\
                        br100,\
                        br125,\
                        br160,\
                        br200,\
                        br250,\
                        br320,\
                        totaltn,\
                        info,\
                        obs,\
                        usuario)\
                        values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (t[1],i[1],t[3],t[4],t[5],t[6],t[7],t[8],t[9],t[10],t[11],t[12],t[13],t[14],t[15],t[16],t[17],t[18],t[19],t[20],t[21],t[22],t[23],t[24],t[25],t[26],t[27],t[28],t[29],t[30],t[31],t[32],t[33],t[34],t[35],t[36],t[37],t[38],t[39],t[40],t[42],t[43],usuario_logado,))
                entrou = 1
        if entrou == 0:
            cursor.execute("INSERT INTO carregamento_temp1 (\
                ot,\
                status,\
                ncliente,\
                dtprog,\
                hrprog,\
                dtcarreg,\
                cidade,\
                estado,\
                dtlibemb,\
                hrlibemb,\
                dtiniemb,\
                hriniemb,\
                dtfimemb,\
                hrfimemb,\
                dtexped,\
                hrexped,\
                dtfat,\
                nf,\
                pedido,\
                dtentrega,\
                transp,\
                cliente,\
                fm055,\
                fm065,\
                fm070,\
                rl063,\
                rl080,\
                rl100,\
                rl125,\
                rl160,\
                br060,\
                br063,\
                br080,\
                br100,\
                br125,\
                br160,\
                br200,\
                br250,\
                br320,\
                totaltn,\
                usuario)\
                values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (i[0].lstrip('0'),i[1],i[2],i[3],i[4],i[5],i[6],i[7],i[8],i[9],i[10],i[11],i[12],i[13],i[14],i[15],i[16],i[17],i[18],i[19],i[20],i[21],i[22],i[23],i[24],i[25],i[26],i[27],i[28],i[29],i[30],i[31],i[32],i[33],i[34],i[35],i[36],i[37],i[38],i[39],usuario_logado,))
    db.commit()
    #// CHAMA A FUNÇÃO FILTRO PRINCIPAL
    filtro_principal()
    root.mainloop()

def filtro_home():
    global filtro_ativo
    filtro_ativo = 2
    atualizar_lista_filtro(data, data)

def soma_toneladas():
    
    soma = 0
    teste = tree_principal.selection()
    for i in teste:
        valor_lista = tree_principal.item(i, "values")[23]
        soma = soma + int(valor_lista)
    lbl_peso.configure(text=str(soma) +' TN')

def detalhes(event):

    def setup():
        cursor.execute("SELECT * FROM carregamento where ot = %s",(ot,))
        consulta = cursor.fetchone()
        if consulta != None:
            opt_hora.set(consulta[42])    
            textbox.insert("0.0", consulta[43])            
        else:
            opt_hora.set('') 
            textbox.delete("0.0", END)            

    def salvar():
        info = opt_hora.get() + ' | ' + ent_entrada.get()
        obs = textbox.get("0.0", "end")
       
        cursor.execute("SELECT * FROM carregamento where ot = %s",(ot,))
        consulta = cursor.fetchone()
        if consulta == None:
            cursor.execute("INSERT INTO carregamento(\
            ot,\
            status,\
            ncliente,\
            dtprog,\
            hrprog,\
            dtcarreg,\
            cidade,\
            estado,\
            dtlibemb,\
            hrlibemb,\
            dtiniemb,\
            hriniemb,\
            dtfimemb,\
            hrfimemb,\
            dtexped,\
            hrexped,\
            dtfat,\
            nf,\
            pedido,\
            dtentrega,\
            transp,\
            cliente,\
            fm055,\
            fm065,\
            fm070,\
            rl063,\
            rl080,\
            rl100,\
            rl125,\
            rl160,\
            br060,\
            br063,\
            br080,\
            br100,\
            br125,\
            br160,\
            br200,\
            br250,\
            br320,\
            totaltn,\
            usuario,\
            info,\
            obs)\
            values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (resultado[1],resultado[2],resultado[3],resultado[4],resultado[5],resultado[6],resultado[7],resultado[8],resultado[9],resultado[10],resultado[11],resultado[12],resultado[13],resultado[14],resultado[15],resultado[16],resultado[17],resultado[18],resultado[19],resultado[20],resultado[21],resultado[22],resultado[23],resultado[24],resultado[25],resultado[26],resultado[27],resultado[28],resultado[29],resultado[30],resultado[31],resultado[32],resultado[33],resultado[34],resultado[35],resultado[36],resultado[37],resultado[38],resultado[39],resultado[40],usuario_logado,info,obs,))
        else:
            cursor.execute("UPDATE carregamento set info = %s, obs = %s where id = %s",(info, obs, consulta[0],))

        db.commit()

        root2.destroy()
        if filtro_ativo == 1:
            atualizar_lista_filtro(dt_ini_filtro, dt_fim_filtro)
        else:
            atualizar_lista()
    
    def excluir():
        cursor.execute("SELECT * FROM carregamento where ot = %s",(ot,))
        consulta = cursor.fetchone()
        if consulta != None:
            confirma = messagebox.askyesno('Programação de Carregamento', "Confirma a exclusão destas informações?", parent=root2)
            if confirma == True:
                cursor.execute("DELETE from carregamento where id = %s",(consulta[0],))
                db.commit()
                root2.destroy()
                if filtro_ativo == 1:
                    atualizar_lista_filtro(dt_ini_filtro, dt_fim_filtro)
                else:
                    atualizar_lista()
        else:
            pass
        

    ot_select = tree_principal.focus()
    ot = tree_principal.item(ot_select, "values")[1]
    cursor.execute("SELECT * FROM simec_carregamento.carregamento_temp1 where ot = %s",(ot,))
    resultado = cursor.fetchone()
    root2 = Toplevel(root)
    root2.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
    root2.unbind_class("Button", "<Key-space>")
    root2.focus_force()
    root2.grab_set()
    root2.configure(background='#ffffff')
    window_width = 500
    window_height = 500
    screen_width = root2.winfo_screenwidth()
    screen_height = root2.winfo_screenheight() - 70
    x_cordinate = int((screen_width / 2) - (window_width / 2))
    y_cordinate = int((screen_height / 2) - (window_height / 2))
    root2.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
    root2.resizable(0, 0)
    root2.configure(bg='#ffffff')
    root2.title(titulos)
    #root2.iconbitmap('img\\icone.ico')

    frame0 = customtkinter.CTkFrame(root2, bg_color='#ffffff',corner_radius=20, fg_color='#ffffff', border_width=2, border_color='#000033')
    frame0.pack(side=TOP, fill=BOTH, expand=TRUE, padx=10, pady=10)

    frame1 = Frame(frame0, bg='#ffffff')
    frame1.pack(side=TOP, fill=X, expand=False, padx=10, pady=10)
    frame2 = Frame(frame0, bg='#000033') #/////LINHA
    frame2.pack(side=TOP, fill=X, expand=False, padx=10, pady=2)
    frame3 = Frame(frame0, bg='#ffffff')
    frame3.pack(side=TOP, fill=X, expand=False, padx=10, pady=10)
    frame4 = Frame(frame0, bg='#000033') #/////LINHA
    frame4.pack(side=TOP, fill=X, expand=False, padx=10, pady=2)
    frame5 = Frame(frame0, bg='#ffffff')
    frame5.pack(side=TOP, fill=BOTH, expand=True, padx=10, pady=10)

    lbl_titulo = Label(frame1, text=f'Adicionar informações na OT: {ot}', font=fonte_padrao_titulo_janelas, bg='#ffffff', fg='#1d366c')
    lbl_titulo.grid(row=0, column=2)

    frame1.grid_columnconfigure(0, weight=1)
    frame1.grid_columnconfigure(4, weight=1)

    #//FRAME 2 - LINHA
    
    #//FRAME 3
    lbl_titulo = Label(frame3, text='Horário programado:', font=fonte_padrao_bold, bg='#ffffff', fg='#1d366c')
    lbl_titulo.grid(row=0, column=1, sticky=W)
    
    lista_hora= [
        '00:00:00',
        '00:30:00',
        '01:00:00',
        '01:30:00',
        '02:00:00',
        '02:30:00',
        '03:00:00',
        '03:30:00',
        '04:00:00',
        '04:30:00',
        '05:00:00',
        '05:30:00',
        '06:00:00',
        '06:30:00',
        '07:00:00',
        '07:30:00',
        '08:00:00',
        '08:30:00',
        '09:00:00',
        '09:30:00',
        '10:00:00',
        '10:30:00',
        '11:00:00',
        '11:30:00',
        '12:00:00',
        '12:30:00',
        '13:00:00',
        '13:30:00',
        '14:00:00',
        '14:30:00',
        '15:00:00',
        '15:30:00',
        '16:00:00',
        '16:30:00',
        '17:00:00',
        '17:30:00',
        '18:00:00',
        '18:30:00',
        '19:00:00',
        '19:30:00',
        '20:00:00',
        '20:30:00',
        '21:00:00',
        '21:30:00',
        '22:00:00',
        '22:30:00',
        '23:00:00',
        '23:30:00']
    
    opt_hora = ttk.Combobox(frame3, textvariable='clique_inicio', values=lista_hora, width=47, font=fonte_padrao, state='readonly')
    opt_hora.grid(row=1, column=1)

    lbl_titulo = Label(frame3, text='Horário de entrada:', font=fonte_padrao_bold, bg='#ffffff', fg='#1d366c')
    lbl_titulo.grid(row=2, column=1, sticky=W, pady=(10,0))
    
    ent_entrada = customtkinter.CTkEntry(frame3, width=400, height=25, border_width=1) 
    ent_entrada.grid(row=3, column=1)    

    lbl_titulo = Label(frame3, text='Observações:', font=fonte_padrao_bold, bg='#ffffff', fg='#1d366c')
    lbl_titulo.grid(row=4, column=1, sticky=W, pady=(10,0))

    textbox = customtkinter.CTkTextbox(frame3, width=400, height=200,
    font= fonte_padrao, wrap=WORD, border_width=1, border_color='#808080',
    fg_color="#ffffff",text_color='#333333')
    
    textbox.grid(row=5, column=1, sticky=W)

    frame3.grid_columnconfigure(0, weight=1)
    frame3.grid_columnconfigure(2, weight=1)    

    #//FRAME 4 - LINHA

    #//FRAME 5
    bt1 = customtkinter.CTkButton(frame5,font=fonte_padrao_bold, text='Salvar',text_color='#ffffff', hover_color='#3DC2FF', bg_color='#ffffff', fg_color='#0275d8', corner_radius=5, command=salvar)
    bt1.grid(row=0, column=1, padx=10)

    bt2 = customtkinter.CTkButton(frame5, text='Remover Info.',text_color='#ffffff', hover_color='#3DC2FF', font=fonte_padrao_bold, bg_color='#ffffff', fg_color='#EB445A', corner_radius=5, command=excluir)
    bt2.grid(row=0, column=2)

    frame5.grid_columnconfigure(0, weight=1)
    frame5.grid_columnconfigure(3, weight=1)        

  
    setup()
    root2.mainloop()

def calendario_inicio():
    root2 = Toplevel(root)
    window_width = 360
    window_height = 258
    screen_width = root2.winfo_screenwidth()
    screen_height = root2.winfo_screenheight() - 70
    x_cordinate = int((screen_width / 2) - (window_width / 2))
    y_cordinate = int((screen_height / 2) - (window_height / 2))
    root2.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
    root2.resizable(0, 0)
    root2.configure(bg='#ffffff')
    root2.title(titulos)
    root2.overrideredirect(True)
    root2.focus_force()
    root2.grab_set()

    def escolher_data_bind(event):
        escolher_data()
    def escolher_data():
        ent_data_inicial.configure(state='normal')
        ent_data_inicial.delete(0, END)
        ent_data_inicial.insert(0, cal.get_date())
        ent_data_inicial.configure(state='readonly')
        root2.destroy()
        root.focus_force()
        root.grab_set()
    
    hoje = date.today()
    cal = Calendar(root2, font=fonte_padrao, selectmode='day', locale='pt_BR',
                cursor="hand1", background='#0275d8', foreground='#ffffff', bordercolor='#2a2d2e', headersbackground='#ffffff', headersforeground='#1c1c1c' )
    
    cal.pack(fill="both", expand=True)

    root2.bind('<Double-1>',escolher_data_bind) # Escolhe a data ao clicar 2x com o mouse
    root2,mainloop()

def calendario_final():
    root2 = Toplevel(root)
    window_width = 360
    window_height = 258
    screen_width = root2.winfo_screenwidth()
    screen_height = root2.winfo_screenheight() - 70
    x_cordinate = int((screen_width / 2) - (window_width / 2))
    y_cordinate = int((screen_height / 2) - (window_height / 2))
    root2.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
    root2.resizable(0, 0)
    root2.configure(bg='#ffffff')
    root2.title(titulos)
    root2.overrideredirect(True)
    root2.focus_force()
    root2.grab_set()

    def escolher_data_bind(event):
        escolher_data()
    def escolher_data():
        ent_data_final.configure(state='normal')
        ent_data_final.delete(0, END)
        ent_data_final.insert(0, cal.get_date())
        ent_data_final.configure(state='readonly')
        root2.destroy()
        root.focus_force()
        root.grab_set()
    
    hoje = date.today()
    cal = Calendar(root2, font=fonte_padrao, selectmode='day', locale='pt_BR',
                cursor="hand1", background='#0275d8', foreground='#ffffff', bordercolor='#2a2d2e', headersbackground='#ffffff', headersforeground='#1c1c1c' )
    
    cal.pack(fill="both", expand=True)

    root2.bind('<Double-1>',escolher_data_bind) # Escolhe a data ao clicar 2x com o mouse
    root2,mainloop()

def atualizar_lista_filtro(data_inicial, data_final):
    #// QUERY DIRETA DO PROTHEUS - TROCAR A LOGICA AQUI POR UMA API 
    cursor2.execute("SELECT \
            SZ2.Z2_NUMERO, SZ2.Z2_STATUS, SZ2.Z2_CLIENTE, (CASE WHEN SZ2.Z2_STATUS = '0' THEN SZ2.Z2_EMISSAO ELSE SZ2.Z2_DTCARR END) Z2_DTCARR, \
			(CASE WHEN SZ2.Z2_STATUS = '0' THEN '  :  ' ELSE SZ2.Z2_HORA END) Z2_HRCARR, SZ2.Z2_INIEMB, SZ2.Z2_MUNENT, SZ2.Z2_UFENT, \
			SZ2.Z2_DTLBEMB, SZ2.Z2_HRLBEMB, SZ2.Z2_INIEMB, SZ2.Z2_HINIEMB, SZ2.Z2_FIMEMB, SZ2.Z2_HFIMEMB, SZ2.Z2_DTEXP, SZ2.Z2_HREXP, \
			SZ2.Z2_DTNF, SZ2.Z2_NOTA, SZ3.Z3_PEDIDO, SC6.C6_ENTREG, ISNULL(SA4.A4_NREDUZ,'') A4_NREDUZ, SA1.A1_NREDUZ, \
			SUM(CASE WHEN SB1.B1__TPPA = 'F' AND SB1.B1__BITOLA = 5.5  THEN SZ3.Z3_QUANT ELSE 0 END) QTDFM055, \
			SUM(CASE WHEN SB1.B1__TPPA = 'F' AND SB1.B1__BITOLA = 6.5  THEN SZ3.Z3_QUANT ELSE 0 END) QTDFM065, \
			SUM(CASE WHEN SB1.B1__TPPA = 'F' AND SB1.B1__BITOLA = 7    THEN SZ3.Z3_QUANT ELSE 0 END) QTDFM070, \
			SUM(CASE WHEN SB1.B1__TPPA = 'R' AND SB1.B1__BITOLA = 6.3  THEN SZ3.Z3_QUANT ELSE 0 END) QTDRL063, \
			SUM(CASE WHEN SB1.B1__TPPA = 'R' AND SB1.B1__BITOLA = 8    THEN SZ3.Z3_QUANT ELSE 0 END) QTDRL080, \
			SUM(CASE WHEN SB1.B1__TPPA = 'R' AND SB1.B1__BITOLA = 10   THEN SZ3.Z3_QUANT ELSE 0 END) QTDRL100, \
			SUM(CASE WHEN SB1.B1__TPPA = 'R' AND SB1.B1__BITOLA = 12.5 THEN SZ3.Z3_QUANT ELSE 0 END) QTDRL125, \
			SUM(CASE WHEN SB1.B1__TPPA = 'R' AND SB1.B1__BITOLA = 16   THEN SZ3.Z3_QUANT ELSE 0 END) QTDRL160, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 6    THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR060, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 6.3  THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR063, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 8    THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR080, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 10   THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR100, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 12.5 THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR125, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 16   THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR160, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 20   THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR200, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 25   THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR250, \
			SUM(CASE WHEN SB1.B1__TPPA = 'V' AND SB1.B1__BITOLA = 32   THEN SZ3.Z3_QUANT ELSE 0 END) QTDBR320, \
			SUM(SZ3.Z3_QUANT) QTDTOT \
            FROM SZ2010 SZ2 \
                INNER JOIN SA1010 SA1 \
                    ON SA1.A1_FILIAL  = '  ' \
                    AND SA1.A1_COD     = SZ2.Z2_CLIENTE \
                    AND SA1.A1_LOJA    = SZ2.Z2_LOJA \
                    AND SA1.D_E_L_E_T_ = ' ' \
                INNER JOIN SZ3010 SZ3 \
                    ON SZ3.Z3_FILIAL  = SZ2.Z2_FILIAL \
                    AND SZ3.Z3_NUMERO  = SZ2.Z2_NUMERO \
                    AND SZ3.D_E_L_E_T_ = ' ' \
                INNER JOIN SC6010 SC6 \
                    ON SC6.C6_FILIAL  = SZ3.Z3_FILIAL \
                    AND SC6.C6_NUM     = SZ3.Z3_PEDIDO \
                    AND SC6.C6_ITEM    = SZ3.Z3_ITEM \
                    AND SC6.D_E_L_E_T_ = ' ' \
                INNER JOIN SB1010 SB1 \
                    ON SB1.B1_FILIAL  = '01' \
                    AND SB1.B1_COD     = SZ3.Z3_PRODUTO \
                    AND SB1.D_E_L_E_T_ = ' ' \
                LEFT JOIN SA4010 SA4 \
                    ON SA4.A4_FILIAL  = '  ' \
                    AND SA4.A4_COD     = SZ2.Z2_TRANSP \
                    AND SA4.D_E_L_E_T_ = ' ' \
            WHERE   SZ2.Z2_FILIAL  = '01' \
                AND ((SZ2.Z2_STATUS IN ('1','2','3','4','5','6','7','8') \
                AND   SZ2.Z2_DTCARR BETWEEN ? AND ?) \
                OR  SZ2.Z2_STATUS  =  '0') \
                AND SZ2.D_E_L_E_T_ = ' ' \
            GROUP BY SZ2.Z2_NUMERO, SZ2.Z2_STATUS, SZ2.Z2_CLIENTE, (CASE WHEN SZ2.Z2_STATUS = '0' THEN SZ2.Z2_EMISSAO ELSE SZ2.Z2_DTCARR END), \
                        (CASE WHEN SZ2.Z2_STATUS = '0' THEN '  :  ' ELSE SZ2.Z2_HORA END), SZ2.Z2_INIEMB, SZ2.Z2_MUNENT, SZ2.Z2_UFENT, SZ2.Z2_DTLBEMB, \
                        SZ2.Z2_HRLBEMB, SZ2.Z2_INIEMB, SZ2.Z2_HINIEMB, SZ2.Z2_FIMEMB, SZ2.Z2_HFIMEMB, SZ2.Z2_DTEXP, SZ2.Z2_HREXP, SZ2.Z2_DTNF, \
                        SZ2.Z2_NOTA, SZ3.Z3_PEDIDO, SC6.C6_ENTREG, ISNULL(SA4.A4_NREDUZ,''), SA1.A1_NREDUZ \
            ORDER BY (CASE WHEN SZ2.Z2_STATUS = '0' THEN SZ2.Z2_EMISSAO ELSE SZ2.Z2_DTCARR END), SZ2.Z2_NUMERO", (data_inicial, data_final))
    
    #// APAGA A TABELA TEMP1 DO USUARIO LOGADO E ZERA O AUTO INCREMENT
    cursor.execute("delete from carregamento_temp1 where usuario = %s",(usuario_logado,))
    #cursor.execute("ALTER TABLE carregamento_temp1 AUTO_INCREMENT = 0")
    db.commit()

    cursor.execute('select * from carregamento')
    tabela_fixa = cursor.fetchall()

    #// VERIFICA SE EXISTE REGISTRO(ALTERAÇÕES DE INFO E OBS ADICIONADOS PELO USUARIO) E ADICIONA NA TABELA TEMP1 CASO TENHA
    for i in cursor2:
        entrou = 0
        for t in tabela_fixa:
            if i[0].lstrip('0') == t[1]: #// VERIFICA SE A OT É IGUAL NA TABELA CARREGAMENTO - CASO SEJA IGUAL, A LINHA SERA ADICIONADA NA TABELA TEMP
                if i[1] == t[2]: #// VERIFICA SE O STATUS DE AMBOS SÃO IGUAIS. CASO NÃO SEJA, A TABELA TEMP1 SERA PREENCHIDA COM O STATUS VINDO DO PROTHEUS
                    cursor.execute("INSERT INTO carregamento_temp1 (\
                        ot,\
                        status,\
                        ncliente,\
                        dtprog,\
                        hrprog,\
                        dtcarreg,\
                        cidade,\
                        estado,\
                        dtlibemb,\
                        hrlibemb,\
                        dtiniemb,\
                        hriniemb,\
                        dtfimemb,\
                        hrfimemb,\
                        dtexped,\
                        hrexped,\
                        dtfat,\
                        nf,\
                        pedido,\
                        dtentrega,\
                        transp,\
                        cliente,\
                        fm055,\
                        fm065,\
                        fm070,\
                        rl063,\
                        rl080,\
                        rl100,\
                        rl125,\
                        rl160,\
                        br060,\
                        br063,\
                        br080,\
                        br100,\
                        br125,\
                        br160,\
                        br200,\
                        br250,\
                        br320,\
                        totaltn,\
                        info,\
                        obs,\
                        usuario)\
                        values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (t[1],t[2],t[3],t[4],t[5],t[6],t[7],t[8],t[9],t[10],t[11],t[12],t[13],t[14],t[15],t[16],t[17],t[18],t[19],t[20],t[21],t[22],t[23],t[24],t[25],t[26],t[27],t[28],t[29],t[30],t[31],t[32],t[33],t[34],t[35],t[36],t[37],t[38],t[39],t[40],t[42],t[43],usuario_logado,))
                else:
                    cursor.execute("INSERT INTO carregamento_temp1 (\
                        ot,\
                        status,\
                        ncliente,\
                        dtprog,\
                        hrprog,\
                        dtcarreg,\
                        cidade,\
                        estado,\
                        dtlibemb,\
                        hrlibemb,\
                        dtiniemb,\
                        hriniemb,\
                        dtfimemb,\
                        hrfimemb,\
                        dtexped,\
                        hrexped,\
                        dtfat,\
                        nf,\
                        pedido,\
                        dtentrega,\
                        transp,\
                        cliente,\
                        fm055,\
                        fm065,\
                        fm070,\
                        rl063,\
                        rl080,\
                        rl100,\
                        rl125,\
                        rl160,\
                        br060,\
                        br063,\
                        br080,\
                        br100,\
                        br125,\
                        br160,\
                        br200,\
                        br250,\
                        br320,\
                        totaltn,\
                        info,\
                        obs,\
                        usuario)\
                        values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (t[1],i[1],t[3],t[4],t[5],t[6],t[7],t[8],t[9],t[10],t[11],t[12],t[13],t[14],t[15],t[16],t[17],t[18],t[19],t[20],t[21],t[22],t[23],t[24],t[25],t[26],t[27],t[28],t[29],t[30],t[31],t[32],t[33],t[34],t[35],t[36],t[37],t[38],t[39],t[40],t[42],t[43],usuario_logado,))
                entrou = 1
        if entrou == 0:
            cursor.execute("INSERT INTO carregamento_temp1 (\
                ot,\
                status,\
                ncliente,\
                dtprog,\
                hrprog,\
                dtcarreg,\
                cidade,\
                estado,\
                dtlibemb,\
                hrlibemb,\
                dtiniemb,\
                hriniemb,\
                dtfimemb,\
                hrfimemb,\
                dtexped,\
                hrexped,\
                dtfat,\
                nf,\
                pedido,\
                dtentrega,\
                transp,\
                cliente,\
                fm055,\
                fm065,\
                fm070,\
                rl063,\
                rl080,\
                rl100,\
                rl125,\
                rl160,\
                br060,\
                br063,\
                br080,\
                br100,\
                br125,\
                br160,\
                br200,\
                br250,\
                br320,\
                totaltn,\
                usuario)\
                values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (i[0].lstrip('0'),i[1],i[2],i[3],i[4],i[5],i[6],i[7],i[8],i[9],i[10],i[11],i[12],i[13],i[14],i[15],i[16],i[17],i[18],i[19],i[20],i[21],i[22],i[23],i[24],i[25],i[26],i[27],i[28],i[29],i[30],i[31],i[32],i[33],i[34],i[35],i[36],i[37],i[38],i[39],usuario_logado,))
    db.commit()

    #// CHAMA A FUNÇÃO FILTRO PRINCIPAL
    filtro_principal()
    root.mainloop()

def filtrar_periodo():
    data_inicial = ent_data_inicial.get()
    data_inicial_convertida = (data_inicial[6:10]+data_inicial[3:5]+data_inicial[0:2])

    data_final = ent_data_final.get()
    data_final_convertida = (data_final[6:10]+data_final[3:5]+data_final[0:2])

    if data_inicial == '' or data_final == '':
        messagebox.showinfo('Programação de Carregamento', 'Preencha o campo (Data Inicial) e (Data final).', parent=root)
    else:
        global dt_ini_filtro
        dt_ini_filtro = data_inicial_convertida
        global dt_fim_filtro
        dt_fim_filtro = data_final_convertida
        global filtro_ativo
        filtro_ativo = 1
        atualizar_lista_filtro(data_inicial_convertida, data_final_convertida)

def bt_atualizar():
    if filtro_ativo == 1:
        atualizar_lista_filtro(dt_ini_filtro, dt_fim_filtro)
    elif filtro_ativo == 2:
        atualizar_lista_filtro(data, data)    
    else:
        atualizar_lista()

def remover_filtro():
    global filtro_ativo
    filtro_ativo = 0
    ent_data_inicial.configure(state='normal')
    ent_data_inicial.delete(0, END)
    ent_data_inicial.insert(0, 'Data inicial ->')
    ent_data_inicial.configure(state='readonly')
    
    ent_data_final.configure(state='normal')
    ent_data_final.delete(0, END)
    ent_data_final.insert(0, 'Data final ->')
    ent_data_final.configure(state='readonly')    
    
    ent_busca.delete(0, END)
    ent_busca.insert(0, '')

    atualizar_lista()

def versao():
    cursor.execute("SELECT versao FROM versao where versao = %s",(titulos,))
    versao = cursor.fetchone()
    if versao == None:
        messagebox.showwarning('Programação de Carregamento', 'Atualize a versão de seu software.', parent=root)
        root.destroy()
    else:
        pass

def login():
        def verifica_senha_protheus(usuario, senha):
            x=''
            autent = usuario+':'+senha
            autent_b = autent.encode()
            autorizacao = base64.b64encode(autent_b)

            url = 'http://192.168.1.18:8683/rest/api/oauth2/v1/token'
            headers = {
                'POST': '/rest/api/oauth2/v1/token',                
                'Host': 'http://192.168.1.18:8683',
                'Accept': 'application/json',
                'Authentication': 'BASIC '+ str(autorizacao),
            }
            parametros = {
                'grant_type':'password',
                'username':f'{usuario}',
                'password':f'{senha}',
            }
            x = requests.post(url, headers=headers, params=parametros)

            x = x.status_code

            if x == 201:
                global usuario_logado
                usuario_logado = usuario
                root2.destroy()
                atualizar_lista()
                
            else:
                messagebox.showwarning('Programação de Carregamento', 'Usuário ou senha inválidos.', parent=root2)

        def entrar_bind(event):
            entrar()

        def entrar():
            usuario = ent_usuario.get()
            senha = ent_senha.get()
            if usuario != '' and senha != '':
                verifica_senha_protheus(usuario, senha)
                
        root2 = Toplevel(root)
        root2.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
        root2.unbind_class("Button", "<Key-space>")
        root2.focus_force()
        root2.grab_set()

        window_width = 500
        window_height = 350
        screen_width = root2.winfo_screenwidth()
        screen_height = root2.winfo_screenheight() - 70
        x_cordinate = int((screen_width / 2) - (window_width / 2))
        y_cordinate = int((screen_height / 2) - (window_height / 2))
        root2.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
        root2.resizable(0, 0)
        root2.configure(bg='#ffffff')
        root2.title(titulos)
        #root2.iconbitmap('img\\icone.ico')

        fr1 = Frame(root2, bg='#ffffff')
        fr1.pack(side=TOP, fill=X, expand=False)
        fr2 = Frame(root2, bg='#eeeeee') #/// LINHA
        fr2.pack(side=TOP, fill=X, expand=False, pady=10)
        fr3 = Frame(root2, bg='#ffffff')
        fr3.pack(side=TOP, fill=X, expand=False)
        fr4 = Frame(root2, bg='#eeeeee') #/// LINHA
        fr4.pack(side=TOP, fill=X, expand=False, pady=10)
        fr5 = Frame(root2, bg='#ffffff')
        fr5.pack(side=TOP, fill=X, expand=False, pady=5)
        fr6 = Frame(root2, bg='#1D366C')
        fr6.pack(side=BOTTOM, fill=X, expand=False)
        
        #// Frame1
        img_phone = Image.open('img\\logo_login.png')
        resize_phone = img_phone.resize((180, 80))
        nova_img_phone = ImageTk.PhotoImage(resize_phone)
        lbl_phone = Label(fr1, image=nova_img_phone, text='', font=fonte_padrao, background="#ffffff", fg="#cdcdcd")
        lbl_phone.photo = nova_img_phone
        lbl_phone.grid(row=0, column=1, pady=10)

        lbl_phone = Label(fr1, text='Utilize o usuário e senha do Protheus!', font=fonte_padrao, background="#ffffff", fg="#363636")
        lbl_phone.photo = nova_img_phone
        lbl_phone.grid(row=1, column=1)

        fr1.grid_columnconfigure(0, weight=1) 
        fr1.grid_columnconfigure(2, weight=1) 
        
        #// Frame2 Linha

        #// Frame3
        lbl = Label(fr3, text='Usuário', font=fonte_padrao_bold, background="#ffffff", fg="#363636")
        lbl.grid(row=0, column=1, sticky=W)
        ent_usuario = customtkinter.CTkEntry(fr3, width=250, height=25, placeholder_text='Usuário do Protheus', **estilo_entry_padrao) 
        ent_usuario.grid(row=1, column=1, pady=(0,10))
        ent_usuario.focus_force()       
        ent_usuario.bind("<Return>", entrar_bind)

        lbl = Label(fr3, text='Senha', font=fonte_padrao_bold, background="#ffffff", fg="#363636")
        lbl.grid(row=2, column=1, sticky=W)
        ent_senha = customtkinter.CTkEntry(fr3, width=250, height=25, show='*', placeholder_text='Senha do Protheus', **estilo_entry_padrao) 
        ent_senha.grid(row=3, column=1, pady=(0,10))       
        ent_senha.bind("<Return>", entrar_bind)
        
        fr3.grid_columnconfigure(0, weight=1) 
        fr3.grid_columnconfigure(3, weight=1) 
        
        ent_usuario.insert(0, 'icasagrande')
        ent_senha.insert(0, '6176Ic12')

        #// Frame4 Linha
        
        #// Frame5
        bt_confirma = customtkinter.CTkButton(fr5, text='Entrar', command=entrar, width=100, text_color="#ffffff", font=fonte_padrao_bold, fg_color="#0275d8")
        bt_confirma.grid(row=0, column=1, padx=10)

        fr5.grid_columnconfigure(0, weight=1) 
        fr5.grid_columnconfigure(3, weight=1) 

        root2.wm_protocol("WM_DELETE_WINDOW", lambda: [root2.destroy(), root.destroy()])
        root2.mainloop()    

def busca_bind(event):
    busca()

def busca():
    busca = ent_busca.get()
    if busca != '':
        obs = ''
        db.cmd_reset_connection()
        cursor.execute("select * from carregamento_temp1 where ot = %s and usuario = %s ",(busca,usuario_logado,))
        result_busca = cursor.fetchall()

        if result_busca != None:
            img0 = Image.open('img\\00.png')
            resize_img0 = img0.resize((15, 15))
            nova_img0 = ImageTk.PhotoImage(resize_img0)

            img1 = Image.open('img\\01.png')
            resize_img1 = img1.resize((15, 15))
            nova_img1 = ImageTk.PhotoImage(resize_img1)
            
            img2 = Image.open('img\\02.png')
            resize_img2 = img2.resize((15, 15))
            nova_img2 = ImageTk.PhotoImage(resize_img2)

            img3 = Image.open('img\\03.png')
            resize_img3 = img3.resize((15, 15))
            nova_img3 = ImageTk.PhotoImage(resize_img3)

            img4 = Image.open('img\\04.png')
            resize_img4 = img4.resize((15, 15))
            nova_img4 = ImageTk.PhotoImage(resize_img4)

            img5 = Image.open('img\\05.png')
            resize_img5 = img5.resize((15, 15))
            nova_img5 = ImageTk.PhotoImage(resize_img5)

            img6 = Image.open('img\\06.png')
            resize_img6 = img6.resize((15, 15))
            nova_img6 = ImageTk.PhotoImage(resize_img6)

            img7 = Image.open('img\\07.png')
            resize_img7 = img7.resize((15, 15))
            nova_img7 = ImageTk.PhotoImage(resize_img7)

            img8 = Image.open('img\\08.png')
            resize_img8 = img8.resize((15, 15))
            nova_img8 = ImageTk.PhotoImage(resize_img8)

            img9 = Image.open('img\\09.png')
            resize_img9 = img9.resize((15, 15))
            nova_img9 = ImageTk.PhotoImage(resize_img9)

            img10 = Image.open('img\\10.png')
            resize_img10 = img10.resize((15, 15))
            nova_img10 = ImageTk.PhotoImage(resize_img10)

            img11 = Image.open('img\\11.png')
            resize_img11 = img11.resize((15, 15))
            nova_img11 = ImageTk.PhotoImage(resize_img11)

            img12 = Image.open('img\\12.png')
            resize_img12 = img12.resize((15, 15))
            nova_img12 = ImageTk.PhotoImage(resize_img12)

            img13 = Image.open('img\\13.png')
            resize_img13 = img13.resize((15, 15))
            nova_img13 = ImageTk.PhotoImage(resize_img13)

            tree_principal.delete(*tree_principal.get_children())
            for i in result_busca:
                if i[2] == '0':
                    status_ot = 'Sem OT'
                    img_status = nova_img0
                elif i[2] == '1':
                    status_ot = 'Aguard. Transp'
                    img_status = nova_img1
                elif i[2] == '2':
                    status_ot = 'Aguard. Caminhão'
                    img_status = nova_img2
                elif i[2] == '3':
                    status_ot = 'Lib. Embarque'
                    img_status = nova_img3
                elif i[2] == '4':
                    status_ot = 'Embarque Iniciado'
                    img_status = nova_img4
                elif i[2] == '5':
                    status_ot = 'Embarque Finalizado'
                    img_status = nova_img5
                elif i[2] == '6':
                    status_ot = 'Expedido'
                    img_status = nova_img6
                elif i[2] == '7':
                    status_ot = 'Faturado'
                    img_status = nova_img7
                elif i[2] == '8':
                    status_ot = 'Devolvido'
                    img_status = nova_img8
                elif i[2] == '9':
                    status_ot = 'Refaturado'
                    img_status = nova_img9            
                elif i[2] == '10':
                    status_ot = 'Refat. Expedido'
                    img_status = nova_img10
                elif i[2] == '11':
                    status_ot = 'Refat. Faturado'
                    img_status = nova_img11            
                elif i[2] == '12':
                    status_ot = 'Cancelado'
                    img_status = nova_img12            
                elif i[2] == '13':
                    status_ot = 'Bloqueado Crédito'
                    img_status = nova_img13          


                if i[42] == None or i[42] == '':
                    info = ''
                    tags_cor = 'par'
                else:
                    info = i[42]
                    tags_cor = 'contem'

                if i[43] == None or i[43] == ' ':
                    obs = ''

                if i[42] == '' and i[43] != '':
                    obs = i[43]
                    tags_cor = 'contem_obs'


                dt_prog = i[4][6:8]+'/'+i[4][4:6]+'/'+i[4][:4]
                dt_carregamento = i[6][6:8]+'/'+i[6][4:6]+'/'+i[6][:4]
                dt_entrega = i[20][6:8]+'/'+i[20][4:6]+'/'+i[20][:4]
                tree_principal.insert('', 'end', text=' ', image=img_status,
                                        values=(
                                        info,i[1].lstrip('0'), i[22].strip(), i[3], i[19], dt_entrega, str(i[23]).replace('.0', ''), str(i[24]).replace('.0', ''), str(i[25]).replace('.0', ''), str(i[26]).replace('.0', ''), str(i[27]).replace('.0', ''), str(i[28]).replace('.0', ''), str(i[29]).replace('.0', ''), str(i[30]).replace('.0', ''), str(i[31]).replace('.0', ''), str(i[32]).replace('.0', ''), str(i[33]).replace('.0', ''), str(i[34]).replace('.0', ''), str(i[35]).replace('.0', ''), str(i[36]).replace('.0', ''), str(i[37]).replace('.0', ''), str(i[38]).replace('.0', ''), str(i[39]).replace('.0', ''), str(i[40]).replace('.0', ''), dt_prog, i[5], dt_carregamento, i[7].strip(), i[8], i[21], obs),
                                        tags=(tags_cor,))
            root.mainloop()
        else:
            messagebox.showwarning('Programação de Carregamento', 'Ordem de transporte não encontrada!', parent=root)

#/////ESTRUTURA
root = customtkinter.CTk()
root.state('zoomed')
root.configure(fg_color='#ffffff')
#root.after(0, login)
root.title(titulos)

frame0 = customtkinter.CTkFrame(root, bg_color='#ffffff',corner_radius=20, fg_color='#ffffff', border_width=2, border_color='#000033')
frame0.pack(side=TOP, fill=BOTH, expand=TRUE, padx=10, pady=10)

frame1 = Frame(frame0, bg='#ffffff')
frame1.pack(side=TOP, fill=X, expand=False, padx=10, pady=10)
frame2 = Frame(frame0, bg='#000033') #/////LINHA
frame2.pack(side=TOP, fill=X, expand=False, padx=10, pady=2)
frame3 = Frame(frame0, bg='#ffffff')
frame3.pack(side=TOP, fill=X, expand=False, padx=10, pady=10)
frame4 = Frame(frame0, bg='#000033') #/////LINHA
frame4.pack(side=TOP, fill=X, expand=False, padx=10, pady=2)
frame5 = Frame(frame0, bg='#ffffff')
frame5.pack(side=TOP, fill=BOTH, expand=True, padx=10, pady=10)

#/////LAYOUT
#/////FRAME1
img_logo = Image.open('img\\logo.png')
resize_logo = img_logo.resize((140, 80))
nova_img_logo = ImageTk.PhotoImage(resize_logo)
lbl_logo = Label(frame1, image=nova_img_logo, background='#ffffff')
lbl_logo.photo = nova_img_logo
lbl_logo.grid(row=0, column=1, padx=6)
lbl_titulo = Label(frame1, text='Programação de Carregamento', font=fonte_padrao_titulo, bg='#ffffff', fg='#1d366c')
lbl_titulo.grid(row=0, column=2)

frame1.grid_columnconfigure(0, weight=1)
frame1.grid_columnconfigure(4, weight=1)

#/////FRAME2 LINHA

#/////FRAME3

img_peso = Image.open('img\\peso.png')
resize_peso = img_peso.resize((30, 30))
nova_img_peso = ImageTk.PhotoImage(resize_peso)
lbl_peso = Button(frame3, image=nova_img_peso, background='#ffffff', borderwidth=0, relief=RIDGE,activebackground="#ffffff", activeforeground="#1d366c", cursor="hand2", command=soma_toneladas)
lbl_peso.photo = nova_img_peso
lbl_peso.grid(row=0, column=1, padx=0)

lbl_peso = Label(frame3, text='0 TN', background='#ffffff', font=fonte_padrao, foreground='#2a2e2e')
lbl_peso.grid(row=0, column=2)

lbl_separacao = Label(frame3, text='|', background='#ffffff', font=fonte_padrao, foreground='#2a2e2e')
lbl_separacao.grid(row=0, column=3, padx=5)

bt1 = customtkinter.CTkButton(frame3, width=100, text='Hoje',text_color='#ffffff', hover_color='#3DC2FF', font=fonte_padrao_bold, bg_color='#ffffff', fg_color='#5cb85c', corner_radius=5, command=filtro_home)
bt1.grid(row=0, column=4)

lbl_separacao = Label(frame3, text='|', background='#ffffff', font=fonte_padrao, foreground='#2a2e2e')
lbl_separacao.grid(row=0, column=5, padx=5)

bt1 = customtkinter.CTkButton(frame3, width=100, text='Atualizar',text_color='#ffffff', hover_color='#3DC2FF', font=fonte_padrao_bold, bg_color='#ffffff', fg_color='#0275d8', corner_radius=5, command=bt_atualizar)
bt1.grid(row=0, column=6)

lbl_separacao = Label(frame3, text='|', background='#ffffff', font=fonte_padrao, foreground='#2a2e2e')
lbl_separacao.grid(row=0, column=7, padx=5)

ent_data_inicial = customtkinter.CTkEntry(frame3, **estilo_entry_padrao, width=100, placeholder_text='Data inicial ->', placeholder_text_color='#2a2d2e')
ent_data_inicial.grid(row=0, column=8, padx=5)
ent_data_inicial.configure(state='readonly')

img_data = Image.open('img\\agenda.png')
resize_data = img_data.resize((25, 25))
nova_img_data = ImageTk.PhotoImage(resize_data)
lbl_data_inicial = Button(frame3, image=nova_img_data, text='', command=calendario_inicio, **bt_icone)
lbl_data_inicial.photo = nova_img_data
lbl_data_inicial.grid(row=0, column=9)

lbl_separacao = Label(frame3, text='--', background='#ffffff', font=fonte_padrao, foreground='#2a2e2e')
lbl_separacao.grid(row=0, column=10, padx=5)

ent_data_final = customtkinter.CTkEntry(frame3, **estilo_entry_padrao, width=100, placeholder_text='Data final ->', placeholder_text_color='#2a2d2e')
ent_data_final.grid(row=0, column=11, padx=(0,5))
ent_data_final.configure(state='readonly')

img_data = Image.open('img\\agenda.png')
resize_data = img_data.resize((25, 25))
nova_img_data = ImageTk.PhotoImage(resize_data)
lbl_data = Button(frame3, image=nova_img_data, text='', command=calendario_final, **bt_icone)
lbl_data.photo = nova_img_data
lbl_data.grid(row=0, column=12)

lbl_separacao = Label(frame3, text='->', background='#ffffff', font=fonte_padrao, foreground='#2a2e2e')
lbl_separacao.grid(row=0, column=13, padx=5)

bt1 = customtkinter.CTkButton(frame3, width=100, text='Filtrar Período',text_color='#ffffff', hover_color='#3DC2FF', font=fonte_padrao_bold, bg_color='#ffffff', fg_color='#0275d8', corner_radius=5, command=filtrar_periodo)
bt1.grid(row=0, column=14)

lbl_separacao = Label(frame3, text='|', background='#ffffff', font=fonte_padrao, foreground='#2a2e2e')
lbl_separacao.grid(row=0, column=15, padx=5)


ent_busca = customtkinter.CTkEntry(frame3, **estilo_entry_padrao, width=100, placeholder_text='Nº da OT', placeholder_text_color='#2a2d2e')
ent_busca.grid(row=0, column=16, padx=(0,5))
ent_busca.bind("<Return>", busca_bind)

img_busca = Image.open('img\\lupa.png')
resize_busca = img_busca.resize((25, 25))
nova_img_busca = ImageTk.PhotoImage(resize_busca)
lbl_busca = Button(frame3, image=nova_img_busca, text='', command=busca, **bt_icone)
lbl_busca.photo = nova_img_busca
lbl_busca.grid(row=0, column=17)


lbl_separacao = Label(frame3, text='|', background='#ffffff', font=fonte_padrao, foreground='#2a2e2e')
lbl_separacao.grid(row=0, column=18, padx=5)


bt1 = customtkinter.CTkButton(frame3, width=100, text='Remover Filtro',text_color='#ffffff', hover_color='#3DC2FF', font=fonte_padrao_bold, bg_color='#ffffff', fg_color='#EB445A', corner_radius=5, command=remover_filtro)
bt1.grid(row=0, column=19)

frame3.grid_columnconfigure(0, weight=1)
frame3.grid_columnconfigure(20, weight=1)

#/////FRAME4 LINHA

#/////FRAME5

lista = []


style = ttk.Style()
# style.theme_use('default')
style.configure('Treeview',
                background='#ffffff',
                rowheight=24,
                fieldbackground='#ffffff',
                font=fonte_padrao)
style.configure("Treeview.Heading",
                foreground='#1d366c',
                background="#ffffff",
                font=fonte_padrao_bold)
style.map('Treeview', background=[('selected', '#0275d8')])

colunas = ("1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26","27","28","29","30","31")
tree_principal = ttk.Treeview(frame5, selectmode='extended', columns=colunas)
vsb = ttk.Scrollbar(frame5, orient="vertical", command=tree_principal.yview)
vsb.pack(side=RIGHT, fill='y')
tree_principal.configure(yscrollcommand=vsb.set)
vsbx = ttk.Scrollbar(frame5, orient="horizontal", command=tree_principal.xview)
vsbx.pack(side=BOTTOM, fill='x')
tree_principal.configure(xscrollcommand=vsbx.set)
tree_principal.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')


tree_principal.column("#0", anchor='w', minwidth=50, width=50)
tree_principal.column("1", anchor='c', minwidth=125, width=90)
tree_principal.column("2", anchor='c', minwidth=80, width=80)
tree_principal.column("3", anchor='c', minwidth=100, width=100)
tree_principal.column("4", anchor='c', minwidth=76, width=76)
tree_principal.column("5", anchor='c', minwidth=60, width=60)
tree_principal.column("6", anchor='c', minwidth=80, width=80)
tree_principal.column("7", anchor='c', minwidth=45, width=45)
tree_principal.column("8", anchor='c', minwidth=45, width=45)
tree_principal.column("9", anchor='c', minwidth=45, width=45)
tree_principal.column("10", anchor='c', minwidth=45, width=45)
tree_principal.column("11", anchor='c', minwidth=45, width=45)
tree_principal.column("12", anchor='c', minwidth=45, width=45)
tree_principal.column("13", anchor='c', minwidth=45, width=45)
tree_principal.column("14", anchor='c', minwidth=45, width=45)
tree_principal.column("15", anchor='c', minwidth=45, width=45)
tree_principal.column("16", anchor='c', minwidth=45, width=45)
tree_principal.column("17", anchor='c', minwidth=45, width=45)
tree_principal.column("18", anchor='c', minwidth=45, width=45)
tree_principal.column("19", anchor='c', minwidth=45, width=45)
tree_principal.column("20", anchor='c', minwidth=45, width=45)
tree_principal.column("21", anchor='c', minwidth=45, width=45)
tree_principal.column("22", anchor='c', minwidth=45, width=45)
tree_principal.column("23", anchor='c', minwidth=45, width=45)
tree_principal.column("24", anchor='c', minwidth=45, width=45)
tree_principal.column("25", anchor='c', minwidth=80, width=80)
tree_principal.column("26", anchor='c', minwidth=54, width=54)
tree_principal.column("27", anchor='c', minwidth=80, width=80)
tree_principal.column("28", anchor='c', minwidth=120, width=120)
tree_principal.column("29", anchor='c', minwidth=40, width=40)
tree_principal.column("30", anchor='c', minwidth=100, width=100)
tree_principal.column("31", anchor='c', minwidth=100, width=100)

tree_principal.heading("#0", text="Status")
tree_principal.heading("1", text="Previsto/Chegada")
tree_principal.heading("2", text="OT")
tree_principal.heading("3", text="Cliente")
tree_principal.heading("4", text="Nº Cliente")
tree_principal.heading("5", text="Pedido")
tree_principal.heading("6", text="Dt.Entrega")
tree_principal.heading("7", text="F5,5")
tree_principal.heading("8", text="F6,5")
tree_principal.heading("9", text="F7,0")
tree_principal.heading("10", text="R6,3")
tree_principal.heading("11", text="R8,0")
tree_principal.heading("12", text="R10,0")
tree_principal.heading("13", text="R12,5")
tree_principal.heading("14", text="R16,0")
tree_principal.heading("15", text="B6,0")
tree_principal.heading("16", text="B6,3")
tree_principal.heading("17", text="B8,0")
tree_principal.heading("18", text="B10,0")
tree_principal.heading("19", text="B12,5")
tree_principal.heading("20", text="B16,0")
tree_principal.heading("21", text="B20,0")
tree_principal.heading("22", text="B25,0")
tree_principal.heading("23", text="B32,0")
tree_principal.heading("24", text="TN")
tree_principal.heading("25", text="Data Prog")
tree_principal.heading("26", text="Hora Prog")
tree_principal.heading("27", text="Data Carreg")
tree_principal.heading("28", text="Cidade")
tree_principal.heading("29", text="UF")
tree_principal.heading("30", text="Transp")
tree_principal.heading("31", text="Obs")

tree_principal.tag_configure('par', background='#e9e9e9')
tree_principal.tag_configure('impar', background='#ffffff')
tree_principal.tag_configure('contem', background='#00FA9A')
tree_principal.tag_configure('contem_obs', font=fonte_padrao_bold)
tree_principal.bind("<Double-1>", detalhes)
frame5.grid_columnconfigure(0, weight=1)
frame5.grid_columnconfigure(3, weight=1)

#//CONEXAO COM O BANCO DE DADOS
server = '192.168.1.18'
database = 'Protheus'
username = 'pcp'
password = 'gv@2020'
conectar = pyodbc.connect('DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
cursor2 =  conectar.cursor()


db = mysql.connector.connect(
    #host="localhost",
    host="192.168.11.125",
    #user="root",
    user="acesso_rede",
    passwd="senha",
    database="simec_carregamento",
)
cursor = db.cursor(buffered=True)

root.mainloop()

