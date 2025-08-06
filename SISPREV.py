import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import threading
import time

ANALISTAS = [
    "VANESSA DE FARIA RAMOS MATES",
    "ZENILDA FERREIRA DOS SANTOS",
    "CARMEM LUCIA CARNEIRO AFONSO DA CUNHA GUIO",
    "ALESSANDRO MONJARDIM OSLEGHER DE ALMEIDA",
    "BRUNO TAMANINI LOPES"
]

def automatizar(lista_entradas, tipo_busca):
    status.set("Iniciando automação...")
    chrome_options = Options()
    chrome_options.debugger_address = "localhost:9222"
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 15)
    fast_wait = WebDriverWait(driver, 2)

    for nome_segurado, analista in lista_entradas:
        try:
            seletor_procura = Select(wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentCampos_ddlProcura"))))
            seletor_procura.select_by_value(tipo_busca)
            time.sleep(1)

            campo_busca = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentCampos_txtPesquisa")))
            campo_busca.clear()
            campo_busca.send_keys(nome_segurado + Keys.ENTER)

            try:
                btn_processo = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_ContentCampos_AccordionPane2_content_grdProcessoSetor_ctl03_imgbtnReceber']"))
                )
                time.sleep(1.5)
                btn_processo.click()
            except:
                btn_processo = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_ContentCampos_AccordionPane2_content_grdProcessoSetor_ctl03_imgbtnEdit']"))
                )
                btn_processo.click()

            time.sleep(1.5)

            btn_tramitar = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentToolBar_btnTramitar")))
            btn_tramitar.click()
            time.sleep(2)

            select_tramite = Select(wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentToolBar_ddlTipoTramite"))))
            select_tramite.select_by_visible_text("Interna")
            time.sleep(0.5)

            select_modelo_despacho = Select(wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentToolBar_ddlModeloDespacho"))))
            select_modelo_despacho.select_by_visible_text("DISTRIBUIÇÃO DE PROCESSOS")
            time.sleep(1)

            select_despacho = Select(wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentToolBar_ddlDespacho"))))
            select_despacho.select_by_visible_text("ENCAMINHADO")
            time.sleep(0.5)

            select_usuario = Select(wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentToolBar_ddlUsuario"))))
            select_usuario.select_by_visible_text(analista)
            time.sleep(0.5)

        except Exception as e:
            status.set(f"Erro ao preencher os campos de tramitação: {str(e)}")
            return

        try:
            btn_enviar = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentToolBar_Button1")))
            original_windows = driver.window_handles
            btn_enviar.click()

            try:
                WebDriverWait(driver, 5).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                alert.accept()
                status.set("Alerta de confirmação aceito.")
            except:
                status.set("Processo tramitado, mas nenhum alerta de confirmação foi detectado.")

            WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > len(original_windows))
            new_windows = driver.window_handles
            nova_janela = [w for w in new_windows if w not in original_windows]

            if nova_janela:
                driver.switch_to.window(nova_janela[0])
                WebDriverWait(driver, 10).until(EC.url_contains("VisualizaRelatorio.aspx"))
                time.sleep(2)
                driver.close()
                driver.switch_to.window(original_windows[0])

            status.set(f"Processo tramitado com sucesso para {analista}.")

        except Exception as e:
            status.set(f"Erro ao clicar no botão Tramitar ou acessar relatório: {str(e)}")


def iniciar_automacao():
    tipo_busca = combo_tipo_busca.get().strip()
    try:
        qtde = int(combo_qtde.get())
    except:
        status.set("Selecione a quantidade de processos.")
        return

    lista_entradas = []
    for i in range(qtde):
        nome = entradas_nome[i].get().strip()
        analista = entradas_analista[i].get().strip()
        if not nome or not analista:
            status.set(f"Preencha todos os campos da linha {i+1}.")
            return
        lista_entradas.append((nome, analista))

    for entrada in entradas_nome + entradas_analista:
        entrada.delete(0, tk.END)
    status.set("")

    threading.Thread(target=automatizar, args=(lista_entradas, tipo_busca), daemon=True).start()


# Interface
root = tk.Tk()
root.title("Tramitação Automática - SISPREV")
root.geometry("800x600")
root.configure(bg="#f0f4f7")
root.attributes("-topmost", False)

titulo = tk.Label(root, text="👋 Olá, Érica! Bem-vinda ao sistema de tramitação automática ❤️", font=("Arial", 14, "bold"), bg="#f0f4f7", fg="#1a1a1a")
titulo.pack(pady=(20, 10))

frame_top = tk.Frame(root, bg="#f0f4f7")
frame_top.pack(pady=5)

tk.Label(frame_top, text="Quantidade:", bg="#f0f4f7").grid(row=0, column=0, padx=5)
combo_qtde = ttk.Combobox(frame_top, values=list(range(1, 11)), width=5, state="readonly")
combo_qtde.grid(row=0, column=1, padx=5)

tk.Label(frame_top, text="Tipo de busca:", bg="#f0f4f7").grid(row=0, column=2, padx=5)
combo_tipo_busca = ttk.Combobox(frame_top, values=["NOME_SEGURADO", "NR_PROCESSO_ADM"], width=20, state="readonly")
combo_tipo_busca.grid(row=0, column=3, padx=5)
combo_tipo_busca.set("NOME_SEGURADO")

fixar_var = tk.IntVar()
fixar_check = tk.Checkbutton(frame_top, text="📌 Fixar na tela", variable=fixar_var, bg="#f0f4f7", command=lambda: root.attributes("-topmost", fixar_var.get()))
fixar_check.grid(row=0, column=4, padx=10)

frame_inputs = tk.Frame(root, bg="#f0f4f7")
frame_inputs.pack(pady=10)

entradas_nome = []
entradas_analista = []

def atualizar_campos(*args):
    for widget in frame_inputs.winfo_children():
        widget.destroy()
    entradas_nome.clear()
    entradas_analista.clear()

    try:
        qtde = int(combo_qtde.get())
    except:
        return

    for i in range(qtde):
        tk.Label(frame_inputs, text=f"{i+1}. Nome ou número:", bg="#f0f4f7").grid(row=i, column=0, sticky="w")
        nome_entry = tk.Entry(frame_inputs, width=40)
        nome_entry.grid(row=i, column=1, padx=5, pady=2)
        entradas_nome.append(nome_entry)

        analista_cb = ttk.Combobox(frame_inputs, values=ANALISTAS, width=35, state="readonly")
        analista_cb.grid(row=i, column=2, padx=5, pady=2)
        entradas_analista.append(analista_cb)

combo_qtde.bind("<<ComboboxSelected>>", atualizar_campos)

btn_enviar = tk.Button(root, text="🚀 Tramitar Todos", bg="#0066cc", fg="white", font=("Arial", 10, "bold"), command=iniciar_automacao)
btn_enviar.pack(pady=15)

status = tk.StringVar()
tk.Label(root, textvariable=status, fg="blue", bg="#f0f4f7").pack()

btn_enviar.focus()
root.mainloop()
