
import time
import webbrowser
import pandas as pd
import phonenumbers
import pyautogui
import streamlit as st
from urllib.parse import quote_plus

# ========= CONFIG PADRÃO =========
COLS_PADRAO = {
    "Nome": ["Nome", "Destinatario", "Cliente", "Contato"],
    "Numero": ["Numero", "Telefone", "Celular", "WhatsApp", "Whatsapp"],
    "Pedido": ["Pedido", "N_Pedido", "OS", "Ordem"],
    "DataEntrega": ["DataEntrega", "Previsao", "Data Prevista", "Entrega Prevista"],
    "OptOut": ["OptOut", "Opt-out", "NaoEnviar", "Bloqueado", "Sair"]
}
WAIT_CHAT_DEFAULT = 10  # segundos
PAUSA_ENTRE_DEFAULT = 3 # segundos
# =================================

# PyAutoGUI segurança
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.25

# ---------- Núcleo ----------
def e164_br(numero_raw: str) -> str:
    s = "".join(ch for ch in str(numero_raw) if ch.isdigit() or ch == "+")
    if not s:
        raise ValueError("Número vazio")
    if not s.startswith("+"):
        s = "+55" + s
    p = phonenumbers.parse(s)
    if not phonenumbers.is_valid_number(p):
        raise ValueError(f"Número inválido: {numero_raw}")
    return phonenumbers.format_number(p, phonenumbers.PhoneNumberFormat.E164)

def montar_msg(nome, pedido, data_prev):
    return (
        "Sua compra está a caminho!\n\n"
        f"Olá, sou Rharianne da transportadora WS TRANSPORTE. Estou falando com {nome}?\n"
        "Estou entrando em contato para confirmar se este número está ativo e se o endereço de destino está correto "
        "para seguirmos com a rota de entrega do seu pedido!\n\n"
        f"📦 Número do pedido: {pedido}\n"
        f"🚚 Previsão de entrega: até {data_prev}\n\n"
        "Se não for você, responda “NÃO”. Para não receber mensagens, responda “SAIR”."
    ).strip()

def garantir_foco_navegador():
    # traz o navegador pra frente antes de apertar ENTER
    pyautogui.hotkey("alt", "tab")
    time.sleep(0.6)
    pyautogui.hotkey("alt", "tab")
    time.sleep(0.4)

def enviar_whatsapp(numero_e164: str, mensagem: str, wait_chat: int, simular: bool = False):
    texto_url = quote_plus(mensagem)
    url = f"https://web.whatsapp.com/send?phone={numero_e164.replace('+','')}&text={texto_url}"
    webbrowser.open(url)
    time.sleep(wait_chat)
    if not simular:
        garantir_foco_navegador()
        pyautogui.press("enter")
        time.sleep(1.2)

# ---------- Utilidades ----------
def tentar_auto_mapear_colunas(df: pd.DataFrame) -> dict:
    mapa = {}
    cols_lower = {c.lower(): c for c in df.columns}
    for alvo, candidatos in COLS_PADRAO.items():
        escolhido = None
        for cand in candidatos:
            cand_lower = cand.lower()
            if cand_lower in cols_lower:
                escolhido = cols_lower[cand_lower]
                break
            # aproximação por “contém”
            for c in df.columns:
                if cand_lower in c.lower():
                    escolhido = c
                    break
            if escolhido:
                break
        mapa[alvo] = escolhido
    return mapa

def aplicar_optout(df, col_optout):
    if not col_optout or col_optout not in df.columns:
        return df.copy()
    return df[df[col_optout].astype(str).str.upper().ne("Y")].copy()

# ============== UI ==============
st.set_page_config(page_title="Envio WhatsApp - WS", page_icon="💬", layout="centered")
st.title("💬 Envio de WhatsApp — WS Transportes")

st.markdown("""
Suba a planilha Excel, confirme o mapeamento das colunas e clique **Iniciar envios**.  
> **Dicas importantes**  
> • Faça login no **WhatsApp Web** no navegador padrão antes de iniciar.  
> • Feche abas desnecessárias; deixe apenas a aba do Streamlit e a do WhatsApp Web abertas.  
> • Use primeiro o modo **Simular (não envia)** para validar mapeamento e mensagens.
""")

uploaded = st.file_uploader("Planilha Excel", type=["xlsx", "xls"])
if uploaded:
    df_raw = pd.read_excel(uploaded).fillna("")
    st.subheader("Prévia da planilha")
    st.dataframe(df_raw.head(20), use_container_width=True)

    # Mapeamento de colunas
    st.subheader("Mapeamento de colunas")
    sugestao = tentar_auto_mapear_colunas(df_raw)
    col1, col2 = st.columns(2)
    with col1:
        col_nome = st.selectbox("Coluna de Nome", [None]+list(df_raw.columns), index=( [None]+list(df_raw.columns) ).index(sugestao.get("Nome")) if sugestao.get("Nome") in ([None]+list(df_raw.columns)) else 0)
        col_num  = st.selectbox("Coluna de Número", [None]+list(df_raw.columns), index=( [None]+list(df_raw.columns) ).index(sugestao.get("Numero")) if sugestao.get("Numero") in ([None]+list(df_raw.columns)) else 0)
        col_ped  = st.selectbox("Coluna de Pedido", [None]+list(df_raw.columns), index=( [None]+list(df_raw.columns) ).index(sugestao.get("Pedido")) if sugestao.get("Pedido") in ([None]+list(df_raw.columns)) else 0)
    with col2:
        col_data = st.selectbox("Coluna de Data de Entrega", [None]+list(df_raw.columns), index=( [None]+list(df_raw.columns) ).index(sugestao.get("DataEntrega")) if sugestao.get("DataEntrega") in ([None]+list(df_raw.columns)) else 0)
        col_out  = st.selectbox("Coluna de Opt-Out (Y = não enviar)", [None]+list(df_raw.columns), index=( [None]+list(df_raw.columns) ).index(sugestao.get("OptOut")) if sugestao.get("OptOut") in ([None]+list(df_raw.columns)) else 0)

    obrigatorias = [col_nome, col_num, col_ped, col_data]
    if any(c is None for c in obrigatorias):
        st.error("Mapeie todas as colunas obrigatórias: Nome, Número, Pedido e Data de Entrega.")
        st.stop()

    df = df_raw[[col_nome, col_num, col_ped, col_data] + ([col_out] if col_out else [])].copy()
    df.columns = ["Nome", "Numero", "Pedido", "DataEntrega"] + (["OptOut"] if col_out else [])

    # Respeita opt-out
    df = aplicar_optout(df, "OptOut" if "OptOut" in df.columns else None)

    st.success(f"Contatos após Opt-out: {len(df)}")
    st.dataframe(df.head(20), use_container_width=True)

    st.subheader("Parâmetros")
    colA, colB, colC = st.columns(3)
    with colA:
        wait_chat = st.number_input("Espera para carregar o chat (s)", min_value=5, max_value=30, value=WAIT_CHAT_DEFAULT)
    with colB:
        pausa_entre = st.number_input("Pausa entre contatos (s)", min_value=1, max_value=30, value=PAUSA_ENTRE_DEFAULT)
    with colC:
        simular = st.checkbox("Simular (não envia)", value=True, help="Abre o WhatsApp com a mensagem preenchida, mas não pressiona Enter.")

    st.markdown("—")
    iniciar = st.button("🚀 Iniciar envios")
    if iniciar:
        if df.empty:
            st.warning("Nenhum contato para enviar.")
            st.stop()

        # Resultado agregado
        resultados = []
        barra = st.progress(0)
        area_log = st.empty()

        enviados = 0
        falhas = 0
        total = len(df)

        st.info("Certifique-se de estar logada no WhatsApp Web no **mesmo navegador padrão**.")
        time.sleep(1)

        for idx, r in df.iterrows():
            nome = str(r["Nome"])
            numero_raw = str(r["Numero"])
            pedido = str(r["Pedido"])
            data_prev = str(r["DataEntrega"])

            try:
                numero = e164_br(numero_raw)
                msg = montar_msg(nome, pedido, data_prev)

                area_log.write(f"Enviando para **{nome}** ({numero})…")
                enviar_whatsapp(numero, msg, wait_chat=wait_chat, simular=simular)
                enviados += 1
                status = "SIMULADO" if simular else "ENVIADO"
                resultados.append({"Nome": nome, "Numero": numero, "Pedido": pedido, "DataEntrega": data_prev, "Status": status, "Erro": ""})
            except Exception as e:
                falhas += 1
                resultados.append({"Nome": nome, "Numero": numero_raw, "Pedido": pedido, "DataEntrega": data_prev, "Status": "FALHA", "Erro": str(e)})
                area_log.write(f"⚠️ Falha com **{numero_raw}**: {e}")

            barra.progress(min(int(((enviados+falhas)/total)*100), 100))
            time.sleep(pausa_entre)

        st.success(f"Finalizado. Enviados: {enviados} | Falhas: {falhas} | Total: {total}")

        df_res = pd.DataFrame(resultados)
        st.dataframe(df_res, use_container_width=True)
        csv = df_res.to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ Baixar log (CSV)", data=csv, file_name="log_envios_whatsapp.csv", mime="text/csv")

        if simular:
            st.warning("Você rodou em **modo Simular**. Se estiver tudo certo, desmarque a opção e rode novamente para enviar de fato.")
