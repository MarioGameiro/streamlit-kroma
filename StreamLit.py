import streamlit as st
import os
import xlwings as xw
from datetime import datetime
import zipfile
import pandas as pd
from io import BytesIO
from datetime import date
from io import TextIOWrapper
import re
import streamlit_authenticator as stauth
import numpy as np
from pathlib import Path
# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

senhas_criptografadas = stauth.Hasher(["123456", "123123", "333333"]).generate()

credenciais = {"usernames": {
    "mario.moura@kromaenergia.com.br": {"name": "Mario", "password": senhas_criptografadas[0]}}}

authenticator = stauth.Authenticate(credenciais, "credenciais_hashco", "fsyfus%$67fs76AH7", cookie_expiry_days=30)


def autenticar_usuario(authenticator):
    nome, status_autenticacao, username = authenticator.login()

    if status_autenticacao:
        return {"nome": nome, "username": username}
    elif status_autenticacao == False:
        st.error("Combinação de usuário e senha inválidas")
    else:
        st.error("Preencha o formulário para fazer login")


def logout(authenticator):
    authenticator.logout()


# autenticar o usuario
dados_usuario = autenticar_usuario(authenticator)

# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
if dados_usuario:

    # Função para processar os dados e executar a macro no script "Gerar Prevs Sensibilidade Matriz"
    def processar_e_executar_macro(lines, rv_file_content, selected_year, selected_month, folder_selected):
        cont = []

        # Parseando o conteúdo do arquivo
        for i in range(0, len(lines)):
            cont.append((lines[i].replace('\n', '')).split(','))

        path_calc_ena = os.path.join(os.getcwd(),'PycharmProjects\StremLit\.venv','Planilha cálculo de ENA - CurtoPrazo.xlsm')

        # Abrir o arquivo Excel
        #wb = xw.Book('Z:\\Middle\\ALVES\\Planilha cálculo de ENA - CurtoPrazo.xlsm')
        wb = xw.Book(path_calc_ena)
        sheet = wb.sheets("Prevs")

        # Colocar o conteúdo do arquivo .rvx na célula B30
        sheet.range("B30").value = rv_file_content

        # Colocar o ano e mês nas células J10 e J11
        sheet.range("J10").value = selected_year
        sheet.range("J11").value = selected_month

        for i in range(0, len(cont)):
            # Atribuindo os valores nas células
            sheet.range("X4").value = cont[i][0]
            sheet.range("X5").value = cont[i][1]
            sheet.range("X6").value = cont[i][2]
            sheet.range("X7").value = cont[i][3]

            # Executar a macro
            your_macro = wb.macro('Calc_1')  # Nome da macro
            your_macro()

            # Criar o nome do arquivo com base nos dados
            arquivo = "SE" + cont[i][0].replace("%", "") + "SUL" + cont[i][1].replace("%", "") + "NE" + cont[i][
                2].replace("%", "") + "NO" + cont[i][3].replace("%", "")

            # Gerar o arquivo de saída
            file_path = os.path.join(folder_selected, f'prevs-{arquivo}.rv0')
            with open(file_path, 'w') as prevs:
                vazoes = sheet.range("J30:J197").value
                for j in range(0, len(vazoes)):
                    prevs.write(str(vazoes[j]) + "\n")

            st.write(f"Prevs.{arquivo}.rv0 gerado com sucesso.")

        wb.save()
        wb.close()


    # ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    # Função para compilar os arquivos .zip e processar "compila_cmo_medio.csv", "compila_ea_inicial.csv" e "compila_ena_mensal_percentual.csv"
    def compila_estudo_lp(zip_files):
        results = []
        current_date = datetime.now().strftime("%Y-%m-%d")  # Data de compilação atual

        for zip_file in zip_files:
            with zipfile.ZipFile(BytesIO(zip_file.read()), 'r') as z:
                # Verificar se os arquivos estão presentes no zip
                if ("compila_cmo_medio.csv" in z.namelist() and
                        "compila_ea_inicial.csv" in z.namelist() and
                        "compila_ena_mensal_percentual.csv" in z.namelist()):

                    # Processar "compila_cmo_medio.csv"
                    with z.open("compila_cmo_medio.csv") as f:
                        df = pd.read_csv(f, sep=';')
                        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                        df[['Mes', 'Semana']] = df['Deck'].str.split('-', expand=True)
                        df.drop(['MEN=0-SEM=1', 'Deck', 'Semana'], axis=1, inplace=True)
                        df = df[~df['Mes'].astype(str).str.startswith('N')]  # Remove os valores de NW
                        df.rename(columns={
                            'SUDESTE': 'PLD SE/CO',
                            'SUL': 'PLD SUL',
                            'NORDESTE': 'PLD NE',
                            'NORTE': 'PLD NO'
                        }, inplace=True)
                        df['Ano'] = datetime.now().year  # Aqui, assumi o ano atual; ajuste se for diferente
                        df['Data de Compilação'] = current_date
                        df['Nome do Arquivo'] = zip_file.name

                        # Calcular a média por Ano, Mes e Submercado
                        df_media = df.groupby(['Ano', 'Mes', 'Data de Compilação', 'Nome do Arquivo'])[
                            ['PLD SE/CO', 'PLD SUL', 'PLD NE', 'PLD NO']].mean().reset_index()

                        # Fazer o melt para transformar os submercados em formato longo
                        df_melted = df_media.melt(
                            id_vars=['Ano', 'Data de Compilação', 'Nome do Arquivo', 'Mes'],
                            value_vars=['PLD SE/CO', 'PLD SUL', 'PLD NE', 'PLD NO'],
                            var_name='Variável',
                            value_name='Valor'
                        )

                    # Processar "compila_ea_inicial.csv"
                    with z.open("compila_ea_inicial.csv") as f_ear:
                        df_ear = pd.read_csv(f_ear, sep=';')
                        df_ear = df_ear.loc[:, ~df_ear.columns.str.contains('^Unnamed')]
                        df_ear[['Mes', 'Semana']] = df_ear['Deck'].str.split('-', expand=True)
                        df_ear = df_ear[df_ear['Deck'].str.endswith('s1')]  # Filtrar onde Deck termina com 's1'
                        df_ear.drop(['MEN=0-SEM=1', 'Sensibilidade', 'Deck', 'Semana'], axis=1, inplace=True)
                        df_ear.set_index('Mes', inplace=True)
                        df_ear.rename(columns={
                            'SUDESTE': 'EAR SE/CO',
                            'SUL': 'EAR SUL',
                            'NORDESTE': 'EAR NE',
                            'NORTE': 'EAR NO'
                        }, inplace=True)

                        # Fazer o melt para transformar os submercados EAR em formato longo
                        df_ear_melted = df_ear.reset_index().melt(
                            id_vars=['Mes'],
                            value_vars=['EAR SE/CO', 'EAR SUL', 'EAR NE', 'EAR NO'],
                            var_name='Variável',
                            value_name='Valor'
                        )

                        # Adicionar as colunas 'Ano', 'Data de Compilação' e 'Nome do Arquivo' ao df_ear_melted
                        df_ear_melted['Ano'] = datetime.now().year
                        df_ear_melted['Data de Compilação'] = current_date
                        df_ear_melted['Nome do Arquivo'] = zip_file.name

                    # Processar "compila_ena_mensal_percentual.csv"
                    with z.open("compila_ena_mensal_percentual.csv") as f_ena:
                        df_ena = pd.read_csv(f_ena, sep=';')
                        df_ena = df_ena.loc[:, ~df_ena.columns.str.contains('^Unnamed')]
                        df_ena[['Mes', 'Semana']] = df_ena['Deck'].str.split('-', expand=True)
                        df_ena.drop(['MEN=0-SEM=1', 'Sensibilidade', 'Deck', 'Semana'], axis=1, inplace=True)
                        df_ena.set_index('Mes', inplace=True)
                        df_ena.rename(columns={
                            'SUDESTE': 'ENA SE/CO',
                            'SUL': 'ENA SUL',
                            'NORDESTE': 'ENA NE',
                            'NORTE': 'ENA NO'
                        }, inplace=True)

                        # Fazer o melt para transformar os submercados ENA em formato longo
                        df_ena_melted = df_ena.reset_index().melt(
                            id_vars=['Mes'],
                            value_vars=['ENA SE/CO', 'ENA SUL', 'ENA NE', 'ENA NO'],
                            var_name='Variável',
                            value_name='Valor'
                        )

                        # Adicionar as colunas 'Ano', 'Data de Compilação' e 'Nome do Arquivo' ao df_ena_melted
                        df_ena_melted['Ano'] = datetime.now().year
                        df_ena_melted['Data de Compilação'] = current_date
                        df_ena_melted['Nome do Arquivo'] = zip_file.name

                    # Concatenar df_melted (PLD), df_ear_melted (EAR), e df_ena_melted (ENA)
                    df_final = pd.concat([df_melted.round(2), df_ear_melted.round(2), df_ena_melted.round(2)], ignore_index=True)

                    # Adicionar o DataFrame final à lista de resultados
                    results.append(df_final)

                else:
                    st.warning(
                        f"Arquivos 'compila_cmo_medio.csv', 'compila_ea_inicial.csv' ou 'compila_ena_mensal_percentual.csv' não encontrados em {zip_file.name}")

        # Concatenar todos os resultados e exibir
        if results:
            result_df = pd.concat(results, ignore_index=True)
            st.write("Dados compilados:")
            st.dataframe(result_df)
        else:
            st.warning("Nenhum dado processado.")


    # --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    def compila_estudo_matriz(zip_files):

        month_dict = {
            1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL", 5: "MAIO", 6: "JUNHO",
            7: "JULHO", 8: "AGOSTO", 9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"}
        resultados = []  # Para armazenar os resultados finais

        # Iterar sobre todos os arquivos .zip carregados
        for zip_file in zip_files:
            # Abrir o arquivo zip
            with zipfile.ZipFile(zip_file, 'r') as zf:
                # COMPILA EAR
                df_ear = pd.read_csv(zf.open('compila_ea_inicial.csv'), sep=';')
                df_ear.drop('MEN=0-SEM=1', axis=1, inplace=True)
                deck = df_ear.iloc[0, 1]  # Pegar a linha do Deck
                ano = deck[2:6]
                mes = deck[6:8]
                nome_do_mes = month_dict[int(mes)]

                # COMPILA CMO
                df_cmo = pd.read_csv(zf.open('compila_cmo_medio.csv'), sep=';')
                df_cmo = df_cmo.loc[:, ~df_cmo.columns.str.contains('^Unnamed')]
                df_cmo.drop('MEN=0-SEM=1', axis=1, inplace=True)

                # COMPILA TH
                df_th = pd.read_csv(zf.open('compila_ena_th_percentual_sse.csv'), sep=';')
                df_th.drop('MEN=0-SEM=1', axis=1, inplace=True)

                th_se = str(round(df_th.iloc[1, 2], 1))
                th_s = str(round(df_th.iloc[1, 3], 1))
                th_ne = str(round(df_th.iloc[1, 4], 1))
                th_no = str(round(df_th.iloc[1, 5], 1))

                # COMPILA EAR
                ear_se = str(round(df_ear.iloc[0, 2], 1))
                ear_s = str(round(df_ear.iloc[0, 3], 1))
                ear_ne = str(round(df_ear.iloc[0, 4], 1))
                ear_no = str(round(df_ear.iloc[0, 5], 1))

                estudo = f"EAR {ear_se}_{ear_s}_{ear_ne}_{ear_no} - TH {th_se}_{th_s}_{th_ne}_{th_no}"

                # FILTRA OS VALORES ABAIXO DO PISO (valor mínimo é 61)
                df_cmo.loc[df_cmo['SUDESTE'] < 61, 'SUDESTE'] = 61
                df_cmo.loc[df_cmo['SUL'] < 61, 'SUL'] = 61
                df_cmo.loc[df_cmo['NORDESTE'] < 61, 'NORDESTE'] = 61
                df_cmo.loc[df_cmo['NORTE'] < 61, 'NORTE'] = 61

                # Remove os valores de NW
                df_cmo = df_cmo[~df_cmo['Deck'].astype(str).str.startswith('N')]

                # Cálculo da média por sensibilidade
                df_final = df_cmo.groupby('Sensibilidade').mean().round(0)
                df_final.columns = ["SE/CO", "SUL", "NE", "NO"]
                df_final.sort_values(by=['SE/CO'], ascending=False, inplace=True)
                df_final.reset_index(inplace=True)

                # Adicionar as colunas de informação
                df_final.insert(0, "ano", ano)
                df_final.insert(1, "mes", nome_do_mes)
                df_final.insert(2, "data de compilacao", date.today().strftime('%d/%m/%Y'))
                df_final.insert(3, "Estudo", estudo)
                df_final.insert(4, "Arquivo", zip_file.name)

                # Adicionar o resultado ao acumulador
                resultados.append(df_final)

        # Concatenar todos os resultados em um único DataFrame
        if resultados:
            df_resultados = pd.concat(resultados, ignore_index=True)
            return df_resultados
        else:
            return None


    # ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    # Função para processar arquivo 'relato.rvX'
    def process_relato_file(file):
        data = []
        start_string = "1  CAMARGOS     #"
        end_string = "314  PIMENTAL     #"
        capture = False

        for line in file:
            if start_string in line:
                capture = True  # Iniciar captura
            if capture:
                n_uhe = line[4:8].strip()  # N UHE
                nome_uhe = line[9:25].strip().replace('#', '').replace('@', '')  # Nome UHE
                ear_ini = line[27:33].strip()  # EAR ini
                ear_fin = line[33:39].strip()  # EAR fin
                qdef = line[71:81].strip()  # Qdef
                data.append([n_uhe, nome_uhe, ear_ini, ear_fin, qdef])

            if end_string in line:
                capture = False  # Parar captura
                break

        return data


    # Função para processar despacho térmico
    def process_thermal_dispatch(file):
        data = []
        start_string = "RELATORIO  DA  OPERACAO  TERMICA E CONTRATOS"
        end_string = "Usina termica GNL com despacho definido anteriormente."
        capture = False
        subsistemas = ["    SE ", "    S ", "    NE ", "    N "]  # Subsistemas esperados

        for line in file:
            if start_string in line:
                capture = True  # Iniciar captura
                continue  # Pular linha de início
            if end_string in line:
                capture = False  # Parar captura
                break

            if capture:
                if any(line.startswith(subsistema) for subsistema in subsistemas):
                    subsistema = line[4:6].strip()  # Subsistema
                    nome_usina = line[8:19].strip()  # Nome usina
                    d1 = line[28:38].strip()  # D1
                    d2 = line[40:50].strip()  # D2
                    d3 = line[52:62].strip()  # D3

                    if d1 and d2 and d3:
                        media = round((float(d1) + float(d2) + float(d3)) / 3, 1)
                        data.append([subsistema, nome_usina, media])

        return data


    # Função principal para processar arquivos dentro do arquivo zipado
    def process_zip_files(zip_file):
        ear_ini_data = []
        ear_fin_data = []
        qdef_data = []
        thermal_data = []
        filenames = []

        # Abrir o arquivo zip principal
        with zipfile.ZipFile(zip_file, 'r') as z:
            # Percorrer os arquivos dentro do zip
            for file_name in z.namelist():
                if file_name.startswith("DC") and file_name.endswith(".zip"):
                    filenames.append(file_name.replace('.zip', ''))  # Adiciona o nome do arquivo
                    with z.open(file_name) as dc_zip:
                        with zipfile.ZipFile(dc_zip, 'r') as dc_z:
                            # Filtrar para processar arquivos que terminam com .rv0, .rv1, .rv2, .rv3, .rv4
                            relato_files = [f for f in dc_z.namelist() if f.endswith(
                                ('relato.rv0', 'relato.rv1', 'relato.rv2', 'relato.rv3', 'relato.rv4'))]

                            # Processar apenas os arquivos relato com as extensões especificadas
                            for relato_file in relato_files:
                                with dc_z.open(relato_file) as relato:
                                    relato_text = TextIOWrapper(relato, encoding='utf-8')

                                    # Processar despacho térmico
                                    thermal_dispatch = process_thermal_dispatch(relato_text)

                                    for idx, row in enumerate(thermal_dispatch):
                                        if len(thermal_data) <= idx:
                                            thermal_data.append([row[0], row[1], row[2]])
                                        else:
                                            thermal_data[idx].append(row[2])

                                    relato.seek(0)  # Resetar a posição do cursor

                                    # Processar EAR ini, EAR fin e Qdef
                                    file_data = process_relato_file(relato_text)

                                    for idx, row in enumerate(file_data):
                                        if len(ear_ini_data) <= idx:
                                            ear_ini_data.append([row[0], row[1], row[2]])  # EAR ini
                                            ear_fin_data.append([row[0], row[1], row[3]])  # EAR fin
                                            qdef_data.append([row[0], row[1], row[4]])  # Qdef
                                        else:
                                            ear_ini_data[idx].append(row[2])
                                            ear_fin_data[idx].append(row[3])
                                            qdef_data[idx].append(row[4])

        # Verificar se o número de colunas corresponde ao número de arquivos processados
        num_columns = len(filenames) + 2  # 2 colunas adicionais para "N UHE" e "Nome UHE"

        # Ajustar ear_ini_data para garantir que todas as linhas tenham o número correto de colunas
        for i in range(len(ear_ini_data)):
            if len(ear_ini_data[i]) < num_columns:
                # Preencher com NaN para garantir que o número de colunas seja correto
                ear_ini_data[i] += [np.nan] * (num_columns - len(ear_ini_data[i]))

        for i in range(len(ear_fin_data)):
            if len(ear_fin_data[i]) < num_columns:
                ear_fin_data[i] += [np.nan] * (num_columns - len(ear_fin_data[i]))

        for i in range(len(qdef_data)):
            if len(qdef_data[i]) < num_columns:
                qdef_data[i] += [np.nan] * (num_columns - len(qdef_data[i]))

        # Criar DataFrames a partir das listas de dados
        ear_ini_df = pd.DataFrame(ear_ini_data, columns=["N UHE", "Nome UHE"] + filenames)
        ear_fin_df = pd.DataFrame(ear_fin_data, columns=["N UHE", "Nome UHE"] + filenames)
        qdef_df = pd.DataFrame(qdef_data, columns=["N UHE", "Nome UHE"] + filenames)
        thermal_df = pd.DataFrame(thermal_data, columns=["Subsistema", "Nome Usina"] + filenames)

        # Remover linhas com valores em branco em EAR ini e EAR fin
        ear_ini_df = ear_ini_df.replace('', pd.NA).dropna(subset=filenames, how='all')
        ear_fin_df = ear_fin_df.replace('', pd.NA).dropna(subset=filenames, how='all')

        return ear_ini_df, ear_fin_df, qdef_df, thermal_df

        # Função para salvar os DataFrames em um arquivo Excel com múltiplas abas


    def save_to_excel(ear_ini_df, ear_fin_df, qdef_df, thermal_df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            ear_ini_df.to_excel(writer, sheet_name='EAR ini', index=False)
            ear_fin_df.to_excel(writer, sheet_name='EAR fin', index=False)
            qdef_df.to_excel(writer, sheet_name='Qdef', index=False)
            thermal_df.to_excel(writer, sheet_name='Despacho Termico', index=False)
        output.seek(0)
        return output


    # ------------------------------------------------------------------------------------------------------------------------------------------------------
    # Interface do Streamlit

    # Menu lateral para escolher entre os scripts
    script_selected = st.sidebar.selectbox(
        "Escolha um script para executar",
        ["Gerar Prevs Sensibilidade Matriz", "Compila Estudo LP", "Compila Estudo Matriz", "Compila Relato DC", "Portfólio"]
    )

    # Se o script selecionado for "Gerar Prevs Sensibilidade Matriz"
    if script_selected == "Gerar Prevs Sensibilidade Matriz":
        st.header("Gerar Prevs Sensibilidade Matriz")

        # Caixa de upload para o arquivo .txt contendo os dados
        uploaded_file = st.file_uploader("Carregue o arquivo de cenários", type="dat")

        # Carregar o arquivo .rv0, .rv1, .rv2, .rv3 ou .rv4
        rv_file = st.file_uploader("Carregue o arquivo .rv0, .rv1, .rv2, .rv3 ou .rv4",
                                   type=["rv0", "rv1", "rv2", "rv3", "rv4"])

        # Seleção do ano atual até dois anos no futuro
        current_year = pd.Timestamp.now().year
        years = [str(current_year + i) for i in range(3)]
        selected_year = st.selectbox("Selecione o ano", years)

        # Seleção do mês (em extenso, com a primeira letra maiúscula)
        months = [
            "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
            "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
        ]
        selected_month = st.selectbox("Selecione o mês", months)

        # Pasta selecionada para salvar os arquivos gerados
        folder_selected = st.text_input("Insira o diretório para salvar os arquivos gerados:")

        if uploaded_file is not None and rv_file is not None and folder_selected:
            # Ler o conteúdo do arquivo .txt de upload
            lines = uploaded_file.read().decode("utf-8").splitlines()

            # Ler o conteúdo do arquivo .rvx
            rv_file_content = rv_file.read().decode("utf-8")

            # Botão para processar o arquivo
            if st.button("Processar e Executar Macro"):
                try:
                    # Processar o conteúdo e executar a macro
                    processar_e_executar_macro(lines, rv_file_content, selected_year, selected_month, folder_selected)
                    st.success("Execução concluída com sucesso!")
                except Exception as e:
                    st.error(f"Erro durante a execução: {str(e)}")
        else:
            if not uploaded_file:
                st.warning("Carregue um arquivo .txt.")
            if not rv_file:
                st.warning("Carregue um arquivo .rv0, .rv1, .rv2, .rv3 ou .rv4.")
            if not folder_selected:
                st.warning("Insira um diretório para salvar os arquivos.")

    # Se o script selecionado for "Compila Estudo LP"
    elif script_selected == "Compila Estudo LP":
        st.header("Compila Estudo LP")

        # Caixa de upload para selecionar múltiplos arquivos .zip
        zip_files = st.file_uploader("Carregue os arquivos .zip", type="zip", accept_multiple_files=True)

        # Botão para compilar
        if st.button("Compilar Estudos"):
            if zip_files:
                # Executar a função que lida com os arquivos zip
                compila_estudo_lp(zip_files)
            else:
                st.warning("Nenhum arquivo .zip foi selecionado.")

    # Se o script selecionado for "Compila Estudo Matriz"
    elif script_selected == "Compila Estudo Matriz":
        st.title("Compila Estudo Matriz")

        # Caixa de upload para selecionar múltiplos arquivos .zip
        zip_files_matriz = st.file_uploader("Carregue os arquivos .zip para Matriz", type="zip",
                                            accept_multiple_files=True)

        # Botão para compilar
        if st.button("Compilar"):
            if zip_files_matriz:
                # Chamar a função que processa os arquivos .zip
                df_resultado_matriz = compila_estudo_matriz(zip_files_matriz)

                if df_resultado_matriz is not None:
                    # Exibir o DataFrame final
                    st.write("Resultados da compilação:")
                    st.dataframe(df_resultado_matriz)
                else:
                    st.warning("Nenhum dado foi processado.")
            else:
                st.warning("Nenhum arquivo .zip foi selecionado.")


    elif script_selected == "Compila Relato DC":
        st.title("Compila Relato DC")

        # Upload de arquivo .zip
        uploaded_file = st.file_uploader("Carregue o arquivo .zip", type="zip")

        # Botão para processar
        if st.button("Processar Relato DC"):
            if uploaded_file:
                # Obter o nome do arquivo zip carregado
                file_name = uploaded_file.name

                # Extrair o número entre "_" do nome do arquivo
                match = re.search(r"Estudo_(\d+)_Compilacao.zip", file_name)
                if match:
                    number = match.group(1)
                else:
                    st.error("O nome do arquivo não segue o padrão 'Estudo_XXXXX_Compilacao.zip'.")
                    st.stop()

                # Processar o arquivo zipado
                ear_ini_df, ear_fin_df, qdef_df, thermal_df = process_zip_files(uploaded_file)

                # Salvar o arquivo em memória
                excel_data = save_to_excel(ear_ini_df, ear_fin_df, qdef_df, thermal_df)

                # Nome do arquivo final
                output_file_name = f"Relato_{number}_Compilacao.xlsx"

                # Botão de download para o arquivo Excel
                st.download_button(
                    label="Baixar arquivo Excel",
                    data=excel_data,
                    file_name=output_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Exibir os DataFrames resultantes
                st.write("EAR Ini DataFrame")
                st.dataframe(ear_ini_df)

                st.write("EAR Fin DataFrame")
                st.dataframe(ear_fin_df)

                st.write("Qdef DataFrame")
                st.dataframe(qdef_df)

                st.write("Despacho Térmico DataFrame")
                st.dataframe(thermal_df)
            else:
                st.warning("Por favor, carregue um arquivo .zip.")


    elif script_selected == "Portfólio":
        st.header("Rodar Portfólio e calcular MTM")
        if st.button("Rodar Portfólio"):

            rodar_portfolio_mtm()