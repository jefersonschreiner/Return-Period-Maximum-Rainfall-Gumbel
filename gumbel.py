import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import gumbel_r
import pandas as pd
from datetime import datetime
import sys
import os

def selecionar_anos(dados):
    """
    Mostra os anos disponíveis pra seleção da análise.
    """
    anos_disponiveis = dados['Ano'].dropna().unique()
    
    print("\nAnos disponíveis para análise:")
    for i, ano in enumerate(sorted(anos_disponiveis), 1):
        print(f"{i}. {int(ano)}")

    print("\nDigite os números dos anos que deseja incluir (separados por vírgula)")
    print("Ou digite 'todos' para incluir todos os anos")
    selecao = input("Sua seleção: ").strip()
    
    if selecao.lower() == 'todos':
        return dados
    else:
        try:
            indices = [int(i.strip()) - 1 for i in selecao.split(',')]
            anos_selecionados = [sorted(anos_disponiveis)[i] for i in indices]
            dados_filtrados = dados[dados['Ano'].isin(anos_selecionados)]
            print(f"\nAnos selecionados: {', '.join(map(str, anos_selecionados))}")
            return dados_filtrados
        except (ValueError, IndexError):
            print("\nSeleção inválida. Serão incluídos todos os anos por padrão.")
            return dados

def analise_recorrencia(totais_anuais, output_prefix=''):
    """
    Realiza a análise de tempo de recorrência com os totais anuais.
    Retorna um dicionário com todos os resultados estatísticos.
    """
    # Cálculos estatísticos
    media = np.mean(totais_anuais)
    desvio_padrao = np.std(totais_anuais, ddof=1)
    print(f"\nAnálise de Tempo de Recorrência (Totais Anuais):")
    print(f"Média dos totais anuais: {media:.2f} mm")
    print(f"Desvio padrão: {desvio_padrao:.2f} mm")
    
    # Ordenação (Weibull)
    n = len(totais_anuais)
    totais_ordenados = np.sort(totais_anuais)[::-1]
    m = np.arange(1, n + 1)
    P = m / (n + 1)
    T_empirico = 1 / P
    
    # Ajuste da distribuição de Gumbel
    loc, scale = gumbel_r.fit(totais_anuais)
    print(f"Parâmetros de Gumbel: loc = {loc:.2f}, scale = {scale:.2f}")
    
    # Cálculo do Tempo de Recorrência
    T_alvo = np.array([2, 5, 10, 25, 50, 100, 1000, 10000])
    P_alvo = 1 / T_alvo
    Q_T = gumbel_r.ppf(1 - P_alvo, loc=loc, scale=scale)
    
    # Gráfico dos Resultados
    plt.figure(figsize=(10, 6))
    plt.scatter(T_empirico, totais_ordenados, color='red', label='Dados Empíricos')
    plt.plot(T_alvo, Q_T, marker='o', color='blue', label='Distribuição de Gumbel')
    plt.xscale('log')
    plt.xlabel('Tempo de Recorrência (anos)')
    plt.ylabel('Total Anual de Chuva (mm)')
    plt.title('Tempo de Recorrência de Totais Anuais de Chuva')
    plt.grid(True, which="both", ls="--")
    plt.legend()
    plt.savefig(f'{output_prefix}grafico_recorrencia.png', dpi=300)
    plt.show()
    
    # Tabela de resultados
    fig, ax = plt.subplots(figsize=(8, 3))
    ax.axis('off')
    tabela_data = [["T (anos)", "Total Anual (mm)"]] + [[f"{T}", f"{Q:.1f}"] for T, Q in zip(T_alvo, Q_T)]
    tabela = ax.table(cellText=tabela_data, loc='center', cellLoc='center')
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(12)
    tabela.scale(1.2, 1.2)
    plt.title("Tempos de Recorrência - Totais Anuais", pad=20)
    plt.savefig(f'{output_prefix}tabela_recorrencia.png', bbox_inches='tight', dpi=300)
    plt.show()
    
    return {
        'media': media,
        'desvio_padrao': desvio_padrao,
        'n_amostras': n,
        'parametros_gumbel': {'loc': loc, 'scale': scale},
        'tempos_recorrencia': T_alvo,
        'totais_estimados': Q_T,
        'totais_observados': totais_anuais
    }

def separar_dados_por_ano(arquivo_entrada, arquivo_saida='DadosCompletos.xlsx'):
    """
    Processa os dados de chuva e realiza análises completas.
    """
    try:
        # Verificar extensão do arquivo de saída
        if not arquivo_saida.lower().endswith('.xlsx'):
            arquivo_saida += '.xlsx'
        
        # Ler os dados originais
        dados = pd.read_excel(arquivo_entrada, sheet_name='Dados')
        
        # Verificar colunas necessárias
        if 'Data' not in dados.columns or 'Total' not in dados.columns:
            raise ValueError("O arquivo deve conter colunas 'Data' e 'Total'")
        
        # Extrair ano e filtrar dados
        dados['Ano'] = pd.to_datetime(dados['Data']).dt.year
        dados_filtrados = selecionar_anos(dados)
        
        # Criar resumo anual
        resumo = dados_filtrados.groupby('Ano').agg({
            'Total': ['sum', 'max', 'count']
        })
        resumo.columns = ['Total Anual (mm)', 'Chuva Máxima (mm)', 'Meses com Dados']
        resumo = resumo.reset_index()
        
        # Análise de tempo de recorrência
        resultados = analise_recorrencia(resumo['Total Anual (mm)'].values, 'recorrencia_')
        
        # Salvar todos os resultados em Excel
        with pd.ExcelWriter(arquivo_saida, engine='xlsxwriter') as writer:
            # 1. Aba de Resumo Anual
            resumo.to_excel(writer, sheet_name='Resumo Anual', index=False)
            
            # 2. Aba de Estatísticas Descritivas
            estatisticas = pd.DataFrame({
                'Estatística': ['Média (mm)', 'Desvio Padrão (mm)', 'Número de Anos', 
                               'Parâmetro Loc (Gumbel)', 'Parâmetro Scale (Gumbel)'],
                'Valor': [resultados['media'], resultados['desvio_padrao'], resultados['n_amostras'],
                          resultados['parametros_gumbel']['loc'], resultados['parametros_gumbel']['scale']]
            })
            estatisticas.to_excel(writer, sheet_name='Estatísticas', index=False)
            
            # 3. Aba de Análise de Recorrência
            pd.DataFrame({
                'Tempo de Recorrência (anos)': resultados['tempos_recorrencia'],
                'Total Anual Estimado (mm)': resultados['totais_estimados']
            }).to_excel(writer, sheet_name='Análise Recorrência', index=False)
            
            # 4. Aba com Dados Completos para Análise
            pd.DataFrame({
                'Ano': resumo['Ano'],
                'Total Anual Observado (mm)': resumo['Total Anual (mm)'],
                'Total Anual Ordenado (mm)': np.sort(resumo['Total Anual (mm)'])[::-1],
                'Tempo de Recorrência Empírico (anos)': 1 / np.arange(1, len(resumo) + 1) * (len(resumo) + 1)
            }).to_excel(writer, sheet_name='Dados para Análise', index=False)
            
            # 5. Abas por ano com dados diários
            for ano in sorted(dados_filtrados['Ano'].unique()):
                dados_ano = dados_filtrados[dados_filtrados['Ano'] == ano][['Data', 'Total']]
                dados_ano.to_excel(writer, sheet_name=str(ano), index=False)
            
            # Formatação
            workbook = writer.book
            formato_num = workbook.add_format({'num_format': '#,##0.00'})
            formato_geral = workbook.add_format({'align': 'center'})
            
            for sheet in writer.sheets:
                worksheet = writer.sheets[sheet]
                if sheet == 'Resumo Anual':
                    worksheet.set_column('A:A', 10, formato_geral)
                    worksheet.set_column('B:D', 18, formato_num)
                elif sheet == 'Estatísticas':
                    worksheet.set_column('A:A', 25)
                    worksheet.set_column('B:B', 20, formato_num)
                elif sheet == 'Análise Recorrência':
                    worksheet.set_column('A:B', 25, formato_num)
                elif sheet == 'Dados para Análise':
                    worksheet.set_column('A:A', 10, formato_geral)
                    worksheet.set_column('B:D', 25, formato_num)
                else:  # Abas por ano
                    worksheet.set_column('A:A', 15)
                    worksheet.set_column('B:B', 15, formato_num)
        
        print(f"\nArquivo '{arquivo_saida}' criado com sucesso contendo:")
        print("- Resumo anual")
        print("- Estatísticas descritivas")
        print("- Análise de tempos de recorrência") 
        print("- Dados completos para análise")
        print("- Dados diários por ano")
        print("\nGráficos salvos como:")
        print("- 'recorrencia_grafico.png' (Gráfico de recorrência)")
        print("- 'recorrencia_tabela.png' (Tabela de recorrência)")
        
        return True
    
    except Exception as e:
        print(f"\nErro durante o processamento: {str(e)}")
        return False

def main():
    print("=== ESTUDO DE GUMBEL DE MÁXIMAS (CHUVA ANUAL) ===")
    print("Este programa realiza:\n"
          "1. Separação dos dados por ano\n"
          "2. Cálculo de totais anuais\n"
          "3. Análise estatística completa\n"
          "4. Estudo de tempos de recorrência\n")
    
    arquivo_entrada = input("Arquivo de entrada (padrão: DadosChuva.xlsx): ") or 'DadosChuva.xlsx'
    arquivo_saida = input("Arquivo de saída (padrão: AnaliseCompleta.xlsx): ") or 'AnaliseCompleta.xlsx'
    
    # Corrigir extensão se necessário
    if not arquivo_saida.lower().endswith('.xlsx'):
        arquivo_saida = os.path.splitext(arquivo_saida)[0] + '.xlsx'
    
    if separar_dados_por_ano(arquivo_entrada, arquivo_saida):
        print("\nProcessamento concluído com sucesso!")
    else:
        print("\nOcorreram erros durante o processamento.")

if __name__ == "__main__":
    main()