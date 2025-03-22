# Gerenciador de Arquivos Excel

Este é um aplicativo em Python que utiliza a biblioteca Tkinter para criar uma interface gráfica para gerenciar arquivos Excel. O programa permite criar, abrir e editar arquivos Excel de forma interativa.

## Tela Inicial

![Tela Inicial](https://github.com/giovanep4mg/Projeto_Caixa/blob/main/imagens/tela%20inicial-programa.PNG)

## Tela Criando um arquivo excel vazio, porém com as colunas já pré-definidas

![Tela Criando um arquivo excel vazio, porém com as colunas já pré-definidas](https://github.com/giovanep4mg/Projeto_Caixa/blob/main/imagens/print-planilha-criada.PNG)

## Tela que abre para o usuário inserir os dados

![Tela que abre para o usuário inserir os dados](https://github.com/giovanep4mg/Projeto_Caixa/blob/main/imagens/print-inserir-dia-na-planilha.PNG)

## Requisitos

Antes de executar o programa, certifique-se de ter instalado os seguintes pacotes Python:

```bash
pip install pandas easygui xlsxwriter
```

## Funcionalidades

- Criar um novo arquivo Excel com colunas predefinidas.
- Abrir um arquivo Excel existente.
- Editar um arquivo Excel preenchendo novos valores financeiros.

## Como os cálculos são feitos

O programa realiza diversas operações matemáticas para atualizar os valores financeiros e gerar novos registros no arquivo Excel. Aqui está um detalhamento das principais operações:

1. **Coleta de dados**
   - O usuário insere os seguintes valores:
     - `Dia`: Número do dia do registro.
     - `Dinheiro Salão`: Valor em dinheiro do salão.
     - `Notas de 2`: Quantidade de notas de R$ 2,00.
     - `Moeda Salão`: Valor total de moedas no salão.
     - `Dinheiro Casa`: Valor total de dinheiro disponível em casa.
     - `Moeda Casa`: Valor total de moedas disponíveis em casa.
     - `Gasto Dia`: Despesas diárias.
     - `Sicoob`, `Sumup`, `Nullbank`, `MercPago`: Valores disponíveis em cada banco.

2. **Cálculos de valores atualizados**
   - `casa = dinhCasa`: Mantém o valor do dinheiro em casa.
   - `novo_moedaCasa = moedaCasa`: Mantém o valor da moeda casa atual.
   - `caixa = dinhSalao + notas2`: Calcula o total de dinheiro físico no salão.
   - `totalbancos = sicoob + sumup + nullbank + mercPago`: Soma dos valores disponíveis nos bancos.

3. **Atualização da Moeda Casa**
   - Se houver um valor anterior para `Moeda/Casa`, ele é somado ao novo valor:
     ```python
     moedaCasa_anterior = df['Moeda/Casa'].iloc[-1] if 'Moeda/Casa' in df.columns and not df.empty else obter_valor_numerico("Digite o valor anterior de Moeda Casa: ")
     moedaCasa_atual = novo_moedaCasa + moedaCasa_anterior
     ```

4. **Cálculo do total geral**
   - `totalsoma = casa + caixa + totalbancos + moedaSalao + moedaCasa_atual`: Soma total de todos os valores disponíveis.
   - O total anterior é recuperado:
     ```python
     total_anterior = df['TotalSoma'].iloc[-1] if 'TotalSoma' in df.columns and not df.empty else obter_valor_numerico("Digite o valor do total anterior: ")
     ```
   - `novo_lucro = totalsoma - total_anterior`: Calcula o lucro do período.
   - `totalDia = novo_lucro + gastoDia`: Atualiza o total considerando o gasto diário.

5. **Registro no DataFrame e salvamento no Excel**
   - O novo conjunto de valores é adicionado como uma nova linha ao arquivo Excel.
   - A planilha é salva e formatada corretamente com `xlsxwriter`.

## Como Executar

Execute o seguinte comando para iniciar a interface gráfica:

```bash
python nome_do_arquivo.py
```

Uma janela será aberta com opções para criar, abrir ou editar um arquivo Excel.

## Licença

Este projeto é de uso livre e pode ser modificado conforme necessário.

