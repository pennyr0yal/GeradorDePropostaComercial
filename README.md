# Gerador de Proposta Comercial em `.docx`
Este projeto gera propostas comerciais personalizadas a partir de uma base de dados com clientes e produtos. Utiliza `pandas` para leitura e agrupamento dos dados e `docxtpl` para renderização dos documentos Word.

---

## Funcionalidades

- Lê uma planilha `.xlsx` contendo informações de clientes, produtos, preços, quantidades e descontos.
- Agrupa os produtos por cliente e gera um documento Word personalizado para cada um, utilizando um modelo .docx.
- Realiza cálculos de:
  - Total bruto por item e por cliente
  - Valor total de desconto
  - Valor líquido final
- Cria uma pasta para cada cliente e salva os arquivos com nomes organizados.
- Exibe uma interface gráfica ao final do processo, com opção de abrir a pasta de saída.

## Como Usar

1. Os dados para teste estão no arquivo `clientes.xlsx` e o template da proposta em `template_proposta.docx`.
2. Execute o arquivo .bat na pasta principal. O programa instalará o venv e as bibliotecas necessárias.
3. O programa processará os dados e criará os arquivos `.docx` na pasta Propostas.
4. Ao final, será exibida uma janela informando que as propostas foram geradas, com opção de abrir os arquivos diretamente.

## Observações

- Esse projeto simula um processo real de geração de documentos padronizados, comum em áreas comerciais, de vendas e atendimento B2B.
- O programa foi projetado para ser facilmente adaptável a outros tipos de documentos, como orçamentos, contratos ou recibos.
- A automação lida com múltiplos registros, garantindo economia de tempo em ambientes com alto fluxo de clientes ou produtos.
- Possíveis melhorias incluem integrar o programa a sistemas maiores, como ERPs ou CRMs, e adicionar o envio automático das propostas por e-mail.
  
# Autora
Desenvolvido por Natalia Junghans

📧 natbjunghans@gmail.com
