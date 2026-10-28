# Separar faturas, boletos e NFs por cliente

Lê 3 PDFs gerais e gera **um PDF por cliente** com a fatura + boletos + notas fiscais, **somente depois de conferir tudo**.

## Entrada (na pasta de cima, ex.: `conth`)
| Arquivo | Padrão do nome |
|---|---|
| Faturamento agrupado | `Emiss*Faturamento*.pdf` (1 arquivo) |
| Boletos Sicredi | `boleto*.pdf` (1 arquivo) |
| Notas fiscais (DANFE) | `DANFE*.pdf` (1 ou mais arquivos) |

## Conferências (antes de gerar)
**Fatura x Boleto**
- Cliente = Pagador (exato, nome cortado em 40 caracteres ou pelas NFs)
- Valor da fatura = soma dos boletos
- Nº da duplicata = nº do documento do boleto
- NFs da fatura x "Ref. Notas Fiscais" do boleto

**Fatura x NF**
- Cada NF da fatura existe no arquivo de NFs
- Destinatário da NF = cliente da fatura
- Valor da NF = valor da NF na fatura
- Soma das NFs − crédito = VALOR FATURA
- QTDE. NOTAS = NFs lidas

**Sobras**: boletos sem fatura e NFs que não estão em nenhuma fatura.

A conferência aparece na tela e é salva em `saida\Conferencia\conferencia_<data>.xlsx`. Os arquivos só são gerados depois de responder `s`.

## Saída
`saida\<CLIENTE>.pdf` = fatura + boletos + NFs (na ordem da fatura).
Cliente sem boleto: `<CLIENTE> - SEM BOLETO.pdf`. No final confere páginas gravadas x esperadas.

## Uso (Windows)
1. `python -m pip install -r requirements.txt`
2. Coloque esta pasta dentro da pasta dos PDFs (ex.: `conth\separar`).
3. Dê dois cliques em `SEPARAR.bat` (ou `python separar_faturas_boletos.py --sem-perguntar` para gerar sem perguntar).

A pasta `saida` é a entrada do `envio-boletos-whats-email`.
