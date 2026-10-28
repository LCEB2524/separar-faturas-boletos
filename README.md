# Separar faturas e boletos por cliente

Lê o PDF de **faturamento agrupado** e o PDF de **boletos (Sicredi)** e gera um PDF por cliente com a fatura + os boletos correspondentes.

## Cruzamento
1. Nome do cliente (campo `CLIENTE` da fatura) = `Pagador` do boleto.
2. Nome truncado: o boleto corta o pagador em 40 caracteres.
3. NFs da fatura (colunas `NOTA`) x `Ref. Notas Fiscais` do boleto (ignorando zeros à esquerda).

Cliente sem boleto gera só a fatura, com ` - SEM BOLETO` no nome. Cliente com vários boletos recebe todos no mesmo arquivo.

## Conferência no final
- Total de páginas de entrada x saída.
- Valor da fatura x soma dos boletos.
- Nº da duplicata x nº do documento do boleto.
- NFs da fatura que não constam na referência do boleto.
- Boletos que ficaram sem fatura.

## Uso (Windows)
1. `python -m pip install -r requirements.txt`
2. Coloque esta pasta dentro da pasta dos PDFs (ex.: `conth\separar`). O script procura, na pasta de cima:
   - `Emiss*Faturamento*.pdf`
   - `boleto*.pdf`
3. Dê dois cliques em `SEPARAR.bat`. Os arquivos saem em `saida`.

Ou informe os arquivos direto:
```
python separar_faturas_boletos.py "C:\pasta\faturamento.pdf" "C:\pasta\boletos.pdf"
```

A pasta `saida` é a entrada do [envio-boletos-whats-email](../envio-boletos-whats-email).
