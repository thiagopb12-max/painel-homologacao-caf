# Painel de Homologação — CAF (DM-209196)

**DATAPREV** — Dados para cidadania

Painel gerencial para acompanhamento em tempo real da homologação do Cadastro Nacional da Agricultura Familiar (CAF).

## Como funciona

O painel lê automaticamente o arquivo `CAF.xlsx` e exibe:
- KPIs de progresso da homologação
- Gráficos de distribuição por classificação, categoria e grupo
- Status de cada homologador
- Tabela filtrável com todos os cenários

## Como atualizar os dados

1. Edite o arquivo `CAF.xlsx` com os dados atualizados
2. Faça commit e push para o GitHub:
   ```bash
   git add CAF.xlsx
   git commit -m "Atualiza dados da homologação"
   git push
   ```
3. O GitHub Pages republica automaticamente em ~1 minuto

## Acesso online

Após o deploy, o painel fica disponível em:

`https://thiagopb12-max.github.io/painel-homologacao-caf/`
