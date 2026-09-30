# Trans RV Portal

Portal web para análise e classificação determinística de viagens, tabelas de frete e minutas TMS.

## Publicação

O projeto é publicado via GitHub Pages:
https://alanmarquesreis1-create.github.io/trans-rv-portal/

## Núcleo atual

- Importação múltipla por arrastar e soltar de XLSX/XLS/CSV.
- Identificação de Amazon, Mercado Livre e Shopee.
- Classificação determinística Amazon: OPB → RTP → LH.
- Classificação determinística Mercado Livre: CPL → DEDICADO → MULTISTOP → LH.
- C195/C195.x → DEDICADO 195; C173/C173.x → DEDICADO 173; C18/C18.x → DEDICADO 18; C14/C14.x → DEDICADO 14.
- Shopee → LINE_HAUL.
- DE/PARA de localidades e tipologias configurável.
- Cadastro de Nodes Amazon.
- Importação do mestre de tabelas de frete e busca de tabela esperada.
- Importação de minutas e pós-validação esperada x TMS.
- Controle de minutas fora da demanda.
- Fila de exceções/tratativas.
- Exportação XLSX para classificação, TMS, minutas e tratativas.
- Dashboard executivo.
- Histórico local de execuções e alterações.

## Arquitetura atual

Esta edição é uma SPA estática executada no navegador. Os arquivos são processados localmente e os cadastros/configurações desta versão ficam no armazenamento local do navegador.

O GitHub Pages hospeda apenas a interface. Não há backend, Firebase, autenticação compartilhada ou banco central nesta versão.

## Próxima evolução

1. Validar as bases reais Week 39/40 no navegador.
2. Completar o motor de matching das tabelas com todos os critérios obrigatórios e vigência/status.
3. Incorporar o cadastro oficial completo de Nodes Amazon.
4. Evoluir pré-validação e pós-validação TMS.
5. Criar persistência central, usuários/permissões e auditoria compartilhada.
6. Adicionar importação/exportação de layouts configuráveis.
7. Evoluir Dashboard Gerencial de Falhas.
8. Homologar com bases reais antes de uso operacional.

## Tecnologia

HTML, CSS, JavaScript e SheetJS para leitura/escrita de planilhas no navegador.