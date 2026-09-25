# Padrão dos manuais G-Code 3D

Referência: `g-code-09-papelarias.html`. Todo manual novo copia o CSS, os componentes e o script de ilustrações dele e troca só o conteúdo.

## Estrutura

Capa (código, mercado, promessa, radar com 5 números) + 12 seções + Cola de balcão.

01 Por que funciona · 02 Quem é a cliente · 03 A dor · 04 Oferta em três níveis · 05 Material e execução · 06 Kit de amostra · 07 Precificação · 08 Abordagem de 40 segundos · 09 Seis objeções · 10 Follow-up · 11 As três perguntas · 12 Missão 72 horas

**Cola de balcão** (anexo, uma página impressa): o que levar na mão, os 40 segundos, as saídas, as 6 objeções em uma linha cada, os números de preço e o bloco "Nunca".

## Imagens

Sem placeholders. Ilustrações em SVG geradas por JavaScript no fim do arquivo (`data-ill="chave"`), na paleta da marca, com peças em verde escuro. De 5 a 7 por manual, cada uma mostrando uma informação do texto (escala, processo, comparação), não decoração. Não desenhar pessoas.

## Regras de texto

- Não afirmar que um material ou item é mais barato ou mais caro que outro.
- Proibido: "atóxico", "certificado", "aprovado para alimento".
- Todo número é hipótese a validar (rodapé).
- Bloco `.aviso` explícito onde houver risco real.
