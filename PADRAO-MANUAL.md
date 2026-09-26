# Padrão dos manuais G-Code 3D

Referência: `g-code-09-papelarias.html`. Todo manual novo copia o CSS, os componentes e o script de ilustrações dele e troca só o conteúdo.

## Estrutura

Capa: logo LR pequena no canto superior esquerdo, logo grande em marca d'água (verde a 7%), código com listras de camada, mercado, promessa, faixa Dia 1 / Dia 2 / Dia 3 / 72h e radar com 5 números. Sem ilustração e sem animação (entrega final em PDF).

Depois da capa: + 12 seções + Cola de balcão.

01 Por que funciona · 02 Quem é a cliente · 03 A dor · 04 Oferta em três níveis · 05 Material e execução · 06 Kit de amostra · 07 Precificação · 08 Abordagem de 40 segundos · 09 Seis objeções · 10 Follow-up · 11 As três perguntas · 12 Missão 72 horas

**Blocos práticos dentro das seções** (adaptar ao nicho): 03 sinais no Instagram antes de visitar · 04 calendário de pedidos do nicho · 05 conferência antes de entregar · 06 quanto custa sair para a rua · 07 conta de um pedido real passo a passo e condições comerciais · 10 mensagem depois da primeira entrega · 12 registro de visitas para preencher à mão.

**PDF:** A4, capa e contracapa de página inteira, cada seção começa em página nova; conferir que nenhuma página fique abaixo de ~60% de ocupação.

**Cola de balcão** (anexo, uma página impressa): o que levar na mão, os 40 segundos, as saídas, as 6 objeções em uma linha cada, os números de preço e o bloco "Nunca".

## Imagens

Sem placeholders. Ilustrações em SVG geradas por JavaScript no fim do arquivo (`data-ill="chave"`), na paleta da marca, com peças em verde escuro. De 5 a 7 por manual, cada uma mostrando uma informação do texto (escala, processo, comparação), não decoração. Não desenhar pessoas.

## Regras de texto

- Não afirmar que um material ou item é mais barato ou mais caro que outro.
- Proibido: "atóxico", "certificado", "aprovado para alimento".
- Todo número é hipótese a validar (rodapé).
- Bloco `.aviso` explícito onde houver risco real.
