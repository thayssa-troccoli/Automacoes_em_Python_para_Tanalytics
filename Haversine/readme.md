# Verificação de Proximidade entre Coordenadas (Auditoria de Campo)

Este script foi desenvolvido para auditar entrevistas de campo utilizando a distância geográfica entre coordenadas coletadas pelos entrevistadores.
Ele identifica automaticamente pontos muito próximos entre si, ajudando a detectar:

- risco de entrevistas duplicadas
- entrevistas feitas no mesmo local
- possíveis inconsistências ou fraudes
- deslocamentos suspeitos no processo de coleta

⚙️ Como funciona

O código utiliza a fórmula de Haversine para calcular a distância entre dois pontos geográficos (latitude/longitude) em metros.

Etapas:

- Calcula a distância entre todas as combinações possíveis de coordenadas
- Compara cada distância com um limite definido (threshold) — padrão de 100 metros
- Retorna pares de coordenadas que estão dentro do raio permitido
- Exibe quais entrevistas foram realizadas muito próximas umas das outras

🛠️ Tecnologias

- Python
- math (funções trigonométricas usadas pelo Haversine)
