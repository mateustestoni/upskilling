# Sistema de Quests

## Estrutura da aba Quests
Colunas:
- id
- name
- description
- reward_type (xp | mult)
- value
- multiplier
- duration_days
- skill_target
- done

## Comportamento
- Quests ativas: done vazio
- Ao completar:
  - Marca done com data
  - Aplica recompensa
