# Plano de migração legado -> sap_core_next

## Fase 1 (já iniciada)
- Core agnóstico mínimo funcional.
- Ports/interfaces definidos.
- Plugin PM inicial.
- Adapters de planilha e SAP (bridge + real mínimo).
- Políticas de retry/IA + logs estruturados.

## Fase 2
- Expandir plugin PM para cobrir fluxos especiais críticos (IW41/IP41/IP42) como comandos explícitos.
- Introduzir plugin PP usando mesma estratégia de comandos.
- Adicionar validação de contrato de plugin em CI.

## Fase 3
- Migrar escrita de status para sink desacoplado com schema estável.
- Adicionar trilha de auditoria completa do auto-heal (antes/depois).
- Adicionar classificação de erros por tipo.

## Fase 4
- Reduzir dependência do bridge legado: mover fluxos para commands nativos da nova base.
- Manter bridge apenas para fallback.

## Fase 5
- Publicação como repositório independente.
- Pipeline de testes (unit + integração com fakes + smoke SAP real controlado).
