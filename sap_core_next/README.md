# sap_core_next

Nova base paralela (independente) para evolução da automação SAP com **core agnóstico** e arquitetura extensível por plugins.

> Esta pasta não altera nem remove a implementação legada.

## Objetivos

- Core sem conhecimento de TCODE específico.
- Regras de negócio desacopladas em plugins de módulo SAP.
- Integração SAP atrás de adapter/port.
- Planilha apenas como entrada/saída (sem lógica de negócio).
- IA opcional e governada por política.
- Retry/auto-heal previsível, configurável e auditável.

## Estrutura

```text
sap_core_next/
  config/default.yaml
  scripts/run_engine.py
  src/sap_core_next/
    core/
      engine.py
      models.py
      context.py
      policies.py
      logging.py
    ports/
      scenario_source.py
      result_sink.py
      sap_gateway.py
      ai_assistant.py
      evidence_store.py
      plugin.py
    adapters/
      sap/
        legacy_bridge_gateway.py
        win32com_gateway.py
      spreadsheet/
        openpyxl_source.py
        openpyxl_sink.py
      ai/
        noop_ai.py
        http_ai_assistant.py
    services/
      evidence_service.py
    registry/
      plugin_registry.py
    plugins/
      pm/plugin.py
    config/settings.py
  tests/
```

## Diagrama textual

```text
[Spreadsheet Source] ---> [ExecutionEngine (Core)] ---> [SAP Gateway Adapter]
         |                        |                           |
         |                        |                           +--> Win32COM (real)
         |                        |                           +--> Legacy Bridge (compat)
         |                        |
         |                        +--> [Plugin Registry] ---> [PM Plugin]
         |                        |
         |                        +--> [Retry Policy + AI Policy]
         |                        |
         |                        +--> [Evidence Store]
         |
[Spreadsheet Sink] <----- [Execution Report + Structured Logs]
```

## Como executar (exemplo)

```bash
python sap_core_next/scripts/run_engine.py \
  --file "caminho/planilha.xlsx" \
  --sheet "Testes"
```

## Plugin inicial implementado

- `PMPlugin`: normalização de comando, enriquecimento de parâmetros por contexto e memória de IDs em sucesso.

## Estratégia de migração

Veja o arquivo `MIGRATION_PLAN.md`.

## Trade-offs de design

- Foi escolhido um **esqueleto mínimo funcional** para reduzir risco e evitar over-engineering.
- Para manter suporte SAP real sem mexer no legado, foi criado `LegacyBridgeGateway`.
- `Win32ComSapGateway` está implementado como base mínima; fluxos complexos devem ser evoluídos nos plugins.
- IA padrão configurada para no-op no bootstrap inicial; ativação pode ser feita via adapter HTTP.
