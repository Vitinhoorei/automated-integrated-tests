from __future__ import annotations

def _clean_text(text: str, max_len: int | None = None) -> str:
    """
    Normaliza texto:
    - remove quebras de linha
    - remove espaços duplicados
    - aplica limite de tamanho se definido
    """
    txt = " ".join((text or "").split())
    if max_len and len(txt) > max_len:
        return txt[: max_len - 3] + "..."
    return txt


def build_status(
    status: str,
    source: str = "",
    message: str = "",
    max_len: int = 180,
) -> str:
    """
    MANTIDO por compatibilidade.
    Constrói a string única de status no formato antigo:

    FAIL | STATUSBAR | Campo Centro é obrigatório
    """

    status_u = (status or "").strip().upper()
    source_u = (source or "").strip().upper()
    msg = _clean_text(message, max_len=max_len)

    parts: list[str] = []
    if status_u:
        parts.append(status_u)
    if source_u:
        parts.append(source_u)
    if msg:
        parts.append(msg)

    return " | ".join(parts)


def build_status_fields(
    *,
    status: str,
    source: str,
    message: str,
    max_message_len: int = 180,
) -> dict[str, str]:
    """
    NOVO MÉTODO – USO PADRÃO DAQUI PRA FRENTE

    Retorna os campos JÁ NORMALIZADOS e SEPARADOS
    para escrita direta nas colunas da planilha.

    Regras:
    - status e source sempre em UPPER
    - message NÃO pode repetir status nem source
    - message curta, técnica e limpa
    """

    status_u = (status or "").strip().upper()
    source_u = (source or "").strip().upper()

    msg = _clean_text(message, max_len=max_message_len)

    # Proteção contra duplicação acidental
    if msg:
        msg_low = msg.lower()
        if status_u.lower() in msg_low:
            msg = msg.replace(status_u, "").strip(" |-")
        if source_u.lower() in msg_low:
            msg = msg.replace(source_u, "").strip(" |-")

    return {
        "status": status_u,
        "status_source": source_u,
        "status_message": msg,
    }