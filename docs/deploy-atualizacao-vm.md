# Deploy com VM e atualização do LegalHub

Este documento deixa o projeto preparado para o modelo correto:

- o usuário interno instala o app uma única vez;
- a VM publica o manifesto `latest.json` e os pacotes `.zip`;
- as chaves sensíveis ficam somente na VM;
- o app desktop chama a API da VM por `LEGALHUB_API_BASE_URL`.

## 1. Variáveis locais do app desktop

No `.env` local do usuário ou no ambiente do Windows, configure somente dados não sensíveis:

```env
LEGALHUB_UPDATE_URL=https://legalhub.seudominio.com/updates/latest.json
LEGALHUB_API_BASE_URL=https://legalhub.seudominio.com
LEGALHUB_CLIENT_TOKEN=token_interno_opcional
LEGALHUB_API_TIMEOUT=180
```

Não coloque `OPENAI_API_KEY` no computador dos usuários internos.

## 2. Variáveis da VM

Na VM, configure as chaves reais:

```env
OPENAI_API_KEY=sua_chave_real
OPENAI_PROMPT_ID=seu_prompt_id
MODEL_ID=seu_model_id
VERSION_MODEL=22
JWT_SECRET=uma_chave_interna_forte
```

A VM deve expor endpoints próprios para as funcionalidades de IA. O desktop deve enviar PDFs/dados para a VM e receber apenas o resultado.

## 3. Gerar release local

Instale as dependências de build:

```powershell
uv add --dev pyinstaller packaging requests
```

Gere o pacote:

```powershell
uv run python scripts/build_release.py --version 1.0.1 --base-url https://legalhub.seudominio.com
```

O script gera:

```text
releases/LegalHub-1.0.1.zip
updates/latest.json
```

## 4. Subir arquivos para a VM

Exemplo:

```powershell
scp .\releases\LegalHub-1.0.1.zip usuario@IP_DA_VM:/opt/legalhub/releases/
scp .\updates\latest.json usuario@IP_DA_VM:/opt/legalhub/updates/latest.json
```

A VM precisa servir estes caminhos:

```text
https://legalhub.seudominio.com/releases/LegalHub-1.0.1.zip
https://legalhub.seudominio.com/updates/latest.json
```

## 5. Formato do latest.json

```json
{
    "version": "1.0.1",
    "url": "https://legalhub.seudominio.com/releases/LegalHub-1.0.1.zip",
    "sha256": "HASH_DO_ARQUIVO",
    "mandatory": true,
    "notes": "Nova versão publicada."
}
```

## 6. Cliente remoto já incluído

Foi criado o arquivo:

```text
src/backend/remote_client.py
```

Uso básico:

```python
from backend.remote_client import remote_client

resultado = remote_client.post_json("/api/minha-rota", {"texto": "teste"})
```

Para enviar PDFs:

```python
from pathlib import Path
from backend.remote_client import remote_client

resultado = remote_client.post_files(
    "/api/ia/pet-inicial",
    [Path("arquivo.pdf")],
    data={"gerar_word": "false"},
)
```

## 7. Próximo passo

Após configurar as rotas na VM, adapte os serviços de IA para chamar `remote_client` quando `LEGALHUB_API_BASE_URL` estiver preenchida. Assim:

- em desenvolvimento local, você ainda pode usar OpenAI direto;
- em produção interna, os usuários chamam a VM;
- as chaves reais ficam fora do executável.
