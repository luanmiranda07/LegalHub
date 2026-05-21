import os
from typing import Optional

# ajuste o import conforme o nome/classe/função do seu cc_engine.py
from engine import LoginEngine  # <- se o seu engine tiver outro nome, troca aqui
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

app = FastAPI()

# Se o front estiver em outra porta (ex: 5500) ou abrir via file://, precisa CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # em produção, restrinja
    allow_credentials=False,
    allow_methods=["POST", "OPTIONS"],
    allow_headers=["*"],
)

engine = LoginEngine(
    host=os.getenv("LDAP_HOST"),
    domain_fqdn=os.getenv("DOMAIN_FQDN"),
    group_dn=os.getenv("GROUP_DN") or None,
)


class LoginIn(BaseModel):
    usuario: str
    senha: str


@app.post("/api/login")
def login(payload: LoginIn):
    usuario = (payload.usuario or "").strip()
    senha = (payload.senha or "").strip()

    if not usuario or not senha:
        raise HTTPException(
            status_code=400, detail={"erro": "Preencha usuário e senha."}
        )

    resp = engine.authenticate(usuario, senha)  # deve retornar dict com code/erro

    code = int(resp.get("code", 500))
    if code != 200:
        raise HTTPException(status_code=code, detail=resp)

    return resp


# uvicorn api:app --reload --host 0.0.0.0 --port 8000
