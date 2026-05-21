from __future__ import annotations

from typing import Any, Dict, Optional

from ldap3 import BASE, SIMPLE, SUBTREE, Connection, Server

import os
from dotenv import load_dotenv

load_dotenv() # Loads variables from .env into os.environ

LDAP_HOST = os.getenv('LDAP_HOST')
DOMAIN_FQDN = os.getenv('DOMAIN_FQDN')

# Valores padrão para executar este arquivo diretamente (sem interface)
LDAP_USER = os.getenv('LDAP_USER')
LDAP_PASS = os.getenv('LDAP_PASS')

# Deixe None por enquanto se você ainda não tem o DN do grupo.
# Exemplo (MODELO): "CN=MeuGrupo,OU=Grupos,DC=caiqueadv,DC=local"
GROUP_DN: str | None = None


class LoginEngine:
    """Motor de login (sem interface).

    Contrato de retorno (dict):
      conexao: "OK" | "ERRO"
      code: int  (200/401/404/500)
      login: str (entrada)
      sam: str   (sAMAccountName usado na busca)
      base_dn: str | None
      dn/cn/display_name/nome/sobrenome/email: str | None
      member_of: list[str]
      grupo_validado: bool | None
      erro: str | None
    """

    def __init__(
        self,
        host: str,
        domain_fqdn: str,
        group_dn: Optional[str] = None,
    ) -> None:
        self.host = host
        self.domain_fqdn = domain_fqdn
        self.group_dn = group_dn

    def authenticate(self, login: str, senha: str) -> Dict[str, Any]:
        result: Dict[str, Any] = {
            "conexao": "ERRO",
            "code": 500,
            "login": login,
            "sam": self._to_sam(login),
            "base_dn": None,
            "dn": None,
            "cn": None,
            "display_name": None,
            "nome": None,
            "sobrenome": None,
            "email": None,
            "member_of": [],
            "grupo_validado": None,  # True/False quando group_dn estiver definido
            "erro": None,
        }

        server = Server(self.host, get_info=None)

        # 1) Conexão / Bind
        try:
            bind_user = self._to_bind_user(login)
            conn = Connection(
                server,
                user=bind_user,
                password=senha,
                authentication=SIMPLE,
                auto_bind=True,
            )
            result["conexao"] = "OK"
            result["code"] = 200
        except Exception as e:
            # Bind falhou: credenciais inválidas / política / AD indisponível etc.
            result["conexao"] = "ERRO"
            result["code"] = 401
            result["erro"] = f"Falha de autenticação: {e}"
            return result

        try:
            # 2) Descobre base_dn no RootDSE
            base_dn = self._get_base_dn(conn)
            result["base_dn"] = base_dn

            # 3) Busca usuário
            user_entry = self._find_user(conn, base_dn, result["sam"])

            result["dn"] = self._attr(user_entry, "distinguishedName")
            result["cn"] = self._attr(user_entry, "cn")
            result["display_name"] = self._attr(user_entry, "displayName")
            result["nome"] = self._attr(user_entry, "givenName")
            result["sobrenome"] = self._attr(user_entry, "sn")
            result["email"] = self._attr(user_entry, "mail")

            member_of = user_entry.memberOf.values if "memberOf" in user_entry else []
            result["member_of"] = [str(x) for x in member_of]

            # 4) (Opcional) Validação de grupo
            if self.group_dn:
                user_dn = result["dn"] or ""
                ok_nested = self._is_member_of_nested(conn, base_dn, user_dn, self.group_dn)
                result["grupo_validado"] = bool(ok_nested)

            result["code"] = 200
            return result

        except LookupError as e:
            # Usuário não encontrado
            result["code"] = 404
            result["erro"] = str(e)
            return result

        except Exception as e:
            result["code"] = 500
            result["erro"] = f"Erro durante consulta LDAP: {e}"
            return result

        finally:
            conn.unbind()

    @staticmethod
    def _attr(entry: Any, attr: str) -> Optional[str]:
        if attr in entry and getattr(entry, attr, None) is not None:
            value = getattr(entry, attr).value
            return str(value) if value is not None else None
        return None

    def _to_bind_user(self, login: str) -> str:
        # Se já vier como UPN (user@domain), usa direto.
        if "@" in login:
            return login
        return f"{login}@{self.domain_fqdn}"

    @staticmethod
    def _to_sam(login: str) -> str:
        # Se vier como email/UPN, para busca por sAMAccountName pegamos o lado esquerdo.
        return login.split("@", 1)[0].strip()

    @staticmethod
    def _get_base_dn(conn: Connection) -> str:
        conn.search(
            search_base="",
            search_filter="(objectClass=*)",
            search_scope=BASE,
            attributes=["defaultNamingContext"],
        )
        if not conn.entries:
            raise RuntimeError("Não foi possível obter defaultNamingContext via RootDSE.")
        base_dn = conn.entries[0].defaultNamingContext.value
        if not base_dn:
            raise RuntimeError("defaultNamingContext veio vazio no RootDSE.")
        return str(base_dn)

    @staticmethod
    def _find_user(conn: Connection, base_dn: str, sam: str) -> Any:
        conn.search(
            search_base=base_dn,
            search_filter=f"(sAMAccountName={sam})",
            search_scope=SUBTREE,
            attributes=[
                "distinguishedName",
                "cn",
                "displayName",
                "givenName",
                "sn",
                "mail",
                "memberOf",
            ],
        )
        if not conn.entries:
            raise LookupError(f"Usuário não encontrado: sAMAccountName={sam}")
        return conn.entries[0]

    @staticmethod
    def _is_member_of_nested(conn: Connection, base_dn: str, user_dn: str, group_dn: str) -> bool:
        # Regra do AD: LDAP_MATCHING_RULE_IN_CHAIN (grupos aninhados)
        filter_nested = (
            f"(&"
            f"(distinguishedName={user_dn})"
            f"(memberOf:1.2.840.113556.1.4.1941:={group_dn})"
            f")"
        )
        conn.search(
            search_base=base_dn,
            search_filter=filter_nested,
            search_scope=SUBTREE,
            attributes=["distinguishedName"],
        )
        return bool(conn.entries)


def main() -> None:
    engine = LoginEngine(
        host=LDAP_HOST,
        domain_fqdn=DOMAIN_FQDN,
        group_dn=GROUP_DN,
    )

    resp = engine.authenticate(LDAP_USER, LDAP_PASS)

    # Prints que você pediu
    print("Conexão:", resp["conexao"])
    print("code:", resp["code"])

    if resp["code"] != 200:
        print(resp["erro"])
        return

    # Dados úteis
    print("Base DN:", resp["base_dn"])
    print("DN:", resp["dn"])
    print("CN:", resp["cn"])
    print("Display Name:", resp["display_name"])
    print("Nome:", resp["nome"])
    print("Sobrenome:", resp["sobrenome"])
    print("Email:", resp["email"])

    if GROUP_DN:
        print("\nValidação de grupo:")
        print("Group DN:", GROUP_DN)
        print("Membro (aninhado / in-chain):", resp["grupo_validado"])
    else:
        print("\nValidação de grupo: GROUP_DN=None (ainda não configurado).")


if __name__ == "__main__":
    main()
