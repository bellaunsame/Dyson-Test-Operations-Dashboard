from flask import Flask, request, jsonify
from ldap3 import Server, Connection, ALL, ALL_ATTRIBUTES, NTLM, core
import os

app = Flask(__name__)

# LDAP configuration (replace with your actual LDAP server details)
LDAP_SERVER = os.getenv("LDAP_SERVER", "ldap://localhost:389")
LDAP_USER_DN = os.getenv("LDAP_USER_DN", "cn=admin,dc=example,dc=com")
LDAP_PASSWORD = os.getenv("LDAP_PASSWORD", "adminpassword")
LDAP_SEARCH_BASE = os.getenv("LDAP_SEARCH_BASE", "dc=example,dc=com")

# Optional: Use SSL/TLS if your LDAP server supports it
USE_SSL = os.getenv("LDAP_USE_SSL", "false").lower() == "true"


def ldap_authenticate(username: str, password: str):
    """
    Authenticate a user against the LDAP server.
    Returns user attributes if successful, None otherwise.
    """
    if not username or not password:
        return None

    try:
        # Connect to LDAP server
        server = Server(LDAP_SERVER, get_info=ALL, use_ssl=USE_SSL)
        
        # First bind with admin account to search for the user DN
        with Connection(server, LDAP_USER_DN, LDAP_PASSWORD, auto_bind=True) as admin_conn:
            search_filter = f"(uid={username})"
            admin_conn.search(
                search_base=LDAP_SEARCH_BASE,
                search_filter=search_filter,
                attributes=ALL_ATTRIBUTES
            )

            if not admin_conn.entries:
                return None  # User not found

            user_dn = admin_conn.entries[0].entry_dn

        # Now try binding with the found user DN and provided password
        with Connection(server, user_dn, password, auto_bind=True) as user_conn:
            return {
                "dn": user_dn,
                "attributes": user_conn.extend.standard.who_am_i()
            }

    except core.exceptions.LDAPBindError:
        return None  # Invalid credentials
    except Exception as e:
        app.logger.error(f"LDAP error: {e}")
        return None
