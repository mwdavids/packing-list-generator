"""Tests for the packing list generator app.

Run with:  pytest test_app.py -v
"""

import os
import sqlite3
import tempfile

import pytest
from fastapi.testclient import TestClient


@pytest.fixture(autouse=True)
def _isolated_db(tmp_path, monkeypatch):
    """Each test gets its own fresh database and JWT secret."""
    db_path = str(tmp_path / "test.db")
    monkeypatch.setenv("DATABASE_PATH", db_path)
    monkeypatch.setenv("JWT_SECRET", "test-secret-key-for-testing-only")
    monkeypatch.setenv("LISTS_FILE", str(tmp_path / "nonexistent.json"))  # no migration
    monkeypatch.setenv("BASE_USER", "baseuser")
    monkeypatch.setenv("BASE_PASS", "basepass123")
    monkeypatch.setenv("ANTHROPIC_API_KEY", "")
    monkeypatch.setenv("INVITE_CODE", "")  # disabled by default in tests

    # Re-import to pick up new env vars and reinit DB
    import importlib
    import main as main_mod
    main_mod.DATABASE_PATH = db_path
    main_mod.JWT_SECRET = "test-secret-key-for-testing-only"
    main_mod.LISTS_FILE = str(tmp_path / "nonexistent.json")
    main_mod.BASE_USER = "baseuser"
    main_mod.BASE_PASS = "basepass123"
    main_mod.INVITE_CODE = ""
    main_mod._login_attempts.clear()  # reset rate limiter between tests
    main_mod.init_db()


@pytest.fixture
def client():
    from main import app
    return TestClient(app)


def _register(client, username="alice", password="pass1234", display_name="Alice", invite_code=""):
    return client.post("/api/register", json={
        "username": username, "password": password, "display_name": display_name,
        "invite_code": invite_code,
    })


def _login(client, username="alice", password="pass1234"):
    return client.post("/api/login", json={"username": username, "password": password})


# =========================================================================
# Auth: Register
# =========================================================================

class TestRegister:
    def test_register_success(self, client):
        r = _register(client)
        assert r.status_code == 200
        body = r.json()
        assert body["ok"] is True
        assert body["username"] == "alice"
        assert "token" in r.cookies

    def test_register_duplicate_username(self, client):
        _register(client)
        r = _register(client)
        assert r.status_code == 409
        assert "already taken" in r.json()["detail"]

    def test_register_username_case_insensitive(self, client):
        _register(client, username="Alice")
        r = _register(client, username="alice")
        assert r.status_code == 409

    def test_register_short_username(self, client):
        r = _register(client, username="ab")
        assert r.status_code == 422

    def test_register_short_password(self, client):
        r = _register(client, username="validuser", password="12345")
        assert r.status_code == 422

    def test_register_invalid_username_chars(self, client):
        r = _register(client, username="bad user!")
        assert r.status_code == 422

    def test_register_empty_display_name(self, client):
        r = client.post("/api/register", json={
            "username": "validuser", "password": "pass1234", "display_name": "",
        })
        assert r.status_code == 422


# =========================================================================
# Auth: Login / Logout / Me
# =========================================================================

class TestLoginLogoutMe:
    def test_login_success(self, client):
        _register(client)
        r = _login(client)
        assert r.status_code == 200
        assert r.json()["ok"] is True
        assert "token" in r.cookies

    def test_login_wrong_password(self, client):
        _register(client)
        r = _login(client, password="wrong")
        assert r.status_code == 401

    def test_login_nonexistent_user(self, client):
        r = _login(client, username="nobody")
        assert r.status_code == 401

    def test_me_unauthenticated(self, client):
        r = client.get("/api/me")
        assert r.status_code == 200
        assert r.json()["authenticated"] is False

    def test_me_authenticated(self, client):
        _register(client)
        r = client.get("/api/me")
        assert r.status_code == 200
        body = r.json()
        assert body["authenticated"] is True
        assert body["username"] == "alice"
        assert body["display_name"] == "Alice"

    def test_logout(self, client):
        _register(client)
        r = client.post("/api/logout")
        assert r.status_code == 200
        # After logout, /api/me should be unauthenticated
        r2 = client.get("/api/me")
        assert r2.json()["authenticated"] is False


# =========================================================================
# Lists CRUD
# =========================================================================

class TestLists:
    def test_unauthenticated_returns_401(self, client):
        r = client.get("/api/lists")
        assert r.status_code == 401

    def test_create_and_get_lists(self, client):
        _register(client)
        r = client.post("/api/lists", json={
            "name": "Rainier 2024", "type": "alpine climbing", "content": "Pack\nRope\nHelmet",
        })
        assert r.status_code == 200
        body = r.json()
        assert body["name"] == "Rainier 2024"
        assert "id" in body

        lists = client.get("/api/lists").json()
        assert len(lists) == 1
        assert lists[0]["name"] == "Rainier 2024"

    def test_delete_list(self, client):
        _register(client)
        created = client.post("/api/lists", json={
            "name": "Test", "type": "test", "content": "item",
        }).json()
        r = client.delete(f"/api/lists/{created['id']}")
        assert r.status_code == 200
        assert client.get("/api/lists").json() == []

    def test_delete_nonexistent_list(self, client):
        _register(client)
        r = client.delete("/api/lists/9999999")
        assert r.status_code == 404

    def test_user_isolation(self, client):
        """User A cannot see or delete User B's lists."""
        _register(client, username="alice")
        client.post("/api/lists", json={"name": "Alice list", "type": "t", "content": "x"})
        alice_list_id = client.get("/api/lists").json()[0]["id"]

        # Logout and register as Bob
        client.post("/api/logout")
        _register(client, username="bob", display_name="Bob")
        bob_lists = client.get("/api/lists").json()
        assert len(bob_lists) == 0

        # Bob can't delete Alice's list
        r = client.delete(f"/api/lists/{alice_list_id}")
        assert r.status_code == 404


# =========================================================================
# Fork base lists
# =========================================================================

class TestForkBaseLists:
    def _setup_base_user(self, client):
        """Register the base user and add some lists."""
        _register(client, username="baseuser", password="basepass123", display_name="Base")
        client.post("/api/lists", json={"name": "Base list 1", "type": "backpacking", "content": "Pack"})
        client.post("/api/lists", json={"name": "Base list 2", "type": "climbing", "content": "Rope"})
        # Mark as base user
        from main import get_db
        db = get_db()
        db.execute("UPDATE users SET is_base_user = 1 WHERE username = 'baseuser'")
        db.commit()
        db.close()
        client.post("/api/logout")

    def test_fork_copies_base_lists(self, client):
        self._setup_base_user(client)
        _register(client, username="newuser", password="pass1234", display_name="New")
        me = client.get("/api/me").json()
        assert me["base_library_available"] is True

        r = client.post("/api/fork-base-lists")
        assert r.status_code == 200
        assert r.json()["copied"] == 2

        lists = client.get("/api/lists").json()
        assert len(lists) == 2

    def test_fork_no_base_user(self, client):
        _register(client)
        r = client.post("/api/fork-base-lists")
        assert r.json()["copied"] == 0


# =========================================================================
# Generations: save, list, get, delete
# =========================================================================

class TestGenerations:
    def test_unauthenticated_returns_401(self, client):
        r = client.get("/api/generations")
        assert r.status_code == 401

    def test_save_and_list_generation(self, client):
        _register(client)
        r = client.post("/api/generations", json={
            "trip_type": "Backpacking",
            "location": "Enchantments",
            "markdown": "## Shelter\n| Item | Priority | Notes |\n|---|---|---|\n| Tent | | |",
            "title": "Enchantments trip",
        })
        assert r.status_code == 200
        body = r.json()
        assert "id" in body
        assert "share_token" in body

        gens = client.get("/api/generations").json()
        assert len(gens) == 1
        assert gens[0]["title"] == "Enchantments trip"
        assert gens[0]["share_token"] == body["share_token"]

    def test_get_generation_by_id(self, client):
        _register(client)
        saved = client.post("/api/generations", json={
            "markdown": "## Test", "title": "Test trip",
        }).json()
        r = client.get(f"/api/generations/{saved['id']}")
        assert r.status_code == 200
        assert r.json()["markdown"] == "## Test"

    def test_delete_generation(self, client):
        _register(client)
        saved = client.post("/api/generations", json={
            "markdown": "## Delete me", "title": "Del",
        }).json()
        r = client.delete(f"/api/generations/{saved['id']}")
        assert r.status_code == 200
        assert client.get("/api/generations").json() == []

    def test_delete_nonexistent_generation(self, client):
        _register(client)
        r = client.delete("/api/generations/9999")
        assert r.status_code == 404

    def test_generation_user_isolation(self, client):
        _register(client, username="alice")
        saved = client.post("/api/generations", json={
            "markdown": "## Alice trip", "title": "Alice trip",
        }).json()

        client.post("/api/logout")
        _register(client, username="bob", display_name="Bob")
        assert client.get("/api/generations").json() == []
        assert client.get(f"/api/generations/{saved['id']}").status_code == 404
        assert client.delete(f"/api/generations/{saved['id']}").status_code == 404


# =========================================================================
# Share
# =========================================================================

class TestShare:
    def test_share_link_works_without_auth(self, client):
        _register(client)
        saved = client.post("/api/generations", json={
            "trip_type": "Day hiking",
            "location": "Mt. Si",
            "markdown": "## Shared list content",
            "title": "Mt. Si day hike",
        }).json()
        token = saved["share_token"]

        # Logout to prove no auth needed
        client.post("/api/logout")

        r = client.get(f"/api/share/{token}")
        assert r.status_code == 200
        body = r.json()
        assert body["title"] == "Mt. Si day hike"
        assert body["markdown"] == "## Shared list content"
        assert body["trip_type"] == "Day hiking"
        assert body["location"] == "Mt. Si"

    def test_share_invalid_token(self, client):
        r = client.get("/api/share/nonexistent-token")
        assert r.status_code == 404

    def test_each_generation_gets_unique_token(self, client):
        _register(client)
        t1 = client.post("/api/generations", json={"markdown": "a", "title": "a"}).json()["share_token"]
        t2 = client.post("/api/generations", json={"markdown": "b", "title": "b"}).json()["share_token"]
        assert t1 != t2


# =========================================================================
# Export xlsx
# =========================================================================

class TestExportXlsx:
    def test_unauthenticated_returns_401(self, client):
        r = client.post("/api/export-xlsx", json={"markdown": "## Test"})
        assert r.status_code == 401

    def test_export_returns_xlsx(self, client):
        _register(client)
        md = "## Shelter\n| Item | Priority | Notes |\n|---|---|---|\n| Tent | | Freestanding |\n| Tarp | OPTIONAL | |"
        r = client.post("/api/export-xlsx", json={
            "markdown": md, "title": "Test Export", "group_size": "2",
        })
        assert r.status_code == 200
        assert "spreadsheetml" in r.headers["content-type"]
        assert len(r.content) > 100  # non-trivial file


# =========================================================================
# Page routes
# =========================================================================

class TestPageRoutes:
    def test_root_returns_html(self, client):
        r = client.get("/")
        assert r.status_code == 200

    def test_share_page_returns_html(self, client):
        r = client.get("/share/some-token")
        assert r.status_code == 200


# =========================================================================
# Migration from lists.json
# =========================================================================

class TestMigration:
    def test_migration_imports_json(self, tmp_path, monkeypatch):
        import json
        lists_file = tmp_path / "lists.json"
        lists_file.write_text(json.dumps([
            {"id": "1", "name": "Trip A", "type": "backpacking", "date_added": "2025-01-01", "content": "item1"},
            {"id": "2", "name": "Trip B", "type": "hiking", "date_added": "2025-06-01", "content": "item2"},
        ]))
        db_path = tmp_path / "migration_test.db"

        import main as main_mod
        main_mod.DATABASE_PATH = str(db_path)
        main_mod.LISTS_FILE = str(lists_file)
        main_mod.BASE_USER = "miguser"
        main_mod.BASE_PASS = "migpass123"
        main_mod.init_db()

        db = sqlite3.connect(str(db_path))
        db.row_factory = sqlite3.Row
        users = db.execute("SELECT * FROM users WHERE username = 'miguser'").fetchall()
        assert len(users) == 1
        assert users[0]["is_base_user"] == 1

        lists_rows = db.execute("SELECT * FROM lists").fetchall()
        assert len(lists_rows) == 2
        db.close()


# =========================================================================
# Rate limiting
# =========================================================================

class TestRateLimit:
    def test_rate_limit_blocks_after_max_attempts(self, client):
        _register(client, username="ratelimituser")
        client.post("/api/logout")
        for _ in range(5):
            r = _login(client, username="ratelimituser", password="wrongpass")
            assert r.status_code == 401
        # 6th attempt should be blocked
        r = _login(client, username="ratelimituser", password="wrongpass")
        assert r.status_code == 429
        assert "Too many" in r.json()["detail"]

    def test_successful_login_clears_attempts(self, client):
        _register(client, username="clearuser")
        client.post("/api/logout")
        for _ in range(3):
            _login(client, username="clearuser", password="wrongpass")
        # Successful login should clear counter
        r = _login(client, username="clearuser", password="pass1234")
        assert r.status_code == 200
        # Should be able to fail again without hitting limit
        client.post("/api/logout")
        for _ in range(4):
            r = _login(client, username="clearuser", password="wrongpass")
            assert r.status_code == 401


# =========================================================================
# Invite code
# =========================================================================

class TestInviteCode:
    def test_no_invite_code_required_when_not_set(self, client):
        """When INVITE_CODE is empty, registration should work without one."""
        r = _register(client)
        assert r.status_code == 200

    def test_invite_code_required_when_set(self, client):
        import main as main_mod
        main_mod.INVITE_CODE = "secret-invite-123"
        r = _register(client, username="noinvite")
        assert r.status_code == 403
        assert "invite" in r.json()["detail"].lower()

    def test_wrong_invite_code_rejected(self, client):
        import main as main_mod
        main_mod.INVITE_CODE = "secret-invite-123"
        r = _register(client, username="wrongcode", invite_code="bad-code")
        assert r.status_code == 403

    def test_correct_invite_code_accepted(self, client):
        import main as main_mod
        main_mod.INVITE_CODE = "secret-invite-123"
        r = _register(client, username="goodcode", invite_code="secret-invite-123")
        assert r.status_code == 200
        assert r.json()["ok"] is True
