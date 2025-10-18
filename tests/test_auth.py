import pytest


@pytest.mark.unit
def test_api_key_middleware():
    from mcp_excel.auth import APIKeyMiddleware
    from starlette.applications import Starlette
    from starlette.responses import PlainTextResponse
    from starlette.testclient import TestClient

    app = Starlette()

    @app.route("/test")
    async def test_route(request):
        return PlainTextResponse("success")

    app.add_middleware(APIKeyMiddleware, api_key="secret-key")

    client = TestClient(app)

    response = client.get("/test", headers={"Authorization": "Bearer secret-key"})
    assert response.status_code == 200
    assert response.text == "success"

    response = client.get("/test")
    assert response.status_code == 401

    response = client.get("/test", headers={"Authorization": "Bearer wrong-key"})
    assert response.status_code == 401

    response = client.get("/test", headers={"Authorization": "InvalidFormat"})
    assert response.status_code == 401
