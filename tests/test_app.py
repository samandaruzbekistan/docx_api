import json

import httpx
import pytest
from fastapi.testclient import TestClient

import main
from test_docx_parser import docx, row


def question(number: int) -> dict:
    return {
        "row": number + 1,
        "question": {"text": f"Savol {number}", "image": None},
        "correct": {"text": "a", "image": None},
        "wrong1": {"text": "b", "image": None},
        "wrong2": {"text": "c", "image": None},
        "wrong3": {"text": "", "image": None},
    }


def send_body(count: int) -> dict:
    return {
        "test_id": 1, "language": "uz", "subject_id": 2, "grade_id": 7,
        "author": " Karimova ", "questions": [question(n) for n in range(1, count + 1)],
    }


@pytest.fixture
def api(monkeypatch):
    """Reyting tizimi API'sining soxta nusxasi: so'rovlarni yozib boradi."""
    calls: list[httpx.Request] = []
    state = {"handler": None}

    def transport(request: httpx.Request) -> httpx.Response:
        calls.append(request)
        return state["handler"](request)

    monkeypatch.setattr(main, "API_BASE_URL", "https://reyting.test/api/public")
    monkeypatch.setattr(main, "API_KEY", "maxfiy-kalit")
    monkeypatch.setattr(main, "APP_PASSWORD", "")
    monkeypatch.setattr(
        main, "_client",
        lambda timeout: httpx.AsyncClient(transport=httpx.MockTransport(transport), headers=main._headers()),
    )
    state["calls"] = calls
    return state


def saved(request: httpx.Request) -> httpx.Response:
    count = len(json.loads(request.content)["questions"])
    return httpx.Response(200, json={"success": True, "saved_count": count, "skipped_count": 0, "failed_count": 0, "errors": []})


def test_parse_returns_questions_without_sending_anything(api):
    upload = docx(row("2+2 necha?", "4", "3", "5", "6", "Karimova"))

    response = TestClient(main.app).post("/api/parse", files=[("files", ("test.docx", upload)), ("files", ("yomon.docx", b"x"))])

    good, bad = response.json()["files"]
    assert good["name"] == "test.docx" and good["author"] == "Karimova" and len(good["questions"]) == 1
    assert bad["error"] and bad["questions"] == []
    assert api["calls"] == []


def test_lists_are_proxied_with_the_api_key(api):
    api["handler"] = lambda request: httpx.Response(200, json={"success": True, "data": [{"id": 3, "name": "Matematika"}]})

    response = TestClient(main.app).get("/api/subjects", params={"language": "ru", "test_id": 5})

    assert response.json()["data"][0]["name"] == "Matematika"
    request = api["calls"][0]
    assert str(request.url) == "https://reyting.test/api/public/subjects?language=ru&test_id=5"
    assert request.headers["X-API-KEY"] == "maxfiy-kalit"


def test_subjects_can_be_narrowed_to_classes_and_the_cookie_is_forwarded(api, monkeypatch):
    monkeypatch.setattr(main, "API_COOKIE", "usenew=1")
    api["handler"] = lambda request: httpx.Response(200, json={"success": True, "data": []})
    client = TestClient(main.app)

    client.get("/api/subjects", params=[("language", "uz"), ("class_ids", 4), ("class_ids", 9)])
    client.get("/api/classes", params={"level": 7})

    subjects, classes = api["calls"]
    assert subjects.url.params.get_list("class_ids[]") == ["4", "9"]
    assert subjects.headers["Cookie"] == "usenew=1"
    assert str(classes.url) == "https://reyting.test/api/public/classes?level=7"


def test_a_wrong_api_key_is_explained(api):
    api["handler"] = lambda request: httpx.Response(401, json={"message": "Unauthorized"})

    response = TestClient(main.app).get("/api/tests")

    assert response.status_code == 401
    assert "API kaliti" in response.json()["error"]


def test_send_splits_questions_into_chunks_and_sums_the_result(api, monkeypatch):
    monkeypatch.setattr(main, "CHUNK_QUESTIONS", 2)
    api["handler"] = saved

    result = TestClient(main.app).post("/api/send", json=send_body(5)).json()

    assert result == {"success": True, "saved": 5, "skipped": 0, "linked": 0, "class_ids": [], "failed": 0,
                      "unsent": 0, "errors": [], "error": None}
    bodies = [json.loads(request.content) for request in api["calls"]]
    assert [len(body["questions"]) for body in bodies] == [2, 2, 1]
    assert bodies[0]["class_ids"] == []
    assert bodies[0]["author"] == "Karimova" and bodies[0]["grade_id"] == 7
    assert "row" not in bodies[0]["questions"][0]
    assert api["calls"][0].headers["X-API-KEY"] == "maxfiy-kalit"


def test_send_forwards_selected_classes(api):
    api["handler"] = lambda request: httpx.Response(200, json={
        "saved_count": 2, "skipped_count": 1, "linked_count": 1, "failed_count": 0, "errors": []})

    result = TestClient(main.app).post("/api/send", json={**send_body(3), "class_ids": [38, 34, 38]}).json()

    assert json.loads(api["calls"][0].content)["class_ids"] == [34, 38]
    assert result["success"] is True and result["linked"] == 1 and result["class_ids"] == [34, 38]


def test_send_stops_when_the_rating_system_ignores_classes(api, monkeypatch):
    # Eski reyting tizimi linked_count qaytarmaydi: savollar butun darajaga tushgan bo'ladi.
    monkeypatch.setattr(main, "CHUNK_QUESTIONS", 2)
    api["handler"] = saved

    result = TestClient(main.app).post("/api/send", json={**send_body(5), "class_ids": [34]}).json()

    assert result["success"] is False and "sinf bo'yicha" in result["error"]
    assert len(api["calls"]) == 1 and result["unsent"] == 3


def test_send_reports_duplicates_and_points_errors_at_word_rows(api, monkeypatch):
    monkeypatch.setattr(main, "CHUNK_QUESTIONS", 2)
    answers = iter([
        {"saved_count": 1, "skipped_count": 1, "failed_count": 0, "errors": []},
        {"saved_count": 1, "skipped_count": 0, "failed_count": 1, "errors": ["2-savol: Rasm hajmi 2 MB dan oshmasligi kerak."]},
    ])
    api["handler"] = lambda request: httpx.Response(200, json=next(answers))

    result = TestClient(main.app).post("/api/send", json=send_body(4)).json()

    assert result["success"] is False
    assert (result["saved"], result["skipped"], result["failed"]) == (2, 1, 1)
    assert result["errors"] == ["5-qator: Rasm hajmi 2 MB dan oshmasligi kerak."]


def test_send_stops_when_the_api_rejects_the_request(api, monkeypatch):
    monkeypatch.setattr(main, "CHUNK_QUESTIONS", 2)
    answers = iter([saved, lambda request: httpx.Response(422, json={"success": False, "error": "Test yopilgan."})])
    api["handler"] = lambda request: next(answers)(request)

    result = TestClient(main.app).post("/api/send", json=send_body(5)).json()

    assert result["success"] is False and result["error"] == "Test yopilgan."
    assert result["saved"] == 2 and result["unsent"] == 3
    assert len(api["calls"]) == 2


def test_big_questions_are_split_by_size(api, monkeypatch):
    monkeypatch.setattr(main, "CHUNK_BYTES", 3000)
    api["handler"] = saved
    body = send_body(3)
    for item in body["questions"]:
        item["question"]["image"] = "data:image/png;base64," + "A" * 2000

    TestClient(main.app).post("/api/send", json=body)

    assert [len(json.loads(request.content)["questions"]) for request in api["calls"]] == [1, 1, 1]


def test_the_page_can_be_protected_with_a_password(api, monkeypatch):
    client = TestClient(main.app)
    assert client.get("/").status_code == 200

    monkeypatch.setattr(main, "APP_PASSWORD", "sir")
    assert client.get("/").status_code == 401
    assert client.post("/api/send", json=send_body(1)).status_code == 401
    assert client.get("/", auth=("admin", "xato")).status_code == 401
    assert client.get("/", auth=("admin", "sir")).status_code == 200
