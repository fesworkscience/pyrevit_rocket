# -*- coding: utf-8 -*-
import json
import os
import uuid
try:
    import sqlite3
except Exception:
    sqlite3 = None

from . import config, logger, task_models


def connect(settings=None):
    settings = settings or config.load()
    if sqlite3 is None:
        return None
    con = sqlite3.connect(settings.get("LOCAL_CACHE_PATH"))
    con.row_factory = sqlite3.Row
    init(con)
    return con


def init(con):
    con.execute("""create table if not exists projects (
        id text primary key,
        project_json text not null,
        updated_at text not null
    )""")
    con.execute("""create table if not exists tasks (
        id text primary key,
        project_id text not null,
        task_json text not null,
        updated_at text not null
    )""")
    con.execute("""create table if not exists pending_queue (
        id integer primary key autoincrement,
        action text not null,
        project_id text not null,
        payload_json text not null,
        created_at text not null,
        attempts integer default 0,
        last_error text
    )""")
    con.execute("create table if not exists sync_state (project_id text primary key, last_sync text)")
    con.commit()


def _json_path(name):
    return os.path.join(config.appdata_dir(), name)


def _read_json(path, fallback):
    if not os.path.exists(path):
        return fallback
    try:
        with open(path, "rb") as fp:
            return json.loads(fp.read().decode("utf-8"))
    except Exception:
        logger.exception("Failed to read json cache: %s" % path)
        return fallback


def _write_json(path, data):
    raw = json.dumps(data, indent=2, ensure_ascii=False, sort_keys=True)
    with open(path, "wb") as fp:
        fp.write(raw.encode("utf-8"))


def local_projects_path():
    return _json_path("local_projects.json")


def list_local_projects():
    con = connect()
    if con is not None:
        rows = con.execute("select project_json from projects order by updated_at desc").fetchall()
        con.close()
        return [json.loads(r["project_json"]) for r in rows]
    return _read_json(local_projects_path(), {"items": []}).get("items") or []


def save_local_project(payload, last_error=""):
    project_id = payload.get("id") or payload.get("project_id") or "local-project-" + str(uuid.uuid4())
    project = {
        "id": project_id,
        "company_id": payload.get("company_id") or "",
        "name": payload.get("name") or payload.get("project_name") or u"Локальный проект",
        "code": payload.get("code") or payload.get("project_code") or "",
        "customer": payload.get("customer") or "",
        "address": payload.get("address") or "",
        "description": payload.get("description") or "",
        "created_at": payload.get("created_at") or task_models.now_iso(),
        "updated_at": task_models.now_iso(),
        "user_role": payload.get("user_role") or "project_coordinator",
        "is_local": True,
        "is_active": bool(payload.get("is_active", True)),
        "last_error": last_error or "",
    }
    cache_project(project)
    return project


def cache_project(project):
    project_id = project.get("id") or project.get("project_id")
    if not project_id:
        return
    project["updated_at"] = project.get("updated_at") or task_models.now_iso()
    con = connect()
    if con is not None:
        con.execute(
            "insert or replace into projects(id, project_json, updated_at) values(?,?,?)",
            (project_id, json.dumps(project, ensure_ascii=False), project["updated_at"]),
        )
        con.commit()
        con.close()
        return
    projects = [x for x in list_local_projects() if x.get("id") != project_id]
    projects.append(project)
    _write_json(local_projects_path(), {"items": projects})


def cache_task(task):
    task_id = task.get("id") or task.get("temporary_id")
    project_id = task.get("project_id")
    if not task_id or not project_id:
        return
    task["updated_at"] = task.get("updated_at") or task_models.now_iso()
    con = connect()
    if con is not None:
        con.execute(
            "insert or replace into tasks(id, project_id, task_json, updated_at) values(?,?,?,?)",
            (task_id, project_id, json.dumps(task, ensure_ascii=False), task["updated_at"]),
        )
        con.commit()
        con.close()
        return
    path = _json_path("local_tasks.json")
    items = _read_json(path, {"items": []}).get("items") or []
    items = [x for x in items if (x.get("id") or x.get("temporary_id")) != task_id]
    items.append(task)
    _write_json(path, {"items": items})


def list_tasks(project_id):
    if not project_id:
        return []
    con = connect()
    if con is not None:
        rows = con.execute("select task_json from tasks where project_id=? order by updated_at desc", (project_id,)).fetchall()
        con.close()
        return [json.loads(r["task_json"]) for r in rows]
    items = _read_json(_json_path("local_tasks.json"), {"items": []}).get("items") or []
    return [x for x in items if x.get("project_id") == project_id]


def get_task(task_id, project_id=None):
    for task in list_tasks(project_id) if project_id else _all_tasks():
        if task.get("id") == task_id or task.get("temporary_id") == task_id:
            return task
    return None


def _all_tasks():
    con = connect()
    if con is not None:
        rows = con.execute("select task_json from tasks order by updated_at desc").fetchall()
        con.close()
        return [json.loads(r["task_json"]) for r in rows]
    return _read_json(_json_path("local_tasks.json"), {"items": []}).get("items") or []


def update_task(project_id, task_id, updates):
    task = get_task(task_id, project_id)
    if not task:
        return None
    task.update(updates or {})
    task["updated_at"] = task_models.now_iso()
    cache_task(task)
    return task


def add_comment(project_id, task_id, comment):
    task = get_task(task_id, project_id)
    if not task:
        return None
    comments = task.get("comments") or []
    comments.append({"text": comment, "created_at": task_models.now_iso()})
    task["comments"] = comments
    if comment:
        existing = task.get("comment") or ""
        task["comment"] = (existing + "\n" + comment).strip() if existing else comment
    task["updated_at"] = task_models.now_iso()
    cache_task(task)
    return task


def enqueue(action, payload):
    project_id = payload.get("project_id") or payload.get("id")
    if not project_id:
        return
    con = connect()
    if con is None:
        return
    con.execute(
        "insert into pending_queue(action, project_id, payload_json, created_at) values(?,?,?,?)",
        (action, project_id, json.dumps(payload, ensure_ascii=False), task_models.now_iso()),
    )
    con.commit()
    con.close()


def pending(project_id, limit=50):
    con = connect()
    if con is None:
        return []
    rows = con.execute("select * from pending_queue where project_id=? order by id limit ?", (project_id, limit)).fetchall()
    con.close()
    return [dict(r) for r in rows]


def pending_count(project_id):
    con = connect()
    if con is None:
        return 0
    row = con.execute("select count(*) as c from pending_queue where project_id=?", (project_id,)).fetchone()
    con.close()
    return int(row["c"] or 0)


def delete_pending(item_id):
    con = connect()
    if con is None:
        return
    con.execute("delete from pending_queue where id=?", (item_id,))
    con.commit()
    con.close()


def mark_pending_error(item_id, error):
    con = connect()
    if con is None:
        return
    con.execute("update pending_queue set attempts=attempts+1, last_error=? where id=?", (error, item_id))
    con.commit()
    con.close()
    logger.write("Pending queue error: %s" % error)


def get_last_sync(project_id):
    con = connect()
    if con is None:
        return ""
    row = con.execute("select last_sync from sync_state where project_id=?", (project_id,)).fetchone()
    con.close()
    return row["last_sync"] if row else ""


def set_last_sync(project_id, value):
    con = connect()
    if con is None:
        return
    con.execute("insert or replace into sync_state(project_id,last_sync) values(?,?)", (project_id, value))
    con.commit()
    con.close()

