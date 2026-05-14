# -*- coding: utf-8 -*-
try:
    basestring
except NameError:
    basestring = str

import json

from . import api_client, config, local_cache, logger


def normalize_items(data):
    if isinstance(data, dict) and "items" in data:
        return data.get("items") or []
    if isinstance(data, list):
        return data
    return []


def merge_projects(server_projects, local_projects):
    result = []
    seen = set()
    for project in (server_projects or []) + (local_projects or []):
        pid = project.get("id") or project.get("project_id")
        if pid and pid not in seen:
            seen.add(pid)
            result.append(project)
    return result


def current_project_from_settings(settings):
    project_id = settings.get("CURRENT_PROJECT_ID") or ""
    if not project_id:
        return None
    return {
        "id": project_id,
        "name": settings.get("CURRENT_PROJECT_NAME") or project_id,
        "code": settings.get("CURRENT_PROJECT_CODE") or "",
        "user_role": settings.get("CURRENT_USER_ROLE") or "",
        "is_local": config.is_local_mode(settings),
    }


def cached_projects(settings=None):
    settings = settings or config.load()
    projects = local_cache.list_local_projects()
    current = current_project_from_settings(settings)
    if current:
        projects = merge_projects(projects, [current])
    return projects


def load_projects(use_server=True):
    settings = config.load()
    local_projects = cached_projects(settings)
    if not use_server:
        if config.is_local_mode(settings):
            return local_projects, u"Локальный режим"
        return local_projects, u"Показан локальный кэш. Нажмите «Обновить» для синхронизации"
    if config.is_local_mode(settings):
        return local_projects, u"Сервер не настроен, используется локальный режим"
    try:
        projects = normalize_items(api_client.ApiClient(settings).get_projects())
        for project in projects:
            local_cache.cache_project(project)
        return merge_projects(projects, local_projects), u""
    except Exception as exc:
        logger.exception("Failed to load projects")
        return local_projects, u"Сервер недоступен, используется локальный режим"


def choose_project(project):
    if not project:
        return config.load()
    return config.set_current_project(project)


def create_project(payload):
    settings = config.load()
    if config.is_local_mode(settings):
        result = local_cache.save_local_project(payload)
        config.set_current_project(result)
        return result, u"Проект создан локально"
    try:
        result = api_client.ApiClient(settings).create_project(payload)
        if isinstance(result, dict):
            local_cache.cache_project(result)
            config.set_current_project(result)
        return result, u"Проект создан"
    except Exception as exc:
        logger.exception("Failed to create project via API")
        result = local_cache.save_local_project(payload, str(exc))
        config.set_current_project(result)
        return result, u"Сервер недоступен, проект создан локально"


def require_project(settings=None):
    settings = settings or config.load()
    if not settings.get("CURRENT_PROJECT_ID"):
        raise ValueError(u"Сначала выберите проект")
    return settings


def sync_pending(project_id):
    settings = config.load()
    if config.is_local_mode(settings):
        return 0
    client = api_client.ApiClient(settings)
    sent = 0
    for item in local_cache.pending(project_id):
        try:
            payload = item.get("payload_json")
            payload = json.loads(payload) if isinstance(payload, basestring) else payload
            if item.get("action") == "create_task":
                res = client.create_task(payload)
                if isinstance(res, dict):
                    local_cache.cache_task(res)
            elif item.get("action") == "patch_task":
                client.patch_task(payload.get("id") or payload.get("temporary_id"), payload.get("project_id"), payload)
            elif item.get("action") == "comment_task":
                client.add_comment(payload.get("task_id"), payload.get("project_id"), payload.get("comment"))
            local_cache.delete_pending(item.get("id"))
            sent += 1
        except Exception as exc:
            local_cache.mark_pending_error(item.get("id"), str(exc))
    return sent
