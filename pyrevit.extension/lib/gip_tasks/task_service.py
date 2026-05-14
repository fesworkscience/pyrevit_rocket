# -*- coding: utf-8 -*-
from . import api_client, config, local_cache, logger, task_models


def list_tasks(project_id, filters=None, use_server=True):
    if not project_id:
        raise ValueError(u"Сначала выберите проект")
    settings = config.load()
    if not use_server:
        if config.is_local_mode(settings):
            return local_cache.list_tasks(project_id), u"Локальный режим"
        return local_cache.list_tasks(project_id), u"Показан локальный кэш. Нажмите «Обновить» для синхронизации"
    if config.is_local_mode(settings):
        return local_cache.list_tasks(project_id), u"Сервер не настроен, используется локальный режим"
    try:
        data = api_client.ApiClient(settings).list_tasks(project_id, filters or {})
        items = data.get("items") if isinstance(data, dict) else data
        items = items or []
        for task in items:
            local_cache.cache_task(task)
        return items, u""
    except Exception:
        logger.exception("Failed to load tasks")
        return local_cache.list_tasks(project_id), u"Сервер недоступен, показан локальный кэш"


def create_task(payload):
    settings = config.load()
    if not payload.get("project_id"):
        raise ValueError(u"Сначала выберите проект")
    if config.is_local_mode(settings):
        local_cache.cache_task(payload)
        return payload, u"Задание сохранено локально"
    try:
        result = api_client.ApiClient(settings).create_task(payload)
        if isinstance(result, dict):
            payload.update(result)
        local_cache.cache_task(payload)
        return payload, u"Задание отправлено"
    except Exception:
        logger.exception("Failed to create task via API")
        local_cache.cache_task(payload)
        local_cache.enqueue("create_task", payload)
        return payload, u"Сервер недоступен, задание сохранено локально"


def update_status(task, marker_status):
    project_id = task.get("project_id")
    task_id = task.get("id") or task.get("temporary_id")
    updates = {
        "marker_status": marker_status,
        "status": task_models.marker_to_server_status(marker_status),
        "updated_at": task_models.now_iso(),
    }
    settings = config.load()
    if config.is_local_mode(settings) or not task.get("id"):
        updated = local_cache.update_task(project_id, task_id, updates)
        return updated or task, u"Статус сохранен локально"
    try:
        updated = api_client.ApiClient(settings).patch_task(task.get("id"), project_id, updates)
        if isinstance(updated, dict):
            local_cache.cache_task(updated)
            return updated, u"Статус обновлен"
    except Exception:
        logger.exception("Failed to update status via API")
        local_cache.enqueue("patch_task", dict(task, **updates))
    updated = local_cache.update_task(project_id, task_id, updates)
    return updated or task, u"Сервер недоступен, статус сохранен локально"


def add_comment(task, comment):
    project_id = task.get("project_id")
    task_id = task.get("id") or task.get("temporary_id")
    settings = config.load()
    if config.is_local_mode(settings) or not task.get("id"):
        updated = local_cache.add_comment(project_id, task_id, comment)
        return updated or task, u"Комментарий сохранен локально"
    try:
        api_client.ApiClient(settings).add_comment(task.get("id"), project_id, comment)
    except Exception:
        logger.exception("Failed to send comment via API")
        local_cache.enqueue("comment_task", {"project_id": project_id, "task_id": task.get("id"), "comment": comment})
    updated = local_cache.add_comment(project_id, task_id, comment)
    return updated or task, u"Комментарий сохранен"
