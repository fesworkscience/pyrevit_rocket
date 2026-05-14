# -*- coding: utf-8 -*-
from . import api_client, local_cache, task_models


def poll_project(project_id):
    if not project_id:
        raise ValueError(u"Сначала выберите проект")
    client = api_client.ApiClient()
    since = local_cache.get_last_sync(project_id)
    data = client.task_changes(project_id, since)
    items = data.get("items") if isinstance(data, dict) else data
    count = 0
    for task in items or []:
        local_cache.cache_task(task)
        count += 1
    local_cache.set_last_sync(project_id, task_models.now_iso())
    return count

