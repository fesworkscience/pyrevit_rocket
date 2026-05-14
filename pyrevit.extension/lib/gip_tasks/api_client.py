# -*- coding: utf-8 -*-
import json
import sys
try:
    import urllib2
    from urllib import urlencode
except ImportError:
    import urllib.request as urllib2
    from urllib.parse import urlencode

from . import config


DEFAULT_TIMEOUT_SECONDS = 3


class ApiError(Exception):
    pass


class ApiClient(object):
    def __init__(self, settings=None):
        self.settings = settings or config.load()
        self.base_url = (self.settings.get("API_BASE_URL") or "").rstrip("/")
        self.token = self.settings.get("AUTH_TOKEN") or ""
        self.timeout = self._timeout()

    def _timeout(self):
        try:
            return max(1, int(self.settings.get("API_TIMEOUT_SECONDS") or DEFAULT_TIMEOUT_SECONDS))
        except Exception:
            return DEFAULT_TIMEOUT_SECONDS

    def request(self, method, path, payload=None, query=None):
        if not self.base_url:
            raise ApiError("API_BASE_URL is empty")
        url = self.base_url + path
        if query:
            clean = {}
            for key, value in query.items():
                if value is not None and value != "":
                    clean[key] = value
            if clean:
                url += "?" + urlencode(clean)
        body = None
        headers = {"Accept": "application/json", "Content-Type": "application/json"}
        if self.token:
            headers["Authorization"] = "Bearer " + self.token
        if payload is not None:
            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        req = urllib2.Request(url, data=body, headers=headers)
        req.get_method = lambda: method
        try:
            res = urllib2.urlopen(req, timeout=self.timeout)
            raw = res.read()
            if not raw:
                return {}
            if sys.version_info[0] < 3:
                raw = raw.decode("utf-8")
            return json.loads(raw)
        except Exception as exc:
            raise ApiError(str(exc))

    def get_projects(self):
        return self.request("GET", "/projects")

    def create_project(self, payload):
        return self.request("POST", "/projects", payload)

    def list_tasks(self, project_id, filters=None):
        if not project_id:
            raise ApiError(u"Сначала выберите проект")
        query = {"project_id": project_id}
        query.update(filters or {})
        return self.request("GET", "/tasks", query=query)

    def create_task(self, payload):
        if not payload.get("project_id"):
            raise ApiError(u"Сначала выберите проект")
        return self.request("POST", "/tasks", payload)

    def patch_task(self, task_id, project_id, payload):
        if not project_id:
            raise ApiError(u"Сначала выберите проект")
        payload = payload or {}
        payload["project_id"] = project_id
        return self.request("PATCH", "/tasks/%s" % task_id, payload)

    def add_comment(self, task_id, project_id, comment):
        if not project_id:
            raise ApiError(u"Сначала выберите проект")
        return self.request("POST", "/tasks/%s/comments" % task_id, {"project_id": project_id, "comment": comment})

    def task_changes(self, project_id, since=None):
        if not project_id:
            raise ApiError(u"Сначала выберите проект")
        return self.request("GET", "/tasks/changes", query={"project_id": project_id, "since": since})
