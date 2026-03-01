# coding: utf-8
import json
import os

import clr
clr.AddReference("System")

from System import IO
from System import Text
from System.Net import HttpWebRequest, WebException

API_BASE = "https://api-cpsk-superapp.gip.su/api/gip-vision/v1"


def _read_response_text(response):
    stream = response.GetResponseStream()
    reader = IO.StreamReader(stream, Text.Encoding.UTF8)
    text = reader.ReadToEnd()
    reader.Close()
    stream.Close()
    return text


def _json_or_raw(text):
    try:
        return json.loads(text), None
    except Exception:
        return None, text


def _build_multipart_body(file_path, field_name):
    boundary = "---------------------------gipvision{0}".format(str(abs(hash(file_path))))
    file_name = os.path.basename(file_path)
    file_bytes = IO.File.ReadAllBytes(file_path)

    preamble = (
        "--{0}\r\n"
        "Content-Disposition: form-data; name=\"{1}\"; filename=\"{2}\"\r\n"
        "Content-Type: application/octet-stream\r\n\r\n"
    ).format(boundary, field_name, file_name)
    ending = "\r\n--{0}--\r\n".format(boundary)

    pre_bytes = Text.Encoding.UTF8.GetBytes(preamble)
    end_bytes = Text.Encoding.UTF8.GetBytes(ending)

    body = IO.MemoryStream()
    body.Write(pre_bytes, 0, pre_bytes.Length)
    body.Write(file_bytes, 0, file_bytes.Length)
    body.Write(end_bytes, 0, end_bytes.Length)

    content = body.ToArray()
    body.Close()

    return boundary, content


def create_session_by_plane(model_file_path, api_key):
    url = API_BASE + "/session/by_plane/"
    boundary, body = _build_multipart_body(model_file_path, "model_file")

    req = HttpWebRequest.Create(url)
    req.Method = "POST"
    req.ContentType = "multipart/form-data; boundary={0}".format(boundary)
    req.ContentLength = body.Length
    req.Timeout = 1000 * 60 * 5

    if api_key:
        req.Headers.Add("X-API-Key", api_key)

    req_stream = req.GetRequestStream()
    req_stream.Write(body, 0, body.Length)
    req_stream.Close()

    try:
        resp = req.GetResponse()
        text = _read_response_text(resp)
        resp.Close()
        data, raw = _json_or_raw(text)
        return {
            "ok": True,
            "status": 201,
            "data": data,
            "raw": raw
        }
    except WebException as ex:
        status = 0
        text = ""
        if ex.Response:
            try:
                status = int(ex.Response.StatusCode)
                text = _read_response_text(ex.Response)
                ex.Response.Close()
            except Exception:
                pass
        data, raw = _json_or_raw(text)
        return {
            "ok": False,
            "status": status,
            "data": data,
            "raw": raw,
            "error": str(ex)
        }


def resolve_onetime_code(onetime_code, api_key):
    url = API_BASE + "/session/resolve-by-onetime-code/"
    payload = json.dumps({"onetime_code": onetime_code})
    payload_bytes = Text.Encoding.UTF8.GetBytes(payload)

    req = HttpWebRequest.Create(url)
    req.Method = "POST"
    req.ContentType = "application/json"
    req.ContentLength = payload_bytes.Length
    req.Timeout = 1000 * 30

    if api_key:
        req.Headers.Add("X-API-Key", api_key)

    req_stream = req.GetRequestStream()
    req_stream.Write(payload_bytes, 0, payload_bytes.Length)
    req_stream.Close()

    try:
        resp = req.GetResponse()
        text = _read_response_text(resp)
        resp.Close()
        data, raw = _json_or_raw(text)
        return {
            "ok": True,
            "status": 200,
            "data": data,
            "raw": raw
        }
    except WebException as ex:
        status = 0
        text = ""
        if ex.Response:
            try:
                status = int(ex.Response.StatusCode)
                text = _read_response_text(ex.Response)
                ex.Response.Close()
            except Exception:
                pass
        data, raw = _json_or_raw(text)
        return {
            "ok": False,
            "status": status,
            "data": data,
            "raw": raw,
            "error": str(ex)
        }
