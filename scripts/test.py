#!/usr/bin/env python3
"""
测试 WPS 统一 Skill 桥接服务连接和基本功能（Excel + PPT + Word）
用法：python test.py [host] [port]
"""

import sys
import os
import json
import urllib.request

BRIDGE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "bridge")
if BRIDGE_DIR not in sys.path:
    sys.path.insert(0, BRIDGE_DIR)

from service_lifecycle import inspect_health

HOST = sys.argv[1] if len(sys.argv) > 1 else "127.0.0.1"
PORT = sys.argv[2] if len(sys.argv) > 2 else "58891"
BASE_URL = f"http://{HOST}:{PORT}"
LOOPBACK_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def api_get(path):
    try:
        req = urllib.request.Request(f"{BASE_URL}{path}")
        with LOOPBACK_OPENER.open(req, timeout=10) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        return {"success": False, "error": str(e)}


def api_post(path, data, service_health):
    try:
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(
            f"{BASE_URL}{path}", data=body,
            headers={
                "Content-Type": "application/json",
                "X-WPS-Bridge-Project-Id": service_health["projectId"],
                "X-WPS-Bridge-Instance-Id": service_health["instanceId"],
            },
        )
        with LOOPBACK_OPENER.open(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        return {"success": False, "error": str(e)}


def main():
    print("=" * 60)
    print("  WPS 统一 Skill 桥接服务测试（Excel + PPT + Word）")
    print("=" * 60)
    print(f"  目标: {BASE_URL}")
    print()

    print("[1] 健康检查 GET /health")
    r = api_get("/health")
    print(f"    结果: {json.dumps(r, ensure_ascii=False)}")
    if r.get("status") != "ok":
        print("    [失败] 桥接服务未运行，请先执行: python scripts/start.py")
        return
    if not (r.get("projectId") and r.get("instanceId")):
        print("    [失败] 这是缺少实例身份的旧版 bridge，请保存工作后停止旧服务并重试")
        return
    inspection = inspect_health(r)
    if not inspection.reusable:
        print(f"    [失败] {inspection.error}")
        return
    service_health = r
    print()

    print("[2] 获取 action 列表 GET /actions")
    r = api_get("/actions")
    if r.get("actions") is not None:
        print(f"    可用 action 数量: {r.get('count')}")
        by_app = {}
        for a in r["actions"]:
            by_app[a["app"]] = by_app.get(a["app"], 0) + 1
        print(f"    按应用分布: {by_app}")
    else:
        print(f"    [失败] {r}")
    print()

    print("[3] 测试跨应用连接 POST /dispatch {action: ping}")
    r = api_post("/dispatch", {"action": "ping", "params": {}}, service_health)
    print(f"    结果: {json.dumps(r, ensure_ascii=False)}")
    print()

    print("[4] Excel 连接 GET AppInfo {action: getAppInfo, app: excel}")
    r = api_post(
        "/dispatch",
        {"action": "getAppInfo", "params": {"app": "excel"}},
        service_health,
    )
    print(f"    结果: {json.dumps(r, ensure_ascii=False)}")
    print()

    print("[5] PPT 连接 GET AppInfo {action: getAppInfo, app: ppt}")
    r = api_post(
        "/dispatch",
        {"action": "getAppInfo", "params": {"app": "ppt"}},
        service_health,
    )
    print(f"    结果: {json.dumps(r, ensure_ascii=False)}")
    print()

    print("[6] Word 连接 GET AppInfo {action: getAppInfo, app: word}")
    r = api_post(
        "/dispatch",
        {"action": "getAppInfo", "params": {"app": "word"}},
        service_health,
    )
    print(f"    结果: {json.dumps(r, ensure_ascii=False)}")
    print()

    print("=" * 60)
    print("  测试完成")
    print("=" * 60)


if __name__ == "__main__":
    main()
