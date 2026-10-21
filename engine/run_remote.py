"""GitHub Actions 用：向 Google 試算表（Apps Script）拿資料 → 排課 → 把結果送回試算表。

需要兩個環境變數（放在 GitHub 的 Secrets）：
  SCHED_URL  Apps Script 網頁應用程式網址（…/exec）
  SCHED_KEY  通關密語（試算表「初次設定」時產生）

隱私：試算表送來的老師姓名已換成代號（T001…），這裡看不到真名；程式也不印出任何資料內容。
"""
import json
import os
import sys
import traceback
import urllib.parse
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from export import insights  # noqa: E402
from scheduler import load_tables, run  # noqa: E402

TIMEOUT = 120


def call(url, key, payload=None):
    """GET（payload=None）或 POST JSON；Apps Script 會 302 轉址，urllib 會自動跟過去。"""
    if payload is None:
        req = urllib.request.Request(url + "?" + urllib.parse.urlencode({"key": key, "action": "data"}))
    else:
        body = json.dumps(dict(payload, key=key), ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=TIMEOUT) as res:
        text = res.read().decode("utf-8")
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        raise RuntimeError("試算表沒有回傳正確資料：請確認網頁應用程式已部署、存取權為「所有人」。") from None


def main():
    url, key = os.environ.get("SCHED_URL", "").strip(), os.environ.get("SCHED_KEY", "").strip()
    if not url or not key:
        sys.exit("缺少 SCHED_URL 或 SCHED_KEY，請到 GitHub → Settings → Secrets and variables → Actions 設定。")
    job = call(url, key)
    if not job.get("ok"):
        sys.exit(f"試算表拒絕提供資料：{job.get('error', '未知原因')}")
    job_id = job["jobId"]
    try:
        data = load_tables(job["tables"])
        result = run(data, job.get("source", "Google 試算表"), time_limit=int(job.get("timeLimit", 60)),
                     log=lambda m: None)  # 不印任何內容，避免公開的執行紀錄出現資料
        if "rules" in result:
            result["insights"] = insights(result)
        reply = call(url, key, {"action": "result", "jobId": job_id, "result": result})
    except Exception as e:  # 任何錯誤都回報給試算表，讓 Bing 在試算表上就看得到
        msg = str(e) if isinstance(e, (ValueError, RuntimeError)) else f"程式錯誤：{type(e).__name__}"
        traceback.print_exception(type(e), e, None)  # 只印錯誤類型與位置，不印資料
        call(url, key, {"action": "error", "jobId": job_id, "error": msg})
        sys.exit(1)
    if not reply.get("ok"):
        sys.exit(f"結果送回試算表失敗：{reply.get('error', '未知原因')}")
    print(f"完成：{result['meta'].get('status')}，{result['meta'].get('seconds')} 秒")


if __name__ == "__main__":
    main()
