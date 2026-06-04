"""
版本号自动递增工具
用法:
    python bump_version.py minor    # v3.0 → v3.1  (功能更新)
    python bump_version.py patch    # v3.0 → v3.0.1 (Bug修复)
    python bump_version.py major    # v3.0 → v4.0  (大版本)

工作流程:
    1. 读取当前版本号 (gui.py 中的 VERSION)
    2. 按指定规则递增
    3. 更新 gui.py 和 version.json
    4. 自动 git add + commit + tag
    5. 提示 git push
"""

import re
import sys
import json
import subprocess
from pathlib import Path

ROOT = Path(__file__).parent

def read_current_version():
    """从 gui.py 读取当前版本号"""
    src = (ROOT / "gui.py").read_text(encoding="utf-8")
    m = re.search(r'^VERSION\s*=\s*"([^"]+)"', src, re.MULTILINE)
    if not m:
        print("[ERROR] 无法从 gui.py 读取 VERSION")
        sys.exit(1)
    return m.group(1)

def bump_version(current, mode):
    """递增版本号"""
    v = current.lstrip("v")
    parts = v.split(".")

    if mode == "major":
        # v3.0 → v4.0
        parts = [str(int(parts[0]) + 1), "0"]
    elif mode == "minor":
        if len(parts) == 2:
            # v3.0 → v3.1
            parts = [parts[0], str(int(parts[1]) + 1)]
        else:
            # v3.0.1 → v3.1
            parts = [parts[0], str(int(parts[1]) + 1)]
    elif mode == "patch":
        if len(parts) == 2:
            # v3.0 → v3.0.1
            parts = [parts[0], parts[1], "1"]
        else:
            # v3.0.1 → v3.0.2
            parts[-1] = str(int(parts[-1]) + 1)
    else:
        print(f"[ERROR] 未知模式: {mode}，可用: major / minor / patch")
        sys.exit(1)

    return "v" + ".".join(parts)

def update_files(new_version):
    """更新 gui.py 和 version.json 中的版本号"""
    # 更新 gui.py
    gui_path = ROOT / "gui.py"
    src = gui_path.read_text(encoding="utf-8")
    src = re.sub(
        r'^(VERSION\s*=\s*)"[^"]*"',
        lambda m: f'{m.group(1)}"{new_version}"',
        src,
        count=1,
        flags=re.MULTILINE
    )
    gui_path.write_text(src, encoding="utf-8")
    print(f"[OK] gui.py → {new_version}")

    # 更新 version.json
    ver_path = ROOT / "version.json"
    if ver_path.exists():
        ver = json.loads(ver_path.read_text(encoding="utf-8"))
        ver["version"] = new_version
        ver["updateDate"] = __import__("datetime").date.today().isoformat()
        ver_path.write_text(
            json.dumps(ver, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8"
        )
        print(f"[OK] version.json → {new_version}")

def git_commit_and_tag(new_version):
    """自动 git commit + tag"""
    try:
        subprocess.run(
            ["git", "-C", str(ROOT), "add", "gui.py", "version.json"],
            check=True, capture_output=True
        )
        subprocess.run(
            ["git", "-C", str(ROOT), "commit", "-m", f"chore: bump version to {new_version}"],
            check=True, capture_output=True
        )
        subprocess.run(
            ["git", "-C", str(ROOT), "tag", "-a", new_version, "-m", f"Version {new_version}"],
            check=True, capture_output=True
        )
        print(f"[OK] git commit + tag {new_version} 已创建")
        print()
        print("下一步推送:")
        print(f"  git push origin {new_version}")
        print(f"  git push origin main")
    except subprocess.CalledProcessError as e:
        stderr = e.stderr.decode() if e.stderr else ""
        if "nothing to commit" in stderr:
            print("[!] 没有变更需要提交")
        elif "already exists" in stderr:
            print(f"[!] Tag {new_version} 已存在")
        else:
            print(f"[!] Git 操作警告: {stderr[:200]}")

if __name__ == "__main__":
    if len(sys.argv) < 2 or sys.argv[1] not in ("major", "minor", "patch"):
        print(__doc__)
        sys.exit(1)

    mode = sys.argv[1]
    current = read_current_version()
    new_ver = bump_version(current, mode)

    print(f"  {current} → {new_ver}  ({mode})")
    print()

    update_files(new_ver)
    git_commit_and_tag(new_ver)

    print()
    print(f"如需回退到此版本: git checkout tags/{new_ver}")
