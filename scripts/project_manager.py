#!/usr/bin/env python3
"""
项目管理脚本 — 新建、归档、恢复项目，并同步 project.json

用法：
  python3 scripts/project_manager.py new <name> [description]
  python3 scripts/project_manager.py archive <name> [reason]
  python3 scripts/project_manager.py restore <name>
  python3 scripts/project_manager.py list
  python3 scripts/project_manager.py sync

命令说明：
  new       新建项目文件夹结构 + 注册到 project.json，自动设为 active
  archive   软归档：将文件夹移至 _archived/，记录保留在 project.json → archived 块
  restore   恢复归档：将文件夹移回 projects/，重新注册为 active
  list      列出所有活跃项目和归档记录
  sync      扫描 projects/ 文件夹，将未注册的项目补入 project.json
"""

import sys
import os
import json
import shutil
from datetime import date

_BRAND      = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
_PROJECTS   = os.path.join(_BRAND, 'projects')
_ARCHIVED   = os.path.join(_BRAND, '_archived')
_PROJ_JSON  = os.path.join(_BRAND, 'project.json')


# ── 工具函数 ──────────────────────────────────────────────────────────────────

def load_json():
    with open(_PROJ_JSON, encoding='utf-8') as f:
        return json.load(f)


def save_json(data):
    with open(_PROJ_JSON, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"  ✓ project.json 已更新")


def today():
    return date.today().isoformat()


# ── 模板 ─────────────────────────────────────────────────────────────────────

CONTEXT_TEMPLATE = """\
# {name} · context.md

## 角色定义
（描述本项目的 AI 角色，如：你是 ArktechX 的品牌内容编辑，负责...）

## 目标受众
（目标受众描述，如：面向企业高管、投资人...）

## 输出风格
- 语言：中文 / 英文 / 双语
- 语气：专业、简洁、结果导向

## 领域知识
（填写本项目相关的行业背景知识）

## 输出约束
- 不夸大未验证的数字
- 保持品牌调性：前沿、专业、落地、负责任
- 硬编码联系方式禁止写入脚本（使用 BT.COMPANY_* 常量）
"""

MANIFEST_TEMPLATE = """\
# 素材清单 MANIFEST.md

> 本文件人工维护，列出 media/ 下每个素材的用途。

| 文件名 | 用于 | 备注 |
|---|---|---|
| （示例）cover.jpg | docs/output.pptx 封面 | 1920×1080 |
"""


# ── 命令实现 ──────────────────────────────────────────────────────────────────

def cmd_new(args):
    if not args:
        print("用法：python3 scripts/project_manager.py new <name> [description]")
        sys.exit(1)

    name        = args[0]
    description = " ".join(args[1:]) if len(args) > 1 else ""
    data        = load_json()

    if name in data.get('projects', {}):
        print(f"✗ 项目 '{name}' 已存在于 project.json")
        sys.exit(1)

    if name in data.get('archived', {}):
        print(f"✗ 项目 '{name}' 在归档记录中，请先执行 restore 恢复")
        sys.exit(1)

    project_dir = os.path.join(_PROJECTS, name)

    if os.path.exists(project_dir):
        print(f"  ⚠ 文件夹 projects/{name}/ 已存在，跳过创建，仅注册到 project.json")
    else:
        for sub in ['docs/img', 'jobs', 'media', 'references']:
            os.makedirs(os.path.join(project_dir, sub), exist_ok=True)
        with open(os.path.join(project_dir, 'context.md'), 'w', encoding='utf-8') as f:
            f.write(CONTEXT_TEMPLATE.format(name=name))
        with open(os.path.join(project_dir, 'media', 'MANIFEST.md'), 'w', encoding='utf-8') as f:
            f.write(MANIFEST_TEMPLATE)
        print(f"  ✓ 已创建 projects/{name}/ 文件夹结构")

    data.setdefault('projects', {})[name] = {
        "name": name,
        "description": description,
        "created": today()
    }
    data['active'] = name
    save_json(data)

    print(f"  ✓ 已新建项目 '{name}'，并设为 active")
    print(f"  → 下一步：编辑 projects/{name}/context.md 填写项目上下文")


def cmd_archive(args):
    if not args:
        print("用法：python3 scripts/project_manager.py archive <name> [reason]")
        sys.exit(1)

    name   = args[0]
    reason = " ".join(args[1:]) if len(args) > 1 else ""
    data   = load_json()

    if name not in data.get('projects', {}):
        if name in data.get('archived', {}):
            print(f"✗ 项目 '{name}' 已经在归档记录中")
        else:
            print(f"✗ project.json 中找不到项目 '{name}'")
        sys.exit(1)

    project_dir = os.path.join(_PROJECTS, name)
    archived_dir = os.path.join(_ARCHIVED, name)
    os.makedirs(_ARCHIVED, exist_ok=True)

    if os.path.exists(project_dir):
        if os.path.exists(archived_dir):
            print(f"✗ _archived/{name}/ 已存在，请手动处理")
            sys.exit(1)
        shutil.move(project_dir, archived_dir)
        print(f"  ✓ 已移动 projects/{name}/ → _archived/{name}/")
    else:
        print(f"  ⚠ projects/{name}/ 文件夹不存在，仅更新 project.json")

    entry = data['projects'].pop(name)
    data.setdefault('archived', {})[name] = {
        **entry,
        "archived_at": today(),
        "reason": reason
    }

    if data.get('active') == name:
        data['active'] = ""
        print(f"  ⚠ '{name}' 是当前 active 项目，active 已清空，请手动指定新项目")

    save_json(data)
    print(f"  ✓ 已归档项目 '{name}'")
    print(f"    记录保存在 project.json → archived.{name}")


def cmd_restore(args):
    if not args:
        print("用法：python3 scripts/project_manager.py restore <name>")
        sys.exit(1)

    name = args[0]
    data = load_json()

    if name not in data.get('archived', {}):
        print(f"✗ 归档记录中找不到项目 '{name}'")
        sys.exit(1)

    archived_dir = os.path.join(_ARCHIVED, name)
    project_dir  = os.path.join(_PROJECTS, name)

    if os.path.exists(archived_dir):
        if os.path.exists(project_dir):
            print(f"✗ projects/{name}/ 已存在，请手动处理冲突")
            sys.exit(1)
        shutil.move(archived_dir, project_dir)
        print(f"  ✓ 已恢复 _archived/{name}/ → projects/{name}/")
    else:
        print(f"  ⚠ _archived/{name}/ 文件夹不存在，仅更新 project.json")

    entry = data['archived'].pop(name)
    entry.pop('archived_at', None)
    entry.pop('reason', None)
    data.setdefault('projects', {})[name] = entry
    data['active'] = name

    save_json(data)
    print(f"  ✓ 已恢复项目 '{name}'，并设为 active")


def cmd_list(args):
    data     = load_json()
    active   = data.get('active', '')
    projects = data.get('projects', {})
    archived = data.get('archived', {})

    width = 52
    print(f"\n{'─' * width}")
    print(f"  活跃项目（{len(projects)} 个）")
    print(f"{'─' * width}")

    if not projects:
        print("  （无）")
    for key, p in projects.items():
        is_active = key == active
        marker    = " ← active" if is_active else ""
        prefix    = "★" if is_active else "·"
        print(f"  {prefix} {key}{marker}")
        if p.get('description'):
            desc = p['description']
            print(f"      {desc[:58]}{'…' if len(desc) > 58 else ''}")
        print(f"      创建：{p.get('created', '未知')}")

    print(f"\n{'─' * width}")
    print(f"  已归档项目（{len(archived)} 个）")
    print(f"{'─' * width}")

    if not archived:
        print("  （无）")
    for key, p in archived.items():
        print(f"  ✗ {key}")
        reason_str = f"  原因：{p['reason']}" if p.get('reason') else ""
        print(f"      归档日期：{p.get('archived_at', '未知')}{reason_str}")

    print(f"{'─' * width}\n")


def cmd_sync(args):
    data = load_json()
    registered = set(data.get('projects', {}).keys()) | set(data.get('archived', {}).keys())

    added = []
    for name in sorted(os.listdir(_PROJECTS)):
        if name.startswith('.'):
            continue
        if name not in registered:
            data.setdefault('projects', {})[name] = {
                "name": name,
                "description": "",
                "created": today()
            }
            added.append(name)

    if added:
        save_json(data)
        print(f"  ✓ 已补充 {len(added)} 个未注册项目：{', '.join(added)}")
    else:
        print("  ✓ project.json 已与文件系统同步，无需更新")


# ── 入口 ─────────────────────────────────────────────────────────────────────

COMMANDS = {
    'new':     cmd_new,
    'archive': cmd_archive,
    'restore': cmd_restore,
    'list':    cmd_list,
    'sync':    cmd_sync,
}


def main():
    if len(sys.argv) < 2 or sys.argv[1] not in COMMANDS:
        print(__doc__)
        print(f"可用命令：{', '.join(COMMANDS.keys())}")
        sys.exit(1)
    COMMANDS[sys.argv[1]](sys.argv[2:])


if __name__ == '__main__':
    main()
