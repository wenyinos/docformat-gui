#!/usr/bin/env python3
"""
打包脚本 - 生成 Windows/Linux/macOS 可执行文件
用法：python build.py [windows|linux|macos|all|clean]
"""

import os
import sys
import shutil
import subprocess
import platform
import plistlib
from pathlib import Path

# 配置
APP_NAME = "公文格式处理工具"
APP_NAME_EN = "DocFormatter"
VERSION = "1.0.0"
MAIN_SCRIPT = "docformat_gui.py"

# 输出目录
DIST_DIR = Path("dist")
BUILD_DIR = Path("build")

# macOS 签名 / 公证配置（通过环境变量提供；本地开发可不设，会自动退回 ad-hoc 签名）
#   MACOS_SIGN_IDENTITY   —— "Developer ID Application: Your Name (TEAMID)" 或其证书指纹
#   MACOS_NOTARY_PROFILE  —— 预先用 notarytool store-credentials 存好的 keychain profile 名
#   （或改用 MACOS_NOTARY_APPLE_ID / MACOS_NOTARY_TEAM_ID / MACOS_NOTARY_PASSWORD 三件套）
MACOS_ENTITLEMENTS = Path("packaging/macos/entitlements.plist")


def check_pyinstaller():
    """检查 PyInstaller 是否安装"""
    try:
        import PyInstaller
        print(f"✓ PyInstaller {PyInstaller.__version__} 已安装")
        return True
    except ImportError:
        print("✗ PyInstaller 未安装")
        print("  请运行: pip install pyinstaller")
        return False


def clean():
    """清理构建目录"""
    print("\n清理旧构建文件...")
    for d in [DIST_DIR, BUILD_DIR]:
        if d.exists():
            shutil.rmtree(d)
            print(f"  删除 {d}/")
    
    # 删除 spec 文件
    for f in Path(".").glob("*.spec"):
        f.unlink()
        print(f"  删除 {f}")


def _get_docx_templates_path():
    """获取 python-docx 模板文件路径"""
    try:
        import docx
        templates_dir = Path(docx.__file__).parent / "templates"
        if templates_dir.exists():
            return str(templates_dir)
    except ImportError:
        pass
    return None


def _get_tkinterdnd_pyinstaller_args(platform_name):
    """收集 tkinterdnd2 的 tkdnd 运行库，确保打包后拖拽可用。"""
    try:
        from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs
    except Exception as e:
        print(f"  [警告] 无法加载 PyInstaller hook 工具，跳过 tkinterdnd2 资源收集: {e}")
        return []

    try:
        data_files = collect_data_files("tkinterdnd2")
        dynamic_libs = collect_dynamic_libs("tkinterdnd2")
    except Exception as e:
        print(f"  [警告] 无法收集 tkinterdnd2 资源，拖拽功能在打包版中可能不可用: {e}")
        return []

    sep = ";" if platform_name == "windows" else ":"
    args = ["--hidden-import=tkinterdnd2"]

    for src, dest in data_files:
        args.append(f"--add-data={src}{sep}{dest}")
    for src, dest in dynamic_libs:
        args.append(f"--add-binary={src}{sep}{dest}")

    if data_files or dynamic_libs:
        print(f"  ✓ 已收集 tkinterdnd2 资源: data={len(data_files)} binary={len(dynamic_libs)}")

    return args


def _patch_macos_info_plist(app_path):
    """修正 macOS app bundle 的基础元数据。"""
    plist_path = Path(app_path) / "Contents" / "Info.plist"
    if not plist_path.exists():
        return

    with open(plist_path, "rb") as f:
        data = plistlib.load(f)

    data["CFBundleName"] = APP_NAME
    data["CFBundleDisplayName"] = APP_NAME
    data["CFBundleIdentifier"] = "com.kagurananaga.docformat-gui"
    data["CFBundleDevelopmentRegion"] = "zh_CN"
    data["CFBundleLocalizations"] = ["zh_CN", "en"]

    with open(plist_path, "wb") as f:
        plistlib.dump(data, f)


def _codesign_macos(app_path):
    """对 .app 进行签名。

    若设置了 MACOS_SIGN_IDENTITY，则用真实 Developer ID 证书签名，并启用
    hardened runtime + entitlements + secure timestamp（公证的硬性前提）；
    否则退回 ad-hoc 签名（仅供本地自测，无法通过公证 / 仍会被门禁拦截）。

    返回 True 表示使用了可公证的真实签名。
    """
    identity = os.environ.get("MACOS_SIGN_IDENTITY", "").strip()
    if not identity:
        print("  [提示] 未设置 MACOS_SIGN_IDENTITY，使用 ad-hoc 签名（不会通过公证）")
        result = subprocess.run(
            ["codesign", "--force", "--deep", "--sign", "-", str(app_path)],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            print(f"  [警告] ad-hoc 签名失败：{result.stderr.strip() or result.stdout.strip()}")
        return False

    print(f"  正在用 Developer ID 证书签名: {identity}")
    cmd = [
        "codesign", "--force", "--deep",
        "--options", "runtime",          # 启用 hardened runtime（公证必需）
        "--timestamp",                    # 安全时间戳（公证必需）
        "--sign", identity,
        str(app_path),
    ]
    if MACOS_ENTITLEMENTS.exists():
        # PyInstaller 的 Python 程序在 hardened runtime 下需要 entitlements，否则运行时崩溃
        cmd[1:1] = ["--entitlements", str(MACOS_ENTITLEMENTS)]
    else:
        print(f"  [警告] 找不到 entitlements 文件: {MACOS_ENTITLEMENTS}，公证后可能无法运行")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"  [错误] 证书签名失败：{result.stderr.strip() or result.stdout.strip()}")
        return False
    print("  ✓ 证书签名完成（hardened runtime + timestamp）")
    return True


def _notarytool_auth_args():
    """组装 notarytool 的鉴权参数。优先用 keychain profile，其次用 Apple ID 三件套。"""
    profile = os.environ.get("MACOS_NOTARY_PROFILE", "").strip()
    if profile:
        return ["--keychain-profile", profile]
    apple_id = os.environ.get("MACOS_NOTARY_APPLE_ID", "").strip()
    team_id = os.environ.get("MACOS_NOTARY_TEAM_ID", "").strip()
    password = os.environ.get("MACOS_NOTARY_PASSWORD", "").strip()
    if apple_id and team_id and password:
        return ["--apple-id", apple_id, "--team-id", team_id, "--password", password]
    return None


def _notarize_and_staple(target_path):
    """把 target_path（.app 或 .dmg）提交公证并钉上票据。返回 True 表示成功。"""
    auth = _notarytool_auth_args()
    if auth is None:
        print("  [提示] 未配置公证凭据（MACOS_NOTARY_*），跳过公证")
        return False

    target = Path(target_path)
    submit_path = target
    tmp_zip = None
    # .app 是目录，notarytool 需要先压成 zip 再提交（.dmg 可直接提交）
    if target.is_dir():
        tmp_zip = target.with_suffix(".notarize.zip")
        zip_res = subprocess.run(
            ["ditto", "-c", "-k", "--keepParent", str(target), str(tmp_zip)],
            capture_output=True, text=True,
        )
        if zip_res.returncode != 0:
            print(f"  [错误] 打包待公证 zip 失败：{zip_res.stderr.strip()}")
            return False
        submit_path = tmp_zip

    print(f"  正在提交公证（等待 Apple 返回，可能需要几分钟）: {target.name}")
    submit = subprocess.run(
        ["xcrun", "notarytool", "submit", str(submit_path), "--wait"] + auth,
        capture_output=True, text=True,
    )
    if tmp_zip is not None:
        tmp_zip.unlink(missing_ok=True)
    print(submit.stdout.strip())
    if submit.returncode != 0 or "status: Accepted" not in submit.stdout:
        print(f"  [错误] 公证未通过：{submit.stderr.strip() or submit.stdout.strip()}")
        return False

    # 钉票据：必须钉在原始的 .app / .dmg 上（不是提交用的 zip）
    staple = subprocess.run(
        ["xcrun", "stapler", "staple", str(target)],
        capture_output=True, text=True,
    )
    if staple.returncode != 0:
        print(f"  [错误] 钉票据失败：{staple.stderr.strip() or staple.stdout.strip()}")
        return False
    print(f"  ✓ 公证 + 钉票据完成: {target.name}")
    return True


def build_windows():
    """构建 Windows 版本"""
    print("\n" + "=" * 50)
    print("构建 Windows 版本")
    print("=" * 50)
    
    output_name = f"docformat_windows"
    
    # 获取 docx 模板路径
    docx_tpl = _get_docx_templates_path()
    dnd_args = _get_tkinterdnd_pyinstaller_args("windows")
    
    cmd = [
        "pyinstaller",
        "--onefile",           # 单文件
        "--windowed",          # 无控制台窗口
        f"--name={output_name}",
        "--clean",
        # 添加数据文件
        "--add-data=scripts;scripts",
        # python-docx 模板文件（页眉页脚等必需）
        f"--add-data={docx_tpl};docx/templates" if docx_tpl else "--collect-data=docx",
        "--hidden-import=docx",
        "--hidden-import=lxml",
        MAIN_SCRIPT
    ]
    cmd[-1:-1] = dnd_args
    
    print(f"运行: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=False)
    
    if result.returncode == 0:
        exe_path = DIST_DIR / f"{output_name}.exe"
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"\n✓ Windows 版本构建成功!")
            print(f"  文件: {exe_path}")
            print(f"  大小: {size_mb:.1f} MB")
            return True
    
    print("\n✗ Windows 版本构建失败")
    return False


def build_linux():
    """构建 Linux 版本"""
    print("\n" + "=" * 50)
    print("构建 Linux 版本")
    print("=" * 50)
    
    output_name = f"docformat_linux"
    
    # 获取 docx 模板路径
    docx_tpl = _get_docx_templates_path()
    dnd_args = _get_tkinterdnd_pyinstaller_args("linux")
    
    cmd = [
        "pyinstaller",
        "--onefile",
        f"--name={output_name}",
        "--clean",
        "--add-data=scripts:scripts",
        # python-docx 模板文件
        f"--add-data={docx_tpl}:docx/templates" if docx_tpl else "--collect-data=docx",
        "--hidden-import=docx",
        "--hidden-import=lxml",
        MAIN_SCRIPT
    ]
    cmd[-1:-1] = dnd_args
    
    print(f"运行: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=False)
    
    if result.returncode == 0:
        exe_path = DIST_DIR / output_name
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"\n✓ Linux 版本构建成功!")
            print(f"  文件: {exe_path}")
            print(f"  大小: {size_mb:.1f} MB")
            return True
    
    print("\n✗ Linux 版本构建失败")
    return False


def build_macos():
    """构建 macOS 版本"""
    print("\n" + "=" * 50)
    print("构建 macOS 版本")
    print("=" * 50)
    
    machine = platform.machine().lower()
    is_apple_silicon = any(token in machine for token in ("arm64", "aarch64"))
    output_name = "docformat_macos_apple_silicon" if is_apple_silicon else "docformat_macos_intel"
    
    # 获取 docx 模板路径
    docx_tpl = _get_docx_templates_path()
    # 注意：macOS 打包版当前默认关闭拖拽（见 docformat_gui._should_enable_drag_drop），
    # 因此不收集 tkdnd 运行库，避免把用不到的二进制打进 .app。

    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",          # macOS 生成 .app bundle
        f"--name={output_name}",
        "--clean",
        "--osx-bundle-identifier=com.kagurananaga.docformat-gui",
        # macOS 路径分隔符与 Linux 相同
        "--add-data=scripts:scripts",
        # python-docx 模板文件
        f"--add-data={docx_tpl}:docx/templates" if docx_tpl else "--collect-data=docx",
        "--hidden-import=docx",
        "--hidden-import=lxml",
        MAIN_SCRIPT
    ]
    
    print(f"运行: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=False)
    
    if result.returncode == 0:
        # --windowed 在 macOS 上会生成 .app 目录
        app_path = DIST_DIR / f"{output_name}.app"
        bin_path = DIST_DIR / output_name
        
        if app_path.exists():
            print(f"\n✓ macOS 版本构建成功!")
            print(f"  文件: {app_path}")
            _patch_macos_info_plist(app_path)

            # 1) 签名（有证书则启用 hardened runtime，可公证；否则 ad-hoc）
            real_signed = _codesign_macos(app_path)

            # 2) 若为真实签名，先公证并钉票据到 .app 本身
            if real_signed:
                _notarize_and_staple(app_path)

            # 3) 生成 DMG（此时里面装的是已签名/已公证的 .app）
            dmg_path = DIST_DIR / f"{output_name}.dmg"
            dmg_cmd = [
                "hdiutil", "create",
                "-volname", "DocFormatter",
                "-srcfolder", str(app_path),
                "-ov", "-format", "UDZO",
                str(dmg_path)
            ]
            print(f"  正在生成 DMG...")
            dmg_result = subprocess.run(dmg_cmd, capture_output=True)
            if dmg_result.returncode == 0 and dmg_path.exists():
                size_mb = dmg_path.stat().st_size / (1024 * 1024)
                print(f"  DMG: {dmg_path} ({size_mb:.1f} MB)")
                # 4) DMG 容器本身也要公证 + 钉票据，用户从网上下载双击才不被拦
                if real_signed:
                    _notarize_and_staple(dmg_path)
            return True
        elif bin_path.exists():
            size_mb = bin_path.stat().st_size / (1024 * 1024)
            print(f"\n✓ macOS 版本构建成功!")
            print(f"  文件: {bin_path}")
            print(f"  大小: {size_mb:.1f} MB")
            return True
    
    print("\n✗ macOS 版本构建失败")
    return False


def create_release_notes():
    """生成发布说明"""
    notes = f"""# {APP_NAME} v{VERSION}

## 下载

- **Windows**: `docformat_windows.exe` - 双击运行
- **Linux (麒麟/UOS)**: `docformat_linux` - 添加执行权限后运行
- **macOS (Intel)**: `docformat_macos_intel.dmg` - 双击挂载后拖入应用程序文件夹
- **macOS (Apple Silicon)**: `docformat_macos_apple_silicon.dmg` - 双击挂载后拖入应用程序文件夹

## 功能

- ✅ 智能一键处理（标点修复 + 格式统一）
- ✅ 格式诊断
- ✅ 标点符号修复
- ✅ 支持 GB/T 公文标准、学术论文、法律文书格式

## 系统要求

- Windows 10/11 或
- 麒麟 V10 / 统信 UOS 或其他 Linux 发行版 或
- macOS 12 (Monterey) 或更高版本

## 使用说明

1. 下载对应系统的文件
2. 双击运行（Linux 需先添加执行权限）
3. 选择要处理的 .docx 文件
4. 点击「开始处理」

## 注意事项

- 仅支持 .docx 格式，不支持旧版 .doc
- 处理后的文件会另存为新文件，不会覆盖原文件
- macOS 和 Linux 版本不支持 .doc/.wps 格式转换
"""
    
    release_file = DIST_DIR / "RELEASE_NOTES.md"
    release_file.write_text(notes, encoding="utf-8")
    print(f"\n✓ 发布说明已生成: {release_file}")


def main():
    """主函数"""
    print(f"""
╔══════════════════════════════════════════╗
║     {APP_NAME} 打包工具            ║
║     版本: {VERSION}                          ║
╚══════════════════════════════════════════╝
    """)
    
    # 检查依赖
    if not check_pyinstaller():
        sys.exit(1)
    
    # 检查主脚本
    if not Path(MAIN_SCRIPT).exists():
        print(f"✗ 找不到主脚本: {MAIN_SCRIPT}")
        sys.exit(1)
    
    # 解析参数
    target = sys.argv[1] if len(sys.argv) > 1 else "all"
    
    if target not in ["windows", "linux", "macos", "all", "clean"]:
        print(f"用法: python {sys.argv[0]} [windows|linux|macos|all|clean]")
        sys.exit(1)
    
    # 清理
    if target == "clean":
        clean()
        return
    
    clean()
    
    # 构建
    success = True
    
    if target in ["windows", "all"]:
        if sys.platform == "win32":
            success = build_windows() and success
        else:
            print("\n⚠ 跳过 Windows 构建（需要在 Windows 系统上执行）")
    
    if target in ["linux", "all"]:
        if sys.platform.startswith("linux"):
            success = build_linux() and success
        else:
            print("\n⚠ 跳过 Linux 构建（需要在 Linux 系统上执行）")
    
    if target in ["macos", "all"]:
        if sys.platform == "darwin":
            success = build_macos() and success
        else:
            print("\n⚠ 跳过 macOS 构建（需要在 macOS 系统上执行）")
    
    # 生成发布说明
    if DIST_DIR.exists():
        create_release_notes()
    
    # 总结
    print("\n" + "=" * 50)
    if success:
        print("✓ 构建完成!")
        print(f"\n输出目录: {DIST_DIR.absolute()}")
        if DIST_DIR.exists():
            print("\n生成的文件:")
            for f in DIST_DIR.iterdir():
                if f.is_file():
                    size = f.stat().st_size / (1024 * 1024)
                    print(f"  - {f.name} ({size:.1f} MB)")
    else:
        print("✗ 构建过程中出现错误")
        sys.exit(1)


if __name__ == "__main__":
    main()
