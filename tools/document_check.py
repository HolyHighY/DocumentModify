#!/usr/bin/env python3
"""Offline checks only: no network access, decryption, or WPS automation."""

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path, PurePosixPath
import platform
import re
import sys
import xml.etree.ElementTree as ET
import zipfile


def write_json(path, data):
    path = Path(path)
    if path.suffix.lower() != ".json" or path.is_symlink():
        raise ValueError("Output must be a regular .json path")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n",
                    encoding="utf-8")


def file_info(path):
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            size += len(chunk)
            digest.update(chunk)
    return {"size": size, "sha256": digest.hexdigest()}


def inventory(root, excluded):
    root = Path(root).resolve(strict=True)
    if not root.is_dir():
        raise ValueError("Root must be a directory")
    result = {}

    def fail(error):
        raise error

    for current, dirs, files in os.walk(root, followlinks=False, onerror=fail):
        for name in sorted(dirs + files):
            path = Path(current) / name
            if path.is_symlink():
                raise ValueError("Symlinks are not supported in writing packages")
            if path.resolve() == excluded:
                continue
            if name in dirs:
                continue
            if not path.is_file():
                raise ValueError("Package contains a non-regular file")
            result[path.relative_to(root).as_posix()] = file_info(path)
    return dict(sorted(result.items()))


def manifest(args):
    output = Path(args.output)
    # The chosen manifest itself is excluded so regeneration is repeatable.
    files = inventory(args.root, output.resolve())
    write_json(output, {"format": "DocumentModify-manifest-v1",
                        "algorithm": "sha256", "files": files})
    print(f"PASS: wrote manifest for {len(files)} files")
    return 0


def verify(args):
    path = Path(args.manifest)
    data = json.loads(path.read_text(encoding="utf-8"))
    if (data.get("format") != "DocumentModify-manifest-v1"
            or data.get("algorithm") != "sha256"
            or not isinstance(data.get("files"), dict)):
        raise ValueError("Unsupported manifest format")
    expected = data["files"]
    for name, info in expected.items():
        p = PurePosixPath(name)
        if (not name or p.is_absolute() or ".." in p.parts
                or "\\" in name or ":" in name or p.as_posix() != name):
            raise ValueError("Manifest contains an unsafe or non-canonical path")
        if (not isinstance(info, dict)
                or type(info.get("size")) is not int or info["size"] < 0
                or not isinstance(info.get("sha256"), str)
                or not re.fullmatch(r"[0-9a-f]{64}", info["sha256"])):
            raise ValueError("Invalid file information in manifest")
    actual = inventory(args.root, path.resolve())
    missing = sorted(expected.keys() - actual.keys())
    extra = sorted(actual.keys() - expected.keys())
    changed = sorted(k for k in expected.keys() & actual.keys()
                     if expected[k] != actual[k])
    if missing or extra or changed:
        print(json.dumps({"status": "FAIL", "missing": missing,
                          "extra": extra, "changed": changed},
                         ensure_ascii=False, indent=2))
        return 1
    print(f"PASS: {len(actual)} files match; no extra files")
    return 0


def probe(args):
    modules = {}
    for name in ("docx", "win32com", "PIL", "matplotlib", "pypdf"):
        try:
            modules[name] = importlib.util.find_spec(name) is not None
        except (ImportError, ValueError):
            modules[name] = False
    write_json(args.output, {"system": platform.system(),
                             "release": platform.release(),
                             "machine": platform.machine(),
                             "python": platform.python_version(),
                             "python_bits": 64 if sys.maxsize > 2**32 else 32,
                             "modules_detected_not_verified": modules,
                             "wps_automation": "NOT_TESTED",
                             "encrypted_document_access": "NOT_TESTED",
                             "visual_layout": "NOT_TESTED"})
    print("PASS: environment inventory saved; application capabilities untested")
    return 0


def inspect_docx(args):
    source, output = Path(args.input), Path(args.output)
    if source.resolve() == output.resolve():
        raise ValueError("Inspection output must not replace the input")
    result = {"input_name": source.name, **file_info(source),
              "status": "INCONCLUSIVE", "visual_layout": "NOT_TESTED"}
    try:
        with zipfile.ZipFile(source) as archive:
            names = archive.namelist()
            if len(names) != len(set(names)):
                raise ValueError("Duplicate ZIP members")
            required = ("[Content_Types].xml", "_rels/.rels", "word/document.xml")
            for name in required:
                member = archive.getinfo(name)
                if member.file_size > 16 * 1024 * 1024:
                    raise ValueError("XML exceeds this basic check's 16 MiB limit")
                raw = archive.read(name)
                if b"<!DOCTYPE" in raw.upper() or b"<!ENTITY" in raw.upper():
                    raise ValueError("DTD/entity declarations are not supported")
                element = ET.fromstring(raw)
                if name == "word/document.xml":
                    document = element
            w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
            if document.tag != w + "document":
                raise ValueError("Unsupported main document XML namespace/root")
            result.update({"status": "PARSEABLE_BASIC_DOCX",
                           "paragraphs": len(list(document.iter(w + "p"))),
                           "tables": len(list(document.iter(w + "tbl"))),
                           "media_members": sum(n.startswith("word/media/")
                                                and not n.endswith("/") for n in names),
                           "headers": sum(bool(re.fullmatch(r"word/header\d+\.xml", n))
                                          for n in names),
                           "footers": sum(bool(re.fullmatch(r"word/footer\d+\.xml", n))
                                          for n in names)})
    except (zipfile.BadZipFile, KeyError, ET.ParseError, RuntimeError, ValueError,
            NotImplementedError) as error:
        result["reason"] = str(error)
        result["note"] = "Cannot establish ordinary DOCX structure; encryption is not diagnosed"
    write_json(output, result)
    if result["status"] != "PARSEABLE_BASIC_DOCX":
        print("INCONCLUSIVE: inspect local JSON; use an authorized application route")
        return 2
    print("PASS: basic DOCX structure parsed; semantic preservation/layout untested")
    return 0


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    command = commands.add_parser("probe", help="Record non-secret platform/module metadata")
    command.add_argument("--output", required=True)
    command.set_defaults(run=probe)
    command = commands.add_parser("manifest", help="Hash package files; omit this manifest")
    command.add_argument("--root", required=True)
    command.add_argument("--output", required=True)
    command.set_defaults(run=manifest)
    command = commands.add_parser("verify", help="Reject missing, changed, or extra files")
    command.add_argument("--root", required=True)
    command.add_argument("--manifest", required=True)
    command.set_defaults(run=verify)
    command = commands.add_parser("inspect-docx", help="Inspect ordinary DOCX without extracting text")
    command.add_argument("--input", required=True)
    command.add_argument("--output", required=True)
    command.set_defaults(run=inspect_docx)
    args = parser.parse_args()
    try:
        return args.run(args)
    except (OSError, ValueError, TypeError, AttributeError) as error:
        print(f"FAIL: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
