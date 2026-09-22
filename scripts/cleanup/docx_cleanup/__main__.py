"""Data cleanup stage: docx in -> docx out, driven by a profile.

  python -m docx_cleanup input.docx -o input.clean.docx --profile ns
  python -m docx_cleanup input.docx -o input.clean.docx --config my.json --mode Report

Exit code: 0 = ok (also when nothing matched), 1 = error.
"""
from __future__ import annotations

import argparse
import csv
import json
import shutil
import sys
from pathlib import Path
from typing import Any, Optional

from .model import MODES, Finding
from .package import DocxPackage
from .registry import RULES

PROFILES_DIR = Path(__file__).resolve().parents[2] / "profiles"


def load_config(profile: Optional[str], config: Optional[str]) -> dict[str, Any]:
    if config:
        path = Path(config)
    else:
        path = PROFILES_DIR / f"{profile or 'default'}.json"
    if not path.is_file():
        raise ValueError(f"Config not found: {path}")
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    rules = data.get("rules", {})
    for rule_id, cfg in rules.items():
        if rule_id not in RULES:
            raise ValueError(f"{path.name}: unknown rule '{rule_id}'; known: {sorted(RULES)}")
        if cfg.get("mode", "Off") not in MODES:
            raise ValueError(f"{path.name}: rule '{rule_id}' has invalid mode '{cfg.get('mode')}'; use {MODES}")
    data["_path"] = str(path)
    return data


def enabled_rules(config: dict[str, Any], mode_override: Optional[str]) -> list[tuple[str, str, dict]]:
    """(rule_id, mode, params) for every rule that is not Off. The override never enables an Off rule."""
    out = []
    for rule_id, cfg in config.get("rules", {}).items():
        mode = cfg.get("mode", "Off")
        if mode == "Off":
            continue
        if mode_override:
            mode = mode_override
        params = {k: v for k, v in cfg.items() if k != "mode"}
        out.append((rule_id, mode, params))
    return out


def write_manifest(path: Path, findings: list[Finding]) -> None:
    fields = ["RuleId", "Action", "Confidence", "Part", "Paragraph", "ShapeId", "ShapeName",
              "Container", "Reason", "Location", "Features"]
    # utf-8-sig so that Excel shows Vietnamese/Japanese text correctly
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=fields)
        writer.writeheader()
        for f in findings:
            writer.writerow(f.to_row())


def run(input_path: Path, output_path: Path, config: dict[str, Any],
        mode_override: Optional[str] = None, manifest: Optional[Path] = None) -> list[Finding]:
    if input_path.resolve() == output_path.resolve():
        raise ValueError("Output must differ from input.")
    rules = enabled_rules(config, mode_override)
    pkg = DocxPackage(input_path)
    findings: list[Finding] = []

    for rule_id, mode, params in rules:
        rule = RULES[rule_id](params)
        found = rule.detect(pkg)
        if mode == "Apply":
            rule.apply(pkg, found)
        findings.extend(found)
        changed = sum(1 for f in found if not f.action.startswith("reported"))
        print(f"  {rule_id} [{mode}]: {len(found)} found, {changed} changed")

    if pkg.dirty_parts:
        pkg.save(output_path)
    else:
        shutil.copyfile(input_path, output_path)

    manifest = manifest or output_path.with_name(output_path.stem + ".cleanup-manifest.csv")
    write_manifest(manifest, findings)
    print(f"Output:   {output_path}")
    print(f"Manifest: {manifest}")
    return findings


def main(argv: Optional[list[str]] = None) -> int:
    ap = argparse.ArgumentParser(prog="docx_cleanup", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", type=Path, nargs="?")
    ap.add_argument("-o", "--output", type=Path,
                    help="default: <input>.clean.docx next to the input")
    src = ap.add_mutually_exclusive_group()
    src.add_argument("--profile", help=f"profile name in {PROFILES_DIR} (default: 'default')")
    src.add_argument("--config", help="path to a profile JSON file")
    ap.add_argument("--mode", choices=["Report", "Apply"],
                    help="override the mode of every enabled rule (e.g. Report for a dry run)")
    ap.add_argument("--manifest", type=Path, help="default: <output>.cleanup-manifest.csv")
    ap.add_argument("--list-rules", action="store_true", help="list known rules and their defaults")
    args = ap.parse_args(argv)

    if args.list_rules:
        for rule_id, cls in RULES.items():
            print(f"{rule_id}: {cls.description}")
            print("  defaults: " + json.dumps(cls.defaults))
        return 0
    if args.input is None:
        ap.error("the following arguments are required: input")

    try:
        config = load_config(args.profile, args.config)
        out = args.output or args.input.with_name(args.input.stem + ".clean.docx")
        print(f"Cleanup:  {args.input}  (config: {config['_path']})")
        if not enabled_rules(config, None):
            print("  no rules enabled - output is a copy of the input")
        run(args.input, out, config, args.mode, args.manifest)
        return 0
    except Exception as exc:  # noqa: BLE001 - report any failure as exit code 1
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
