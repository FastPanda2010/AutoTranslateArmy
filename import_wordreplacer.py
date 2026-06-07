from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any


def parse_js_regex(value: str) -> tuple[str, str] | None:
    if not (value.startswith("/") and value.count("/") >= 2):
        return None
    last_slash = value.rfind("/")
    pattern = value[1:last_slash]
    flags = value[last_slash + 1 :]
    return pattern, flags


def simple_pattern(value: str) -> str:
    return re.escape(value)


def convert_replacement(item: dict[str, Any]) -> dict[str, str] | None:
    if not item.get("active", True):
        return None
    source = str(item.get("repA", ""))
    replacement = str(item.get("repB", ""))
    if not source:
        return None

    if item.get("type") == "RegEx":
        parsed = parse_js_regex(source)
        if parsed:
            pattern, flags = parsed
        else:
            pattern, flags = source, "i"
    else:
        pattern, flags = simple_pattern(source), "i"

    return {
        "pattern": pattern,
        "flags": flags,
        "replacement": replacement,
        "source": source,
        "type": str(item.get("type", "")),
    }


def convert(input_path: Path, output_path: Path) -> int:
    data = json.loads(input_path.read_text(encoding="utf-8-sig"))
    rules = []
    seen = set()
    for item in data.get("replacements", []):
        rule = convert_replacement(item)
        if not rule:
            continue
        key = (rule["pattern"], rule["flags"], rule["replacement"])
        if key in seen:
            continue
        seen.add(key)
        rules.append(rule)

    output = {
        "source_format": "WordReplacer II",
        "source_version": data.get("version"),
        "rules": rules,
    }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    return len(rules)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert a WordReplacer II export into local translation rules.")
    parser.add_argument("input", type=Path, help="WordReplacer II exported txt/json file.")
    parser.add_argument("output", type=Path, nargs="?", default=Path("wordreplacer_rules.json"))
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    count = convert(args.input, args.output)
    print(f"wrote {count} rules to {args.output}")


if __name__ == "__main__":
    main()
