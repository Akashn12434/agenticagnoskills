#!/usr/bin/env python3
"""Verify excel-skill structure."""

import sys
from pathlib import Path


def main():
    skill_dir = Path(__file__).parent.parent

    required_files = [
        skill_dir / "SKILL.md",
        skill_dir / "references" / "guide.md",
        skill_dir / "scripts" / "create_xlsx.py",
        skill_dir / "scripts" / "read_xlsx.py",
    ]

    missing_files = [str(path.relative_to(skill_dir)) for path in required_files if not path.exists()]
    if missing_files:
        print(f"ERROR: Missing files: {missing_files}")
        sys.exit(1)

    skill_content = (skill_dir / "SKILL.md").read_text(encoding="utf-8")
    guide_content = (skill_dir / "references" / "guide.md").read_text(encoding="utf-8")

    required_skill_sections = [
        "Excel Spreadsheet Skill",
        "When to Use",
        "Workbook Creation Rules",
        "Reading and Extraction Rules",
        "Design Requirements",
        "Validation Checklist",
        "Script Usage",
        "Best Practices",
    ]
    missing_skill = [section for section in required_skill_sections if section not in skill_content]
    if missing_skill:
        print(f"ERROR: Missing SKILL.md sections: {missing_skill}")
        sys.exit(1)

    required_guide_sections = [
        "Standard Workbook Structure",
        "Table Formatting",
        "Number Formatting",
        "Formula Guidelines",
        "Chart Guidelines",
        "Spreadsheet Templates",
        "Spreadsheet Guardrails",
    ]
    missing_guide = [section for section in required_guide_sections if section not in guide_content]
    if missing_guide:
        print(f"ERROR: Missing guide.md sections: {missing_guide}")
        sys.exit(1)

    print("OK: excel-skill ready")
    sys.exit(0)


if __name__ == "__main__":
    main()
