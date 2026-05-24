#!/usr/bin/env python3
"""Verify word-skill structure."""

import sys
from pathlib import Path


def main():
    skill_dir = Path(__file__).parent.parent

    required_files = [
        skill_dir / "SKILL.md",
        skill_dir / "references" / "guide.md",
        skill_dir / "scripts" / "create_docx.py",
        skill_dir / "scripts" / "read_docx.py",
    ]

    missing_files = [str(path.relative_to(skill_dir)) for path in required_files if not path.exists()]
    if missing_files:
        print(f"ERROR: Missing files: {missing_files}")
        sys.exit(1)

    skill_content = (skill_dir / "SKILL.md").read_text(encoding="utf-8")
    guide_content = (skill_dir / "references" / "guide.md").read_text(encoding="utf-8")

    required_skill_sections = [
        "Word Document Skill",
        "When to Use",
        "Document Creation Rules",
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
        "Standard Document Structure",
        "Content Density",
        "Heading Guidelines",
        "Professional Formatting",
        "Document Templates",
        "Readability Guardrails",
    ]
    missing_guide = [section for section in required_guide_sections if section not in guide_content]
    if missing_guide:
        print(f"ERROR: Missing guide.md sections: {missing_guide}")
        sys.exit(1)

    print("OK: word-skill ready")
    sys.exit(0)


if __name__ == "__main__":
    main()
