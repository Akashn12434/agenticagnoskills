#!/usr/bin/env python3
"""Verify ppt-skill structure."""

import sys
from pathlib import Path


def main():
    skill_dir = Path(__file__).parent.parent

    skill_md = skill_dir / "SKILL.md"
    guide_md = skill_dir / "references" / "guide.md"
    create_ppt_py = skill_dir / "scripts" / "create_ppt.py"

    if not skill_md.exists():
        print("ERROR: SKILL.md not found")
        sys.exit(1)

    if not guide_md.exists():
        print("ERROR: references/guide.md not found")
        sys.exit(1)

    if not create_ppt_py.exists():
        print("ERROR: scripts/create_ppt.py not found")
        sys.exit(1)

    skill_content = skill_md.read_text(encoding="utf-8")
    guide_content = guide_md.read_text(encoding="utf-8")

    required_skill_sections = [
        "PowerPoint Presentation Skill",
        "When to Use",
        "Presentation Creation Rules",
        "Design Requirements",
        "Validation Checklist",
        "Script Usage",
        "Best Practices",
    ]

    missing_skill = [
        section for section in required_skill_sections
        if section not in skill_content
    ]

    if missing_skill:
        print(f"ERROR: Missing SKILL.md sections: {missing_skill}")
        sys.exit(1)

    required_guide_sections = [
        "Standard Presentation Layout",
        "Content Density",
        "Font Guidelines",
        "Color Guidelines",
        "Layout Rules and Recommendations",
        "Presentation Flow",
        "Business Presentation Template",
        "Educational Presentation Template",
        "Technical Presentation Template",
    ]

    missing_guide = [
        section for section in required_guide_sections
        if section not in guide_content
    ]

    if missing_guide:
        print(f"ERROR: Missing guide.md sections: {missing_guide}")
        sys.exit(1)

    print("OK: ppt-skill ready")
    sys.exit(0)


if __name__ == "__main__":
    main()
