"""Known rules. A new rule: add a module under rules/ and list its class here."""
from __future__ import annotations

from .model import Rule
from .rules.reviewer_boxes import ReviewerBoxes

RULES: dict[str, type[Rule]] = {cls.id: cls for cls in (ReviewerBoxes,)}
