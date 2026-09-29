"""Known rules. A new rule: add a module under rules/ and list its class here.

Rules run in the order they appear in the profile, not in this list.
"""
from __future__ import annotations

from .model import Rule
from .rules.expand_simple_fields import ExpandSimpleFields
from .rules.reviewer_boxes import ReviewerBoxes
from .rules.unwrap_layout_tables import UnwrapLayoutTables

RULES: dict[str, type[Rule]] = {
    cls.id: cls for cls in (ReviewerBoxes, UnwrapLayoutTables, ExpandSimpleFields)
}
