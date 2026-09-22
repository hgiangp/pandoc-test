"""Common contract of all cleanup rules."""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any, ClassVar, Optional

from .package import DocxPackage, ShapeRef

MODES = ("Off", "Report", "Apply")
CONFIDENCES = ("high", "medium", "low")


@dataclass
class Finding:
    rule_id: str
    part: str
    paragraph_index: int
    shape_id: str
    shape_name: str
    container: str
    confidence: str
    reason: str
    features: dict[str, Any] = field(default_factory=dict)
    location: str = ""
    action: str = "reported"
    ref: Optional[ShapeRef] = field(default=None, repr=False)

    def to_row(self) -> dict[str, Any]:
        return {
            "RuleId": self.rule_id,
            "Action": self.action,
            "Confidence": self.confidence,
            "Part": self.part,
            "Paragraph": self.paragraph_index,
            "ShapeId": self.shape_id,
            "ShapeName": self.shape_name,
            "Container": self.container,
            "Reason": self.reason,
            "Location": self.location,
            "Features": json.dumps(self.features, ensure_ascii=False, sort_keys=True),
        }


class Rule:
    """A cleanup rule. Subclasses set `id` and `defaults`, and implement detect/apply.

    detect() must not modify the package; apply() gets the findings of detect().
    """

    id: ClassVar[str] = ""
    description: ClassVar[str] = ""
    defaults: ClassVar[dict[str, Any]] = {}

    def __init__(self, params: Optional[dict[str, Any]] = None):
        params = dict(params or {})
        unknown = set(params) - set(self.defaults)
        if unknown:
            raise ValueError(f"{self.id}: unknown parameter(s) {sorted(unknown)}; "
                             f"allowed: {sorted(self.defaults)}")
        self.params = {**self.defaults, **params}

    def detect(self, pkg: DocxPackage) -> list[Finding]:
        raise NotImplementedError

    def apply(self, pkg: DocxPackage, findings: list[Finding]) -> None:
        raise NotImplementedError
