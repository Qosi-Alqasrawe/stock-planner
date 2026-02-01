# src/config.py
from dataclasses import dataclass, field

# =========================
# Alerts thresholds (كما هو)
# =========================
@dataclass
class Thresholds:
    red_lt: float = 7
    orange_lt: float = 21
    yellow_lt: float = 60  # green >= yellow_lt

@dataclass
class Settings:
    working_days: int = 26
    thresholds: Thresholds = field(default_factory=Thresholds)

# =========================
# Production Plan defaults (جديد)
# =========================
WORKING_DAYS = 26

DEMAND_CLASSES = ["VERY_HIGH", "HIGH", "MEDIUM", "LOW", "VERY_LOW"]

PLAN_DEFAULTS = {
    # كل ما رفعتهم بتكبر الكميات
    "target_months_map": {
        "VERY_HIGH": 3,
        "HIGH": 4,
        "MEDIUM": 6,
        "LOW": 9,
        "VERY_LOW": 12,
    },
    # سيفتي بالأيام (زيادة فوق أفق التخطيط)
    "safety_days_map": {
        "VERY_HIGH": 14,
        "HIGH": 10,
        "MEDIUM": 7,
        "LOW": 5,
        "VERY_LOW": 3,
    },
    # حد أدنى دفعات للبطيء (حتى لو تغطيته كويسة)
    "min_batch_months_map": {
        "LOW": 4,
        "VERY_LOW": 6,
    },
    # تقريب الكميات لدفعات كبيرة (0 = off)
    "batch_round_to": 1000,  # غيرها 500/2000 حسب وضعكم
}
