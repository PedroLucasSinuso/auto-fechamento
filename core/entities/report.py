from dataclasses import dataclass, field
from typing import Dict, List

@dataclass
class Report:
    terminals: List[str] = field(default_factory=list)
    gross_sales: Dict[str, float] = field(default_factory=dict)
    gross_add: float = 0.0
    discounts: float = 0.0
    exchanged_items: float = 0.0
    shipping: float = 0.0
    cash_inflow: Dict[str, float] = field(default_factory=dict)
    cash_outflow: Dict[str, float] = field(default_factory=dict)
    total_cash_inflow: float = 0.0
    total_cash_outflow: float = 0.0
    expenses: float = 0.0
    credsystem: float = 0.0
    omnichannel: float = 0.0
    payment_methods: Dict[str, float] = field(default_factory=dict)