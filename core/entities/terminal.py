from dataclasses import dataclass

@dataclass
class Terminal:
    id: str
    gross_sales: float = 0.0
    net_sales: float = 0.0
    exchanged_items: float = 0.0
    items_count: int = 0
    tickets_count: int = 0