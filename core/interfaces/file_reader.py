from abc import ABC, abstractmethod
from typing import List

class FileReader(ABC):
    @abstractmethod
    def read(self, file_path: str) -> List[str]:
        pass