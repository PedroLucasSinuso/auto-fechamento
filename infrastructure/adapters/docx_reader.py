import re
import aspose.words as aw
from core.interfaces.file_reader import FileReader
from infrastructure.config.settings import DATE_PATTERN
from typing import List

class DocxReader(FileReader):
    def read(self, file_path: str) -> List[str]:
        #Lê um arquivo DOCX e retorna linhas filtradas
        doc = aw.Document(file_path)
        lines = [p.get_text().strip() 
                for p in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True) 
                if p.get_text().strip()]
                
        return [line for line in lines 
               if "Filial" not in line 
               and "Aspose.Words" not in line 
               and not re.search(DATE_PATTERN, line)]