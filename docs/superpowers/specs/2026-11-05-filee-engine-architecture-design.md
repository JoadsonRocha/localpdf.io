# Design: Arquitetura de Motores Estilo Filee no LocalPDF.io (Python)

**Data:** 2026-11-05  
**Autor:** Antigravity & JoadsonRocha  
**Status:** Aprovado para Planejamento  

---

## 1. Contexto e Motivação

O **LocalPDF.io** possui mais de 15 ferramentas de manipulação e conversão de PDFs rodando 100% localmente. No entanto, atualmente toda a lógica de negócio, inicialização de dependências nativas (Ghostscript, Tesseract) e rotas HTTP estão centralizadas de forma monolítica em `app.py` (>5.000 linhas).

Inspirando-se na arquitetura do **Filee** ([KnifeLemon/Filee](https://github.com/KnifeLemon/Filee)), este projeto visa desacoplar o motor de conversão em uma arquitetura de pipeline extensível baseada em:
1. **Contrato de Conversor (`BaseConverter` / `IConverter`)**: Cada ferramenta encapsula seus próprios formatos de entrada/saída (*Edges*), identificador único (`id`), verificação de dependências nativas (`is_available`) e método de execução com reporte de progresso.
2. **Roteador Inteligente (`ConversionRouter`)**: Gerencia o despacho por `tool_id` direto (para ferramentas específicas como compressão, marca d'água, etc.) e o grafo de conversões por formatos (`from_ext -> to_ext`). Se uma rota direta existe (ex: `docx -> pdf`), ele a executa; se for composta (ex: `docx -> pdf -> png`), ele encadeia os conversores de forma transparente.
3. **Gerenciador de Motores (`EngineManager`)**: Monitora o status de dependências nativas (Tesseract OCR, Ghostscript) e gerencia caminhos locais/vendor.
4. **Preservação do Stack Atual**: Mantém o ecossistema Python 3.11+, Flask, PyMuPDF, Waitress e compatibilidade total com o desktop `windows_launcher.py` e Docker.

---

## 2. Objetivos, Não-Objetivos e Anti-Objetivos

### Metas (Goals)
- Modularizar todas as operações de conversão e edição de PDF em classes independentes herdando de `BaseConverter`.
- Isolar a verificação e configuração de Ghostscript e Tesseract em um `EngineManager` centralizado.
- Implementar um `ConversionRouter` híbrido: busca prioritária por `tool_id` e resolução por grafo de formatos (*Route Planning*) para conversões automáticas ou encadeadas.
- Extrair o editor visual de páginas para um serviço dedicado (`services/editor.py`).
- Manter 100% de retrocompatibilidade com as rotas web (`/convert`, `/editor`, etc.) e com o launcher desktop (`windows_launcher.py`).
- Criar suíte de testes unitários dedicada para cada conversor e para o roteador.

### Não-Objetivos (Non-Goals)
- Reescrever o frontend para frameworks SPA (React/Vue/Svelte) neste ciclo. Os templates HTML existentes serão mantidos integralmente.
- Migrar o backend para C#/.NET (o usuário definiu explicitamente manter o stack Python).
- Alterar `windows_launcher.py` neste estágio, garantindo que o build do PyInstaller (`localpdf.spec`) permaneça 100% estável.

### Anti-Objetivos (Anti-Goals)
- Quebrar os contratos das APIs atuais (`/convert`, `/editor/export`, etc.) usadas pela interface web.
- Alterar as convenções de nomes de arquivos gerados (ex: `_comprimido.pdf`, `_mesclado.pdf`), mantendo conformidade estrita com os testes existentes.
- Introduzir dependências de nuvem ou serviços externos que comprometam a privacidade 100% offline.

---

## 3. Arquitetura e Estrutura de Diretórios

```text
localpdf/
├── core/
│   ├── __init__.py
│   ├── converter.py       # BaseConverter, ConversionStep, ConversionResult
│   ├── router.py          # ConversionRouter (Despacho por Tool ID e Grafo de Rotas)
│   └── engine_manager.py  # Detecção e status de Tesseract, Ghostscript, etc.
├── converters/
│   ├── __init__.py
│   ├── documents.py       # WordToPdfConverter, ExcelToPdfConverter, TxtToPdfConverter, PdfToWordConverter, PdfToExcelConverter
│   ├── images.py          # ImagesToPdfConverter, PdfToImagesConverter
│   ├── pdf_ops.py         # MergePdfConverter, SplitPdfConverter, CompressPdfConverter, ProtectPdfConverter, WatermarkPdfConverter, PageNumbersConverter
│   ├── ocr.py             # OcrPdfConverter (com fallback nativo inteligente), PdfToTextConverter
│   └── pdfa.py            # PdfToPdfAConverter (Ghostscript)
├── services/
│   ├── __init__.py
│   └── editor.py          # Lógica isolada do Editor Visual (thumbnails, preview, export)
├── web/
│   ├── __init__.py
│   └── routes.py          # Adaptador Flask conectando as requisições HTTP ao ConversionRouter
├── tests/
│   ├── test_converters.py # Testes unitários de cada conversor
│   ├── test_router.py     # Testes do grafo de roteamento e encadeamento
│   └── ... (testes existentes preservados)
├── app.py                 # Ponto de entrada enxuto e compatível
└── windows_launcher.py    # Desktop launcher (preservado e intocado)
```

---

## 4. Especificação Técnica dos Componentes

### 4.1. `core/converter.py`
Define a interface base que todo conversor deve implementar:

```python
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import List, Tuple, Optional, Callable, Dict, Any

@dataclass
class ConversionStep:
    input_paths: List[str]
    output_dir: str
    target_format: Optional[str] = None
    options: Dict[str, Any] = field(default_factory=dict)
    original_filenames: List[str] = field(default_factory=list)
    progress_callback: Optional[Callable[[float, str], None]] = None

@dataclass
class ConversionResult:
    success: bool
    output_files: List[str]
    suggested_filename: Optional[str] = None
    error: Optional[str] = None
    metadata: Dict[str, Any] = field(default_factory=dict)

class BaseConverter(ABC):
    id: str
    display_name: str
    edges: List[Tuple[str, str]]  # ex: [("docx", "pdf")]
    max_parallelism: int = 4

    @abstractmethod
    def is_available(self) -> bool:
        """Verifica se dependências e bibliotecas estão disponíveis."""
        pass

    @abstractmethod
    def convert(self, step: ConversionStep) -> ConversionResult:
        """Executa a conversão."""
        pass
```

### 4.2. `core/router.py` (Despacho Híbrido)
O `ConversionRouter` resolve o conflito de operações com o mesmo formato de entrada e saída (`pdf -> pdf`):
1. **Despacho por Tool ID (Prioritário):** Quando uma requisição vem com `tool_id` específico (ex: `compress-pdf`, `watermark-pdf`), o roteador despacha diretamente para o conversor correspondente.
2. **Despacho por Grafo de Formatos (Fallback / Automático):** Quando apenas o formato de destino é informado (ex: upload arrastado onde se pede `png`), o roteador encontra o caminho no grafo (ex: `docx -> pdf -> png` via BFS).
3. **API do Router:**
   ```python
   def convert(self, input_paths: List[str], target_format: Optional[str] = None, 
               tool_id: Optional[str] = None, options: Dict[str, Any] = None,
               output_dir: Optional[str] = None) -> ConversionResult: ...
   ```

### 4.3. `core/engine_manager.py` & Fallback de OCR
- **Ghostscript e Tesseract:** Centraliza a detecção de binários instalados no sistema ou na pasta `vendor/`.
- **OCR Fallback:** Preserva o comportamento atual do `ocr_pdf`: se o Tesseract não estiver disponível, o `OcrPdfConverter` realiza fallback transparente para a extração do texto digital nativo via PyMuPDF, garantindo zero regressões nos testes.
- **Exposição de Status:** Disponibiliza o método `get_status()` e rota `GET /api/engines/status` para consumo opcional.

### 4.4. Convenção de Nomenclatura e Resposta
- Cada conversor define seu `suggested_filename` (ex: `{base}_comprimido.pdf`, `{base}_mesclado.pdf`), mantendo compatibilidade com os testes de `test_web_features.py`.
- O utilitário `build_response` em `web/routes.py` usa `suggested_filename` como nome de download padrão.

### 4.5. `services/editor.py`
Isola os métodos de renderização de thumbnails (`render_pdf_page`), descritores de página (`editor_page_descriptor`), inserção de imagens e exportação customizada de páginas (`export_editor_pages`).

---

## 5. Decisões do Construtor (Builder Calls)
1. **Tool ID como chave de despacho:** Resolve a ambiguidade entre múltiplos conversores `pdf -> pdf` (compress, watermark, protect, etc.).
2. **Editor como Serviço Independente:** O editor visual não herda de `BaseConverter` para evitar forçar estruturas de payload JSON e thumbnails em contratos de arquivo.
3. **Preservação de Fallback do OCR:** O conversor de OCR continua funcional mesmo sem Tesseract instalado, usando extração nativa PyMuPDF.
4. **Isolamento de `windows_launcher.py`:** Mantido intacto para evitar qualquer impacto no empacotamento PyInstaller.
