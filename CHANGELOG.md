# Changelog

Todas as mudanças notáveis neste projeto serão documentadas neste arquivo.

O formato é baseado em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.0.0/),
e este projeto adere ao [Versionamento Semântico](https://semver.org/lang/pt-BR/).

## [Não lançado]

### Adicionado
- 🌐 Opção de interface em inglês com preferência salva localmente
- 🎨 Identidade visual azul, favicon e layout responsivo para mobile
- 🛡️ Testes automatizados de segurança para entradas inválidas e limites de upload
- 💧 Marca d'água de texto em PDF com posição configurável
- 🔢 Numeração de páginas com posições de cabeçalho e rodapé
- 🔐 Proteção local de PDF com senha e criptografia AES-256
- 📄 Conversão de PDF para texto nativo em arquivo TXT
- 🖥️ Editor visual de páginas PDF inspirado em ferramentas de escritório
- 🔀 Reordenação de páginas por arrastar e soltar
- ➕ Inserção de páginas de outro PDF, imagens ou páginas em branco
- 📋 Duplicação, rotação, extração e exclusão de páginas
- 💾 Exportação do PDF reorganizado em um único arquivo
- 📚 Plano técnico e critérios de aceite em [ROADMAP.md](ROADMAP.md)

### Segurança
- 🧹 Respostas de erro não expõem exceções internas ou caminhos do sistema
- ✅ Validação de extensões, limite de upload e dados inválidos do editor

## [1.0.0] - 2025-11-17

### Adicionado
- ✨ Conversão de PDF para imagens (PNG)
- ✨ Conversão de múltiplas imagens para PDF
- ✨ Conversão de Word (DOCX) para PDF (suporta múltiplos arquivos)
- ✨ Conversão de Excel (XLSX) para PDF
- ✨ Conversão de texto (TXT) para PDF
- ✨ Conversão de PDF para Word (DOCX)
- ✨ Conversão de PDF para Excel (XLSX)
- ✨ Conversão de PDF para texto (TXT)
- ✨ Mesclar múltiplos PDFs em um único arquivo
- ✨ Dividir PDF em páginas individuais
- ✨ Comprimir PDF mantendo qualidade
- 🐳 Suporte a Docker para deploy fácil
- 🎨 Interface web responsiva e intuitiva
- 🔒 Processamento 100% local (privacidade garantida)
- 📝 Documentação completa (README, CONTRIBUTING, CODE_OF_CONDUCT)
- 📋 Templates para Issues e Pull Requests

### Tecnologias
- Flask 2.3.3
- PyMuPDF 1.23.8
- Pillow 10.0.1
- python-docx 0.8.11
- ReportLab 4.0.5
- OpenPyXL 3.1.2

---

**Formato das versões:** [Major.Minor.Patch]
- **Major**: Mudanças incompatíveis na API
- **Minor**: Novas funcionalidades compatíveis
- **Patch**: Correções de bugs
