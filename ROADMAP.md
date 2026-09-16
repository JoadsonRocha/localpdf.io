# LocalPDF.io Roadmap

Este documento descreve a evolução planejada do LocalPDF.io para uma experiência de ferramentas PDF inspirada em produtos como iLovePDF e WPS Office, mantendo identidade própria, código aberto e processamento local.

## Objetivo

Criar uma interface web clara para descobrir ferramentas PDF e executar operações em um fluxo curto: selecionar arquivos, configurar a operação, processar localmente e baixar o resultado.

## Editor visual de páginas PDF

### Escopo do MVP

O editor permitirá abrir um PDF e modificar a estrutura das páginas diretamente na interface:

- Visualizar miniaturas de todas as páginas.
- Reordenar páginas por arrastar e soltar.
- Selecionar e inserir páginas de outro PDF.
- Inserir imagens JPG ou PNG como novas páginas.
- Inserir páginas em branco.
- Duplicar páginas.
- Girar páginas em 90 graus.
- Excluir páginas.
- Extrair páginas selecionadas para um novo PDF.
- Desfazer e refazer operações.
- Exportar o documento final em PDF.

O MVP será um editor estrutural de páginas. Alterar livremente textos, fontes, imagens ou elementos internos de uma página exige um motor de edição diferente e fica fora desta primeira entrega.

### Fluxo do usuário

1. O usuário escolhe `Editar PDF` na página inicial.
2. Faz upload de um PDF.
3. A aplicação gera miniaturas das páginas.
4. O usuário reorganiza ou modifica as páginas na área de trabalho.
5. A interface mostra as alterações pendentes.
6. O usuário confirma `Salvar PDF`.
7. O backend gera um novo arquivo sem alterar o original.
8. O usuário baixa o PDF e pode iniciar outra edição.

### Interface planejada

- Barra superior com nome do arquivo, desfazer, refazer, zoom e salvar.
- Barra lateral ou painel de inserção para PDF, imagem e página em branco.
- Área central com miniaturas em grade ou lista.
- Seleção múltipla de páginas.
- Ações por página: girar, duplicar, extrair e excluir.
- Indicadores visuais para páginas selecionadas e alterações não salvas.
- Confirmação antes de excluir páginas ou sair com alterações pendentes.
- Layout responsivo: grade compacta no celular e painel completo no desktop.

### Contrato técnico proposto

Inicialmente, o fluxo pode continuar síncrono e usar o endpoint atual. A API deverá evoluir para aceitar uma lista explícita de operações:

```json
{
  "operations": [
    {"type": "move", "from": 4, "to": 1},
    {"type": "rotate", "page": 2, "degrees": 90},
    {"type": "delete", "page": 6}
  ]
}
```

Endpoints previstos:

```text
POST /api/editor/preview
POST /api/editor/export
GET  /api/editor/download/<job_id>
```

Para arquivos pequenos, `preview` poderá retornar miniaturas geradas pelo PyMuPDF. Para documentos grandes, o processamento deverá ser convertido para jobs assíncronos com status consultável.

### Backend

- Criar um serviço isolado para operações de páginas.
- Usar PyMuPDF para inserir, excluir, copiar, mover e rotacionar páginas.
- Preservar o arquivo original e gerar um novo PDF.
- Validar que os índices de página existem antes da operação.
- Rejeitar documentos sem páginas após exclusões.
- Limitar quantidade de páginas, tamanho do arquivo e quantidade de operações.
- Limpar todos os arquivos temporários após o download ou falha.
- Retornar erros com códigos estáveis, sem expor exceções internas.

### Critérios de aceite do editor

- Um PDF de várias páginas pode ser reorganizado e baixado com a nova ordem.
- Inserções de PDF, imagem e página em branco aparecem na posição escolhida.
- Rotação é aplicada apenas às páginas selecionadas.
- Exclusão e duplicação preservam a ordem das páginas restantes.
- Desfazer e refazer funcionam antes da exportação.
- O arquivo original nunca é sobrescrito.
- O resultado abre em leitores PDF comuns.
- O fluxo funciona com teclado e em telas pequenas.
- Os arquivos permanecem no ambiente local configurado pelo usuário.

## Evolução da interface geral

### Fase 1: base da aplicação

- Extrair o HTML, CSS e JavaScript de `app.py` para `templates/` e `static/`.
- Criar um catálogo único de ferramentas compartilhado pelo frontend e backend.
- Criar uma página própria para cada ferramenta.
- Padronizar upload, validação, processamento, erro e download.

### Fase 2: experiência inspirada em ferramentas PDF online

- Home organizada por categorias.
- Busca de ferramentas.
- Drag and drop consistente.
- Lista de arquivos com remoção e ordenação.
- Opções específicas de cada operação.
- Estados de carregamento, sucesso e erro.
- Design responsivo e acessível.

### Fase 3: editor visual de páginas

- Miniaturas e seleção de páginas.
- Reordenação, inserção, duplicação, rotação e exclusão.
- Histórico de desfazer/refazer.
- Exportação do PDF reorganizado.
- Testes de integração para diferentes combinações de operações.

### Fase 4: edição avançada

Possíveis evoluções, depois do editor estrutural:

- Inserção e edição de texto.
- Inclusão de imagens e formas.
- Anotações, marcações e destaques.
- Assinatura digital ou visual.
- Redação permanente de conteúdo sensível.

Esses recursos devem ser avaliados separadamente, pois aumentam bastante a complexidade de renderização, fontes, compatibilidade e segurança.

## Qualidade e segurança

Antes de publicar o editor:

- Adicionar testes unitários para cada operação de página.
- Adicionar testes de rota para upload, preview e exportação.
- Testar PDFs protegidos, corrompidos, grandes e com muitas páginas.
- Validar extensões e conteúdo real dos arquivos.
- Tratar limite de upload e erros de dependências externas.
- Não retornar `str(exception)` diretamente ao usuário.
- Atualizar `SECURITY.md` com os limites do editor.
- Atualizar os READMEs quando uma ferramenta estiver realmente implementada.

## Deploy no Railway

O aplicativo aceita a variável de ambiente `PORT` exigida pelo Railway e pode ser construído a partir do `dockerfile` existente. O deploy público muda o modelo de privacidade: os arquivos deixam de ficar exclusivamente na máquina do usuário e passam pelo container hospedado. Para produção, adicionar autenticação, limites de uso, política de retenção e revisão de logs antes de expor o serviço.

## Estado atual

O MVP do editor visual está implementado em `/editor/preview` e `/editor/export`. Ele oferece miniaturas, reordenação, inserção de PDF e imagens, páginas em branco, duplicação, rotação, exclusão, desfazer/refazer e exportação. A edição livre de texto e elementos internos da página continua planejada para uma fase posterior.