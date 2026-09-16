# Security Policy

## Versões Suportadas

Atualmente, estamos suportando a seguinte versão com atualizações de segurança:

| Versão | Suportada          |
| ------ | ------------------ |
| 1.0.x  | :white_check_mark: |

## Reportando uma Vulnerabilidade

A segurança dos usuários é nossa prioridade. Se você descobriu uma vulnerabilidade de segurança, por favor, nos ajude mantendo a comunidade segura.

### Como Reportar

**NÃO** abra uma issue pública para vulnerabilidades de segurança.

Em vez disso, envie um email para: **virgilio.junior94@gmail.com**

Inclua:
- Descrição da vulnerabilidade
- Passos para reproduzir
- Versão afetada
- Impacto potencial
- Sugestões de correção (se houver)

### O que Esperar

1. **Confirmação**: Você receberá uma confirmação em até 48 horas
2. **Avaliação**: Investigaremos o problema em até 7 dias
3. **Correção**: Trabalharemos em uma correção prioritária
4. **Divulgação**: Coordenaremos a divulgação responsável com você

### Política de Divulgação

- Manteremos você informado sobre o progresso
- Creditaremos você pela descoberta (se desejar)
- Divulgaremos publicamente após a correção estar disponível

## Boas Práticas de Segurança

Ao usar LocalPDF.io:

✅ **Recomendado:**
- Use em ambiente local/privado
- Mantenha as dependências atualizadas
- Execute em containers Docker isolados
- Faça backup de arquivos importantes

⚠️ **Evite:**
- Expor a aplicação publicamente na internet
- Processar arquivos de fontes não confiáveis sem verificação
- Executar com permissões elevadas desnecessárias

## Considerações de Privacidade

- No modo local, os arquivos são processados no computador do usuário
- A aplicação não envia dados para serviços externos por conta própria
- Arquivos temporários são automaticamente deletados após processamento

### Deploy público

O Railway e outros servidores públicos não fazem parte do modo local. Quando implantada em um servidor público, a aplicação processa os arquivos na infraestrutura desse servidor. Esse modo exige controle de acesso, retenção de dados, logs e política de privacidade próprios.

### MSI planejado

O futuro MSI deverá escutar somente em `127.0.0.1`, usar diretórios temporários no perfil local do usuário e não depender de Railway. O instalador ainda não está disponível.

## Testes de segurança

Os testes locais cobrem extensões não permitidas, limite de upload, ferramentas desconhecidas, nomes de arquivo com traversal, senhas fracas, entradas inválidas do editor e ausência de traceback nas respostas.

Execute com:

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

---

**Obrigado por ajudar a manter o LocalPDF.io seguro! 🔒**
