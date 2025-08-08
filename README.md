# Exchange365-DistList-Manager

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Exchange Online](https://img.shields.io/badge/Exchange-Online-green.svg)](https://docs.microsoft.com/en-us/exchange/exchange-online)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Script PowerShell interativo para gerenciamento completo de listas de distribuição e contatos externos no Microsoft 365 / Exchange Online.

## 📋 Índice

- [Funcionalidades](#-funcionalidades)
- [Pré-requisitos](#-pré-requisitos)
- [Instalação](#-instalação)
- [Uso](#-uso)
- [Solução de Problemas](#-solução-de-problemas)
- [Estrutura do Menu](#-estrutura-do-menu)
- [Exemplos de Uso](#-exemplos-de-uso)
- [Contribuindo](#-contribuindo)
- [Autor](#-autor)
- [Licença](#-licença)

## ✨ Funcionalidades

### Gerenciamento de Listas de Distribuição
- ✅ Visualizar todas as listas de distribuição existentes
- ✅ Listar membros de uma lista específica com contagem total
- ✅ Adicionar membros (internos e externos) às listas
- ✅ Remover membros com confirmação de segurança
- ✅ Exportar membros para arquivo CSV
- ✅ Criar novas listas de distribuição

### Gerenciamento de Contatos Externos
- ✅ Criar contatos externos para emails fora da organização
- ✅ Listar todos os contatos externos cadastrados
- ✅ Remover contatos externos
- ✅ Adicionar contatos às listas durante a criação

### Recursos do Sistema
- ✅ Múltiplas opções de autenticação (MFA, credenciais, UPN)
- ✅ Diagnóstico completo do sistema
- ✅ Verificação automática de conexão
- ✅ Instalação automática do módulo Exchange Online
- ✅ Configuração automática de TLS 1.2
- ✅ Status de conexão em tempo real

## 🔧 Pré-requisitos

- **PowerShell 5.1** ou superior
- **Windows 10/11** ou **Windows Server 2016+**
- **Conta com permissões** de administrador do Exchange Online
- **Conexão com a internet**
- **Módulo ExchangeOnlineManagement** (instalado automaticamente pelo script)

## 💻 Instalação

### Opção 1: Git Clone
```powershell
git clone https://github.com/elvisfalmeida/Exchange365-DistList-Manager.git
cd Exchange365-DistList-Manager
```

### Opção 2: Download Direto
```powershell
# Baixar o script diretamente
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/elvisfalmeida/Exchange365-DistList-Manager/main/Exchange365-DistList-Manager.ps1" -OutFile "Exchange365-DistList-Manager.ps1"
```

### Configurar Execution Policy (apenas uma vez)
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## 🚀 Uso

### Executar o Script

1. Abra o PowerShell como **Administrador**
2. Navegue até a pasta do script
3. Execute:

```powershell
.\Exchange365-DistList-Manager.ps1
```

### Primeira Execução

Na primeira execução, o script irá:
1. Verificar e instalar o módulo Exchange Online (se necessário)
2. Configurar o TLS 1.2 automaticamente
3. Solicitar conexão com o Exchange Online
4. Apresentar o menu principal

## 🔍 Solução de Problemas

### Erro: "Ocorreu um erro ao copiar o conteúdo para um fluxo"

**Solução:**
```powershell
# 1. Configure o TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 2. Atualize o módulo
Update-Module -Name ExchangeOnlineManagement -Force

# 3. Limpe o cache
Remove-Module ExchangeOnlineManagement -Force
Import-Module ExchangeOnlineManagement
```

### Erro de Autenticação

**Para contas com MFA:**
- Use a opção 1 (Autenticação Interativa) no menu de conexão

**Para contas sem MFA:**
- Use a opção 3 (Autenticação com UPN específico)

### Módulo não encontrado

O script instalará automaticamente, mas se falhar:
```powershell
Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
```

## 📊 Estrutura do Menu

```
=== LISTAS DE DISTRIBUIÇÃO ===
1. Ver todas as listas de distribuição
2. Listar membros de uma lista
3. Adicionar membro a uma lista
4. Remover membro de uma lista
5. Exportar membros para CSV
6. Criar nova lista de distribuição

=== CONTATOS EXTERNOS ===
7. Criar contato externo
8. Listar contatos externos
9. Remover contato externo

=== SISTEMA ===
C. Conectar/Reconectar ao Exchange Online
D. Diagnóstico do Sistema
0. Sair
```

## 💡 Exemplos de Uso

### Exportar Membros de uma Lista

1. Escolha a opção `5` no menu
2. Digite o nome da lista: `marketing@empresa.com`
3. Pressione Enter para usar o caminho padrão (pasta do script)
4. O arquivo será salvo como: `membros_20250108_143025.csv`

### Adicionar Contato Externo a uma Lista

1. Escolha a opção `7` para criar contato externo
2. Digite o nome: `João Silva`
3. Digite o email: `joao.silva@empresaexterna.com`
4. Confirme para adicionar a uma lista
5. Digite o nome da lista de destino

### Criar Nova Lista de Distribuição

1. Escolha a opção `6`
2. Digite o nome: `Suporte Técnico`
3. Digite o alias: `suporte`
4. O email será: `suporte@seudominio.com`

## 🤝 Contribuindo

Contribuições são bem-vindas! Sinta-se à vontade para:

1. Fazer um Fork do projeto
2. Criar uma branch para sua feature (`git checkout -b feature/NovaFuncionalidade`)
3. Commit suas mudanças (`git commit -m 'Adiciona nova funcionalidade'`)
4. Push para a branch (`git push origin feature/NovaFuncionalidade`)
5. Abrir um Pull Request

### Sugestões de Melhorias

- [ ] Adicionar suporte para grupos do Microsoft 365
- [ ] Implementar importação em massa via CSV
- [ ] Adicionar logs de auditoria
- [ ] Criar modo de operação não-interativo
- [ ] Adicionar suporte para múltiplos idiomas

## 👨‍💻 Autor

**Elvis Almeida**
- Website: [ebyte.net.br](https://ebyte.net.br)
- Email: elvis@ebyte.net.br
- GitHub: [@elvisfalmeida](https://github.com/elvisfalmeida)

## 📄 Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo [LICENSE](LICENSE) para detalhes.

## 🙏 Agradecimentos

- Microsoft pela documentação do Exchange Online PowerShell
- Comunidade PowerShell pelos exemplos e boas práticas
- Todos os contribuidores e usuários do script

## 📞 Suporte

Para reportar bugs ou solicitar novas funcionalidades, por favor:
1. Abra uma [Issue](https://github.com/elvisfalmeida/Exchange365-DistList-Manager/issues)
2. Ou envie um email para: elvis@ebyte.net.br

---

⭐ Se este projeto foi útil para você, considere dar uma estrela no GitHub!

**Desenvolvido com ❤️ por [Elvis Almeida](https://ebyte.net.br)**
