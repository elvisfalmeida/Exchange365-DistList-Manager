# Script de Gerenciamento de Listas de Distribuição - Office 365
# Autor: Elvis Almeida
# Versão: 2.0
# Descrição: Script interativo para gerenciar listas de distribuição e contatos externos

# Função para verificar e instalar o módulo Exchange Online
function Verify-ExchangeModule {
    Write-Host "`n=== Verificando Módulo Exchange Online ===" -ForegroundColor Cyan
    
    # Verifica a versão do PowerShell
    $psVersion = $PSVersionTable.PSVersion
    Write-Host "Versão do PowerShell: $psVersion" -ForegroundColor Gray
    
    if ($psVersion.Major -lt 5) {
        Write-Host "AVISO: PowerShell 5.1 ou superior é recomendado!" -ForegroundColor Yellow
    }
    
    # Verifica se o módulo está instalado
    if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "Módulo Exchange Online não encontrado. Instalando..." -ForegroundColor Yellow
        
        # Verifica se está rodando como administrador
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        
        if (!$isAdmin) {
            Write-Host "AVISO: Executando sem privilégios de administrador. A instalação pode falhar." -ForegroundColor Yellow
        }
        
        try {
            # Instala o NuGet provider se necessário
            if (!(Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue)) {
                Write-Host "Instalando NuGet provider..." -ForegroundColor Yellow
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser
            }
            
            # Define o repositório PSGallery como confiável
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
            
            # Instala o módulo
            Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
            Write-Host "Módulo instalado com sucesso!" -ForegroundColor Green
        }
        catch {
            Write-Host "Erro ao instalar o módulo: $_" -ForegroundColor Red
            Write-Host "`nTente instalar manualmente com:" -ForegroundColor Yellow
            Write-Host "Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser" -ForegroundColor Cyan
            return $false
        }
    }
    else {
        # Verifica a versão do módulo
        $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
        Write-Host "Módulo encontrado - Versão: $($module.Version)" -ForegroundColor Green
        
        # Verifica se há atualizações disponíveis
        if ($module.Version -lt "3.0.0") {
            Write-Host "AVISO: Versão antiga do módulo detectada. Considere atualizar:" -ForegroundColor Yellow
            Write-Host "Update-Module -Name ExchangeOnlineManagement" -ForegroundColor Cyan
        }
    }
    
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Write-Host "Módulo importado com sucesso!" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Erro ao importar o módulo: $_" -ForegroundColor Red
        return $false
    }
}

# Função melhorada para conectar ao Exchange Online
function Connect-ExchangeOnline365 {
    Write-Host "`n=== Conectando ao Exchange Online ===" -ForegroundColor Cyan
    
    # Verifica se já está conectado
    try {
        $testConnection = Get-OrganizationConfig -ErrorAction Stop
        Write-Host "Já conectado ao Exchange Online!" -ForegroundColor Green
        Write-Host "Organização: $($testConnection.DisplayName)" -ForegroundColor Gray
        return $true
    }
    catch {
        Write-Host "Não conectado. Iniciando conexão..." -ForegroundColor Yellow
    }
    
    Write-Host "`nEscolha o método de autenticação:" -ForegroundColor Cyan
    Write-Host "1. Autenticação Interativa (recomendado)"
    Write-Host "2. Autenticação com Credenciais"
    Write-Host "3. Autenticação com UPN específico"
    Write-Host "4. Conectar com parâmetros customizados"
    
    $authMethod = Read-Host "Escolha uma opção (1-4)"
    
    try {
        switch ($authMethod) {
            '1' {
                # Método 1: Autenticação interativa padrão
                Write-Host "Conectando com autenticação interativa..." -ForegroundColor Yellow
                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            }
            '2' {
                # Método 2: Com credenciais
                Write-Host "Digite suas credenciais do Office 365:" -ForegroundColor Yellow
                $UserCredential = Get-Credential -Message "Digite seu email e senha do Office 365"
                Connect-ExchangeOnline -Credential $UserCredential -ShowBanner:$false -ErrorAction Stop
            }
            '3' {
                # Método 3: Com UPN específico
                $upn = Read-Host "Digite seu email/UPN do Office 365"
                Write-Host "Conectando com UPN: $upn" -ForegroundColor Yellow
                Connect-ExchangeOnline -UserPrincipalName $upn -ShowBanner:$false -ErrorAction Stop
            }
            '4' {
                # Método 4: Customizado
                Write-Host "`nOpções de conexão customizada:" -ForegroundColor Cyan
                $upn = Read-Host "Digite seu email/UPN (ou deixe em branco)"
                $useModernAuth = Read-Host "Usar autenticação moderna? (S/N)"
                
                $params = @{
                    ShowBanner = $false
                    ErrorAction = 'Stop'
                }
                
                if ($upn) { $params.UserPrincipalName = $upn }
                
                if ($useModernAuth -eq 'N' -or $useModernAuth -eq 'n') {
                    Write-Host "AVISO: Autenticação básica pode não funcionar dependendo das políticas da organização" -ForegroundColor Yellow
                    $UserCredential = Get-Credential -Message "Digite seu email e senha do Office 365"
                    $params.Credential = $UserCredential
                }
                
                Connect-ExchangeOnline @params
            }
            default {
                Write-Host "Opção inválida. Usando autenticação interativa padrão..." -ForegroundColor Yellow
                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            }
        }
        
        # Testa a conexão
        $org = Get-OrganizationConfig -ErrorAction Stop
        Write-Host "`nConectado com sucesso!" -ForegroundColor Green
        Write-Host "Organização: $($org.DisplayName)" -ForegroundColor Gray
        return $true
    }
    catch {
        Write-Host "`nErro ao conectar: $_" -ForegroundColor Red
        
        # Fornece soluções baseadas no erro
        if ($_.Exception.Message -like "*copying content to a stream*" -or $_.Exception.Message -like "*copiar o conteúdo para um fluxo*") {
            Write-Host "`n=== SOLUÇÕES POSSÍVEIS ===" -ForegroundColor Yellow
            Write-Host "1. Execute o PowerShell como Administrador"
            Write-Host "2. Atualize o módulo: Update-Module ExchangeOnlineManagement -Force"
            Write-Host "3. Limpe o cache do módulo:"
            Write-Host "   Remove-Module ExchangeOnlineManagement -Force" -ForegroundColor Cyan
            Write-Host "   Import-Module ExchangeOnlineManagement" -ForegroundColor Cyan
            Write-Host "4. Verifique se o TLS 1.2 está habilitado:"
            Write-Host "   [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12" -ForegroundColor Cyan
            Write-Host "5. Desinstale e reinstale o módulo completamente"
        }
        elseif ($_.Exception.Message -like "*unauthorized*" -or $_.Exception.Message -like "*não autorizado*") {
            Write-Host "`n=== PROBLEMA DE AUTORIZAÇÃO ===" -ForegroundColor Yellow
            Write-Host "- Verifique se sua conta tem permissões de administrador do Exchange"
            Write-Host "- Verifique se o MFA está configurado corretamente"
            Write-Host "- Tente usar um método de autenticação diferente"
        }
        elseif ($_.Exception.Message -like "*network*" -or $_.Exception.Message -like "*rede*") {
            Write-Host "`n=== PROBLEMA DE REDE ===" -ForegroundColor Yellow
            Write-Host "- Verifique sua conexão com a internet"
            Write-Host "- Verifique se há proxy ou firewall bloqueando"
            Write-Host "- Tente: [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials"
        }
        
        return $false
    }
}

# Função para configurar TLS
function Set-TLSConfiguration {
    Write-Host "`nConfigurando protocolos de segurança..." -ForegroundColor Yellow
    try {
        # Habilita TLS 1.2
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Host "TLS 1.2 configurado com sucesso!" -ForegroundColor Green
    }
    catch {
        Write-Host "Erro ao configurar TLS: $_" -ForegroundColor Red
    }
}

# Função para testar conexão
function Test-ExchangeConnection {
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Função para listar todas as listas de distribuição
function Show-DistributionLists {
    Write-Host "`n=== Listas de Distribuição Existentes ===" -ForegroundColor Cyan
    
    if (!(Test-ExchangeConnection)) {
        Write-Host "Não conectado ao Exchange Online. Por favor, conecte-se primeiro (opção C)." -ForegroundColor Red
        Pause
        return
    }
    
    try {
        Write-Host "Carregando listas..." -ForegroundColor Yellow
        $lists = Get-DistributionGroup | Select-Object DisplayName, PrimarySmtpAddress, @{Name='MemberCount';Expression={(Get-DistributionGroupMember -Identity $_.Identity).Count}}, ManagedBy
        
        if ($lists.Count -eq 0) {
            Write-Host "Nenhuma lista de distribuição encontrada." -ForegroundColor Yellow
        }
        else {
            $lists | Format-Table -AutoSize
            Write-Host "`nTotal de listas: $($lists.Count)" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Erro ao listar: $_" -ForegroundColor Red
    }
    
    Pause
}

# Função para listar membros de uma lista específica
function Show-ListMembers {
    Write-Host "`n=== Listar Membros da Lista ===" -ForegroundColor Cyan
    
    if (!(Test-ExchangeConnection)) {
        Write-Host "Não conectado ao Exchange Online. Por favor, conecte-se primeiro (opção C)." -ForegroundColor Red
        Pause
        return
    }
    
    $listName = Read-Host "Digite o nome ou email da lista de distribuição"
    
    try {
        Write-Host "Carregando membros..." -ForegroundColor Yellow
        $members = Get-DistributionGroupMember -Identity $listName | Select-Object DisplayName, PrimarySmtpAddress, RecipientType
        
        if ($members.Count -eq 0) {
            Write-Host "A lista não possui membros." -ForegroundColor Yellow
        }
        else {
            Write-Host "`nMembros da lista '$listName':" -ForegroundColor Green
            $members | Format-Table -AutoSize
            Write-Host "`nTotal de membros: $($members.Count)" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Erro ao listar membros: $_" -ForegroundColor Red
    }
    
    Pause
}

# Função para adicionar membro à lista
function Add-ListMember {
    Write-Host "`n=== Adicionar Membro à Lista ===" -ForegroundColor Cyan
    
    if (!(Test-ExchangeConnection)) {
        Write-Host "Não conectado ao Exchange Online. Por favor, conecte-se primeiro (opção C)." -ForegroundColor Red
        Pause
        return
    }
    
    $listName = Read-Host "Digite o nome ou email da lista de distribuição"
    $memberEmail = Read-Host "Digite o email do membro a ser adicionado"
    
    try {
        Add-DistributionGroupMember -Identity $listName -Member $memberEmail
        Write-Host "Membro '$memberEmail' adicionado com sucesso à lista '$listName'!" -ForegroundColor Green
    }
    catch {
        if ($_.Exception.Message -like "*already a member*") {
            Write-Host "O usuário já é membro desta lista." -ForegroundColor Yellow
        }
        else {
            Write-Host "Erro ao adicionar membro: $_" -ForegroundColor Red
        }
    }
    
    Pause
}

# Função para remover membro da lista
function Remove-ListMember {
    Write-Host "`n=== Remover Membro da Lista ===" -ForegroundColor Cyan
    
    if (!(Test-ExchangeConnection)) {
        Write-Host "Não conectado ao Exchange Online. Por favor, conecte-se primeiro (opção C)." -ForegroundColor Red
        Pause
        return
    }
    
    $listName = Read-Host "Digite o nome ou email da lista de distribuição"
    $memberEmail = Read-Host "Digite o email do membro a ser removido"
    
    $confirm = Read-Host "Confirma a remoção de '$memberEmail' da lista '$listName'? (S/N)"
    
    if ($confirm -eq 'S' -or $confirm -eq 's') {
        try {
            Remove-DistributionGroupMember -Identity $listName -Member $memberEmail -Confirm:$false
            Write-Host "Membro '$memberEmail' removido com sucesso da lista '$listName'!" -ForegroundColor Green
        }
        catch {
            Write-Host "Erro ao remover membro: $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Operação cancelada." -ForegroundColor Yellow
    }
    
    Pause
}

# Função para exportar membros da lista para CSV
function Export-ListMembers {
    Write-Host "`n=== Exportar Membros da Lista ===" -ForegroundColor Cyan
    
    if (!(Test-ExchangeConnection)) {
        Write-Host "Não conectado ao Exchange Online. Por favor, conecte-se primeiro (opção C)." -ForegroundColor Red
        Pause
        return
    }
    
    $listName = Read-Host "Digite o nome ou email da lista de distribuição"
    
    # Define o caminho padrão como o diretório do script
    $scriptPath = Split-Path -Parent $MyInvocation.PSCommandPath
    if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        # Se não conseguir obter o caminho do script, usa o diretório atual
        $scriptPath = Get-Location
    }
    
    # Sugere um caminho padrão no diretório do script
    $defaultPath = Join-Path $scriptPath "membros_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    Write-Host "Caminho padrão: $defaultPath" -ForegroundColor Gray
    $exportPath = Read-Host "Digite o caminho para salvar o arquivo CSV (ou Enter para usar o padrão)"
    
    if ([string]::IsNullOrWhiteSpace($exportPath)) {
        $exportPath = $defaultPath
    }
    
    try {
        Write-Host "Exportando membros..." -ForegroundColor Yellow
        $members = Get-DistributionGroupMember -Identity $listName | Select-Object DisplayName, PrimarySmtpAddress, RecipientType, Department, Title
        
        if ($members.Count -eq 0) {
            Write-Host "A lista não possui membros para exportar." -ForegroundColor Yellow
        }
        else {
            $members | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
            Write-Host "Membros exportados com sucesso para: $exportPath" -ForegroundColor Green
            Write-Host "Total de membros exportados: $($members.Count)" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Erro ao exportar membros: $_" -ForegroundColor Red
    }
    
    Pause
}

# Função para criar contato externo
function New-ExternalContact {
    Write-Host "`n=== Criar Contato Externo ===" -ForegroundColor Cyan
    
    if (!(Test-ExchangeConnection)) {
        Write-Host "Não conectado ao Exchange Online. Por favor, conecte-se primeiro (opção C)." -ForegroundColor Red
        Pause
        return
    }
    
    $name = Read-Host "Digite o nome do contato"
    $email = Read-Host "Digite o email externo"
    $firstName = Read-Host "Digite o primeiro nome (opcional - pressione Enter para pular)"
    $lastName = Read-Host "Digite o sobrenome (opcional - pressione Enter para pular)"
    
    try {
        $params = @{
            Name = $name
            ExternalEmailAddress = $email
        }
        
        if ($firstName) { $params.FirstName = $firstName }
        if ($lastName) { $params.LastName = $lastName }
        
        New-MailContact @params
        Write-Host "Contato externo '$name' criado com sucesso!" -ForegroundColor Green
        
        # Pergunta se deseja adicionar o contato a uma lista
        $addToList = Read-Host "`nDeseja adicionar este contato a uma lista de distribuição? (S/N)"
        if ($addToList -eq 'S' -or $addToList -eq 's') {
            $listName = Read-Host "Digite o nome ou email da lista"
            try {
                Add-DistributionGroupMember -Identity $listName -Member $email
                Write-Host "Contato adicionado à lista '$listName' com sucesso!" -ForegroundColor Green
            }
            catch {
                Write-Host "Erro ao adicionar à lista: $_" -ForegroundColor Red
            }
        }
    }
    catch {
        if ($_.Exception.Message -like "*already exists*") {
            Write-Host "Um contato com este nome ou email já existe." -ForegroundColor Yellow
        }
        else {
            Write-Host "Erro ao criar contato: $_" -ForegroundColor Red
        }
    }
    
    Pause
}

# Função para listar contatos externos
function Show-ExternalContacts {
    Write-Host "`n=== Contatos Externos ===" -ForegroundColor Cyan
    
    if (!(Test-ExchangeConnection)) {
        Write-Host "Não conectado ao Exchange Online. Por favor, conecte-se primeiro (opção C)." -ForegroundColor Red
        Pause
        return
    }
    
    try {
        Write-Host "Carregando contatos..." -ForegroundColor Yellow
        $contacts = Get-MailContact | Select-Object DisplayName, ExternalEmailAddress, FirstName, LastName
        
        if ($contacts.Count -eq 0) {
            Write-Host "Nenhum contato externo encontrado." -ForegroundColor Yellow
        }
        else {
            $contacts | Format-Table -AutoSize
            Write-Host "`nTotal de contatos externos: $($contacts.Count)" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Erro ao listar contatos: $_" -ForegroundColor Red
    }
    
    Pause
}

# Função para remover contato externo
function Remove-ExternalContact {
    Write-Host "`n=== Remover Contato Externo ===" -ForegroundColor Cyan
    
    if (!(Test-ExchangeConnection)) {
        Write-Host "Não conectado ao Exchange Online. Por favor, conecte-se primeiro (opção C)." -ForegroundColor Red
        Pause
        return
    }
    
    $contactName = Read-Host "Digite o nome ou email do contato externo"
    
    $confirm = Read-Host "Confirma a remoção do contato '$contactName'? (S/N)"
    
    if ($confirm -eq 'S' -or $confirm -eq 's') {
        try {
            Remove-MailContact -Identity $contactName -Confirm:$false
            Write-Host "Contato '$contactName' removido com sucesso!" -ForegroundColor Green
        }
        catch {
            Write-Host "Erro ao remover contato: $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Operação cancelada." -ForegroundColor Yellow
    }
    
    Pause
}

# Função para criar nova lista de distribuição
function New-DistributionList {
    Write-Host "`n=== Criar Nova Lista de Distribuição ===" -ForegroundColor Cyan
    
    if (!(Test-ExchangeConnection)) {
        Write-Host "Não conectado ao Exchange Online. Por favor, conecte-se primeiro (opção C)." -ForegroundColor Red
        Pause
        return
    }
    
    $displayName = Read-Host "Digite o nome da lista"
    $alias = Read-Host "Digite o alias/email da lista (sem @dominio.com)"
    $description = Read-Host "Digite uma descrição (opcional - pressione Enter para pular)"
    
    try {
        $params = @{
            Name = $displayName
            DisplayName = $displayName
            Alias = $alias
            Type = "Distribution"
        }
        
        if ($description) { $params.Notes = $description }
        
        New-DistributionGroup @params
        Write-Host "Lista de distribuição '$displayName' criada com sucesso!" -ForegroundColor Green
        
        # Obtém o domínio primário
        $domain = (Get-AcceptedDomain | Where-Object {$_.Default -eq $true}).DomainName
        Write-Host "Email da lista: $alias@$domain" -ForegroundColor Green
    }
    catch {
        if ($_.Exception.Message -like "*already exists*") {
            Write-Host "Uma lista com este nome ou alias já existe." -ForegroundColor Yellow
        }
        else {
            Write-Host "Erro ao criar lista: $_" -ForegroundColor Red
        }
    }
    
    Pause
}

# Função de diagnóstico
function Show-Diagnostics {
    Write-Host "`n=== DIAGNÓSTICO DO SISTEMA ===" -ForegroundColor Cyan
    
    # Versão do PowerShell
    Write-Host "`nPowerShell:" -ForegroundColor Yellow
    Write-Host "  Versão: $($PSVersionTable.PSVersion)" -ForegroundColor Gray
    Write-Host "  Edição: $($PSVersionTable.PSEdition)" -ForegroundColor Gray
    Write-Host "  SO: $($PSVersionTable.OS)" -ForegroundColor Gray
    
    # Módulo Exchange Online
    Write-Host "`nMódulo Exchange Online:" -ForegroundColor Yellow
    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Select-Object -First 1
    if ($module) {
        Write-Host "  Instalado: Sim" -ForegroundColor Green
        Write-Host "  Versão: $($module.Version)" -ForegroundColor Gray
        Write-Host "  Caminho: $($module.Path)" -ForegroundColor Gray
    }
    else {
        Write-Host "  Instalado: Não" -ForegroundColor Red
    }
    
    # Protocolo TLS
    Write-Host "`nProtocolo de Segurança:" -ForegroundColor Yellow
    Write-Host "  TLS: $([Net.ServicePointManager]::SecurityProtocol)" -ForegroundColor Gray
    
    # Status da conexão
    Write-Host "`nStatus da Conexão:" -ForegroundColor Yellow
    if (Test-ExchangeConnection) {
        Write-Host "  Conectado: Sim" -ForegroundColor Green
        try {
            $org = Get-OrganizationConfig
            Write-Host "  Organização: $($org.DisplayName)" -ForegroundColor Gray
        }
        catch {}
    }
    else {
        Write-Host "  Conectado: Não" -ForegroundColor Red
    }
    
    # Execution Policy
    Write-Host "`nExecution Policy:" -ForegroundColor Yellow
    Write-Host "  CurrentUser: $(Get-ExecutionPolicy -Scope CurrentUser)" -ForegroundColor Gray
    Write-Host "  LocalMachine: $(Get-ExecutionPolicy -Scope LocalMachine)" -ForegroundColor Gray
    
    Pause
}

# Função principal do menu
function Show-MainMenu {
    do {
        Clear-Host
        Write-Host "================================================" -ForegroundColor Cyan
        Write-Host "   GERENCIADOR DE LISTAS DE DISTRIBUIÇÃO 365   " -ForegroundColor Cyan
        Write-Host "================================================" -ForegroundColor Cyan
        
        # Mostra status da conexão
        if (Test-ExchangeConnection) {
            Write-Host "Status: CONECTADO" -ForegroundColor Green -BackgroundColor DarkGreen
        }
        else {
            Write-Host "Status: DESCONECTADO" -ForegroundColor Red -BackgroundColor DarkRed
        }
        
        Write-Host ""
        Write-Host "=== LISTAS DE DISTRIBUIÇÃO ===" -ForegroundColor Yellow
        Write-Host "1. Ver todas as listas de distribuição"
        Write-Host "2. Listar membros de uma lista"
        Write-Host "3. Adicionar membro a uma lista"
        Write-Host "4. Remover membro de uma lista"
        Write-Host "5. Exportar membros para CSV"
        Write-Host "6. Criar nova lista de distribuição"
        Write-Host ""
        Write-Host "=== CONTATOS EXTERNOS ===" -ForegroundColor Yellow
        Write-Host "7. Criar contato externo"
        Write-Host "8. Listar contatos externos"
        Write-Host "9. Remover contato externo"
        Write-Host ""
        Write-Host "=== SISTEMA ===" -ForegroundColor Yellow
        Write-Host "C. Conectar/Reconectar ao Exchange Online"
        Write-Host "D. Diagnóstico do Sistema"
        Write-Host "0. Sair"
        Write-Host ""
        
        $choice = Read-Host "Escolha uma opção"
        
        switch ($choice) {
            '1' { Show-DistributionLists }
            '2' { Show-ListMembers }
            '3' { Add-ListMember }
            '4' { Remove-ListMember }
            '5' { Export-ListMembers }
            '6' { New-DistributionList }
            '7' { New-ExternalContact }
            '8' { Show-ExternalContacts }
            '9' { Remove-ExternalContact }
            'C' { Connect-ExchangeOnline365 }
            'c' { Connect-ExchangeOnline365 }
            'D' { Show-Diagnostics }
            'd' { Show-Diagnostics }
            '0' { 
                Write-Host "`nDesconectando do Exchange Online..." -ForegroundColor Yellow
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                Write-Host "Script finalizado. Até logo!" -ForegroundColor Green
                return 
            }
            default { 
                Write-Host "Opção inválida. Tente novamente." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    } while ($true)
}

# Função para pausar
function Pause {
    Write-Host "`nPressione qualquer tecla para continuar..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# === INÍCIO DO SCRIPT ===
Clear-Host
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "   GERENCIADOR DE LISTAS DE DISTRIBUIÇÃO 365   " -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Configura TLS 1.2
Set-TLSConfiguration

# Verifica e instala o módulo se necessário
if (!(Verify-ExchangeModule)) {
    Write-Host "`nNão foi possível carregar o módulo Exchange Online." -ForegroundColor Red
    Write-Host "Deseja continuar mesmo assim? (S/N)" -ForegroundColor Yellow
    $continue = Read-Host
    
    if ($continue -ne 'S' -and $continue -ne 's') {
        Write-Host "Saindo..." -ForegroundColor Red
        exit
    }
}

# Pergunta se deseja conectar automaticamente
Write-Host "`nDeseja conectar ao Exchange Online agora? (S/N)" -ForegroundColor Cyan
$autoConnect = Read-Host

if ($autoConnect -eq 'S' -or $autoConnect -eq 's') {
    if (!(Connect-ExchangeOnline365)) {
        Write-Host "`nNão foi possível conectar automaticamente." -ForegroundColor Yellow
        Write-Host "Você pode tentar conectar manualmente usando a opção 'C' no menu." -ForegroundColor Yellow
        Write-Host "Pressione qualquer tecla para continuar para o menu..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}
else {
    Write-Host "`nVocê pode conectar a qualquer momento usando a opção 'C' no menu." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
}

# Inicia o menu principal
Show-MainMenu

# Desconecta ao sair
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
