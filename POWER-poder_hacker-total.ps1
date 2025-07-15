# Defina o link pai
$linkPai = "https://www.santander.pt/"

# Função para extrair links de uma página
function Extrair-Links {
    param ($url)
    $response = Invoke-WebRequest -Uri $url
    $links = $response.Links | Where-Object {$_.href -like 'http*'} | Select-Object -ExpandProperty href
    return $links
}

# Função para extrair emails de uma página
function Extrair-Emails {
    param ($url)
    $response = Invoke-WebRequest -Uri $url
    $texto = $response.Content
    $emails = [regex]::Matches($texto, '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
    return $emails.Value
}

# Função para extrair números de telefone de uma página
function Extrair-Telefones {
    param ($url)
    $response = Invoke-WebRequest -Uri $url
    $texto = $response.Content
    # Este padrão regex pode precisar ser ajustado dependendo do formato dos números de telefone no site
    $telefones = [regex]::Matches($texto, '\(?\d{2,3}\)?[-\s.]?\d{3,5}[-\s.]?\d{4}')
    return $telefones.Value
}

# Função para extrair links de documentos de uma página
function Extrair-Documentos {
    param ($url)
    $response = Invoke-WebRequest -Uri $url
    $links = $response.Links | Where-Object {$_.href -like '*.pdf' -or $_.href -like '*.docx' -or $_.href -like '*.doc'}
    return $links | Select-Object -ExpandProperty href
}

# Extrair links do link pai
$linksFilhos = Extrair-Links -url $linkPai

# Mostrar resultados do link pai
Write-Host "Emails no link pai:"
Extrair-Emails -url $linkPai
Write-Host "Telefones no link pai:"
Extrair-Telefones -url $linkPai
Write-Host "Documentos no link pai:"
Extrair-Documentos -url $linkPai

# Processar links filhos
foreach ($link in $linksFilhos) {
    Write-Host "Processando $link"
    Write-Host "Emails:"
    Extrair-Emails -url $link
    Write-Host "Telefones:"
    Extrair-Telefones -url $link
    Write-Host "Documentos:"
    Extrair-Documentos -url $link
}