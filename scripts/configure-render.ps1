param(
    [Parameter(Mandatory = $true)]
    [string]$RenderApiKey,

    [Parameter(Mandatory = $true)]
    [string]$SmtpPassword
)

$headers = @{
    Authorization  = "Bearer $RenderApiKey"
    Accept         = "application/json"
    "Content-Type" = "application/json"
}

$envVars = [ordered]@{
    SMTP_HOST     = "smtp.gmail.com"
    SMTP_PORT     = "587"
    SMTP_USER     = "nicolasd@distribuidoracero.com.ar"
    MAIL_FROM     = "nicolasd@distribuidoracero.com.ar"
    MAIL_TO       = "nicolasd@distribuidoracero.com.ar"
    SMTP_PASSWORD = $SmtpPassword
}

Write-Host "Buscando servicio 'rendiciones' en Render..."
$services = Invoke-RestMethod -Uri "https://api.render.com/v1/services?limit=50" -Headers $headers -Method Get

$service = $null
foreach ($entry in $services) {
    $candidate = $entry.service
    if ($candidate.name -match "rendicion") {
        $service = $candidate
        break
    }
}

if (-not $service) {
    Write-Error "No encontré un servicio con 'rendicion' en el nombre. Servicios disponibles:"
    foreach ($entry in $services) {
        Write-Host " - $($entry.service.name) ($($entry.service.id))"
    }
    exit 1
}

Write-Host "Servicio encontrado: $($service.name) ($($service.id))"

foreach ($key in $envVars.Keys) {
    $value = $envVars[$key]
    Write-Host "Configurando $key..."
    $body = @{ value = $value } | ConvertTo-Json
    Invoke-RestMethod `
        -Uri "https://api.render.com/v1/services/$($service.id)/env-vars/$key" `
        -Headers $headers `
        -Method Put `
        -Body $body | Out-Null
}

Write-Host "Disparando redeploy..."
Invoke-RestMethod `
    -Uri "https://api.render.com/v1/services/$($service.id)/deploys" `
    -Headers $headers `
    -Method Post `
    -Body (@{ clearCache = "do_not_clear" } | ConvertTo-Json) | Out-Null

Write-Host "Listo. Verificá el deploy en el dashboard de Render."
