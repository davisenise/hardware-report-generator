<#
.SYNOPSIS
    Coleta o inventario de hardware da maquina local e gera um
    relatorio HTML visual e autossuficiente (CSS embutido).

.DESCRIPTION
    Junta coleta via CIM (PowerShell) com geracao de front-end
    (HTML/CSS) num arquivo so. O .html abre em qualquer navegador,
    sem internet e sem dependencia. Pensado pra entregar inventario
    com cara de relatorio, nao de bloco de notas.

.EXAMPLE
    .\Gerar-RelatorioHTML.ps1
    Coleta a maquina local, salva o .html na mesma pasta do script e abre.

.NOTES
    Autor : Davi Senise - Suporte TI
    Requer: PowerShell 5.1+
#>

[CmdletBinding()]
param(
    [string] $CaminhoSaida
)

# ---------------------------------------------------------------
#  ONDE SALVAR
#  Por que NAO usar o Desktop por padrao?
#  Em maquina com OneDrive corporativo, o Desktop e redirecionado
#  pra nuvem e o GetFolderPath("Desktop") as vezes devolve um
#  caminho inconsistente (ex: Downloads\Area de Trabalho), que
#  nao existe -> o .html "some". Salvar na pasta do proprio
#  script ($PSScriptRoot) e previsivel e funciona em qualquer PC.
# ---------------------------------------------------------------
if (-not $CaminhoSaida) {
    if ($PSScriptRoot) { $CaminhoSaida = $PSScriptRoot }
    else { $CaminhoSaida = (Get-Location).Path }
}

# ---------------------------------------------------------------
#  HELPER: escapa texto pra nao quebrar o HTML
#  Modelo de disco/placa pode (raramente) ter < > &. Escapar e
#  basico de quem gera HTML a partir de dados externos.
# ---------------------------------------------------------------
function Esc([string]$t) {
    if (-not $t) { return "" }
    return $t.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Trim()
}

# ---------------------------------------------------------------
#  COLETA
# ---------------------------------------------------------------
Write-Host "Coletando dados da maquina..." -ForegroundColor Cyan

$os   = Get-CimInstance Win32_OperatingSystem
$cpu  = Get-CimInstance Win32_Processor | Select-Object -First 1
$cs   = Get-CimInstance Win32_ComputerSystem
$bios = Get-CimInstance Win32_BIOS
$pentes  = Get-CimInstance Win32_PhysicalMemory
$gpus    = Get-CimInstance Win32_VideoController
$nics    = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True"
$logicos = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"

# VRAM real (registro), fallback no WMI que estoura em 4GB
function Get-VRAM($nome) {
    $regBase = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}"
    foreach ($k in (Get-ChildItem $regBase -ErrorAction SilentlyContinue)) {
        $info = Get-ItemProperty $k.PSPath -ErrorAction SilentlyContinue
        if ($info.DriverDesc -eq $nome -and $info.'HardwareInformation.qwMemorySize') {
            return [math]::Round($info.'HardwareInformation.qwMemorySize'/1GB, 1)
        }
    }
    return $null
}

# Tipo de disco real (SSD/HDD) via Get-PhysicalDisk
$fisicos = @{}
try {
    foreach ($d in (Get-PhysicalDisk -ErrorAction Stop)) {
        $fisicos[$d.DeviceId] = @{ Tipo = $d.MediaType; Bus = $d.BusType }
    }
} catch {}

# ---------------------------------------------------------------
#  MONTAGEM DOS CARDS (HTML dinamico)
# ---------------------------------------------------------------

# Linha de dado padrao
function Row($label, $valor) {
    return "<div class='row'><span class='k'>$label</span><span class='v'>$(Esc $valor)</span></div>"
}

# --- Sistema ---
$cardSistema = @"
<article class="card" style="--accent:var(--blue)">
  <div class="tag">SISTEMA</div>
  $(Row 'SO' $os.Caption)
  $(Row 'Versao' "$($os.Version) (build $($os.BuildNumber))")
  $(Row 'Arquitetura' $os.OSArchitecture)
  $(Row 'Fabricante' $cs.Manufacturer)
  $(Row 'Modelo' $cs.Model)
  $(Row 'Serie' $bios.SerialNumber)
</article>
"@

# --- Processador ---
$cardCpu = @"
<article class="card" style="--accent:var(--orange)">
  <div class="tag">PROCESSADOR</div>
  $(Row 'CPU' $cpu.Name)
  $(Row 'Nucleos fisicos' $cpu.NumberOfCores)
  $(Row 'Nucleos logicos' $cpu.NumberOfLogicalProcessors)
  $(Row 'Clock max' "$($cpu.MaxClockSpeed) MHz")
</article>
"@

# --- Memoria ---
$ramTotal = [math]::Round($cs.TotalPhysicalMemory/1GB, 1)
$pentesHtml = ""
$i = 1
foreach ($p in $pentes) {
    $cap = [math]::Round($p.Capacity/1GB, 0)
    $fab = if ($p.Manufacturer) { $p.Manufacturer.Trim() } else { "n/d" }
    $pentesHtml += Row "Slot $i" "${cap}GB - $($p.Speed)MHz - $fab"
    $i++
}
$cardRam = @"
<article class="card" style="--accent:var(--red)">
  <div class="tag">MEMORIA</div>
  <div class="big">$ramTotal <span>GB</span></div>
  $pentesHtml
</article>
"@

# --- Video ---
$gpuHtml = ""
foreach ($g in $gpus) {
    $vram = Get-VRAM $g.Name
    if (-not $vram -and $g.AdapterRAM) { $vram = [math]::Round($g.AdapterRAM/1GB, 1) }
    $gpuHtml += Row 'GPU' $g.Name
    $gpuHtml += Row 'VRAM' $(if ($vram) { "$vram GB" } else { "n/d" })
    $gpuHtml += Row 'Driver' $g.DriverVersion
}
$cardGpu = @"
<article class="card" style="--accent:var(--blue)">
  <div class="tag">VIDEO</div>
  $gpuHtml
</article>
"@

# --- Armazenamento (com medidor segmentado) ---
# Assinatura visual: barra de uso estilo VU meter de fita.
function Medidor($percent) {
    $segs = 24
    $cheios = [math]::Ceiling($percent / 100 * $segs)
    $cor = if ($percent -ge 90) { 'var(--red)' } elseif ($percent -ge 70) { 'var(--orange)' } else { 'var(--green)' }
    $html = "<div class='meter'>"
    for ($s = 0; $s -lt $segs; $s++) {
        if ($s -lt $cheios) { $html += "<i style='background:$cor'></i>" }
        else { $html += "<i></i>" }
    }
    $html += "</div>"
    return $html
}

$discosHtml = ""
foreach ($d in $logicos) {
    $totalGB = [math]::Round($d.Size/1GB, 1)
    $livreGB = [math]::Round($d.FreeSpace/1GB, 1)
    $usadoGB = [math]::Round($totalGB - $livreGB, 1)
    $pct = if ($d.Size -gt 0) { [math]::Round(($d.Size - $d.FreeSpace)/$d.Size * 100, 0) } else { 0 }
    $discosHtml += @"
  <div class="disk">
    <div class="disk-head"><span class="disk-id">$($d.DeviceID)</span><span class="disk-pct">$pct% usado</span></div>
    $(Medidor $pct)
    <div class="disk-foot">$usadoGB GB de $totalGB GB &middot; $livreGB GB livres</div>
  </div>
"@
}
$cardDisco = @"
<article class="card wide" style="--accent:var(--orange)">
  <div class="tag">ARMAZENAMENTO</div>
  $discosHtml
</article>
"@

# --- Rede ---
$redeHtml = ""
foreach ($n in $nics) {
    $ips = ($n.IPAddress | Where-Object { $_ -match '\.' }) -join ', '
    $redeHtml += Row 'Adaptador' $n.Description
    $redeHtml += Row 'IP' $ips
    $redeHtml += Row 'MAC' $n.MACAddress
}
$cardRede = @"
<article class="card wide" style="--accent:var(--blue)">
  <div class="tag">REDE</div>
  $redeHtml
</article>
"@

# ---------------------------------------------------------------
#  TEMPLATE HTML (CSS embutido)
# ---------------------------------------------------------------
$maquina = Esc $env:COMPUTERNAME
$usuario = Esc $cs.UserName
$gerado  = Get-Date -Format 'dd/MM/yyyy HH:mm'

$html = @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Inventario - $maquina</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=JetBrains+Mono:wght@400;500&display=swap');

  :root{
    --bg:#0E0E10; --surface:#16161A; --line:#2A2A30;
    --text:#ECECEC; --muted:#7E7E88;
    --orange:#FF6B00; --red:#E8350C; --blue:#3D7BFF; --green:#36E08A;
    --display:'Space Grotesk','Segoe UI',sans-serif;
    --mono:'JetBrains Mono','Consolas',monospace;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{
    background:var(--bg); color:var(--text);
    font-family:var(--mono); line-height:1.5;
    padding:32px 20px; min-height:100vh;
    /* assinatura lo-fi: scanline bem sutil */
    background-image:repeating-linear-gradient(0deg,rgba(255,255,255,.012) 0px,rgba(255,255,255,.012) 1px,transparent 1px,transparent 3px);
  }
  .wrap{max-width:1100px;margin:0 auto}

  header{border-bottom:2px solid var(--line);padding-bottom:20px;margin-bottom:28px}
  .eyebrow{font-size:12px;letter-spacing:.28em;color:var(--orange);text-transform:uppercase}
  h1{font-family:var(--display);font-weight:700;font-size:clamp(28px,4vw,44px);
     letter-spacing:-.02em;line-height:1;text-transform:uppercase;margin:4px 0 12px}
  .meta{display:flex;flex-wrap:wrap;gap:6px 24px;font-size:13px;color:var(--muted)}
  .meta b{color:var(--text);font-weight:500}

  .grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;align-items:start}
  .card{background:var(--surface);border:1px solid var(--line);border-radius:10px;
        padding:14px;position:relative;overflow:hidden}
  .card::before{content:'';position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--accent)}
  .card.wide{grid-column:1/-1}
  @media(max-width:860px){.grid{grid-template-columns:repeat(2,1fr)}}
  @media(max-width:520px){.grid{grid-template-columns:1fr}}
  .tag{display:inline-block;font-family:var(--display);font-weight:700;font-size:10px;
       letter-spacing:.16em;color:#0E0E10;background:var(--accent);
       padding:3px 9px;border-radius:4px;margin-bottom:12px;text-transform:uppercase}

  .row{display:flex;justify-content:space-between;gap:12px;padding:6px 0;
       border-bottom:1px dashed var(--line);font-size:12px}
  .row:last-child{border-bottom:0}
  .k{color:var(--muted)}
  .v{color:var(--text);text-align:right;font-weight:500;word-break:break-word}

  .big{font-family:var(--display);font-weight:700;font-size:36px;line-height:1;margin:2px 0 10px}
  .big span{font-size:15px;color:var(--muted)}

  .disk{padding:12px 0;border-bottom:1px dashed var(--line)}
  .disk:last-child{border-bottom:0}
  .disk-head{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:8px}
  .disk-id{font-family:var(--display);font-weight:700;font-size:20px}
  .disk-pct{font-size:12px;color:var(--muted)}
  .disk-foot{font-size:12px;color:var(--muted);margin-top:7px}
  .meter{display:flex;gap:3px;height:18px}
  .meter i{flex:1;background:var(--line);border-radius:1px}

  footer{margin-top:32px;padding-top:18px;border-top:1px solid var(--line);
         display:flex;justify-content:space-between;align-items:center;
         font-size:12px;color:var(--muted);flex-wrap:wrap;gap:8px}
  .sig{font-family:var(--display);font-weight:700;letter-spacing:.1em;color:var(--orange)}

  @media print{body{background:#fff;color:#000;-webkit-print-color-adjust:exact;print-color-adjust:exact}}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div class="eyebrow">Inventario de Hardware</div>
    <h1>$maquina</h1>
    <div class="meta">
      <span>Usuario: <b>$usuario</b></span>
      <span>Gerado em: <b>$gerado</b></span>
    </div>
  </header>

  <div class="grid">
    $cardSistema
    $cardCpu
    $cardRam
    $cardGpu
    $cardDisco
    $cardRede
  </div>

  <footer>
    <span>Relatorio gerado via PowerShell</span>
    <span class="sig">DAVI SENISE / TI</span>
  </footer>
</div>
</body>
</html>
"@

# ---------------------------------------------------------------
#  SALVAR E ABRIR
# ---------------------------------------------------------------
$nome = "inventario_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmm').html"
$destino = Join-Path $CaminhoSaida $nome
$html | Out-File -FilePath $destino -Encoding UTF8

# Resolve o caminho absoluto real do arquivo gerado e abre por ele.
# Garante que o navegador abre exatamente o arquivo que foi salvo.
$destinoReal = (Resolve-Path -LiteralPath $destino).Path

Write-Host ""
Write-Host "Relatorio gerado em:" -ForegroundColor Green
Write-Host "  $destinoReal" -ForegroundColor White
Start-Process -FilePath $destinoReal
