Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

try {
    # 1. Localizar Microsoft Edge de forma robusta
    $edge = $null
    $edgeCandidates = @(
        (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe" -ErrorAction SilentlyContinue).'(default)',
        (Get-ItemProperty "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe" -ErrorAction SilentlyContinue).'(default)',
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe"
    )
    foreach ($candidate in $edgeCandidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
            $edge = $candidate
            break
        }
    }

    if (-not $edge) {
        $wsh = New-Object -ComObject WScript.Shell
        [void]$wsh.Popup("Microsoft Edge não foi encontrado neste computador.", 10, "Erro", 16)
        exit
    }

    $inputPaths = $args
    $imageExtensions = @(".jpg", ".jpeg", ".png", ".webp", ".bmp")

    # 2. Se nenhum argumento foi passado (duplo clique), abre seletor de pasta
    if ($null -eq $inputPaths -or $inputPaths.Count -eq 0) {
        $folderDlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDlg.Description = "Selecione a pasta onde estão as imagens do resumo"
        $folderDlg.ShowNewFolderButton = $false
        
        $topForm = New-Object System.Windows.Forms.Form
        $topForm.TopMost = $true
        if ($folderDlg.ShowDialog($topForm) -eq [System.Windows.Forms.DialogResult]::OK) {
            $inputPaths = @($folderDlg.SelectedPath)
        } else {
            exit
        }
        $topForm.Dispose()
    }

    # 3. Identificar imagens e pasta de destino
    $imageFiles = @()
    $targetFolder = ""
    $pdfBaseName = ""

    foreach ($item in $inputPaths) {
        if (Test-Path -LiteralPath $item -PathType Container) {
            $files = Get-ChildItem -LiteralPath $item -File | Where-Object {
                $imageExtensions -contains $_.Extension.ToLower()
            }
            $imageFiles += @($files | Select-Object -ExpandProperty FullName)
            if ([string]::IsNullOrEmpty($targetFolder)) {
                $targetFolder = (Resolve-Path -LiteralPath $item).Path
                $pdfBaseName = (Get-Item -LiteralPath $item).Name
            }
        } elseif (Test-Path -LiteralPath $item -PathType Leaf) {
            $ext = [System.IO.Path]::GetExtension($item).ToLower()
            if ($imageExtensions -contains $ext) {
                $fullPath = (Resolve-Path -LiteralPath $item).Path
                $imageFiles += $fullPath
                if ([string]::IsNullOrEmpty($targetFolder)) {
                    $targetFolder = [System.IO.Path]::GetDirectoryName($fullPath)
                    $parentName = (Get-Item -LiteralPath $targetFolder).Name
                    if ($parentName -and $parentName -notin @("Desktop", "Área de Trabalho", "Downloads", "Imagens", "Pictures")) {
                        $pdfBaseName = $parentName
                    } else {
                        $pdfBaseName = [System.IO.Path]::GetFileNameWithoutExtension($fullPath)
                    }
                }
            }
        }
    }

    $imageFiles = @($imageFiles | Select-Object -Unique)

    if ($imageFiles.Count -eq 0) {
        $wsh = New-Object -ComObject WScript.Shell
        [void]$wsh.Popup("Nenhuma imagem válida (.jpg, .jpeg, .png, .webp, .bmp) foi encontrada.", 10, "Aviso", 48)
        exit
    }

    if ([string]::IsNullOrWhiteSpace($pdfBaseName)) {
        $pdfBaseName = "Resumo_Organizado"
    }

    # 4. Ordenação natural inicial (ex: 1, 2, 3... 10) baseada no nome do arquivo
    $sortedFiles = [System.Collections.ArrayList]@(
        $imageFiles | Sort-Object { [regex]::Replace([System.IO.Path]::GetFileName($_), '\d+', { $args[0].Value.PadLeft(12, '0') }) }
    )

    # Função para carregar imagem sem travar o arquivo no disco
    function LoadImageNoLock($path) {
        if (-not [System.IO.File]::Exists($path)) { return $null }
        try {
            $bytes = [System.IO.File]::ReadAllBytes($path)
            $ms = New-Object System.IO.MemoryStream($bytes, 0, $bytes.Length)
            $img = [System.Drawing.Image]::FromStream($ms)
            $bmp = New-Object System.Drawing.Bitmap($img)
            $img.Dispose()
            $ms.Dispose()
            return $bmp
        } catch {
            try {
                $bmp = New-Object System.Drawing.Bitmap(230, 290)
                $g = [System.Drawing.Graphics]::FromImage($bmp)
                $g.Clear([System.Drawing.Color]::FromArgb(241, 245, 249))
                $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(100, 116, 139))
                $font = New-Object System.Drawing.Font("Segoe UI", 8.5)
                $format = New-Object System.Drawing.StringFormat
                $format.Alignment = [System.Drawing.StringAlignment]::Center
                $format.LineAlignment = [System.Drawing.StringAlignment]::Center
                $rect = New-Object System.Drawing.RectangleF(10, 10, 210, 270)
                $g.DrawString("Prévia não suportada pelo Windows Forms`n(será incluída perfeitamente no PDF)", $font, $brush, $rect, $format)
                $font.Dispose()
                $brush.Dispose()
                $g.Dispose()
                return $bmp
            } catch {
                return $null
            }
        }
    }

    # 5. Interface Gráfica Interativa (Preview com Bate o Olho & Troca)
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Organizador de Resumos PDF"
    $form.Size = New-Object System.Drawing.Size(780, 840)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)

    # Identificar diretório do script para localizar recursos
    $scriptDir = $PSScriptRoot
    if (-not $scriptDir) {
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    }
    if (-not $scriptDir) {
        $scriptDir = (Get-Location).Path
    }

    # Carregar ícone da aplicação
    $iconPath = Join-Path $scriptDir "app.ico"
    if (Test-Path -LiteralPath $iconPath) {
        try {
            $form.Icon = New-Object System.Drawing.Icon($iconPath)
        } catch {}
    }

    # Garante que a janela abre em primeiro plano
    $form.TopMost = $true
    $form.Add_Shown({
        $form.Activate()
        $form.TopMost = $false
    })

    # Variáveis de estado
    $script:currentPage = 0
    $script:totalPages = [Math]::Max(1, [int][Math]::Ceiling($sortedFiles.Count / 4.0))
    $script:selectedGlobalIndex = -1

    # Título e Instrução
    $titleLbl = New-Object System.Windows.Forms.Label
    $titleLbl.Text = "Prévia do Resumo em PDF"
    $titleLbl.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $titleLbl.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
    $titleLbl.Location = New-Object System.Drawing.Point(20, 10)
    $titleLbl.Size = New-Object System.Drawing.Size(400, 26)
    $form.Controls.Add($titleLbl)

    $helpLbl = New-Object System.Windows.Forms.Label
    $helpLbl.Text = "Dica: Para inverter a ordem, clique em uma imagem e depois na outra com quem quer trocar."
    $helpLbl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $helpLbl.ForeColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
    $helpLbl.Location = New-Object System.Drawing.Point(22, 38)
    $helpLbl.Size = New-Object System.Drawing.Size(720, 22)
    $form.Controls.Add($helpLbl)

    # Controles de Navegação de Página (se tiver > 4 imagens)
    $pageNavPanel = New-Object System.Windows.Forms.Panel
    $pageNavPanel.Location = New-Object System.Drawing.Point(490, 10)
    $pageNavPanel.Size = New-Object System.Drawing.Size(250, 30)
    $pageNavPanel.Visible = ($script:totalPages -gt 1)

    $btnPrev = New-Object System.Windows.Forms.Button
    $btnPrev.Text = "< Anterior"
    $btnPrev.Size = New-Object System.Drawing.Size(75, 26)
    $btnPrev.Location = New-Object System.Drawing.Point(0, 2)
    $btnPrev.FlatStyle = "Flat"
    $pageNavPanel.Controls.Add($btnPrev)

    $pageLbl = New-Object System.Windows.Forms.Label
    $pageLbl.Text = "Página 1 de $script:totalPages"
    $pageLbl.Size = New-Object System.Drawing.Size(95, 26)
    $pageLbl.Location = New-Object System.Drawing.Point(77, 5)
    $pageLbl.TextAlign = "MiddleCenter"
    $pageLbl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $pageNavPanel.Controls.Add($pageLbl)

    $btnNext = New-Object System.Windows.Forms.Button
    $btnNext.Text = "Próxima >"
    $btnNext.Size = New-Object System.Drawing.Size(75, 26)
    $btnNext.Location = New-Object System.Drawing.Point(173, 2)
    $btnNext.FlatStyle = "Flat"
    $pageNavPanel.Controls.Add($btnNext)
    $form.Controls.Add($pageNavPanel)

    # Card A4 Central (simulando a folha)
    $a4Card = New-Object System.Windows.Forms.Panel
    $a4Card.Location = New-Object System.Drawing.Point(135, 62)
    $a4Card.Size = New-Object System.Drawing.Size(500, 670)
    $a4Card.BackColor = [System.Drawing.Color]::White
    $a4Card.BorderStyle = "FixedSingle"
    $form.Controls.Add($a4Card)

    # Configuração dos 4 slots no modelo A4 (2 colunas)
    $slotConfigs = @(
        @{ BaseTitle = "[1] Sup. Esquerdo (Início)"; X = 8;   Y = 8;   W = 238; H = 320 },
        @{ BaseTitle = "[2] Inf. Esquerdo";          X = 8;   Y = 338; W = 238; H = 320 },
        @{ BaseTitle = "[3] Sup. Direito";           X = 252; Y = 8;   W = 238; H = 320 },
        @{ BaseTitle = "[4] Inf. Direito (Fim)";     X = 252; Y = 338; W = 238; H = 320 }
    )

    $slotPanels = @()
    $slotPictures = @()
    $slotHeaders = @()

    # Handler unificado de clique para troca usando $this.Tag
    $slotClickHandler = {
        $clickedSlot = [int]$this.Tag
        $globalIdx = ($script:currentPage * 4) + $clickedSlot
        if ($globalIdx -ge $sortedFiles.Count) { return }

        if ($script:selectedGlobalIndex -eq -1) {
            # Primeiro clique: seleciona a imagem
            $script:selectedGlobalIndex = $globalIdx
            $helpLbl.Text = "👉 Imagem selecionada! Agora clique na OUTRA imagem que você quer trocar de lugar com ela."
            $helpLbl.ForeColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
        } elseif ($script:selectedGlobalIndex -eq $globalIdx) {
            # Clicou na mesma imagem: cancela seleção
            $script:selectedGlobalIndex = -1
            $helpLbl.Text = "Seleção cancelada. Clique em uma imagem e depois na outra para inverter a posição."
            $helpLbl.ForeColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
        } else {
            # Segundo clique: troca as duas imagens de lugar!
            $temp = $sortedFiles[$script:selectedGlobalIndex]
            $sortedFiles[$script:selectedGlobalIndex] = $sortedFiles[$globalIdx]
            $sortedFiles[$globalIdx] = $temp

            $script:selectedGlobalIndex = -1
            $helpLbl.Text = "✅ Imagens trocadas! Bateu o olho e ficou certo? Pressione Enter para Gerar PDF."
            $helpLbl.ForeColor = [System.Drawing.Color]::FromArgb(22, 163, 74)
        }
        UpdateSlotDisplay
    }

    for ($s = 0; $s -lt 4; $s++) {
        $cfg = $slotConfigs[$s]

        $pnl = New-Object System.Windows.Forms.Panel
        $pnl.Location = New-Object System.Drawing.Point($cfg.X, $cfg.Y)
        $pnl.Width = $cfg.W
        $pnl.Height = $cfg.H
        $pnl.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 252)
        $pnl.BorderStyle = "FixedSingle"
        $pnl.Cursor = [System.Windows.Forms.Cursors]::Hand
        $pnl.Tag = $s

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $cfg.BaseTitle
        $lbl.Location = New-Object System.Drawing.Point(4, 4)
        $lbl.Width = $cfg.W - 8
        $lbl.Height = 20
        $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
        $lbl.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
        $lbl.TextAlign = "MiddleCenter"
        $lbl.Cursor = [System.Windows.Forms.Cursors]::Hand
        $lbl.Tag = $s
        $pnl.Controls.Add($lbl)

        $pic = New-Object System.Windows.Forms.PictureBox
        $pic.Location = New-Object System.Drawing.Point(4, 26)
        $pic.Width = $cfg.W - 8
        $pic.Height = $cfg.H - 30
        $pic.SizeMode = "Zoom"
        $pic.BackColor = [System.Drawing.Color]::White
        $pic.Cursor = [System.Windows.Forms.Cursors]::Hand
        $pic.Tag = $s
        $pnl.Controls.Add($pic)

        # Eventos de clique (em qualquer parte do card troca de lugar)
        $pnl.Add_Click($slotClickHandler)
        $lbl.Add_Click($slotClickHandler)
        $pic.Add_Click($slotClickHandler)

        # Duplo clique na imagem abre o arquivo original no visualizador de fotos do Windows
        $pic.Add_DoubleClick({
            $clickedSlot = [int]$this.Tag
            $globalIdx = ($script:currentPage * 4) + $clickedSlot
            if ($globalIdx -lt $sortedFiles.Count) {
                Start-Process $sortedFiles[$globalIdx]
            }
        })

        $a4Card.Controls.Add($pnl)
        $slotPanels += $pnl
        $slotPictures += $pic
        $slotHeaders += $lbl
    }

    # Função para atualizar a exibição dos slots
    function UpdateSlotDisplay() {
        $startIndex = $script:currentPage * 4
        $pageLbl.Text = "Página $($script:currentPage + 1) de $script:totalPages"
        $btnPrev.Enabled = ($script:currentPage -gt 0)
        $btnNext.Enabled = ($script:currentPage -lt ($script:totalPages - 1))

        for ($s = 0; $s -lt 4; $s++) {
            $idx = $startIndex + $s
            $pnl = $slotPanels[$s]
            $pic = $slotPictures[$s]
            $lbl = $slotHeaders[$s]

            # Limpar imagem antiga
            if ($null -ne $pic.Image) {
                $old = $pic.Image
                $pic.Image = $null
                $old.Dispose()
            }

            if ($idx -lt $sortedFiles.Count) {
                $filePath = $sortedFiles[$idx]
                $fileName = [System.IO.Path]::GetFileName($filePath)
                $pnl.Visible = $true

                if ($script:selectedGlobalIndex -eq $idx) {
                    # Destaque de seleção (Azul)
                    $lbl.Text = "👉 SELECIONADA: $fileName"
                    $lbl.ForeColor = [System.Drawing.Color]::FromArgb(29, 78, 216)
                    $pnl.BackColor = [System.Drawing.Color]::FromArgb(219, 234, 254)
                    $pnl.BorderStyle = "Fixed3D"
                } else {
                    $lbl.Text = "$($slotConfigs[$s].BaseTitle) - $fileName"
                    $lbl.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
                    $pnl.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 252)
                    $pnl.BorderStyle = "FixedSingle"
                }

                $pic.Image = LoadImageNoLock $filePath
            } else {
                $lbl.Text = "$($slotConfigs[$s].BaseTitle) - (Vazio)"
                $lbl.ForeColor = [System.Drawing.Color]::FromArgb(156, 163, 175)
                $pnl.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)
                $pnl.BorderStyle = "FixedSingle"
            }
        }
    }

    $btnPrev.Add_Click({
        if ($script:currentPage -gt 0) {
            $script:currentPage--
            UpdateSlotDisplay
        }
    })

    $btnNext.Add_Click({
        if ($script:currentPage -lt ($script:totalPages - 1)) {
            $script:currentPage++
            UpdateSlotDisplay
        }
    })

    # Rodapé: Nome do PDF e Botões
    $botPanel = New-Object System.Windows.Forms.Panel
    $botPanel.Location = New-Object System.Drawing.Point(20, 742)
    $botPanel.Size = New-Object System.Drawing.Size(725, 55)

    $nameLbl = New-Object System.Windows.Forms.Label
    $nameLbl.Text = "Nome do PDF:"
    $nameLbl.Location = New-Object System.Drawing.Point(5, 14)
    $nameLbl.Size = New-Object System.Drawing.Size(100, 24)
    $nameLbl.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $nameLbl.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
    $botPanel.Controls.Add($nameLbl)

    $nameBox = New-Object System.Windows.Forms.TextBox
    $nameBox.Text = $pdfBaseName
    $nameBox.Location = New-Object System.Drawing.Point(110, 10)
    $nameBox.Size = New-Object System.Drawing.Size(260, 28)
    $botPanel.Controls.Add($nameBox)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Gerar PDF (Enter)"
    $btnOk.Location = New-Object System.Drawing.Point(390, 6)
    $btnOk.Size = New-Object System.Drawing.Size(190, 42)
    $btnOk.BackColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.Font = New-Object System.Drawing.Font("Segoe UI", 10.5, [System.Drawing.FontStyle]::Bold)
    $btnOk.FlatStyle = "Flat"
    $btnOk.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $botPanel.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancelar"
    $btnCancel.Location = New-Object System.Drawing.Point(595, 6)
    $btnCancel.Size = New-Object System.Drawing.Size(115, 42)
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(226, 232, 240)
    $btnCancel.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
    $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $botPanel.Controls.Add($btnCancel)

    $form.Controls.Add($botPanel)
    $form.AcceptButton = $btnOk
    $form.CancelButton = $btnCancel

    # Primeira carga das imagens
    UpdateSlotDisplay

    # Exibir janela
    $dialogResult = $form.ShowDialog()

    # Limpar imagens da memória
    for ($s = 0; $s -lt 4; $s++) {
        if ($null -ne $slotPictures[$s].Image) {
            $slotPictures[$s].Image.Dispose()
        }
    }
    $form.Dispose()

    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        exit
    }

    # 6. Atualizar nome do PDF com o que foi digitado (sanitizando caracteres inválidos)
    if (-not [string]::IsNullOrWhiteSpace($nameBox.Text)) {
        $pdfBaseName = $nameBox.Text.Trim()
    }
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($ch in $invalidChars) {
        $pdfBaseName = $pdfBaseName.Replace($ch.ToString(), "-")
    }
    if ($pdfBaseName.EndsWith(".pdf", [System.StringComparison]::OrdinalIgnoreCase)) {
        $pdfBaseName = $pdfBaseName.Substring(0, $pdfBaseName.Length - 4)
    }
    if ([string]::IsNullOrWhiteSpace($pdfBaseName)) {
        $pdfBaseName = "Resumo_Organizado"
    }

    # 7. Montar HTML com a ordem confirmada
    $htmlBuilder = New-Object System.Text.StringBuilder
    [void]$htmlBuilder.AppendLine("<!DOCTYPE html>")
    [void]$htmlBuilder.AppendLine("<html><head><meta charset='utf-8'>")
    [void]$htmlBuilder.AppendLine("<style>")
    [void]$htmlBuilder.AppendLine("  @page { size: A4 portrait; margin: 4mm; }")
    [void]$htmlBuilder.AppendLine("  * { box-sizing: border-box; margin: 0; padding: 0; }")
    [void]$htmlBuilder.AppendLine("  html, body { background: #fff; width: 100%; }")
    [void]$htmlBuilder.AppendLine("  .page {")
    [void]$htmlBuilder.AppendLine("    width: 202mm;")
    [void]$htmlBuilder.AppendLine("    height: 289mm;")
    [void]$htmlBuilder.AppendLine("    display: grid;")
    [void]$htmlBuilder.AppendLine("    grid-template-columns: 1fr 1fr;")
    [void]$htmlBuilder.AppendLine("    grid-template-rows: 1fr 1fr;")
    [void]$htmlBuilder.AppendLine("    page-break-after: always;")
    [void]$htmlBuilder.AppendLine("    break-after: page;")
    [void]$htmlBuilder.AppendLine("  }")
    [void]$htmlBuilder.AppendLine("  .page:last-child { page-break-after: avoid; break-after: avoid; }")
    [void]$htmlBuilder.AppendLine("  .page img { width: 100%; height: 100%; object-fit: fill; display: block; }")
    [void]$htmlBuilder.AppendLine("</style></head><body>")

    # Processar em blocos de 4 imagens por página
    for ($i = 0; $i -lt $sortedFiles.Count; $i += 4) {
        [void]$htmlBuilder.AppendLine("  <div class='page'>")
        
        # 1: Superior Esquerdo (col 1, row 1)
        if ($i -lt $sortedFiles.Count) {
            $uri = [System.Uri]::new($sortedFiles[$i]).AbsoluteUri.Replace("'", "%27").Replace('"', "%22")
            [void]$htmlBuilder.AppendLine("    <img src=`"$uri`" style='grid-column: 1; grid-row: 1;'>")
        }
        # 2: Inferior Esquerdo (col 1, row 2)
        if (($i + 1) -lt $sortedFiles.Count) {
            $uri = [System.Uri]::new($sortedFiles[$i + 1]).AbsoluteUri.Replace("'", "%27").Replace('"', "%22")
            [void]$htmlBuilder.AppendLine("    <img src=`"$uri`" style='grid-column: 1; grid-row: 2;'>")
        }
        # 3: Superior Direito (col 2, row 1)
        if (($i + 2) -lt $sortedFiles.Count) {
            $uri = [System.Uri]::new($sortedFiles[$i + 2]).AbsoluteUri.Replace("'", "%27").Replace('"', "%22")
            [void]$htmlBuilder.AppendLine("    <img src=`"$uri`" style='grid-column: 2; grid-row: 1;'>")
        }
        # 4: Inferior Direito (col 2, row 2)
        if (($i + 3) -lt $sortedFiles.Count) {
            $uri = [System.Uri]::new($sortedFiles[$i + 3]).AbsoluteUri.Replace("'", "%27").Replace('"', "%22")
            [void]$htmlBuilder.AppendLine("    <img src=`"$uri`" style='grid-column: 2; grid-row: 2;'>")
        }

        [void]$htmlBuilder.AppendLine("  </div>")
    }

    [void]$htmlBuilder.AppendLine("</body></html>")

    # Gravar HTML temporário
    $tempHtml = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "resumo_" + [System.Guid]::NewGuid().ToString() + ".html")
    [System.IO.File]::WriteAllText($tempHtml, $htmlBuilder.ToString(), [System.Text.Encoding]::UTF8)

    # Caminho final do PDF
    $outPdf = [System.IO.Path]::Combine($targetFolder, "$pdfBaseName.pdf")
    $counter = 1
    while (Test-Path -LiteralPath $outPdf) {
        $outPdf = [System.IO.Path]::Combine($targetFolder, "${pdfBaseName}_${counter}.pdf")
        $counter++
    }

    # Gerar PDF via Edge Headless
    $proc = Start-Process -FilePath $edge -ArgumentList "--headless", "--disable-gpu", "--allow-file-access-from-files", "--no-pdf-header-footer", "--print-to-pdf=`"$outPdf`"", "`"$tempHtml`"" -Wait -PassThru

    # Limpar temporário
    if (Test-Path -LiteralPath $tempHtml) {
        Remove-Item -LiteralPath $tempHtml -Force -ErrorAction SilentlyContinue
    }

    if (Test-Path -LiteralPath $outPdf) {
        $wsh = New-Object -ComObject WScript.Shell
        $result = $wsh.Popup("PDF gerado com sucesso!`n`nArquivo: $outPdf`n`nDeseja abrir o PDF agora?", 15, "Gerador de Resumos PDF", 4 + 64)
        if ($result -eq 6) {
            Start-Process $outPdf
        }
    } else {
        $wsh = New-Object -ComObject WScript.Shell
        [void]$wsh.Popup("Ocorreu um problema ao gerar o PDF.", 10, "Erro", 16)
    }
} catch {
    $wsh = New-Object -ComObject WScript.Shell
    [void]$wsh.Popup("Erro inesperado: $($_.Exception.Message)", 15, "Gerador de Resumos PDF", 16)
}
