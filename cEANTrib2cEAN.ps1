Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ForÃ§ar o PowerShell a usar UTF-8 para exibir textos
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# FunÃ§Ã£o para criar a interface grÃ¡fica
function Show-UI {
    # Criar a janela principal
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Atualizar cEAN com cEANTrib"
    $form.Size = New-Object System.Drawing.Size(600, 250)
    $form.StartPosition = "CenterScreen"

    # Criar os componentes da interface
    $labelInput = New-Object System.Windows.Forms.Label
    $labelInput.Text = "Arquivo XML de entrada:"
    $labelInput.Location = New-Object System.Drawing.Point(10, 20)
    $labelInput.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($labelInput)

    $textBoxInput = New-Object System.Windows.Forms.TextBox
    $textBoxInput.Location = New-Object System.Drawing.Point(160, 20)
    $textBoxInput.Size = New-Object System.Drawing.Size(300, 20)
    $form.Controls.Add($textBoxInput)

    $buttonInput = New-Object System.Windows.Forms.Button
    $buttonInput.Text = "Selecionar Arquivo"
    $buttonInput.Location = New-Object System.Drawing.Point(470, 20)
    $buttonInput.Size = New-Object System.Drawing.Size(100, 25)
    $buttonInput.Add_Click({
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Filter = "Arquivos XML (*.xml)|*.xml"
        if ($fileDialog.ShowDialog() -eq "OK") {
            $textBoxInput.Text = $fileDialog.FileName
        }
    })
    $form.Controls.Add($buttonInput)

    $labelOutput = New-Object System.Windows.Forms.Label
    $labelOutput.Text = "DiretÃ³rio de saÃ­da:"
    $labelOutput.Location = New-Object System.Drawing.Point(10, 60)
    $labelOutput.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($labelOutput)

    $textBoxOutput = New-Object System.Windows.Forms.TextBox
    $textBoxOutput.Location = New-Object System.Drawing.Point(160, 60)
    $textBoxOutput.Size = New-Object System.Drawing.Size(300, 20)
    $form.Controls.Add($textBoxOutput)

    $buttonOutput = New-Object System.Windows.Forms.Button
    $buttonOutput.Text = "Selecionar DiretÃ³rio"
    $buttonOutput.Location = New-Object System.Drawing.Point(470, 60)
    $buttonOutput.Size = New-Object System.Drawing.Size(100, 25)
    $buttonOutput.Add_Click({
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($folderDialog.ShowDialog() -eq "OK") {
            $textBoxOutput.Text = $folderDialog.SelectedPath
        }
    })
    $form.Controls.Add($buttonOutput)

    $buttonProcess = New-Object System.Windows.Forms.Button
    $buttonProcess.Text = "Processar Arquivo"
    $buttonProcess.Location = New-Object System.Drawing.Point(240, 120)
    $buttonProcess.Size = New-Object System.Drawing.Size(120, 30)
    $buttonProcess.BackColor = "Green"
    $buttonProcess.ForeColor = "White"
    $buttonProcess.Add_Click({
        $inputPath = $textBoxInput.Text
        $outputDir = $textBoxOutput.Text

        if (-not (Test-Path $inputPath)) {
            [System.Windows.Forms.MessageBox]::Show("Por favor, selecione um arquivo XML vÃ¡lido.", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        if (-not (Test-Path $outputDir)) {
            [System.Windows.Forms.MessageBox]::Show("Por favor, selecione um diretÃ³rio de saÃ­da vÃ¡lido.", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        try {
            # Criar o caminho de saÃ­da
            $originalFilename = [System.IO.Path]::GetFileName($inputPath)
            $originalExtension = [System.IO.Path]::GetExtension($inputPath)
            $outputPath = Join-Path $outputDir $originalFilename

            # Verificar se o arquivo jÃ¡ existe
            if (Test-Path $outputPath) {
                $result = [System.Windows.Forms.MessageBox]::Show("O arquivo jÃ¡ existe no diretÃ³rio de saÃ­da. Deseja substituir?", "Arquivo Existente", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($result -eq [System.Windows.Forms.DialogResult]::No) {
                    # Solicitar um novo nome para o arquivo
                    $newNameForm = New-Object System.Windows.Forms.Form
                    $newNameForm.Text = "Novo Nome para o Arquivo"
                    $newNameForm.Size = New-Object System.Drawing.Size(400, 150)
                    $newNameForm.StartPosition = "CenterScreen"

                    $labelNewName = New-Object System.Windows.Forms.Label
                    $labelNewName.Text = "Digite o novo nome (sem extensÃ£o):"
                    $labelNewName.Location = New-Object System.Drawing.Point(10, 20)
                    $labelNewName.Size = New-Object System.Drawing.Size(300, 20)
                    $newNameForm.Controls.Add($labelNewName)

                    $textBoxNewName = New-Object System.Windows.Forms.TextBox
                    $textBoxNewName.Location = New-Object System.Drawing.Point(10, 50)
                    $textBoxNewName.Size = New-Object System.Drawing.Size(360, 20)
                    $newNameForm.Controls.Add($textBoxNewName)

                    $buttonOk = New-Object System.Windows.Forms.Button
                    $buttonOk.Text = "OK"
                    $buttonOk.Location = New-Object System.Drawing.Point(100, 80)
                    $buttonOk.Size = New-Object System.Drawing.Size(80, 30)
                    $buttonOk.Add_Click({
                        if ($textBoxNewName.Text -ne "") {
                            $newNameForm.Tag = $textBoxNewName.Text
                            $newNameForm.Close()
                        }
                    })
                    $newNameForm.Controls.Add($buttonOk)

                    $buttonCancel = New-Object System.Windows.Forms.Button
                    $buttonCancel.Text = "Cancelar"
                    $buttonCancel.Location = New-Object System.Drawing.Point(200, 80)
                    $buttonCancel.Size = New-Object System.Drawing.Size(80, 30)
                    $buttonCancel.Add_Click({
                        $newNameForm.Tag = $null
                        $newNameForm.Close()
                    })
                    $newNameForm.Controls.Add($buttonCancel)

                    $newNameForm.ShowDialog()

                    if ($newNameForm.Tag -eq $null) {
                        return
                    }

                    $newFilename = $newNameForm.Tag + $originalExtension
                    $outputPath = Join-Path $outputDir $newFilename
                } elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
                    return
                }
            }

            # Carregar o XML
            [xml]$xml = Get-Content $inputPath -Encoding UTF8

            # Gerenciar namespaces
            $namespaceManager = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $namespaceManager.AddNamespace("nfe", "http://www.portalfiscal.inf.br/nfe")

            # Selecionar todos os blocos <prod> dentro de <det>
            $prodNodes = $xml.SelectNodes("//nfe:det/nfe:prod", $namespaceManager)

            # Substituir o valor de <cEAN> pelo valor de <cEANTrib>
            foreach ($prod in $prodNodes) {
                $cEAN = $prod.SelectSingleNode("nfe:cEAN", $namespaceManager)
                $cEANTrib = $prod.SelectSingleNode("nfe:cEANTrib", $namespaceManager)
                if ($cEAN -and $cEANTrib) {
                    $cEAN.InnerText = $cEANTrib.InnerText
                }
            }

            # Salvar o arquivo atualizado com UTF-8 BOM
            $utf8Bom = New-Object System.Text.UTF8Encoding $true
            [System.IO.File]::WriteAllText($outputPath, $xml.OuterXml, $utf8Bom)

            [System.Windows.Forms.MessageBox]::Show("Arquivo processado e salvo em:`n$outputPath", "Sucesso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Ocorreu um erro ao processar o arquivo:`n$($_.Exception.Message)", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $form.Controls.Add($buttonProcess)

    # Mostrar a janela
    $form.ShowDialog()
}

# ForÃ§ar o PowerShell a usar UTF-8 para exibir textos
[System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Executar a interface
Show-UI
