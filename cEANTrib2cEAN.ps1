Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Forçar o PowerShell a usar UTF-8 para exibir textos
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Função para criar a interface gráfica
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
    $labelOutput.Text = "Diretório de saída:"
    $labelOutput.Location = New-Object System.Drawing.Point(10, 60)
    $labelOutput.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($labelOutput)

    $textBoxOutput = New-Object System.Windows.Forms.TextBox
    $textBoxOutput.Location = New-Object System.Drawing.Point(160, 60)
    $textBoxOutput.Size = New-Object System.Drawing.Size(300, 20)
    $form.Controls.Add($textBoxOutput)

    $buttonOutput = New-Object System.Windows.Forms.Button
    $buttonOutput.Text = "Selecionar Diretório"
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
            [System.Windows.Forms.MessageBox]::Show("Por favor, selecione um arquivo XML válido.", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        if (-not (Test-Path $outputDir)) {
            [System.Windows.Forms.MessageBox]::Show("Por favor, selecione um diretório de saída válido.", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        try {
            # Criar o caminho de saída
            $originalFilename = [System.IO.Path]::GetFileName($inputPath)
            $outputPath = Join-Path $outputDir $originalFilename

            # Processar o arquivo XML (lógica omitida para foco na UI)

            [System.Windows.Forms.MessageBox]::Show("Arquivo processado e salvo em:`n$outputPath", "Sucesso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Ocorreu um erro ao processar o arquivo:`n$($_.Exception.Message)", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $form.Controls.Add($buttonProcess)

    # Mostrar a janela
    $form.ShowDialog()
}

# Forçar o PowerShell a usar UTF-8 para exibir textos
[System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Executar a interface
Show-UI