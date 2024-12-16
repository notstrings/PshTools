# ユーザーフォームを作る　- ラジオボタン編 -

# アセンブリのロード
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$form = New-Object System.Windows.Forms.Form
$form.Text = $Title                                     # タイトル
$form.Size = New-Object System.Drawing.Size(300,200)    # ウィンドウサイズ
$form.StartPosition = 'CenterScreen'                    # 表示位置
$form.Topmost = $true                                   # TopMost
$form.Add_Closing({
    switch ($form.Text) {
        $ButtonA { $form.Text = $ButtonA }
        $ButtonB { $form.Text = $ButtonB }
        Default  { $form.Text = "" }
    }
})
$tableLayoutPanel1 = New-Object System.Windows.Forms.TableLayoutPanel
    $panel1  = New-Object System.Windows.Forms.TableLayoutPanel
        $label   = New-Object System.Windows.Forms.Label
        $listbox = New-Object System.Windows.Forms.ListBox
    $panel2  = New-Object System.Windows.Forms.Panel
        $button1 = New-Object System.Windows.Forms.Button
        $button2 = New-Object System.Windows.Forms.Button

$tableLayoutPanel1.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel1.RowCount = 2
$null = $tableLayoutPanel1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$null = $tableLayoutPanel1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50))) # ボタン高さはコレ
$null = $tableLayoutPanel1.Controls.Add($panel1, 0, 0)
$null = $tableLayoutPanel1.Controls.Add($panel2, 0, 1)
$null = $form.Controls.Add($tableLayoutPanel1)

    $panel1.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panel1.RowCount = 2
    $null = $panel1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle))
    $null = $panel1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 100)))
    $null = $panel1.Controls.Add($label, 0, 0)
    $null = $panel1.Controls.Add($listbox, 0, 1)

        $label.Dock = [System.Windows.Forms.DockStyle]::Fill
        $label.Text = $Message
        $label.AutoSize = $true
        $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $listbox.Dock = [System.Windows.Forms.DockStyle]::Fill
        $listbox.AllowDrop = $True
        $null = $listbox.Add_DragEnter({
            $_.Effect = "All"
        })
        $null = $listbox.Add_DragDrop({
            foreach($elm in @($_.Data.GetData("FileDrop"))) {
                if( [System.IO.Path]::GetFileName($elm) -match $Filter ){
                    [void]$Listbox.Items.Add($elm)
                }
            }
        })

    $null = $panel2.Controls.Add($button2)
    $null = $panel2.Controls.Add($button1)
    $panel2.Dock = [System.Windows.Forms.DockStyle]::Fill

        $button1.Dock = [System.Windows.Forms.DockStyle]::Right
        $button1.Size = New-Object System.Drawing.Size(128, 36) # ボタン巾のみ指定可能
        $button1.Text = $ButtonA
        $button1.UseVisualStyleBackColor = $true
        $null = $button1.Add_Click({
            $form.Text = $ButtonA
            $form.Close()
        })

        $button2.Dock = [System.Windows.Forms.DockStyle]::Right
        $button2.Size = New-Object System.Drawing.Size(128, 36) # ボタン巾のみ指定可能
        $button2.Text = $ButtonB
        $button2.UseVisualStyleBackColor = $true
        $null = $button2.Add_Click({
            $form.Text = $ButtonB
            $form.Close()
        })

if ($null -ne $List) {
    foreach($elm in $List) {
        if( [System.IO.Path]::GetFileName($elm) -match $Filter ){
            [void]$Listbox.Items.Add($elm)
        }
    }
}
$null = $form.ShowDialog()
