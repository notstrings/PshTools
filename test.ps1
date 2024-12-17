# ユーザーフォームを作る　- ラジオボタン編 -

# アセンブリのロード
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

 $options = @("a","b")


 function ShowDDSelect {
    param (
        [Parameter(Mandatory = $true)]  [string]   $Title,
        [Parameter(Mandatory = $true)]  [string]   $Message,
        [Parameter(Mandatory = $false)] [string]   $FileFilter  = ".*",
        [Parameter(Mandatory = $false)] [string[]] $FileList,
        [Parameter(Mandatory = $false)] [string[]] $Options
    )
    begin {}
    process {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = $Title                                     # タイトル
        $form.Size = New-Object System.Drawing.Size(480,320)    # ウィンドウサイズ
        $form.StartPosition = 'CenterScreen'                    # 表示位置
        $form.Topmost = $true                                   # TopMost
        $form.Add_Closing({
            switch ($form.Text) {
                "OK"     { $form.Text = "OK"     }
                "Cancel" { $form.Text = "Cancel" }
                Default  { $form.Text = ""       }
            }
        })
        $tableLayoutPanel1 = New-Object System.Windows.Forms.TableLayoutPanel
            $panel1  = New-Object System.Windows.Forms.Panel
                $label   = New-Object System.Windows.Forms.Label
                $listbox = New-Object System.Windows.Forms.ListBox
                $GroupBox = New-Object System.Windows.Forms.GroupBox
                    $FlowLayoutPanel = New-Object System.Windows.Forms.FlowLayoutPanel
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
            $null = $panel1.Controls.Add($GroupBox)
            $null = $panel1.Controls.Add($listbox)
            $null = $panel1.Controls.Add($label)
        
                $label.Dock = [System.Windows.Forms.DockStyle]::Top
                $label.Text = $Message
                $label.AutoSize = $true
                $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                $listbox.Dock = [System.Windows.Forms.DockStyle]::Fill
                $listbox.AllowDrop = $True
                $null = $listbox.Add_DragEnter({
                    $_.Effect = "All"
                })
                $null = $listbox.Add_DragDrop({
                    @($_.Data.GetData("FileDrop")) | ForEach-Object {
                        if( [System.IO.Path]::GetFileName($_) -match $FileFilter ){
                            [void]$Listbox.Items.Add($_)
                        }
                    }
                })
                $GroupBox.Dock = [System.Windows.Forms.DockStyle]::Bottom
                $GroupBox.Text = "Options"
                $GroupBox.Height = 50
                $null = $GroupBox.Controls.Add($FlowLayoutPanel)
                    $FlowLayoutPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
                    $Checked = $true
                    $Options | ForEach-Object {
                        $RadioButton = New-Object System.Windows.Forms.RadioButton
                        $RadioButton.Text = $_
                        $RadioButton.Checked = $Checked
                        $RadioButton.AutoSize = $true
                        $FlowLayoutPanel.Controls.Add($Radiobutton)
                        $Checked = $false
                    }
        
            $panel2.Dock = [System.Windows.Forms.DockStyle]::Fill
            $null = $panel2.Controls.Add($button2)
            $null = $panel2.Controls.Add($button1)
        
                $button1.Dock = [System.Windows.Forms.DockStyle]::Right
                $button1.Size = New-Object System.Drawing.Size(128, 36) # ボタン巾のみ指定可能
                $button1.Text = "OK"
                $button1.UseVisualStyleBackColor = $true
                $null = $button1.Add_Click({
                    $form.Text = "OK"
                    $form.Close()
                })
        
                $button2.Dock = [System.Windows.Forms.DockStyle]::Right
                $button2.Size = New-Object System.Drawing.Size(128, 36) # ボタン巾のみ指定可能
                $button2.Text = "Cancel"
                $button2.UseVisualStyleBackColor = $true
                $null = $button2.Add_Click({
                    $form.Text = "Cancel"
                    $form.Close()
                })
        
        if ($null -ne $FileList) {
            $FileList | ForEach-Object {
                if( [System.IO.Path]::GetFileName($_) -match $FileFilter ){
                    [void]$Listbox.Items.Add($_)
                }
            }
        }
        $DUMY = New-Object System.Windows.Forms.Form
        $DUMY.TopMost = $true
        $null = $form.ShowDialog($DUMY)
        return @($form.Text, $listbox.Items, ($FlowLayoutPanel.Controls | Where-Object {$_.Checked} | Select-Object -ExpandProperty Text) )
    }
    end {}
}

Write-host (ShowDDSelect "Title" "Message" ".*" @("aa") @("aa","bbb") )