using System;
using System.IO;
using System.Windows.Forms;
using Desafio.Controllers;

namespace Desafio
{
    public partial class MainForm : Form
    {
        private DocumentProcessor _processor;
        
        public MainForm()
        {
            InitializeComponent();
            _processor = new DocumentProcessor();
        }
        
        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Configurações do formulário
            this.Text = "Quick Filler - Processador de Documentos";
            this.Size = new System.Drawing.Size(500, 400);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
            
            // Título
            var lblTitle = new Label
            {
                Text = "Quick Filler",
                Font = new System.Drawing.Font("Segoe UI", 18, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(200, 40),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            
            var lblSubtitle = new Label
            {
                Text = "Processador de Documentos PDF",
                Font = new System.Drawing.Font("Segoe UI", 10),
                ForeColor = System.Drawing.Color.FromArgb(102, 102, 102),
                Location = new System.Drawing.Point(20, 60),
                Size = new System.Drawing.Size(300, 25)
            };
            
            // Seleção de arquivo PDF
            var lblPdf = new Label
            {
                Text = "Selecionar arquivo PDF:",
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(20, 120),
                Size = new System.Drawing.Size(200, 25)
            };
            
            var txtPdfPath = new TextBox
            {
                Location = new System.Drawing.Point(20, 150),
                Size = new System.Drawing.Size(350, 25),
                Font = new System.Drawing.Font("Segoe UI", 9),
                ReadOnly = true,
                BackColor = System.Drawing.Color.White
            };
            
            var btnSelectPdf = new Button
            {
                Text = "Procurar...",
                Location = new System.Drawing.Point(380, 148),
                Size = new System.Drawing.Size(80, 30),
                Font = new System.Drawing.Font("Segoe UI", 9),
                BackColor = System.Drawing.Color.FromArgb(0, 120, 215),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnSelectPdf.FlatAppearance.BorderSize = 0;
            btnSelectPdf.Click += (s, e) => SelectPdfFile(txtPdfPath);
            
            // Checkbox para OCR
            var chkUseOCR = new CheckBox
            {
                Text = "Usar OCR (para PDFs escaneados/imagens)",
                Location = new System.Drawing.Point(20, 200),
                Size = new System.Drawing.Size(300, 25),
                Font = new System.Drawing.Font("Segoe UI", 9),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51)
            };
            
            // Botão de processar
            var btnProcess = new Button
            {
                Text = "Processar Documento",
                Location = new System.Drawing.Point(20, 235),
                Size = new System.Drawing.Size(200, 40),
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold),
                BackColor = System.Drawing.Color.FromArgb(16, 124, 16),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Enabled = false
            };
            btnProcess.FlatAppearance.BorderSize = 0;
            btnProcess.Click += (s, e) => ProcessDocument(txtPdfPath.Text, chkUseOCR.Checked);
            
            // Barra de progresso
            var progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(20, 285),
                Size = new System.Drawing.Size(440, 20),
                Style = ProgressBarStyle.Marquee,
                Visible = false
            };
            
            // Label de status
            var lblStatus = new Label
            {
                Text = "Selecione um arquivo PDF para começar",
                Font = new System.Drawing.Font("Segoe UI", 9),
                ForeColor = System.Drawing.Color.FromArgb(102, 102, 102),
                Location = new System.Drawing.Point(20, 320),
                Size = new System.Drawing.Size(440, 25)
            };
            
            // Eventos
            txtPdfPath.TextChanged += (s, e) => btnProcess.Enabled = !string.IsNullOrEmpty(txtPdfPath.Text);
            
            // Adicionar controles ao formulário
            this.Controls.AddRange(new Control[] {
                lblTitle, lblSubtitle, lblPdf, txtPdfPath, btnSelectPdf, 
                chkUseOCR, btnProcess, progressBar, lblStatus
            });
            
            this.ResumeLayout(false);
        }
        
        private void SelectPdfFile(TextBox txtPath)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Arquivos PDF (*.pdf)|*.pdf|Todos os arquivos (*.*)|*.*";
                openFileDialog.Title = "Selecionar arquivo PDF";
                openFileDialog.InitialDirectory = Path.Combine(Application.StartupPath, "Exemplos");
                
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtPath.Text = openFileDialog.FileName;
                }
            }
        }
        
        private async void ProcessDocument(string pdfPath, bool useOCR)
        {
            try
            {
                // Mostrar progresso
                var progressBar = (ProgressBar)this.Controls[7];
                var lblStatus = (Label)this.Controls[8];
                
                progressBar.Visible = true;
                lblStatus.Text = useOCR ? "Processando documento com OCR..." : "Processando documento...";
                lblStatus.ForeColor = System.Drawing.Color.FromArgb(0, 120, 215);
                
                // Gerar nome do arquivo de saída
                var outputPath = Path.ChangeExtension(pdfPath, ".xlsx");
                var directory = Path.GetDirectoryName(pdfPath) ?? "";
                var suffix = useOCR ? "_ocr" : "_processado";
                outputPath = Path.Combine(directory, 
                    Path.GetFileNameWithoutExtension(pdfPath) + suffix + ".xlsx");
                
                // Processar em thread separada para não travar a UI
                if (useOCR)
                {
                    await Task.Run(() => _processor.ProcessDocumentWithOCR(pdfPath, outputPath));
                }
                else
                {
                    await Task.Run(() => _processor.ProcessDocument(pdfPath, outputPath));
                }
                
                // Sucesso
                progressBar.Visible = false;
                lblStatus.Text = $"Documento processado com sucesso!\nArquivo salvo em: {outputPath}";
                lblStatus.ForeColor = System.Drawing.Color.FromArgb(16, 124, 16);
                
                MessageBox.Show($"Documento processado com sucesso!\n\nArquivo Excel salvo em:\n{outputPath}", 
                    "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                var progressBar = (ProgressBar)this.Controls[7];
                var lblStatus = (Label)this.Controls[8];
                
                progressBar.Visible = false;
                lblStatus.Text = "Erro ao processar documento";
                lblStatus.ForeColor = System.Drawing.Color.FromArgb(196, 43, 28);
                
                MessageBox.Show($"Erro ao processar documento:\n\n{ex.Message}", 
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
