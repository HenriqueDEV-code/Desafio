using System;
using System.IO;
using Desafio.Models;
using Desafio.Services;
using UglyToad.PdfPig;

namespace Desafio.Controllers
{
    public class DocumentProcessor
    {
        private readonly TimeCardPdfService _timeCardService;
        private readonly PayrollPdfService _payrollService;
        private readonly ImageBasedPdfService _imageBasedService;
        private readonly ExcelGenerator _excelGenerator;

        public DocumentProcessor()
        {
            _timeCardService = new TimeCardPdfService();
            _payrollService = new PayrollPdfService();
            _imageBasedService = new ImageBasedPdfService();
            _excelGenerator = new ExcelGenerator();
        }

        public void ProcessDocument(string pdfPath, string outputPath)
        {
            try
            {
                // Validar arquivos
                if (!File.Exists(pdfPath))
                {
                    throw new FileNotFoundException($"Arquivo PDF não encontrado: {pdfPath}");
                }

                // Detectar tipo de documento
                var documentType = DetectDocumentType(pdfPath);

                Console.WriteLine($"Processando documento: {Path.GetFileName(pdfPath)}");
                Console.WriteLine($"Tipo detectado: {documentType}");

                // Processar baseado no tipo
                switch (documentType)
                {
                    case DocumentType.TimeCard:
                        ProcessTimeCard(pdfPath, outputPath);
                        break;
                    case DocumentType.Payroll:
                        ProcessPayroll(pdfPath, outputPath);
                        break;
                    case DocumentType.Unknown:
                        Console.WriteLine("⚠️  Tipo de documento não identificado automaticamente.");
                        Console.WriteLine("🔄 Tentando processar com OCR como fallback...");
                        
                        // Tentar detectar tipo usando OCR
                        var ocrDocumentType = DetectDocumentTypeWithOCR(pdfPath);
                        Console.WriteLine($"OCR detectou: {ocrDocumentType}");
                        
                        switch (ocrDocumentType)
                        {
                            case DocumentType.TimeCard:
                                ProcessTimeCardWithOCR(pdfPath, outputPath);
                                break;
                            case DocumentType.Payroll:
                                ProcessPayrollWithOCR(pdfPath, outputPath);
                                break;
                            default:
                                throw new NotSupportedException($"Não foi possível identificar o tipo de documento. " +
                                    $"Verifique se o PDF contém um cartão de ponto ou holerite válido. " +
                                    $"Arquivo: {Path.GetFileName(pdfPath)}");
                        }
                        break;
                    default:
                        throw new NotSupportedException($"Tipo de documento não suportado: {documentType}");
                }
                Console.WriteLine($"Arquivo gerado com sucesso: {outputPath}");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao processar documento: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// Processa documento usando OCR (para PDFs escaneados/imagens)
        /// </summary>
        public void ProcessDocumentWithOCR(string pdfPath, string outputPath)
        {
            try
            {
                // Validar arquivos
                if (!File.Exists(pdfPath))
                {
                    throw new FileNotFoundException($"Arquivo PDF não encontrado: {pdfPath}");
                }

                // Detectar tipo de documento
                var documentType = DetectDocumentType(pdfPath);

                Console.WriteLine($"Processando documento com OCR: {Path.GetFileName(pdfPath)}");
                Console.WriteLine($"Tipo detectado: {documentType}");

                // Processar baseado no tipo usando OCR
                switch (documentType)
                {
                    case DocumentType.TimeCard:
                        ProcessTimeCardWithOCR(pdfPath, outputPath);
                        break;
                    case DocumentType.Payroll:
                        ProcessPayrollWithOCR(pdfPath, outputPath);
                        break;
                    case DocumentType.Unknown:
                        Console.WriteLine("⚠️  Tipo de documento não identificado automaticamente.");
                        Console.WriteLine("🔄 Tentando processar com OCR como fallback...");
                        
                        // Tentar detectar tipo usando OCR
                        var ocrDocumentType = DetectDocumentTypeWithOCR(pdfPath);
                        Console.WriteLine($"OCR detectou: {ocrDocumentType}");
                        
                        switch (ocrDocumentType)
                        {
                            case DocumentType.TimeCard:
                                ProcessTimeCardWithOCR(pdfPath, outputPath);
                                break;
                            case DocumentType.Payroll:
                                ProcessPayrollWithOCR(pdfPath, outputPath);
                                break;
                            default:
                                // Se OCR falhou, tentar processar como holerite por padrão (mais comum)
                                Console.WriteLine("🤔 OCR não conseguiu identificar o tipo. Tentando processar como HOLERITE...");
                                try
                                {
                                    ProcessPayroll(pdfPath, outputPath);
                                    Console.WriteLine("✅ Processado como holerite com sucesso!");
                                }
                                catch
                                {
                                    // Se falhar como holerite, tentar como cartão de ponto
                                    Console.WriteLine("🤔 Falhou como holerite. Tentando processar como CARTÃO DE PONTO...");
                                    try
                                    {
                                        ProcessTimeCard(pdfPath, outputPath);
                                        Console.WriteLine("✅ Processado como cartão de ponto com sucesso!");
                                    }
                                    catch
                                    {
                                        throw new NotSupportedException($"Não foi possível identificar o tipo de documento. " +
                                            $"Verifique se o PDF contém um cartão de ponto ou holerite válido. " +
                                            $"Arquivo: {Path.GetFileName(pdfPath)}");
                                    }
                                }
                                break;
                        }
                        break;
                    default:
                        throw new NotSupportedException($"Tipo de documento não suportado: {documentType}");
                }
                Console.WriteLine($"Arquivo gerado com sucesso usando OCR: {outputPath}");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao processar documento com OCR: {ex.Message}");
                throw;
            }
        }

        private DocumentType DetectDocumentType(string pdfPath)
        {
            try
            {
                using (var document = UglyToad.PdfPig.PdfDocument.Open(pdfPath))
                {
                    string text = "";
                    foreach (var page in document.GetPages())
                    {
                        text += page.Text + "\n";
                    }

                    // Normalizar texto para busca case-insensitive
                    var normalizedText = text.ToUpperInvariant();

                    Console.WriteLine($"Texto extraído do PDF (primeiros 200 caracteres): {text.Substring(0, Math.Min(200, text.Length))}...");

                    // Detectar holerite primeiro (mais específico)
                    var payrollKeywords = new[] { 
                        "PROVENTOS", "DESCONTOS", "HORAS NORMAIS", "SALÁRIO", "SALARIO",
                        "0020HORAS NORMAIS", "0060DESC", "FOLHA", "HOLERITE", "VENCIMENTOS",
                        "0020", "0060", "FOLHA DE PAGAMENTO", "CONTRA CHEQUE", "CONTRA-CHEQUE",
                        "VALOR", "QUANTIDADE", "TOTAL", "LIQUIDO", "LÍQUIDO", "BRUTO"
                    };
                    var payrollMatches = 0;
                    foreach (var keyword in payrollKeywords)
                    {
                        if (normalizedText.Contains(keyword))
                        {
                            payrollMatches++;
                            Console.WriteLine($"Palavra-chave de holerite encontrada: {keyword}");
                        }
                    }

                    // Se encontrar pelo menos 3 palavras-chave de holerite
                    if (payrollMatches >= 3)
                    {
                        Console.WriteLine($"Documento detectado como HOLERITE ({payrollMatches} palavras-chave encontradas)");
                        return DocumentType.Payroll;
                    }

                    // Detectar Cartão de Ponto
                    var timeCardKeywords = new[] { 
                        "MÊS/AN", "MES/AN", "ENTRADA", "SAÍDA", "SAIDA", 
                        "CARTÃO DE PONTO", "CARTAO DE PONTO", "PONTO", "HORÁRIO", "HORARIO",
                        "SEG", "TER", "QUA", "QUI", "SEX", "SAB", "DOM",
                        "SEGUNDA", "TERÇA", "TERCA", "QUARTA", "QUINTA", "SEXTA", "SÁBADO", "SABADO", "DOMINGO",
                        "CHECK-IN", "CHECK-OUT", "INTERVALO", "ATN", "HE", "HORAS EXTRAS"
                    };
                    var timeCardMatches = 0;
                    foreach (var keyword in timeCardKeywords)
                    {
                        if (normalizedText.Contains(keyword))
                        {
                            timeCardMatches++;
                            Console.WriteLine($"Palavra-chave de cartão de ponto encontrada: {keyword}");
                        }
                    }

                    // Se encontrar pelo menos 2 palavras-chave de cartão de ponto
                    if (timeCardMatches >= 2)
                    {
                        Console.WriteLine($"Documento detectado como CARTÃO DE PONTO ({timeCardMatches} palavras-chave encontradas)");
                        return DocumentType.TimeCard;
                    }

                    // Se não encontrou texto suficiente, pode ser uma imagem/PDF escaneado
                    if (string.IsNullOrWhiteSpace(text) || text.Length < 50)
                    {
                        Console.WriteLine("PDF parece ser uma imagem escaneada (pouco texto extraído)");
                        return DocumentType.Unknown;
                    }

                    Console.WriteLine($"Tipo de documento não identificado. Texto extraído: {text.Length} caracteres");
                    return DocumentType.Unknown;
                }
            }
            catch (Exception ex) 
            { 
                Console.WriteLine($"Erro ao detectar tipo de documento: {ex.Message}");
                return DocumentType.Unknown;
            }
        }
        
        /// <summary>
        /// Detecta tipo de documento usando OCR como fallback
        /// </summary>
        private DocumentType DetectDocumentTypeWithOCR(string pdfPath)
        {
            try
            {
                Console.WriteLine("🔍 Tentando detectar tipo usando OCR...");
                
                using (var document = UglyToad.PdfPig.PdfDocument.Open(pdfPath))
                {
                    foreach (var page in document.GetPages())
                    {
                        // Tentar extrair imagens da página
                        var images = _imageBasedService.ExtractImagesFromPage(page);
                        
                        foreach (var image in images)
                        {
                            // Processar imagem com OCR
                            var ocrText = _imageBasedService.ProcessImageWithOCR(image);
                            
                            if (!string.IsNullOrEmpty(ocrText))
                            {
                                var normalizedText = ocrText.ToUpperInvariant();
                                Console.WriteLine($"OCR extraiu: {ocrText.Substring(0, Math.Min(100, ocrText.Length))}...");
                                
                                // Detectar holerite
                                var payrollKeywords = new[] { "PROVENTOS", "DESCONTOS", "HORAS NORMAIS", "SALÁRIO", "FOLHA", "HOLERITE" };
                                foreach (var keyword in payrollKeywords)
                                {
                                    if (normalizedText.Contains(keyword))
                                    {
                                        Console.WriteLine($"OCR detectou HOLERITE (palavra-chave: {keyword})");
                                        return DocumentType.Payroll;
                                    }
                                }
                                
                                // Detectar cartão de ponto
                                var timeCardKeywords = new[] { "MÊS/ANO", "ENTRADA", "SAÍDA", "PONTO", "SEG", "TER", "QUA", "QUI", "SEX" };
                                var matches = 0;
                                foreach (var keyword in timeCardKeywords)
                                {
                                    if (normalizedText.Contains(keyword))
                                    {
                                        matches++;
                                    }
                                }
                                
                                if (matches >= 2)
                                {
                                    Console.WriteLine($"OCR detectou CARTÃO DE PONTO ({matches} palavras-chave)");
                                    return DocumentType.TimeCard;
                                }
                            }
                        }
                    }
                }
                
                Console.WriteLine("OCR não conseguiu identificar o tipo de documento");
                return DocumentType.Unknown;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao detectar tipo com OCR: {ex.Message}");
                return DocumentType.Unknown;
            }
        }

        private void ProcessTimeCard(string pdfPath, string outputPath)
        {
            Console.WriteLine("Processando cartão de ponto...");

            // Extrair dados do PDF
            var timeCardData = _timeCardService.ProcessPdf(pdfPath);

            Console.WriteLine($"Dados extraídos: {timeCardData.WorkDays.Count} dias de trabalho");
            Console.WriteLine($"Período: {timeCardData.MonthYear}");


            // Gerar planilha Excel
            _excelGenerator.GenerateTimeCardExcel(timeCardData, outputPath);

            Console.WriteLine("Cartão de ponto processado com sucesso!");
        }

        private void ProcessPayroll(string pdfPath, string outputPath)
        {
            Console.WriteLine("Processando holerite...");

            // Extrair dados do PDF
            var payrollData = _payrollService.ProcessPdf(pdfPath);

            Console.WriteLine($"Dados extraídos: {payrollData.Earnings.Count} proventos, {payrollData.Deductions.Count} descontos");
            Console.WriteLine($"Funcionário: {payrollData.EmployeeName}");
            Console.WriteLine($"Período: {payrollData.Period}");

            // Gerar planilha Excel
            _excelGenerator.GeneratePayrollExcel(payrollData, outputPath);

            Console.WriteLine("Holerite processado com sucesso!");
        }
        
        private void ProcessTimeCardWithOCR(string pdfPath, string outputPath)
        {
            Console.WriteLine("Processando cartão de ponto com OCR...");

            // Extrair dados do PDF usando OCR
            var timeCardData = _imageBasedService.ProcessTimeCardWithOCR(pdfPath);

            Console.WriteLine($"Dados extraídos com OCR: {timeCardData.WorkDays.Count} dias de trabalho");
            Console.WriteLine($"Período: {timeCardData.MonthYear}");

            // Gerar planilha Excel
            _excelGenerator.GenerateTimeCardExcel(timeCardData, outputPath);

            Console.WriteLine("Cartão de ponto processado com sucesso usando OCR!");
        }
        
        private void ProcessPayrollWithOCR(string pdfPath, string outputPath)
        {
            Console.WriteLine("Processando holerite com OCR...");

            // Extrair dados do PDF usando OCR
            var payrollData = _imageBasedService.ProcessPayrollWithOCR(pdfPath);

            Console.WriteLine($"Dados extraídos com OCR: {payrollData.Earnings.Count} proventos, {payrollData.Deductions.Count} descontos");
            Console.WriteLine($"Funcionário: {payrollData.EmployeeName}");
            Console.WriteLine($"Período: {payrollData.Period}");

            // Gerar planilha Excel
            _excelGenerator.GeneratePayrollExcel(payrollData, outputPath);

            Console.WriteLine("Holerite processado com sucesso usando OCR!");
        }
    }


    // Emun
    public enum DocumentType
    {
      TimeCard,
      Payroll,
      Unknown
    }

}