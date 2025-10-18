using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Desafio.Models;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using Tesseract;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using PdfPage = UglyToad.PdfPig.Content.Page;
using TesseractPage = Tesseract.Page;
using System.Globalization;
using PdfiumViewer;
using System.Drawing;
using System.Drawing.Imaging;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using PdfiumDocument = PdfiumViewer.PdfDocument;
using DrawingImageFormat = System.Drawing.Imaging.ImageFormat;

namespace Desafio.Services
{
    /// <summary>
    /// Serviço para processar PDFs com imagens usando OCR (Tesseract)
    /// </summary>
    public class ImageBasedPdfService
    {
        private readonly string _tessDataPath;
        
        public ImageBasedPdfService()
        {
            // Caminho para os dados de treinamento do Tesseract
            _tessDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
            
            // Se não existir, usar o caminho padrão
            if (!Directory.Exists(_tessDataPath))
            {
                _tessDataPath = @"C:\Program Files\Tesseract-OCR\tessdata";
            }
        }
        
        /// <summary>
        /// Processa PDF com imagens usando OCR
        /// </summary>
        public TimeCardData ProcessTimeCardWithOCR(string pdfPath)
        {
            var timeCardData = new TimeCardData();
            
            using (var document = PdfPigDocument.Open(pdfPath))
            {
                foreach (var page in document.GetPages())
                {
                    // Extrair imagens da página
                    var images = ExtractImagesFromPage(page);
                    
                    foreach (var image in images)
                    {
                        // Processar cada imagem com OCR
                        var ocrText = ProcessImageWithOCR(image);
                        
                        if (!string.IsNullOrEmpty(ocrText))
                        {
                            // Processar o texto extraído do OCR
                            ProcessOCRText(ocrText, timeCardData);
                        }
                    }
                }
            }
            
            return timeCardData;
        }
        
        /// <summary>
        /// Processa holerite com imagens usando OCR
        /// </summary>
        public PayrollData ProcessPayrollWithOCR(string pdfPath)
        {
            var payrollData = new PayrollData();
            
            Console.WriteLine($"🔍 Iniciando processamento OCR do holerite: {Path.GetFileName(pdfPath)}");
            
            try
            {
                // Tentar nova abordagem: converter PDF para imagem
                var images = ConvertPdfPagesToImages(pdfPath);
                
                if (images.Count == 0)
                {
                    Console.WriteLine("⚠️ Não foi possível converter PDF para imagens");
                    return payrollData;
                }
                
                Console.WriteLine($"📄 PDF convertido em {images.Count} imagem(ns)");
                
                foreach (var image in images)
                {
                    Console.WriteLine($"\n🖼️ Processando imagem de {image.Length} bytes...");
                    
                    // Processar cada imagem com OCR
                    var ocrText = ProcessImageWithOCR(image);
                    
                    if (!string.IsNullOrEmpty(ocrText))
                    {
                        Console.WriteLine($"📝 Texto OCR extraído: {ocrText.Length} caracteres");
                        
                        // Processar o texto extraído do OCR
                        ProcessPayrollOCRText(ocrText, payrollData);
                    }
                    else
                    {
                        Console.WriteLine("⚠️ OCR não extraiu texto desta imagem");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Erro ao processar PDF com OCR: {ex.Message}");
                
                // Fallback: tentar método antigo
                Console.WriteLine("🔄 Tentando método de extração de imagens...");
                
                using (var document = PdfPigDocument.Open(pdfPath))
                {
                    var pageCount = document.NumberOfPages;
                    Console.WriteLine($"📄 PDF tem {pageCount} página(s)");
                    
                    int pageNumber = 1;
                    foreach (var page in document.GetPages())
                    {
                        Console.WriteLine($"\n=== PROCESSANDO PÁGINA {pageNumber} ===");
                        
                        // Extrair imagens da página
                        var images = ExtractImagesFromPage(page);
                        
                        if (images.Count == 0)
                        {
                            Console.WriteLine("⚠️ Nenhuma imagem encontrada nesta página");
                            pageNumber++;
                            continue;
                        }
                        
                        foreach (var image in images)
                        {
                            Console.WriteLine($"\n🖼️ Processando imagem de {image.Length} bytes...");
                            
                            // Processar cada imagem com OCR
                            var ocrText = ProcessImageWithOCR(image);
                            
                            if (!string.IsNullOrEmpty(ocrText))
                            {
                                Console.WriteLine($"📝 Texto OCR extraído: {ocrText.Length} caracteres");
                                
                                // Processar o texto extraído do OCR
                                ProcessPayrollOCRText(ocrText, payrollData);
                            }
                            else
                            {
                                Console.WriteLine("⚠️ OCR não extraiu texto desta imagem");
                            }
                        }
                        
                        pageNumber++;
                    }
                }
            }
            
            Console.WriteLine($"\n📊 RESUMO FINAL OCR:");
            Console.WriteLine($"   Proventos encontrados: {payrollData.Earnings.Count}");
            Console.WriteLine($"   Descontos encontrados: {payrollData.Deductions.Count}");
            Console.WriteLine($"   Período: {payrollData.Period}");
            Console.WriteLine($"   Funcionário: {payrollData.EmployeeName}");
            
            return payrollData;
        }
        
        /// <summary>
        /// Converte páginas do PDF em imagens usando PdfiumViewer
        /// </summary>
        public List<byte[]> ConvertPdfPagesToImages(string pdfPath)
        {
            var images = new List<byte[]>();
            
            try
            {
                Console.WriteLine($"🔄 Convertendo PDF para imagens: {Path.GetFileName(pdfPath)}");
                
                using (var document = PdfiumDocument.Load(pdfPath))
                {
                    var pageCount = document.PageCount;
                    Console.WriteLine($"📄 PDF tem {pageCount} página(s)");
                    
                    for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
                    {
                        Console.WriteLine($"🖼️ Convertendo página {pageIndex + 1}...");
                        
                        // Renderizar página como bitmap com alta resolução
                        using (var bitmap = document.Render(pageIndex, 300, 300, true))
                        {
                            if (bitmap != null)
                            {
                                Console.WriteLine($"✅ Página {pageIndex + 1} renderizada: {bitmap.Width}x{bitmap.Height} pixels");
                                
                                // Converter bitmap para array de bytes
                                using (var memoryStream = new MemoryStream())
                                {
                                    bitmap.Save(memoryStream, DrawingImageFormat.Png);
                                    var imageBytes = memoryStream.ToArray();
                                    images.Add(imageBytes);
                                    
                                    Console.WriteLine($"💾 Imagem salva: {imageBytes.Length} bytes");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"⚠️ Falha ao renderizar página {pageIndex + 1}");
                            }
                        }
                    }
                }
                
                Console.WriteLine($"✅ Conversão concluída: {images.Count} imagem(ns) gerada(s)");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Erro ao converter PDF para imagens: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }
            
            return images;
        }
        
        /// <summary>
        /// Extrai imagens de uma página do PDF
        /// </summary>
        public List<byte[]> ExtractImagesFromPage(PdfPage page)
        {
            var images = new List<byte[]>();
            
            try
            {
                Console.WriteLine($"Extraindo imagens da página...");
                
                // Tentar extrair imagens da página
                var imageObjects = page.GetImages();
                Console.WriteLine($"Encontradas {imageObjects.Count()} imagens na página");
                
                int imageIndex = 1;
                foreach (var imageObject in imageObjects)
                {
                    try
                    {
                        Console.WriteLine($"Processando imagem {imageIndex}...");
                        
                        // Tentar diferentes métodos de extração
                        byte[]? imageBytes = null;
                        
                        // Método 1: RawBytes
                        if (imageObject.RawBytes != null && imageObject.RawBytes.Length > 0)
                        {
                            imageBytes = imageObject.RawBytes.ToArray();
                            Console.WriteLine($"Imagem extraída via RawBytes: {imageBytes.Length} bytes");
                        }
                        
                        // Método 2: TryGetPng (converter para PNG)
                        if (imageBytes == null || imageBytes.Length == 0)
                        {
                            if (imageObject.TryGetPng(out var pngBytes))
                            {
                                imageBytes = pngBytes.ToArray();
                                Console.WriteLine($"Imagem extraída via TryGetPng: {imageBytes.Length} bytes");
                            }
                        }
                        
                        if (imageBytes != null && imageBytes.Length > 0)
                        {
                            images.Add(imageBytes);
                            Console.WriteLine($"✅ Imagem adicionada com sucesso: {imageBytes.Length} bytes");
                        }
                        else
                        {
                            Console.WriteLine($"⚠️ Não foi possível extrair bytes da imagem {imageIndex}");
                        }
                        
                        imageIndex++;
                    }
                    catch (Exception imgEx)
                    {
                        Console.WriteLine($"Erro ao processar imagem individual: {imgEx.Message}");
                    }
                }
                
                Console.WriteLine($"Total de imagens extraídas: {images.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao extrair imagens: {ex.Message}");
            }
            
            return images;
        }
        
        /// <summary>
        /// Processa imagem com OCR usando Tesseract
        /// </summary>
        public string ProcessImageWithOCR(byte[] imageBytes)
        {
            try
            {
                Console.WriteLine($"Iniciando OCR em imagem de {imageBytes.Length} bytes...");
                
                // Verificar se o Tesseract está disponível
                if (!IsTesseractAvailable())
                {
                    Console.WriteLine("⚠️  Tesseract OCR não está instalado ou não está disponível.");
                    Console.WriteLine("💡 Para usar OCR, instale o Tesseract OCR em: https://github.com/UB-Mannheim/tesseract/wiki");
                    return string.Empty;
                }

                using (var engine = new TesseractEngine(_tessDataPath, "por", EngineMode.Default))
                {
                    // Configurar parâmetros do OCR para melhor qualidade
                    engine.SetVariable("tessedit_char_whitelist", "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzÁÉÍÓÚáéíóúÂÊÎÔÛâêîôûÀÈÌÒÙàèìòùÃÕãõÇç:/-.,() ");
                    engine.SetVariable("tessedit_pageseg_mode", "6"); // Assume uniform block of text
                    engine.SetVariable("preserve_interword_spaces", "1"); // Preservar espaços entre palavras
                    engine.SetVariable("tessedit_ocr_engine_mode", "2"); // LSTM OCR Engine
                    
                    Console.WriteLine("Configurações OCR aplicadas");
                    
                    using (var img = Pix.LoadFromMemory(imageBytes))
                    {
                        Console.WriteLine($"Imagem carregada: {img.Width}x{img.Height} pixels");
                        
                        using (var tesseractPage = engine.Process(img))
                        {
                            var result = tesseractPage.GetText();
                            Console.WriteLine($"OCR extraiu {result.Length} caracteres");
                            
                            // Log do texto extraído para debug
                            if (!string.IsNullOrEmpty(result))
                            {
                                Console.WriteLine($"Primeiros 300 caracteres OCR: {result.Substring(0, Math.Min(300, result.Length))}...");
                                
                                // Verificar se encontrou palavras-chave de holerite
                                var normalizedResult = result.ToUpperInvariant();
                                var keywords = new[] { "PROVENTOS", "DESCONTOS", "HORAS", "SALÁRIO", "INSS", "IRRF", "TOTAL", "LÍQUIDO" };
                                var foundKeywords = keywords.Where(k => normalizedResult.Contains(k)).ToList();
                                
                                if (foundKeywords.Any())
                                {
                                    Console.WriteLine($"✅ Palavras-chave de holerite encontradas: {string.Join(", ", foundKeywords)}");
                                }
                                else
                                {
                                    Console.WriteLine("⚠️ Nenhuma palavra-chave de holerite encontrada no OCR");
                                }
                            }
                            else
                            {
                                Console.WriteLine("⚠️ OCR não extraiu nenhum texto");
                            }
                            
                            return result;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro no OCR: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                Console.WriteLine("💡 Certifique-se de que o Tesseract OCR está instalado corretamente");
                return string.Empty;
            }
        }
        
        /// <summary>
        /// Verifica se o Tesseract está disponível
        /// </summary>
        private bool IsTesseractAvailable()
        {
            try
            {
                // Tentar criar um engine para verificar se está disponível
                using (var engine = new TesseractEngine(_tessDataPath, "por", EngineMode.Default))
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Processa texto extraído do OCR para cartão de ponto
        /// </summary>
        private void ProcessOCRText(string ocrText, TimeCardData timeCardData)
        {
            // Limpar e normalizar o texto do OCR
            var cleanText = CleanOCRText(ocrText);
            
            // Extrair mês/ano
            var monthYearMatch = Regex.Match(cleanText, @"Mês/Ano:\s*(\d+/\d+)", RegexOptions.IgnoreCase);
            if (monthYearMatch.Success)
            {
                timeCardData.MonthYear = monthYearMatch.Groups[1].Value;
            }
            
            // Extrair dados dos dias usando regex mais flexível para OCR
            var dayPattern = @"(\d{1,2})\s+(SAB|DOM|SEG|TER|QUA|QUI|SEX|SÁB|DOMINGO|SEGUNDA|TERÇA|QUARTA|QUINTA|SEXTA|SÁBADO)";
            var matches = Regex.Matches(cleanText, dayPattern, RegexOptions.IgnoreCase);
            
            foreach (Match match in matches)
            {
                if (match.Success)
                {
                    var workDay = ParseWorkDayFromOCR(match.Value, cleanText);
                    if (workDay != null)
                    {
                        timeCardData.WorkDays.Add(workDay);
                    }
                }
            }
        }
        
        /// <summary>
        /// Processa texto extraído do OCR para holerite
        /// </summary>
        private void ProcessPayrollOCRText(string ocrText, PayrollData payrollData)
        {
            Console.WriteLine($"OCR extraiu texto: {ocrText.Substring(0, Math.Min(500, ocrText.Length))}...");
            
            var cleanText = CleanOCRText(ocrText);
            
            // Extrair período (Mês/Ano) - mais flexível para OCR
            var periodMatch = Regex.Match(cleanText, @"Mês/Ano:\s*(\d{2}/\d{4})", RegexOptions.IgnoreCase);
            if (periodMatch.Success)
            {
                payrollData.Period = periodMatch.Groups[1].Value;
                Console.WriteLine($"Período extraído via OCR: {payrollData.Period}");
            }
            
            // Extrair nome do funcionário (se existir)
            var employeeMatch = Regex.Match(cleanText, @"Funcionário:\s*([^\n\r]+)", RegexOptions.IgnoreCase);
            if (employeeMatch.Success)
            {
                payrollData.EmployeeName = employeeMatch.Groups[1].Value.Trim();
                Console.WriteLine($"Funcionário extraído via OCR: {payrollData.EmployeeName}");
            }
            
            // Padrão mais flexível para códigos (aceita /314, M200, etc) - otimizado para OCR
            // Formato: CÓDIGO + DESCRIÇÃO + QUANTIDADE + VALOR
            var itemPattern = @"([A-Z]?/?\d{3,4})\s*([A-Za-z\s\.\-ÁÉÍÓÚáéíóúÂÊÎÔÛâêîôûÀÈÌÒÙàèìòùÃÕãõÇç]+?)\s*(\d+(?:,\d+)?)\s+(\d+(?:,\d+)?)";
            var itemMatches = Regex.Matches(cleanText, itemPattern, RegexOptions.IgnoreCase);
            
            Console.WriteLine($"Itens encontrados via OCR: {itemMatches.Count}");
            
            foreach (Match match in itemMatches)
            {
                if (match.Success)
                {
                    var code = match.Groups[1].Value.Trim();
                    var description = match.Groups[2].Value.Trim();
                    var quantityStr = match.Groups[3].Value;
                    var valueStr = match.Groups[4].Value;
                    
                    Console.WriteLine($"Item OCR encontrado: Código='{code}', Descrição='{description}', Qtd='{quantityStr}', Valor='{valueStr}'");
                    
                    // Usar a mesma função de parsing do PayrollPdfService
                    var quantity = ParseBrazilianNumber(quantityStr);
                    var value = ParseBrazilianNumber(valueStr);
                    
                    var item = new PayrollItem
                    {
                        Code = code,
                        Description = description,
                        Quantity = quantity,
                        Value = value
                    };
                    
                    // Determinar se é provento ou desconto baseado no código
                    if (IsEarningCode(code))
                    {
                        payrollData.Earnings.Add(item);
                        Console.WriteLine($"Provento OCR adicionado: {description} - Qtd: {quantity}, Valor: {value}");
                    }
                    else
                    {
                        payrollData.Deductions.Add(item);
                        Console.WriteLine($"Desconto OCR adicionado: {description} - Qtd: {quantity}, Valor: {value}");
                    }
                }
            }
            
            // Tentar extrair totais da linha TOTAL
            var totalMatch = Regex.Match(cleanText, @"TOTAL\s*\.+\s*(\d+(?:,\d+)?)\s+(\d+(?:,\d+)?)", RegexOptions.IgnoreCase);
            if (totalMatch.Success)
            {
                payrollData.TotalEarnings = ParseBrazilianNumber(totalMatch.Groups[1].Value);
                payrollData.TotalDeductions = ParseBrazilianNumber(totalMatch.Groups[2].Value);
                Console.WriteLine($"Totais OCR extraídos: Proventos={payrollData.TotalEarnings}, Descontos={payrollData.TotalDeductions}");
            }
            else
            {
                Console.WriteLine("⚠️ Linha TOTAL não encontrada via OCR. Calculando a partir dos itens.");
                // Se não encontrar linha TOTAL, calcular a partir dos itens
                payrollData.TotalEarnings = payrollData.Earnings.Sum(x => x.Total);
                payrollData.TotalDeductions = payrollData.Deductions.Sum(x => x.Total);
                Console.WriteLine($"Totais OCR calculados: Proventos={payrollData.TotalEarnings}, Descontos={payrollData.TotalDeductions}");
            }
            
            payrollData.NetSalary = payrollData.TotalEarnings - payrollData.TotalDeductions;
            
            Console.WriteLine($"Resumo OCR final: {payrollData.Earnings.Count} proventos, {payrollData.Deductions.Count} descontos");
        }
        
        /// <summary>
        /// Converte números brasileiros corretamente (vírgula como decimal, ponto como milhares)
        /// </summary>
        private decimal ParseBrazilianNumber(string numberStr)
        {
            if (string.IsNullOrEmpty(numberStr)) return 0;
            
            // Remover pontos de milhares e substituir vírgula por ponto
            var cleanNumber = numberStr.Replace(".", "").Replace(",", ".");
            
            if (decimal.TryParse(cleanNumber, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out decimal result))
            {
                return result;
            }
            
            Console.WriteLine($"⚠️ Erro ao converter número OCR: '{numberStr}' -> '{cleanNumber}'");
            return 0;
        }
        
        /// <summary>
        /// Determina se um código é de provento ou desconto
        /// </summary>
        private bool IsEarningCode(string code)
        {
            // Códigos de proventos geralmente começam com 0, 1 ou são específicos como M200
            var earningCodes = new[] { "0020", "0060", "1000", "1510", "1550", "1554", "1540", "M200", "0043", "0044" };
            return earningCodes.Contains(code) || code.StartsWith("0") || code.StartsWith("1");
        }
        
        /// <summary>
        /// Limpa e normaliza texto extraído do OCR
        /// </summary>
        private string CleanOCRText(string ocrText)
        {
            if (string.IsNullOrEmpty(ocrText))
                return string.Empty;
            
            Console.WriteLine($"Texto OCR original: {ocrText.Substring(0, Math.Min(100, ocrText.Length))}...");
            
            // Limpeza mais inteligente para holerites
            var cleaned = ocrText
                .Replace("\n", " ") // Quebras de linha para espaços
                .Replace("\r", " ") // Retornos de carro
                .Replace("\t", " ") // Tabs para espaços
                .Replace("  ", " ") // Espaços duplos para simples
                .Trim();
            
            // Corrigir erros comuns do OCR em holerites
            cleaned = Regex.Replace(cleaned, @"\b0(\d{3})\b", "$1"); // 0020 -> 0020 (manter códigos)
            cleaned = Regex.Replace(cleaned, @"\bO(\d{3})\b", "0$1"); // O020 -> 0020
            cleaned = Regex.Replace(cleaned, @"\bI(\d{3})\b", "1$1"); // I020 -> 1020
            
            // Corrigir vírgulas e pontos em números
            cleaned = Regex.Replace(cleaned, @"(\d+)[,\.](\d{2})\b", "$1,$2"); // Padronizar formato decimal
            
            // Remover caracteres estranhos mas manter estrutura
            cleaned = Regex.Replace(cleaned, @"[^\w\s:/\-.,()ÁÉÍÓÚáéíóúÂÊÎÔÛâêîôûÀÈÌÒÙàèìòùÃÕãõÇç]", " ");
            
            // Normalizar espaços
            cleaned = Regex.Replace(cleaned, @"\s+", " ");
            
            Console.WriteLine($"Texto OCR limpo: {cleaned.Substring(0, Math.Min(100, cleaned.Length))}...");
            
            return cleaned;
        }
        
        /// <summary>
        /// Processa linha de dia de trabalho extraída do OCR
        /// </summary>
        private WorkDay? ParseWorkDayFromOCR(string dayText, string fullText)
        {
            var dayMatch = Regex.Match(dayText, @"(\d{1,2})\s+(SAB|DOM|SEG|TER|QUA|QUI|SEX|SÁB)", RegexOptions.IgnoreCase);
            if (!dayMatch.Success) return null;
            
            if (!int.TryParse(dayMatch.Groups[1].Value, out int day)) return null;
            
            var workDay = new WorkDay
            {
                Day = day,
                DayOfWeek = dayMatch.Groups[2].Value.ToUpper()
            };
            
            // Procurar horários próximos ao dia no texto completo
            var contextPattern = $@"{day}\s+{dayMatch.Groups[2].Value}[^0-9]*?(\d{{2}}:\d{{2}})\s*-\s*(\d{{2}}:\d{{2}})";
            var timeMatch = Regex.Match(fullText, contextPattern, RegexOptions.IgnoreCase);
            
            if (timeMatch.Success)
            {
                workDay.CheckIn = timeMatch.Groups[1].Value;
                workDay.CheckOut = timeMatch.Groups[2].Value;
            }
            
            // Procurar situação (S/N)
            var situationPattern = $@"{day}\s+{dayMatch.Groups[2].Value}[^0-9]*?([SN])";
            var situationMatch = Regex.Match(fullText, situationPattern, RegexOptions.IgnoreCase);
            
            if (situationMatch.Success)
            {
                workDay.Situation = situationMatch.Groups[1].Value.ToUpper();
            }
            
            return workDay;
        }
    }
}
