using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Desafio.Models;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace Desafio.Services
{
    public class PayrollPdfService
    {

        public PayrollData ProcessPdf(string pdfPath)
        {
            var payrollData = new PayrollData();
            
            using (PdfDocument document = PdfDocument.Open(pdfPath))
            {
                string fullText = ExtractTextFromPdf(document);
                ParsePayrollText(fullText, payrollData);
            }

            return payrollData;
        }

        private string ExtractTextFromPdf(PdfDocument document)
        {
            string text = "";
            foreach (Page page in document.GetPages())
            {
                text += page.Text + "\n";
            }
            return text;
        }


        /// <summary>
        /// Converte números brasileiros corretamente (vírgula como decimal, ponto como milhares)
        /// </summary>
        private decimal ParseBrazilianNumber(string numberStr)
        {
            if (string.IsNullOrEmpty(numberStr)) return 0;
            
            // Remover pontos de milhares e substituir vírgula por ponto
            var cleanNumber = numberStr.Replace(".", "").Replace(",", ".");
            
            if (decimal.TryParse(cleanNumber, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal result))
            {
                return result;
            }
            
            Console.WriteLine($"⚠️ Erro ao converter número: '{numberStr}' -> '{cleanNumber}'");
            return 0;
        }

        private void ParsePayrollText(string text, PayrollData payrollData) 
        {
            Console.WriteLine($"Texto completo extraído: {text.Substring(0, Math.Min(500, text.Length))}...");
            
            // Extrair período (Mês/Ano)
            var periodMatch = Regex.Match(text, @"Mês/Ano:\s*(\d{2}/\d{4})", RegexOptions.IgnoreCase);
            if (periodMatch.Success)
            {
                payrollData.Period = periodMatch.Groups[1].Value;
                Console.WriteLine($"Período extraído: {payrollData.Period}");
            }
            
            // Extrair nome do funcionário (se existir)
            var employeeMatch = Regex.Match(text, @"Funcionário:\s*([^\n\r]+)", RegexOptions.IgnoreCase);
            if (employeeMatch.Success)
            {
                payrollData.EmployeeName = employeeMatch.Groups[1].Value.Trim();
                Console.WriteLine($"Funcionário extraído: {payrollData.EmployeeName}");
            }
            
            // Padrão mais flexível para códigos (aceita /314, M200, etc.)
            // Formato: CÓDIGO + DESCRIÇÃO + QUANTIDADE + VALOR
            var itemPattern = @"([A-Z0-9/]{3,4})([A-Za-z\s\.\-ÁÉÍÓÚáéíóúÂÊÎÔÛâêîôûÀÈÌÒÙàèìòùÃÕãõÇç]+?)(\d+(?:,\d+)?)\s+(\d+(?:,\d+)?)";
            var itemMatches = Regex.Matches(text, itemPattern);
            
            Console.WriteLine($"Itens encontrados: {itemMatches.Count}");
            
            foreach (Match match in itemMatches)
            {
                if (match.Success)
                {
                    var code = match.Groups[1].Value.Trim();
                    var description = match.Groups[2].Value.Trim();
                    var quantityStr = match.Groups[3].Value;
                    var valueStr = match.Groups[4].Value;
                    
                    Console.WriteLine($"Item encontrado: Código='{code}', Descrição='{description}', Qtd='{quantityStr}', Valor='{valueStr}'");
                    
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
                        Console.WriteLine($"Provento adicionado: {description} - Qtd: {quantity}, Valor: {value}");
                    }
                    else
                    {
                        payrollData.Deductions.Add(item);
                        Console.WriteLine($"Desconto adicionado: {description} - Qtd: {quantity}, Valor: {value}");
                    }
                }
            }
            
            // Tentar extrair totais da linha TOTAL
            var totalMatch = Regex.Match(text, @"TOTAL\s*\.+\s*(\d+(?:,\d+)?)\s+(\d+(?:,\d+)?)", RegexOptions.IgnoreCase);
            if (totalMatch.Success)
            {
                payrollData.TotalEarnings = ParseBrazilianNumber(totalMatch.Groups[1].Value);
                payrollData.TotalDeductions = ParseBrazilianNumber(totalMatch.Groups[2].Value);
                Console.WriteLine($"Totais extraídos: Proventos={payrollData.TotalEarnings}, Descontos={payrollData.TotalDeductions}");
            }
            else
            {
                // Se não encontrar linha TOTAL, calcular a partir dos itens
                payrollData.TotalEarnings = payrollData.Earnings.Sum(x => x.Total);
                payrollData.TotalDeductions = payrollData.Deductions.Sum(x => x.Total);
                Console.WriteLine($"Totais calculados: Proventos={payrollData.TotalEarnings}, Descontos={payrollData.TotalDeductions}");
            }
            
            payrollData.NetSalary = payrollData.TotalEarnings - payrollData.TotalDeductions;
            
            Console.WriteLine($"Resumo final: {payrollData.Earnings.Count} proventos, {payrollData.Deductions.Count} descontos");
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

    }
}
