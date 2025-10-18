using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using Desafio.Models;

// Instalar pelo Nuget com isso sera possivel abrir o Arquivo PDF
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;



namespace Desafio.Services
{
    public class TimeCardPdfService
    {

         // O que esse metodo faz:  Abre o PDF, extrai o texto e processa os dados
        public TimeCardData ProcessPdf(String pdfPath)
        {
            var timeCardData = new TimeCardData();

            using (PdfDocument document = PdfDocument.Open(pdfPath))
            {
                string fullText = ExtractTextFromPdf(document);
                ParseTimeCardText(fullText, timeCardData);
            }
            return timeCardData;
        }

        // Pega o texto do PDF de toda pagina e joga para a variavel *text*

        private string ExtractTextFromPdf(PdfDocument document)
        {
            string text = "";
            foreach (Page page in document.GetPages())
            {
                text += page.Text + "\n";
            }
            return text;
        }

         // Encontra o Mes/Ano e processa cada linha de dados.
        private void ParseTimeCardText(string text, TimeCardData timeCardData)
        {
            Console.WriteLine($"Texto completo extraído: {text}");

            // Encontrar TODOS os meses/anos no texto
            var monthYearPattern = @"Mês/Ano:\s*(\d+/\d+)";
            var monthYearMatches = Regex.Matches(text, monthYearPattern);
            
            Console.WriteLine($"Meses encontrados no PDF: {monthYearMatches.Count}");
            
            if (monthYearMatches.Count == 0)
            {
                Console.WriteLine("⚠️ Nenhum mês/ano encontrado no PDF!");
                return;
            }

                // Se há apenas um mês, usar a lógica original
                if (monthYearMatches.Count == 1)
                {
                    timeCardData.MonthYear = monthYearMatches[0].Groups[1].Value;
                    Console.WriteLine($"Processando único mês: {timeCardData.MonthYear}");
                    ProcessSingleMonth(text, timeCardData, timeCardData.MonthYear);
                }
            else
            {
            // Se há múltiplos meses, processar cada um
            Console.WriteLine($"Processando {monthYearMatches.Count} meses...");
            
            for (int monthIndex = 0; monthIndex < monthYearMatches.Count; monthIndex++)
            {
                var monthYear = monthYearMatches[monthIndex].Groups[1].Value;
                Console.WriteLine($"\n=== PROCESSANDO MÊS {monthIndex + 1}: {monthYear} ===");
                
                // Determinar início e fim do texto deste mês
                var startIndex = monthYearMatches[monthIndex].Index;
                var endIndex = (monthIndex + 1 < monthYearMatches.Count) 
                    ? monthYearMatches[monthIndex + 1].Index 
                    : text.Length;
                
                var monthText = text.Substring(startIndex, endIndex - startIndex);
                Console.WriteLine($"Texto do mês {monthYear}: {monthText.Substring(0, Math.Min(200, monthText.Length))}...");
                
                // Processar este mês específico - passar o monthYear correto
                ProcessSingleMonth(monthText, timeCardData, monthYear);
            }
            }
        }
        
        /// <summary>
        /// Processa um único mês do texto
        /// </summary>
        private void ProcessSingleMonth(string monthText, TimeCardData timeCardData, string monthYear)
        {
            // Usar o monthYear passado como parâmetro
            Console.WriteLine($"Processando mês: {monthYear}");
            
            // Só definir o MonthYear se ainda não foi definido (primeiro mês)
            if (string.IsNullOrEmpty(timeCardData.MonthYear))
            {
                timeCardData.MonthYear = monthYear;
                Console.WriteLine($"Mês/Ano extraído (primeiro): {timeCardData.MonthYear}");
            }
            else
            {
                Console.WriteLine($"Mês/Ano encontrado (adicional): {monthYear}");
            }

            // Padrão para extrair cada linha completa do dia
            var dayPattern = @"(\d{1,2})\s+(SAB|DOM|SEG|TER|QUA|QUI|SEX)";
            var matches = Regex.Matches(monthText, dayPattern);
            
            Console.WriteLine($"Dias encontrados neste mês: {matches.Count}");
            
            for (int i = 0; i < matches.Count; i++)
            {
                var match = matches[i];
                var day = match.Groups[1].Value;
                var dayOfWeek = match.Groups[2].Value;
                
                // Encontrar o início e fim desta linha
                var lineStartIndex = match.Index;
                var lineEndIndex = (i + 1 < matches.Count) ? matches[i + 1].Index : monthText.Length;
                
                // Extrair a linha completa
                var dayText = monthText.Substring(lineStartIndex, lineEndIndex - lineStartIndex).Trim();
                
                Console.WriteLine($"Processando linha completa: {dayText}");
                var workDay = ParseWorkDayLine(dayText, monthYear);
                if (workDay != null)
                {
                    timeCardData.WorkDays.Add(workDay);
                    
                    // Adicionar ao mês correspondente
                    var monthData = timeCardData.Months.FirstOrDefault(m => m.MonthYear == workDay.MonthYear);
                    if (monthData == null)
                    {
                        monthData = new MonthData { MonthYear = workDay.MonthYear };
                        timeCardData.Months.Add(monthData);
                        Console.WriteLine($"Novo mês criado: {workDay.MonthYear}");
                    }
                    monthData.WorkDays.Add(workDay);
                    Console.WriteLine($"Dia {workDay.Day} adicionado ao mês {workDay.MonthYear}");
                }
            }
            
            // Log final para debug
            Console.WriteLine($"=== RESUMO FINAL ===");
            Console.WriteLine($"Total de dias processados: {timeCardData.WorkDays.Count}");
            Console.WriteLine($"Total de meses encontrados: {timeCardData.Months.Count}");
            foreach (var month in timeCardData.Months)
            {
                Console.WriteLine($"Mês {month.MonthYear}: {month.WorkDays.Count} dias");
            }
        }


        private WorkDay? ParseWorkDayLine(string line, string monthYear)
        {
            // Usar regex para extrair dados mais precisamente
            // Exemplo: "03 SEG09:50 - 16:0613:45 - 14:00610100S"
            // Formato: DIA DIA_SEMANAENTRADA - SAÍDAINTERVALO_INICIO - INTERVALO_FIMVALORES_SITUAÇÃO
            var dayMatch = Regex.Match(line, @"(\d{1,2})\s+(SAB|DOM|SEG|TER|QUA|QUI|SEX)");
            if (!dayMatch.Success) return null;
            
            if (!int.TryParse(dayMatch.Groups[1].Value, out int day)) return null;
            
            var workDay = new WorkDay
            {
                Day = day,
                DayOfWeek = dayMatch.Groups[2].Value,
                MonthYear = monthYear // Usar o mês/ano passado como parâmetro
            };
            
            Console.WriteLine($"Linha processada: {line}");
            
            // Verificar se é dia de descanso ou feriado
            if (line.Contains("Descanso Semanal") || line.Contains("Feriado"))
            {
                Console.WriteLine($"Dia {day} é descanso/feriado");
                workDay.Situation = "Descanso";
                return workDay;
            }
            
            // Extrair horários - padrão mais flexível
            // Pode ser: "09:50 - 16:06" ou "09:50-16:06" ou "09:50-16:0613:45-14:00"
            var timePattern = @"(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})";
            var timeMatches = Regex.Matches(line, timePattern);
            
            Console.WriteLine($"Horários encontrados: {timeMatches.Count}");
            
            if (timeMatches.Count >= 2)
            {
                // Primeiro par: Entrada da manhã - Saída da tarde (jornada completa)
                workDay.CheckIn = timeMatches[0].Groups[1].Value;  // Entrada da manhã
                workDay.CheckOut = timeMatches[0].Groups[2].Value;  // Saída da tarde
                
                // Segundo par: Intervalo de almoço
                workDay.Interval1 = timeMatches[1].Groups[1].Value + " - " + timeMatches[1].Groups[2].Value;
                
                Console.WriteLine($"Entrada manhã: {workDay.CheckIn}");
                Console.WriteLine($"Saída tarde: {workDay.CheckOut}");
                Console.WriteLine($"Intervalo almoço: {workDay.Interval1}");
            }
            else if (timeMatches.Count == 1)
            {
                // Se só tem um par, pode ser jornada sem intervalo ou só intervalo
                var firstTime = timeMatches[0].Groups[1].Value;
                var secondTime = timeMatches[0].Groups[2].Value;
                
                // Se o primeiro horário é da manhã (antes de 12:00), é entrada
                if (IsMorningTime(firstTime))
                {
                    workDay.CheckIn = firstTime;
                    workDay.CheckOut = secondTime;
                }
                else
                {
                    // Se não é da manhã, pode ser intervalo
                    workDay.Interval1 = firstTime + " - " + secondTime;
                }
            }
            
            // Extrair valores numéricos (ATN, HE, etc.)
            var numericPattern = @"(\d+(?:,\d+)?)";
            var numericMatches = Regex.Matches(line, numericPattern);
            var numericValues = new List<decimal>();
            
            foreach (Match match in numericMatches)
            {
                if (decimal.TryParse(match.Value.Replace(",", "."), out decimal value))
                {
                    numericValues.Add(value);
                }
            }
            
            // Mapear valores baseado na posição (conforme o PDF)
            if (numericValues.Count >= 3)
            {
                workDay.ATN = numericValues[0];
                workDay.OvertimeDay = numericValues[1];
                workDay.OvertimeNight = numericValues[2];
            }
            
            // Extrair função (padrão: 610)
            var functionMatch = Regex.Match(line, @"(\d{3})");
            if (functionMatch.Success)
            {
                workDay.Function = functionMatch.Groups[1].Value;
            }
            
            // Situação (S/N no final)
            var situationMatch = Regex.Match(line, @"([SN])$");
            if (situationMatch.Success)
            {
                workDay.Situation = situationMatch.Groups[1].Value;
            }
            
            return workDay;
        }
        
        /// <summary>
        /// Verifica se um horário é da manhã (antes de 12:00)
        /// </summary>
        private bool IsMorningTime(string timeStr)
        {
            if (TimeSpan.TryParse(timeStr, out TimeSpan time))
            {
                return time.Hours < 12;
            }
            return false;
        }
    }
}
