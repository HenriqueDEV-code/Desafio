using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using Desafio.Models;

// Biblioteca para o Excel
using ClosedXML.Excel;

namespace Desafio.Services
{
    public class ExcelGenerator
    {

        public void GenerateTimeCardExcel(TimeCardData timeCardData, string outputPath)
        {
            using (var workbook = new XLWorkbook())
            {
                CreateTimeCardWorksheet(workbook, timeCardData);
                workbook.SaveAs(outputPath);
            }
        }


        public void GeneratePayrollExcel(PayrollData payrollData, string outputPath)
        {
            using (var workbook = new XLWorkbook())
            {
                CreatePayrollWorksheet(workbook, payrollData);
                workbook.SaveAs(outputPath);
            }
        }


        private void CreateTimeCardWorksheet(IXLWorkbook workbook, TimeCardData timeCardData)
        {
            var worksheet = workbook.Worksheets.Add("Cartão de Ponto");

            // Cabeçalho principal
            worksheet.Cell("A1").Value = "Cartão de Ponto";
            
            // Determinar quantos meses existem
            var monthsCount = timeCardData.Months.Count;
            Console.WriteLine($"Meses encontrados no PDF: {monthsCount}");
            
            if (monthsCount > 1)
            {
                var monthNames = timeCardData.Months.Select(m => m.MonthYear).OrderBy(m => m).ToList();
                worksheet.Cell("A2").Value = $"Período: {string.Join(" e ", monthNames)}";
                Console.WriteLine($"Excel gerado com múltiplos meses: {string.Join(", ", monthNames)}");
            }
            else
            {
                worksheet.Cell("A2").Value = $"Mês/Ano: {timeCardData.MonthYear}";
            }

            int currentRow = 4;
            
            // Processar cada mês separadamente
            foreach (var monthData in timeCardData.Months.OrderBy(m => m.MonthYear))
            {
                Console.WriteLine($"Processando mês: {monthData.MonthYear} com {monthData.WorkDays.Count} dias");
                
                // Cabeçalho das colunas para este mês
                worksheet.Cell($"A{currentRow}").Value = "Data";
                worksheet.Cell($"B{currentRow}").Value = "Entrada1";
                worksheet.Cell($"C{currentRow}").Value = "Saída1";
                worksheet.Cell($"D{currentRow}").Value = "Entrada2";
                worksheet.Cell($"E{currentRow}").Value = "Saída2";
                worksheet.Cell($"F{currentRow}").Value = "Entrada3";
                worksheet.Cell($"G{currentRow}").Value = "Saída3";
                worksheet.Cell($"H{currentRow}").Value = "Entrada4";
                worksheet.Cell($"I{currentRow}").Value = "Saída4";
                worksheet.Cell($"J{currentRow}").Value = "Entrada5";
                worksheet.Cell($"K{currentRow}").Value = "Saída5";
                worksheet.Cell($"L{currentRow}").Value = "Entrada6";
                worksheet.Cell($"M{currentRow}").Value = "Saída6";
                
                // Formatação do cabeçalho
                worksheet.Range($"A{currentRow}:M{currentRow}").Style.Font.Bold = true;
                currentRow++;

                // Dados deste mês
                foreach (var workDay in monthData.WorkDays)
                {
                    // Formato da data igual ao exemplar (DD/MM/YYYY)
                    var dateStr = FormatDateForTimeCard(workDay.Day, workDay.MonthYear);
                    worksheet.Cell($"A{currentRow}").Value = dateStr;
                    
                    // Entrada1 e Saída1 (manhã)
                    worksheet.Cell($"B{currentRow}").Value = workDay.CheckIn ?? "";
                    worksheet.Cell($"C{currentRow}").Value = ExtractFirstExitTime(workDay) ?? "";
                    
                    // Entrada2 e Saída2 (tarde)
                    worksheet.Cell($"D{currentRow}").Value = ExtractSecondEntryTime(workDay) ?? "";
                    worksheet.Cell($"E{currentRow}").Value = workDay.CheckOut ?? "";
                    
                    // Entrada3-6 e Saída3-6 (vazios no exemplar)
                    worksheet.Cell($"F{currentRow}").Value = "";
                    worksheet.Cell($"G{currentRow}").Value = "";
                    worksheet.Cell($"H{currentRow}").Value = "";
                    worksheet.Cell($"I{currentRow}").Value = "";
                    worksheet.Cell($"J{currentRow}").Value = "";
                    worksheet.Cell($"K{currentRow}").Value = "";
                    worksheet.Cell($"L{currentRow}").Value = "";
                    worksheet.Cell($"M{currentRow}").Value = "";
                    
                    currentRow++;
                }
                
                // Se há múltiplos meses, pular 3 linhas antes do próximo mês
                if (monthsCount > 1 && monthData != timeCardData.Months.OrderBy(m => m.MonthYear).Last())
                {
                    currentRow += 3; // Pular 3 linhas
                    Console.WriteLine($"Pulando 3 linhas antes do próximo mês (linha {currentRow})");
                }
            }
            
            // Adicionar seção de Total no final
            currentRow += 2; // Pular 2 linhas antes do total
            worksheet.Cell($"A{currentRow}").Value = "TOTAL";
            worksheet.Cell($"A{currentRow}").Style.Font.Bold = true;
            currentRow++;
            
            // Calcular totais
            var totalWorkingDays = timeCardData.WorkDays.Count(wd => !string.IsNullOrEmpty(wd.CheckIn));
            var totalHours = timeCardData.WorkDays.Sum(wd => wd.ATN);
            var totalOvertimeDay = timeCardData.WorkDays.Sum(wd => wd.OvertimeDay);
            var totalOvertimeNight = timeCardData.WorkDays.Sum(wd => wd.OvertimeNight);
            
            worksheet.Cell($"A{currentRow}").Value = $"Dias trabalhados: {totalWorkingDays}";
            worksheet.Cell($"B{currentRow}").Value = $"Total ATN: {totalHours:F2}";
            worksheet.Cell($"C{currentRow}").Value = $"HE Diurno: {totalOvertimeDay:F2}";
            worksheet.Cell($"D{currentRow}").Value = $"HE Noturno: {totalOvertimeNight:F2}";
            
            Console.WriteLine($"Total adicionado: {totalWorkingDays} dias, {totalHours:F2} ATN, {totalOvertimeDay:F2} HE Diurno, {totalOvertimeNight:F2} HE Noturno");

            // Formatação final
            worksheet.Columns().AdjustToContents();
        }
        
        /// <summary>
        /// Formata a data no formato DD/MM/YYYY igual ao exemplar
        /// </summary>
        private string FormatDateForTimeCard(int day, string monthYear)
        {
            try
            {
                // Parse do mês/ano (formato: MM/YYYY)
                var parts = monthYear.Split('/');
                if (parts.Length == 2 && int.TryParse(parts[0], out int month) && int.TryParse(parts[1], out int year))
                {
                    // Criar data completa
                    var date = new DateTime(year, month, day);
                    return date.ToString("dd/MM/yyyy");
                }
            }
            catch
            {
                // Se falhar, usar formato simples
            }
            
            return $"{day:D2}/{monthYear}";
        }
        
        /// <summary>
        /// Extrai o primeiro horário de saída (geralmente 13:45 para almoço)
        /// </summary>
        private string ExtractFirstExitTime(WorkDay workDay)
        {
            // Se há intervalo, usar o primeiro horário de saída do intervalo
            if (!string.IsNullOrEmpty(workDay.Interval1))
            {
                var intervalParts = workDay.Interval1.Split('-');
                if (intervalParts.Length >= 1)
                {
                    return intervalParts[0].Trim();
                }
            }
            
            // Se não há intervalo, usar horário padrão de almoço
            return "13:45";
        }
        
        /// <summary>
        /// Extrai o segundo horário de entrada (geralmente 14:00 após almoço)
        /// </summary>
        private string ExtractSecondEntryTime(WorkDay workDay)
        {
            // Se há intervalo, usar o segundo horário de entrada do intervalo
            if (!string.IsNullOrEmpty(workDay.Interval1))
            {
                var intervalParts = workDay.Interval1.Split('-');
                if (intervalParts.Length >= 2)
                {
                    return intervalParts[1].Trim();
                }
            }
            
            // Se não há intervalo, usar horário padrão pós-almoço
            return "14:00";
        }


        private void CreatePayrollWorksheet(IXLWorkbook workbook, PayrollData payrollData)
        {
            var worksheet = workbook.Worksheets.Add("Holerite");


            // Cabecalho
            worksheet.Cell("A1").Value = "Holerite";
            worksheet.Cell("A2").Value = $"Funcionario: {payrollData.EmployeeName}";
            worksheet.Cell("A3").Value = $"Período: {payrollData.Period}";


            // Proventos
            int row = 5;
            worksheet.Cell($"A{row}").Value = "PROVENTOS";
            worksheet.Cell($"A{row}").Style.Font.Bold = true;
            row++;

            worksheet.Cell($"A{row}").Value = "Código";
            worksheet.Cell($"B{row}").Value = "Descrição";
            worksheet.Cell($"C{row}").Value = "Quantidade";
            worksheet.Cell($"D{row}").Value = "Valor";
            worksheet.Cell($"E{row}").Value = "Total";
            worksheet.Range($"A{row}:E{row}").Style.Font.Bold = true;
            row++;


            foreach (var earning in payrollData.Earnings)
            {
                worksheet.Cell($"A{row}").Value = earning.Code;
                worksheet.Cell($"B{row}").Value = earning.Description;
                worksheet.Cell($"C{row}").Value = earning.Quantity;
                worksheet.Cell($"D{row}").Value = earning.Value;
                worksheet.Cell($"E{row}").Value = earning.Total;
                
                // Formatação de moeda brasileira (R$) para colunas de valores
                worksheet.Cell($"D{row}").Style.NumberFormat.Format = "R$ #,##0.00";
                worksheet.Cell($"E{row}").Style.NumberFormat.Format = "R$ #,##0.00";
                
                row++;
            }


            // Descontos
            row++;
            worksheet.Cell($"A{row}").Value = "DESCONTOS";
            worksheet.Cell($"A{row}").Style.Font.Bold = true;
            row++;

            worksheet.Cell($"A{row}").Value = "Código";
            worksheet.Cell($"B{row}").Value = "Descrição";
            worksheet.Cell($"C{row}").Value = "Quantidade";
            worksheet.Cell($"D{row}").Value = "Valor";
            worksheet.Cell($"E{row}").Value = "Total";
            worksheet.Range($"A{row}:E{row}").Style.Font.Bold = true;
            row++;

            foreach (var deduction in payrollData.Deductions)
            {
                worksheet.Cell($"A{row}").Value = deduction.Code;
                worksheet.Cell($"B{row}").Value = deduction.Description;
                worksheet.Cell($"C{row}").Value = deduction.Quantity;
                worksheet.Cell($"D{row}").Value = deduction.Value;
                worksheet.Cell($"E{row}").Value = deduction.Total;
                
                // Formatação de moeda brasileira (R$) para colunas de valores
                worksheet.Cell($"D{row}").Style.NumberFormat.Format = "R$ #,##0.00";
                worksheet.Cell($"E{row}").Style.NumberFormat.Format = "R$ #,##0.00";
                
                row++;
            }


            // Totais
            row++;
            worksheet.Cell($"A{row}").Value = "TOTAL PROVENTOS:";
            worksheet.Cell($"E{row}").Value = payrollData.TotalEarnings;
            worksheet.Cell($"A{row}").Style.Font.Bold = true;
            worksheet.Cell($"E{row}").Style.Font.Bold = true;
            worksheet.Cell($"E{row}").Style.NumberFormat.Format = "R$ #,##0.00";
            row++;


            worksheet.Cell($"A{row}").Value = "TOTAL DESCONTOS:";
            worksheet.Cell($"E{row}").Value = payrollData.TotalDeductions;
            worksheet.Cell($"A{row}").Style.Font.Bold = true;
            worksheet.Cell($"E{row}").Style.Font.Bold = true;
            worksheet.Cell($"E{row}").Style.NumberFormat.Format = "R$ #,##0.00";
            row ++;

            worksheet.Cell($"A{row}").Value = "SALÁRIO LÍQUIDO:";
            worksheet.Cell($"E{row}").Value = payrollData.NetSalary;
            worksheet.Cell($"A{row}").Style.Font.Bold = true;
            worksheet.Cell($"E{row}").Style.Font.Bold = true;
            worksheet.Cell($"E{row}").Style.NumberFormat.Format = "R$ #,##0.00";

            // Formatacao
            worksheet.Range("A4:E4").Style.Font.Bold = true;
            worksheet.Columns().AdjustToContents();
        }
        
        /// <summary>
        /// Extrai o mês de uma data para detectar múltiplos meses
        /// </summary>
        private string ExtractMonthFromDate(int day, string monthYear)
        {
            try
            {
                // Parse do mês/ano (formato: MM/YYYY)
                var parts = monthYear.Split('/');
                if (parts.Length == 2 && int.TryParse(parts[0], out int month) && int.TryParse(parts[1], out int year))
                {
                    // Criar data completa
                    var date = new DateTime(year, month, day);
                    return date.ToString("MM/yyyy");
                }
            }
            catch
            {
                // Se falhar, usar formato simples
            }
            
            return monthYear;
        }
    }
}
