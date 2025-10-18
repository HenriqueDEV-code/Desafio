using System;
using System.Collections.Generic;


namespace Desafio.Models
{
    /// <summary>
    ///    Representa os dados extraidos de um cartao de ponto
    /// </summary>



    public class TimeCardData
    {
        public string MonthYear { get; set; } = string.Empty;
        public List<WorkDay> WorkDays { get; set; } = new List<WorkDay>();
        public List<MonthData> Months { get; set; } = new List<MonthData>(); // Agrupar por mês
        public decimal TotalHours { get; set; }
        public decimal TotalOvertimeDay { get; set; }
        public decimal TotalOvertimeNight {  get; set; }
    }

    public class MonthData
    {
        public string MonthYear { get; set; } = string.Empty;
        public List<WorkDay> WorkDays { get; set; } = new List<WorkDay>();
    }

    /// <summary>
    ///  Representa um dia de trabalho especifico
    /// </summary>
    public class WorkDay
    {
        public int Day { get; set; }
        public string DayOfWeek { get; set; } = string.Empty;
        public string MonthYear { get; set; } = string.Empty; // Mês/Ano específico deste dia
        public string? CheckIn { get; set; }
        public string? CheckOut { get; set; }
        public string? Interval1 { get; set; }
        public string? Interval2 { get; set; }
        public string? Interval3 { get; set; }
        public string Situation { get; set; } = string.Empty;
        public string Function {  get; set; } = string.Empty;
        public decimal ATN { get; set; }
        public decimal OvertimeDay { get; set; }
        public decimal OvertimeNight { get;set; }
        public decimal Conc {  get; set; }
        public decimal Insalub {  get; set; }
    }
}
