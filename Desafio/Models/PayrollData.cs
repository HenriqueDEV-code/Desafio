using System;
using System.Collections.Generic;

namespace Desafio.Models
{

     /// <summary>
    /// Representa os dados extraidos de um holerite
    /// </summary>
    
    public class PayrollData
    {
        public string EmployeeName { get; set; } = string.Empty;
        public string EmployeeId { get; set; } = string.Empty;
        public string Period { get; set; } = string.Empty;
        public List<PayrollItem> Earnings { get; set; } = new List<PayrollItem>();
        public List<PayrollItem> Deductions { get; set; } = new List<PayrollItem>();
        public decimal TotalEarnings { get; set; }
        public decimal TotalDeductions { get; set; }
        public decimal NetSalary { get; set; }
    }

    /// <summary>
    /// Representa um item do holerite (provento ou desconto)
    /// </summary>
    /// 

    public class PayrollItem
    {
        public string Code { get; set; } = string.Empty;
        public string Description {  get; set; } = string.Empty;
        public decimal Quantity { get; set; }
        public decimal Value    { get; set; }
        public decimal Total => Quantity * Value;
    }
}
