using System;
using System.IO;
using System.Windows.Forms;
using Desafio.Controllers;

namespace Desafio
{
    internal static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            // Se não há argumentos, abrir interface gráfica
            if (args.Length == 0)
            {
                Application.Run(new MainForm());
            }
            else
            {
                // Modo linha de comando (para compatibilidade)
                ProcessCommandLine(args);
            }
        }
        
        static void ProcessCommandLine(string[] args)
        {
            Console.WriteLine(" === Quick Filler - Processador de Documentos === ");
            Console.WriteLine();

            // Verificar argumentos de linha de comando
            if (args.Length < 2)
            {
                ShowUsage();
                return;
            }

            try
            {
                ProcessDocuments(args);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro: {ex.Message}");
                Console.WriteLine("Pressione qualquer tecla para sair...");
                Console.ReadKey();
            }
        }

        static void ShowUsage()
        {
            Console.WriteLine("Uso:");
            Console.WriteLine("  Desafio.exe <caminho_pdf> <caminho_saida>");
            Console.WriteLine();
            Console.WriteLine("Exemplos:");
            Console.WriteLine("  Desafio.exe Exemplos\\Exemplo-Cartao-Ponto-01.pdf cartao_ponto_saida.xlsx");
            Console.WriteLine("  Desafio.exe Exemplos\\Exemplo-Holerite-01.pdf holerite_saida.xlsx");
            Console.WriteLine();
            Console.WriteLine("Tipos de documento suportados:");
            Console.WriteLine("  - Cartões de Ponto (PDFs com horários de entrada/saída)");
            Console.WriteLine("  - Holerites (PDFs com proventos e descontos)");
        }

        static void ProcessDocuments(string[] args)
        {
            string pdfPath = args[0];
            string outputPath = args[1];

            // Garantir que o caminho de saída tenha extensão .xlsx
            if (!outputPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                outputPath += ".xlsx";
            }

            Console.WriteLine($"Arquivo de entrada: {pdfPath}");
            Console.WriteLine($"Arquivo de saída: {outputPath}");
            Console.WriteLine();

            // Criar processador e executar
            var processor = new DocumentProcessor();
            processor.ProcessDocument(pdfPath, outputPath);

            Console.WriteLine();
            Console.WriteLine("Processamento concluído com sucesso!");
            Console.WriteLine("Pressione qualquer tecla para sair...");
            Console.ReadKey();
        }
    }
}